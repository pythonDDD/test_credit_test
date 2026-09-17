/* ============================================================================
 * /lease/app.js — リース料の計算と逆算（リース見積診断） 画面の動き
 * 計算は engine.js。ここでは入力を受け取り、結果を描くことだけを担当する。
 *
 * 2つの使い方
 *   calc    … 物件価額・期間・年利から、毎月のリース料を出す
 *   reverse … 見積の月額（またはリース料率）から、実質年率（金利）を出す
 * 物件価額とリース期間は、2つの使い方で共通の入力欄を使う。
 * 残価は入力欄に置かず、結果の中のバーを押して選ぶ（入力を最小限にするため）。
 * ========================================================================== */
import { forward, reverse, residualScenarios, residualPayments } from "./engine.js?v=2";

const $ = (id) => document.getElementById(id);
const yen = (n) => Math.round(n).toLocaleString("ja-JP");
const pct = (x, d = 2) => (x * 100).toFixed(d);
const STEPS = [0, 0.05, 0.1, 0.15, 0.2, 0.25, 0.3];

/* 「例を入れる」の数字。分かりやすさを優先して、キリのよい値にしている */
const EXAMPLE = {
  calc: { price: 5000000, months: 60, rate: 3.0, misc: 1.2 },
  reverse: { price: 5000000, months: 60, pay: 95000 },
};

const state = { mode: "calc", residual: 0, unit: "yen" };

/* 有料版（グラフ＋Excel）は後から読み込む。読み込みに失敗しても、無料の計算は止めない */
let paid = null, lastQuote = null;
function notifyPaid(q) {
  lastQuote = q;
  if (paid) { try { paid.updatePaid(q); } catch (e) { console.warn("[リース見積診断] 有料版の表示を更新できませんでした", e); } }
}
function loadPaid() {
  import("./paid.js?v=1").then((m) => {
    m.initPaid();
    paid = m;
    m.updatePaid(lastQuote);
  }).catch((e) => console.warn("[リース見積診断] 有料版を読み込めませんでした（無料の計算はそのまま使えます）", e));
}

/* ------------------------------------------------------------ 小さな道具 */
function num(id) {
  const el = $(id);
  if (!el || el.value.trim() === "") return NaN;
  return parseFloat(el.value);
}

/** 5000000 → 「500万円」。読み間違い（桁の数え違い）を防ぐための添え書き */
function manYen(n) {
  if (!(n > 0)) return "";
  if (n >= 1e8) {
    const oku = Math.floor(n / 1e8), rest = Math.round((n % 1e8) / 1e4);
    return `${oku}億${rest ? rest.toLocaleString("ja-JP") + "万" : ""}円`;
  }
  if (n >= 1e4) {
    const m = n / 1e4;
    return `${Number.isInteger(m) ? m.toLocaleString("ja-JP") : m.toFixed(1)}万円`;
  }
  return `${yen(n)}円`;
}

function track(name, params) {
  try { if (typeof gtag === "function") gtag("event", name, params); } catch (_) { /* 計測できなくても動作は続ける */ }
}

/* ------------------------------------------------------------ 使い方の切り替え */
function setMode(mode, { scroll = false } = {}) {
  state.mode = mode;
  document.querySelectorAll("[data-mode-tab]").forEach((b) => {
    const on = b.dataset.modeTab === mode;
    b.setAttribute("aria-selected", String(on));
    b.tabIndex = on ? 0 : -1;
  });
  document.querySelectorAll("[data-only]").forEach((el) => { el.hidden = el.dataset.only !== mode; });
  $("formTitle").textContent = mode === "calc" ? "条件を入力する" : "見積の数字を入力する";
  $("formHint").textContent = mode === "calc"
    ? "単位は円。はじめての方は「例を入れる」を押すと、使い方がすぐ分かります。"
    : "見積書に書かれた物件価額・リース期間・月額（またはリース料率）を入れてください。";
  history.replaceState(null, "", location.pathname + location.search + (mode === "reverse" ? "#reverse" : ""));
  render({ showErr: false });
  if (scroll) $("tool").scrollIntoView({ behavior: "smooth", block: "start" });
}

/* ------------------------------------------------------------ 入力の読み取り */
function readShared(errors) {
  const price = num("price"), months = num("months");
  if (!(price > 0)) errors.push("物件価額を入れてください。");
  if (!(months >= 1 && months <= 240 && Number.isInteger(months))) errors.push("リース期間は1〜240か月の整数で入れてください。");
  return { price, months };
}

function readCalc() {
  const errors = [];
  const { price, months } = readShared(errors);
  const rate = num("rate"), miscIn = num("misc");
  const misc = isNaN(miscIn) ? 0 : miscIn;
  if (!(rate >= 0 && rate < 50)) errors.push("年利は0〜50%の範囲で入れてください。");
  if (!(misc >= 0 && misc < 30)) errors.push("その他費用率は0〜30%の範囲で入れてください。");
  return { errors, args: { price, months, annualRate: rate / 100, miscRate: misc / 100 } };
}

function readReverse() {
  const errors = [];
  const { price, months } = readShared(errors);
  const raw = num("pay");
  let monthly = NaN;
  if (state.unit === "yen") {
    if (!(raw > 0)) errors.push("見積の月額リース料を入れてください。");
    monthly = raw;
  } else {
    if (!(raw > 0 && raw < 100)) errors.push("リース料率は0〜100%の範囲で入れてください。");
    monthly = price > 0 ? price * raw / 100 : NaN;
  }
  return { errors, args: { price, months, monthly } };
}

/* ------------------------------------------------------------ 描画 */
function render({ scroll = false, showErr = true } = {}) {
  paintHelpers();
  const read = state.mode === "calc" ? readCalc() : readReverse();
  const err = $("err");
  const box = state.mode === "calc" ? $("resCalc") : $("resRev");
  const other = state.mode === "calc" ? $("resRev") : $("resCalc");
  other.hidden = true;

  if (read.errors.length) {
    box.hidden = true;
    notifyPaid(null);
    err.textContent = read.errors[0];
    err.hidden = !showErr;
    return false;
  }
  err.hidden = true;
  const args = { ...read.args, residual: read.args.price * state.residual };
  const monthly = state.mode === "calc" ? paintCalc(args) : (paintReverse(args), args.monthly);
  box.hidden = false;
  notifyPaid({ price: args.price, months: args.months, monthly, residual: args.residual });
  if (scroll) box.scrollIntoView({ behavior: "smooth", block: "start" });
  return true;
}

/** 入力欄の下の添え書き（500万円、リース料率への換算など） */
function paintHelpers() {
  const price = num("price");
  $("priceHelp").textContent = price > 0 ? `＝ ${manYen(price)}` : "";
  const raw = num("pay");
  let help = "";
  if (raw > 0 && price > 0) {
    help = state.unit === "yen"
      ? `＝ リース料率 ${(raw / price * 100).toFixed(3)}%`
      : `＝ 月額 ${yen(price * raw / 100)}円`;
  }
  $("payHelp").textContent = help;
  document.querySelectorAll("[data-months]").forEach((c) =>
    c.classList.toggle("is-on", +c.dataset.months === num("months")));
}

function residualTag(args) {
  return state.residual > 0
    ? `残価 ${Math.round(state.residual * 100)}%（${yen(args.residual)}円）を設定した場合`
    : "残価なし（フルペイアウト）";
}

function paintCalc(args) {
  const r = forward(args);
  $("cMonthly").innerHTML = `${yen(r.monthly)}<span class="u">円 / 月</span>`;
  $("cTag").textContent = residualTag(args);
  $("cSub").textContent = `リース料率 ${pct(r.ratio, 3)}%（物件価額に対する月額の割合）／ 支払総額 ${yen(r.total)}円`;

  const extra = r.total - args.price;
  $("cKpi").innerHTML =
    kpi("実質年率（金利）", isFinite(r.effective) ? `${pct(r.effective)}%` : "—", "税や保険の費用も含めて、借入の金利に直した値") +
    kpi(extra >= 0 ? "物件価額より多く払う額" : "物件価額より少ない支払", `${yen(Math.abs(extra))}円`, extra >= 0 ? "支払総額 − 物件価額" : "満了時に物件を返すぶん、支払が少ない") +
    kpi("支払総額は物件価額の", `${r.multiple.toFixed(3)}倍`, "1.2倍超は見直しの目安");

  // 支払総額の内訳（横に積み上げたバー）
  const parts = [
    { k: "principal", label: "物件価額の回収（元本）", v: r.principal },
    { k: "interest", label: "金利相当額", v: Math.max(r.interest, 0) },
    { k: "misc", label: "その他費用（税・保険・管理費）", v: r.miscTotal },
  ];
  const sum = parts.reduce((s, p) => s + p.v, 0) || 1;
  $("cStack").innerHTML = parts.map((p) =>
    `<i class="stack__seg stack__seg--${p.k}" style="width:${(p.v / sum * 100).toFixed(2)}%" title="${p.label}"></i>`).join("");

  $("cFigs").innerHTML =
    row(`<i class="dot dot--principal"></i>物件価額の回収（元本）`, `${yen(r.principal)}円`) +
    row(`<i class="dot dot--interest"></i>金利相当額`, `${yen(r.interest)}円`) +
    row(`<i class="dot dot--misc"></i>その他費用（税・保険・管理費）`, `${yen(r.miscTotal)}円`) +
    row("支払総額", `${yen(r.total)}円`, "tot") +
    (args.residual > 0 ? row("満了時の残価", `${yen(args.residual)}円`) : "");

  // 残価を置いたときの月額
  const list = residualPayments({ price: args.price, months: args.months, annualRate: args.annualRate, miscRate: args.miscRate }, STEPS);
  paintBars($("cBars"), list.map((x) => ({ ratio: x.ratio, value: x.monthly, text: yen(x.monthly) })));

  const rem = [];
  if (r.total > 0) rem.push(`支払総額のうち ${(Math.max(r.interest, 0) / r.total * 100).toFixed(1)}% が金利相当額です。`);
  if (isFinite(r.effective)) rem.push(`税や保険などの費用まで含めて金利に直すと、実質年率（金利）は ${pct(r.effective)}% です。入力した年利（${(args.annualRate * 100).toFixed(2)}%）との差が、費用の重さです。`);
  if (args.residual > 0) rem.push("残価を設定した分だけ月額が下がっています。満了時に物件を返すのか、残価で買い取るのかは、契約書で必ず確認してください。");
  if (r.multiple >= 1.2) rem.push("支払総額が物件価額の1.2倍を超えています。期間か金利のどちらかを見直せないか、確かめる余地があります。");
  $("cRemarks").innerHTML = rem.map((t) => `<li>${t}</li>`).join("");
  return r.monthly;
}

function paintReverse(args) {
  const r = reverse(args);
  const ok = isFinite(r.effective);
  $("rRate").innerHTML = ok ? `${pct(r.effective)}<span class="u">% / 年</span>` : `—<span class="u">計算できません</span>`;
  $("rTag").textContent = residualTag(args);
  $("rSub").textContent = `リース料率 ${pct(r.ratio, 3)}% ／ 支払総額 ${yen(r.total)}円`;

  $("rKpi").innerHTML =
    kpi("月額リース料", `${yen(args.monthly)}円`, "見積の月額（税抜）") +
    kpi(r.extra >= 0 ? "物件価額より多く払う額" : "物件価額より少ない支払", `${yen(Math.abs(r.extra))}円`, "支払総額 − 物件価額") +
    kpi("支払総額は物件価額の", `${r.multiple.toFixed(3)}倍`, "期間が長いほど大きくなる");

  const list = residualScenarios({ price: args.price, months: args.months, monthly: args.monthly }, STEPS);
  paintBars($("rBars"), list.map((x) => ({ ratio: x.ratio, value: isFinite(x.effective) ? Math.max(x.effective, 0) : 0,
    text: isFinite(x.effective) ? `${pct(x.effective)}%` : "—" })));

  const rem = [];
  if (!ok) rem.push("この月額と期間では、金利を計算できません。物件価額・月額・期間の桁が合っているか確かめてください。");
  else {
    if (r.effective < 0) rem.push("支払総額が物件価額を下回っているため、金利がマイナスになっています。月額や物件価額の入力を確かめてください。");
    else rem.push(`このリースは、年 ${pct(r.effective)}% の金利でお金を借りて物件を買うのと同じ負担です。`);
    rem.push("自社が銀行から設備資金を借りる金利と比べてみてください。差には、税金・保険・事務の手間と、リース会社の利益が入っています。");
    if (args.residual > 0) rem.push(`残価 ${Math.round(state.residual * 100)}% が設定されていた場合の金利です。同じ月額でも、残価がある分だけ金利に直した値は高くなります。`);
  }
  $("rRemarks").innerHTML = rem.map((t) => `<li>${t}</li>`).join("");
}

function kpi(label, value, note) {
  return `<div class="kpi__item"><b>${label}</b><span>${value}</span><small>${note}</small></div>`;
}
function row(th, td, cls = "") {
  return `<tr${cls ? ` class="${cls}"` : ""}><th>${th}</th><td>${td}</td></tr>`;
}

/** 残価ごとのバー。押すとその残価で計算し直す */
function paintBars(box, items) {
  const max = Math.max(...items.map((x) => x.value), 1e-9);
  box.innerHTML = items.map((x) => {
    const on = Math.abs(x.ratio - state.residual) < 1e-9;
    const h = Math.max(6, x.value / max * 100);
    const label = x.ratio === 0 ? "なし" : `${Math.round(x.ratio * 100)}%`;
    return `<button type="button" class="rbar${on ? " is-on" : ""}" data-r="${x.ratio}" aria-pressed="${on}"
      aria-label="残価${label === "なし" ? "なし" : label}：${x.text}">
      <span class="rbar__val">${x.text}</span>
      <span class="rbar__col"><i style="height:${h.toFixed(1)}%"></i></span>
      <span class="rbar__lab">${label}</span></button>`;
  }).join("");
}

/* ------------------------------------------------------------ はじめに一度だけ */
function init() {
  document.querySelectorAll("[data-mode-tab]").forEach((b) => {
    b.addEventListener("click", () => { setMode(b.dataset.modeTab); track("lease_mode", { mode: b.dataset.modeTab }); });
    b.addEventListener("keydown", (e) => {
      if (e.key !== "ArrowRight" && e.key !== "ArrowLeft") return;
      const next = state.mode === "calc" ? "reverse" : "calc";
      setMode(next);
      document.querySelector(`[data-mode-tab="${next}"]`).focus();
    });
  });

  ["price", "months", "rate", "misc", "pay"].forEach((id) => {
    const el = $(id);
    el.addEventListener("input", () => render({ showErr: false }));
    el.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); onRun(); } });
  });

  document.querySelectorAll("[data-months]").forEach((c) => c.addEventListener("click", () => {
    $("months").value = c.dataset.months;
    render({ showErr: false });
  }));

  document.querySelectorAll("[data-unit]").forEach((b) => b.addEventListener("click", () => setUnit(b.dataset.unit, true)));

  ["cBars", "rBars"].forEach((id) => $(id).addEventListener("click", (e) => {
    const b = e.target.closest("[data-r]");
    if (!b) return;
    state.residual = +b.dataset.r;
    render({ showErr: false });
  }));

  $("run").addEventListener("click", onRun);
  $("sample").addEventListener("click", onSample);

  // 記事の中の「金利を逆算する」などのリンク
  document.querySelectorAll("[data-go]").forEach((a) => a.addEventListener("click", (e) => {
    e.preventDefault();
    setMode(a.dataset.go, { scroll: true });
  }));

  setMode(location.hash === "#reverse" ? "reverse" : "calc");
  loadPaid();
}

/** 月額（円）とリース料率（%）の切り替え。入っている値は換算して残す */
function setUnit(unit, convert) {
  if (unit === state.unit) return;
  const price = num("price"), raw = num("pay");
  if (convert && raw > 0 && price > 0) {
    $("pay").value = unit === "pct" ? (raw / price * 100).toFixed(3) : String(Math.round(price * raw / 100));
  }
  state.unit = unit;
  document.querySelectorAll("[data-unit]").forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.unit === unit)));
  $("payLabel").textContent = unit === "yen" ? "月額リース料" : "リース料率";
  $("payUnit").textContent = unit === "yen" ? "円・税抜。見積書の月額" : "%・月額 ÷ 物件価額";
  $("pay").placeholder = unit === "yen" ? "95000" : "1.9";
  $("pay").step = unit === "yen" ? "1" : "0.001";
  render({ showErr: false });
}

function onRun() {
  const ok = render({ scroll: true, showErr: true });
  if (ok) track("lease_calc", { mode: state.mode, via: "button" });
}

function onSample() {
  const ex = EXAMPLE[state.mode];
  $("price").value = ex.price;
  $("months").value = ex.months;
  if (state.mode === "calc") {
    $("rate").value = ex.rate.toFixed(1);
    $("misc").value = ex.misc.toFixed(1);
  } else {
    setUnit("yen", false);
    $("pay").value = ex.pay;
  }
  state.residual = 0;
  render({ scroll: true, showErr: true });
  track("lease_calc", { mode: state.mode, via: "example" });
}

init();
