/* ============================================================================
 * /credit-pro/app.js — 画面と判定エンジンの接続
 * 計算は engine.js、Excel生成は xlsx-export.js。ここはUIだけを担当する。
 * ========================================================================== */
import { evaluate, emptyInput, INDUSTRIES, CAPITAL_TIERS, LISTING_OPTIONS, POLICY }
  from "./engine.js?v=31";
import { downloadXlsx } from "./xlsx-export.js?v=33";
import { checkLicense, payUrl, payUrlReady, companyFingerprint, forgetOrder, isStripeOrder, hasReturnOrder } from "./license.js?v=32";
import { scanPdf, buildPeriod, validatePeriod, toEngineFields } from "./pdf-extract.js?v=31";
import { renderViz, renderHead, attachTips, renderFigures, readingLines } from "./viz.js?v=31";

const $ = (id) => document.getElementById(id);
const COLS = ["今期（直近）", "前期", "前々期"];

const PL_ROWS = [
  ["sales", "売上高", 0], ["cogs", "売上原価", 0],
  [null, "売上総利益（粗利益）", "grossProfit"],
  ["sga", "販売費及び一般管理費", 0],
  [null, "営業利益", "operatingProfit"],
  ["nonOpInc", "営業外収益", 0], ["nonOpExp", "営業外費用", 0],
  [null, "経常利益", "ordinaryProfit"],
  ["extraInc", "特別利益", 0], ["extraExp", "特別損失", 0],
  [null, "税引前当期純利益", "pretaxProfit"],
  ["tax", "法人税・住民税及び事業税等", 0],
  [null, "当期純利益", "netProfit"],
  ["depreciation", "減価償却費（販管費・製造原価の合計）", 0],
];
const BS_ROWS = [
  ["cash", "現金・預金", 0], ["receivables", "受取手形・売掛金", 0],
  ["inventory", "棚卸資産", 0], ["otherCurrentAssets", "その他流動資産", 0],
  [null, "流動資産合計", "currentAssets"],
  ["tangible", "有形固定資産", 0], ["otherFixedAssets", "無形固定資産・投資その他", 0],
  [null, "固定資産合計", "fixedAssets"],
  ["deferred", "繰延資産", 0],
  [null, "資産合計", "totalAssets"],
  ["payables", "支払手形・買掛金", 0],
  ["shortDebt", "短期借入金（1年内返済分を含む）", 0],
  ["otherCurrentLiab", "その他流動負債", 0],
  [null, "流動負債合計", "currentLiab"],
  ["longDebt", "長期借入金・社債", 0], ["otherFixedLiab", "その他固定負債", 0],
  [null, "固定負債合計", "fixedLiab"],
  [null, "負債合計", "totalLiab"],
  ["equity", "純資産合計（自己資本）", 0],
  [null, "負債・純資産合計（総資本）", "totalCapital"],
  [null, "【検算】資産合計 － 負債・純資産合計", "balanceCheck"],
];

let state = emptyInput();

/* ------------------------------------------------------------------ 表示補助 */
/* ------------------------------------------------------------ 表示単位
 * 内部の計算はすべて百万円で行う。engine.js の規模スコアが
 * 百万円の絶対額（100 / 300 / 1,000 …）を閾値に持っているためで、
 * ここを動かすと採点そのものが変わってしまう。
 *
 * ただし state に入る値は整数に丸めない。
 * 千円単位の決算書の 4,767,955千円 を 4,768百万円 に丸めてしまうと、
 * 千円で表示し直したときに 4,768,000 となり、決算書と下3桁が食い違う。
 * 4767.955 のまま持ち、表示のときだけ選ばれた単位に直して丸める。
 */
const UNITS = {
  million:  { label: "百万円", mul: 1,    step: 1 },
  thousand: { label: "千円",   mul: 1000, step: 1 },
};
let dispUnit = "million";
const U_LABEL = () => UNITS[dispUnit].label;
/** 内部値（百万円）→ 表示単位 */
const toDisp = (v) => (typeof v === "number" ? v * UNITS[dispUnit].mul : v);
/** 表示単位 → 内部値（百万円） */
const fromDisp = (v) => (typeof v === "number" ? v / UNITS[dispUnit].mul : v);
/**
 * 貸借が合っているとみなすか。
 * 内部は百万円で持っているため、千円表示のときは 0.0005 百万円（＝500円）のような
 * 端数が出る。画面に「▲0 のズレ」と出してしまうと、利用者は原因を探せない。
 * そこで「表示している単位に直して四捨五入すると0」なら一致として扱う。
 */
const isBalanced = (v) => Math.round((Number(v) || 0) * UNITS[dispUnit].mul) === 0;

const yen = (n) => (!n ? "0" : (n < 0 ? "▲" : "") + Math.abs(Math.round(n)).toLocaleString());
/** 内部値を表示単位に直して整形する。画面に金額を出すときは必ずこれを通す */
const yenU = (n) => yen(toDisp(n));

/**
 * 表示単位を切り替える。state の値は触らないので、何度切り替えても数値は劣化しない。
 * why は利用者への説明文（「決算書に合わせた」のか「自分で選んだ」のか）。
 */
function setDispUnit(u, why) {
  if (!UNITS[u]) return;
  dispUnit = u;
  const radio = document.querySelector(`input[name="dispUnit"][value="${u}"]`);
  if (radio) radio.checked = true;
  document.querySelectorAll("[data-unit-tag]").forEach((el) => {
    el.textContent = "単位：" + U_LABEL();
  });
  // 表の上の説明文も追随させる。ここが「単位は百万円」のままだと、
  // 見出しの単位表記と食い違って利用者を混乱させる。
  document.querySelectorAll("[data-unit-lead]").forEach((el) => {
    el.textContent = `単位は${U_LABEL()}。`;
  });
  // 上のサンプル表示も同じ単位に揃える
  if (typeof renderDemo === "function" && document.getElementById("demoCol")) {
    try { renderDemo(demoKind); } catch (e) { /* 初期化前は何もしない */ }
  }
  const note = $("unitNote");
  if (note) {
    note.textContent = why === "自動"
      ? `読み込んだ決算書が${U_LABEL()}単位で記載されていたため、${U_LABEL()}に合わせました。必要であれば切り替えてください。`
      : "";
  }
  // 見出しの単位表記は buildTable の中にも埋まっているため、表ごと作り直す
  buildTable($("tPL"), PL_ROWS, true);
  buildTable($("tBS"), BS_ROWS, false);
  paint();
  render();
  if (lastRead) showRead(lastRead);
}
const pct = (n) => (n * 100).toFixed(1) + "%";
const esc = (s) => String(s ?? "").replace(/[&<>"]/g,
  (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

function buildTable(el, rows, withTerms) {
  let h = "<thead><tr><th style='text-align:left'>項　目<span class=\"unit-tag\" data-unit-tag>単位：" + U_LABEL() + "</span></th>";
  COLS.forEach((c) => (h += `<th>${c}</th>`));
  h += "</tr></thead><tbody>";
  if (withTerms) {
    h += "<tr><td class='lb'>決算期</td>";
    for (let i = 0; i < 3; i++)
      h += `<td><input type="text" data-k="terms" data-i="${i}" style="text-align:center"></td>`;
    h += "</tr>";
  }
  for (const [key, label, calc] of rows) {
    const isCalc = key === null;
    h += `<tr class="${isCalc ? "total" : ""}"><td class="lb">${label}</td>`;
    for (let i = 0; i < 3; i++) {
      h += isCalc
        ? `<td class="sum" data-calc="${calc}" data-i="${i}">0</td>`
        : `<td><input type="number" step="any" data-k="${key}" data-i="${i}"></td>`;
    }
    h += "</tr>";
  }
  el.innerHTML = h + "</tbody>";
}

/* ---------------------------------------------------------------- 初期化 */
function init() {
  INDUSTRIES.forEach((r) => $("f_industry").add(new Option(r[0], r[0])));
  CAPITAL_TIERS.forEach((t) => $("f_capitalTier").add(new Option(t[0], t[0])));
  LISTING_OPTIONS.forEach((l) => $("f_listing").add(new Option(l, l)));
  buildTable($("tPL"), PL_ROWS, true);
  buildTable($("tBS"), BS_ROWS, false);
  $("tRepay").innerHTML =
    "<thead><tr><th style='text-align:left'>項　目</th><th>1年目</th><th>2年目</th><th>3年目</th></tr></thead>" +
    "<tbody><tr><td class='lb'>年間約定返済額（元金）</td>" +
    [0, 1, 2].map((i) => `<td><input type="number" step="1" data-k="repayment" data-i="${i}"></td>`).join("") +
    "</tr><tr><td class='lb'>└ 自動見積の前提：長期借入金・社債の残存平均返済年数</td>" +
    `<td><input type="number" step="1" min="1" id="f_repayYears"></td>` +
    `<td colspan="2" class="repay-note"><span id="repayState"></span>` +
    `<button type="button" class="linklike" id="btnRepayAuto">自動見積に戻す</button></td></tr></tbody>`;
  $("btnRepayAuto").addEventListener("click", () => {
    state.repayManual = false; applyAutoRepay(); paint(); render();
  });

  document.addEventListener("input", onInput);
  document.querySelectorAll('input[name="dispUnit"]').forEach((el) =>
    el.addEventListener("change", () => setDispUnit(el.value, "手動")));
  $("btnDemo").addEventListener("click", () => { state = demo(); paint(); render(); openManual(); });
  $("btnSample").addEventListener("click", onSample);
  document.querySelectorAll("[data-demo]").forEach((b) =>
    b.addEventListener("click", () => renderDemo(b.dataset.demo)));
  if ($("demoCol")) renderDemo("good");
  $("btnClear").addEventListener("click", () => {
    state = emptyInput();
    state.baseDate = new Date().toISOString().slice(0, 10);
    paint(); render();
  });
  $("btnXlsx").addEventListener("click", onDownload);
  const rc = $("btnRecheck");
  if (rc) rc.addEventListener("click", () => refreshLicense());
  const fg = $("btnForget");
  if (fg) fg.addEventListener("click", () => {
    if (!confirm("前の会社の購入情報と判定結果を、この端末から消します。\n\nいま読み取っている決算書の数字は消えません。\nまだ使えるお支払いが残っている場合、それも使えなくなります。よろしいですか？")) return;
    forgetOrder(); clearSnap();
    refreshLicense();
    // 消したあとは、そのまま購入に進めるよう購入ボタンまで運ぶ
    const buy = $("btnBuy");
    if (buy) { buy.scrollIntoView({ behavior: "smooth", block: "center" }); }
  });
  $("btnRetry").addEventListener("click", refreshLicense);
  $("btnBuy").addEventListener("click", onBuy);
  initUploader();
  initShots();
  if (FREE_MODE) markFreeMode();

  // 決済から戻ったときだけ、入力内容と判定を復元して段を開く。
  // 初回は空の状態で「決算書を置く」だけに集中してもらう。
  const restored = restoreDraft();
  if (!restored) {
    state = emptyInput();
    state.baseDate = new Date().toISOString().slice(0, 10);
  }
  paint(); render();
  if (restored) {
    showStep("step4", true);
    showStep("step5", true);
    refreshLicense();
  }
}

/* ------------------------------------------------------------ 決済ゲート */
/* ------------------------------------------------------------------------
 * テスト用の無料モード
 *
 * FREE_BUILD は、テスト用リポジトリ（test_credit_test）に置くファイルでだけ true。
 * 本番リポジトリに置くファイルは false のままにすること。
 *
 * それに加えて、ホスト名が kazumono.com のときは何があっても無効にする。
 * 取り違えてテスト用のファイルを本番へ上げてしまっても、課金は外れない。
 * この二重の歯止めがあるので、事故で売上がゼロになることはない。
 * ---------------------------------------------------------------------- */
const FREE_BUILD = true;
const FREE_MODE = FREE_BUILD &&
  !/(^|\.)kazumono\.com$/i.test(String(location.hostname || ""));

let licensed = false;

/* ------------------------------------------------------------------------
 * 購入スナップショット
 *
 * 1回の支払い＝1社分を守るための仕組み。
 * 解錠できた瞬間の入力内容をまるごと控えておき、判定結果もExcelも
 * 「控えた内容」からだけ作る。あとから別の決算書に差し替えても、
 * 画面に出るのは買ったときの会社のままになる。
 *
 * 会社名だけを合図にすると、名前を変えずに数字だけ入れ替えられてしまう。
 * そこを塞ぐのがこの控えの役目。
 * ---------------------------------------------------------------------- */
const SNAP_KEY = "kazumono.credit-pro.paid";
/* 購入ボタンを押したときの入力内容。同じタブで戻れば sessionStorage から、
   PayPay のアプリなどを経由して別のタブで戻ったときは localStorage の控えから復元する */
const DRAFT_KEY = "kazumono.credit-pro.draft";
let paidSnap = null;
/**
 * 貸借が合っていない期の名前。空でなければ購入させない。
 * 資産合計と負債・純資産合計がずれたまま判定すると、自己資本比率も償還余力も
 * 全部おかしくなる。おかしい数字にお金を払わせないための歯止め。
 */
let balanceBad = [];
function loadSnap() { try { return JSON.parse(localStorage.getItem(SNAP_KEY) || "null"); } catch (e) { return null; } }
function saveSnap(o) { try { localStorage.setItem(SNAP_KEY, JSON.stringify(o)); } catch (e) { /* noop */ } }
function clearSnap() { paidSnap = null; try { localStorage.removeItem(SNAP_KEY); } catch (e) { /* noop */ } }

/** 決算の中身が入れ替わっていないかを見るための指紋（会社名は含めない） */
function figuresDigest(inp) {
  const keys = ["sales","cogs","sga","nonOpInc","nonOpExp","extraInc","extraExp","tax","depreciation",
    "cash","receivables","inventory","otherCurrentAssets","tangible","otherFixedAssets","deferred",
    "payables","shortDebt","otherCurrentLiab","longDebt","otherFixedLiab","equity","terms"];
  return JSON.stringify(keys.map((k) => (inp && inp[k]) || null));
}
/** いま画面にある内容が、買ったときの内容と同じか */
function snapMatchesLive() {
  return !!paidSnap && figuresDigest(paidSnap.input) === figuresDigest(state);
}

function showGate(which) {
  ["gateWait", "gateBuy", "gateOk", "gateOffline"].forEach((id) => {
    $(id).hidden = (id !== which);
  });
}

/**
 * テスト用の無料モードでの解錠状態。
 * 本番と同じく「買ったときの会社名と一致するときだけ開く」ようにしてある。
 * 挙動をそろえておかないと、テストの意味がなくなるため。
 */
function applyFreeMode() {
  const snap = loadSnap();
  const nm = state_name().trim();
  if (snap && snap.free && nm && snap.name === nm) { paidSnap = snap; licensed = true; }
  else { paidSnap = null; licensed = false; }
  showLicenseDiag(licensed ? "licensed" : "unlicensed", paidSnap ? paidSnap.order : "", "", 0);
  showGate(licensed ? "gateOk" : "gateBuy");
  try { render(); } catch (e) { /* 入力がまだ無いときは何もしない */ }
}

/** テスト環境であることを画面に出す。本番と取り違えないための目印。 */
function markFreeMode() {
  const bar = document.createElement("div");
  bar.className = "freebar";
  bar.innerHTML = "<b>テスト環境</b><span>お支払いなしで解錠できます。本番（kazumono.com）ではこの表示は出ません。</span>" +
    '<button type="button" id="btnFreeReset">解錠を取り消す</button>';
  document.body.prepend(bar);
  document.body.classList.add("has-freebar");
  const buy = $("btnBuy");
  if (buy) buy.textContent = "テスト用：無料で解錠してダウンロード";
  const bt = $("btnFreeReset");
  if (bt) bt.addEventListener("click", () => { clearSnap(); licensed = false; refreshLicense(); });
}

/** テスト用：お金を払わずに解錠する。本番の購入とまったく同じ控えを作る。 */
function freeUnlock() {
  const nm = state_name().trim();
  if (!nm) {
    $("buyNote").textContent = "会社名をご入力ください（テスト用の解錠でも、1社分の控えを作るため必要です）。";
    $("f_name")?.scrollIntoView({ behavior: "smooth", block: "center" });
    $("f_name")?.focus();
    return;
  }
  paidSnap = {
    free: true, order: "TEST-" + Date.now().toString(36).toUpperCase(), fp: "test",
    name: nm, input: JSON.parse(JSON.stringify(state)), at: Date.now(),
  };
  saveSnap(paidSnap);
  licensed = true;
  $("buyNote").textContent = "";
  showLicenseDiag("licensed", paidSnap.order, "", 0);
  showGate("gateOk");
  render();
  $("gateOk")?.scrollIntoView({ behavior: "smooth", block: "center" });
}

/**
 * ライセンスの確認口。checkLicense を直接呼ばず、必ずここを通す。
 *
 * テスト環境（FREE_MODE）では、サーバーに問い合わせても注文番号が無いため
 * 必ず「未購入」が返る。呼び出し側それぞれに迂回を書くと、今回のように
 * 書き忘れた1か所でダウンロードだけ止まる。入口をひとつにしておく。
 */
async function verifyLicense() {
  if (FREE_MODE) {
    const nm = state_name().trim();
    const ok = !!(paidSnap && paidSnap.free && nm && paidSnap.name === nm);
    return { state: ok ? "licensed" : "unlicensed",
             order: paidSnap ? paidSnap.order : "", reason: "", expiresAt: 0 };
  }
  return checkLicense(await companyFingerprint(state_name()));
}

/** 決済まわりで何が起きているかを画面に出し、切り分けられるようにする */
function showLicenseDiag(st, order, reason, expiresAt) {
  const el = $("licDiag");
  if (!el) return;
  const label = { licensed: "購入済み", unlicensed: "未購入", offline: "確認できず" }[st] || st;
  // 解錠されない理由は、利用者が次に何をすればよいか分かる言葉で出す
  const byReason = {
    other_company: "このお支払いは、別の会社の判定に使われています。1回のお支払いにつき1社分です。この会社の分をご入用の場合は、あらためてお求めください。",
    expired: "お支払いから24時間が過ぎたため、期限切れになりました。あらためてお求めください。",
    not_found: "決済の記録が見つかりません。決済直後の場合、記録が届くまで数十秒かかります。少しお待ちください。",
    no_fingerprint: "会社名が未入力です。会社の基本情報に会社名をご入力ください。",
    not_paid: "お支払いがまだ完了していないようです。完了していれば、少し待ってから「購入状況をもう一度確認する」を押してください。",
    wrong_product: "この注文番号は、財務でポン！のお支払いではないようです。お支払い済みの場合は、領収メールを添えてお問い合わせください。",
    wrong_link: "この注文番号は、財務でポン！のお支払いではないようです。お支払い済みの場合は、領収メールを添えてお問い合わせください。",
    bad_order: "注文番号の形が正しくありません。お支払い済みの場合は、領収メールを添えてお問い合わせください。",
    stripe_auth: "決済の確認先で問題が起きています。時間をおいて「購入状況をもう一度確認する」を押してください。",
    stripe_error: "決済の確認先で問題が起きています。時間をおいて「購入状況をもう一度確認する」を押してください。",
    no_key: "決済の確認先で問題が起きています。時間をおいて「購入状況をもう一度確認する」を押してください。",
  };
  const stripeOrder = isStripeOrder(order);
  if (stripeOrder) {
    // Stripe では通知を待たないので、「見つからない」は届いていないのではなく、その注文が無いという意味になる
    byReason.not_found = "この注文番号のお支払いが見つかりませんでした。お支払い済みの場合は、領収メールを添えてお問い合わせください。";
  }
  // この注文ではもう開けない理由（払い直しが必要、または番号そのものが違う）
  const dead = ["other_company", "expired", "wrong_product", "wrong_link", "bad_order"].includes(reason) ||
               (stripeOrder && reason === "not_found");
  const why =
    st === "licensed"
      ? "この会社の分を、" + (expiresAt ? expiresAt.slice(0, 16).replace("T", " ") + " まで" : "24時間") + "何度でもダウンロードできます。"
    : st === "offline" ? "決済の確認先に接続できませんでした。通信環境をご確認ください。"
    : (reason && byReason[reason]) ||
      (order ? "決済の記録が見つかりません。" : "まだ決済していない状態です。");
  // 状態が色で分かるようにする。緑＝購入済み、橙＝要対応、既定＝未購入
  el.classList.toggle("tip--warn", st === "unlicensed" && reason !== null);
  el.classList.toggle("tip--ok", st === "licensed");
  // すでに支払っている人に、もう一度払わせないための出し分け
  const paidButLocked = st === "unlicensed" && !!order && !dead;
  const buy = $("btnBuy"), recheck = $("btnRecheck");
  if (buy) buy.style.opacity = paidButLocked ? ".45" : "";
  if (recheck) recheck.style.fontSize = paidButLocked ? "16px" : "";
  // 使えない番号（別会社に紐づき済み／期限切れ）が残っていると、
  // 何度開いても橙色の警告が出続ける。その場合だけ消す手段を出す。
  const stale = st === "unlicensed" && !!order && dead;
  const fr = $("forgetRow");
  if (fr) fr.hidden = !stale;
  const dup = $("dupWarn");
  if (dup) {
    dup.hidden = !paidButLocked;
    dup.innerHTML = paidButLocked
      ? `<b>お支払いは受け付けられています。もう一度お支払いなさらないでください</b><span>` +
        `確認が済んでいないだけです。「購入状況をもう一度確認する」を押すか、時間をおいて開き直してください。</span>`
      : "";
  }
  el.innerHTML = `<b>決済の状態：${label}</b><span>注文番号：` +
    `<code>${order ? esc(String(order)) : "（なし）"}</code><br>${why}</span>`;
  el.hidden = false;
}

/**
 * 【Square の注文（移行前）だけ】決済直後は、Squareからの通知がこちらの確認より遅れて届くことがある。
 * 1回で諦めると「払ったのに解錠されない」状態のまま終わってしまうため、
 * 注文番号を持っているのに未購入と出た場合だけ、数秒おきに数回だけ確認し直す。
 */
async function pollLicense(order) {
  // Squareからの通知は、実測で2分ほどかかることがある。
  // 「止まっている」と誤解されないよう、経過秒数を出しながら最大3分待つ。
  const TRIES = 60, WAIT = 3000;
  showGate("gateWait");
  for (let i = 0; i < TRIES; i++) {
    await new Promise((r) => setTimeout(r, WAIT));
    const sec = Math.round(((i + 1) * WAIT) / 1000);
    const el = $("licDiag");
    if (el) {
      el.classList.remove("tip--warn");
      el.innerHTML =
        `<b>お支払いを確認しています… 経過 ${sec} 秒</b><span>` +
        `Square からの入金通知が届くまで、<b>2分ほどかかることがあります</b>。` +
        `この画面を開いたままお待ちください。確認できしだい、ダウンロードボタンが出ます。<br>` +
        `ページを閉じても、あとで開き直せば続きから確認できます。<br>` +
        `注文番号：<code>${esc(String(order))}</code></span>`;
    }
    const wait = $("gateWait");
    if (wait) wait.innerHTML =
      `<p class="note-s"><b>お支払いを確認しています…（経過 ${sec} 秒／最大3分）</b><br>` +
      `Squareからの通知待ちです。この画面のままお待ちください。</p>`;
    const { state, expiresAt } = await verifyLicense();
    if (state === "licensed") {
      licensed = true;
      if (window.gtag) gtag("event", "license_ok", { tool: "credit-pro" });
      refreshLicense();   // 控えの作成とゲートの開閉は refreshLicense に任せる
      return true;
    }
    if (state === "offline") break;
  }
  const el = $("licDiag");
  if (el) {
    el.classList.add("tip--warn");
    el.innerHTML = `<b>3分待っても確認できませんでした</b><span>` +
      `お支払いが完了していれば、記録は必ず後から届きます。<b>二重にお支払いなさらないでください。</b><br>` +
      `少し時間をおいて「購入状況をもう一度確認する」を押すか、このページを開き直してください。<br>` +
      `それでも解錠されない場合は、この注文番号を添えてお問い合わせください：<code>${esc(String(order))}</code></span>`;
    el.hidden = false;
  }
  showGate("gateBuy");
  return false;
}

/** 判定に使っている会社名（Workerには送らず、ハッシュ化して使う） */
function state_name() { return state && state.name ? state.name : ""; }

async function refreshLicense() {
  if (FREE_MODE) { applyFreeMode(); return; }
  showGate("gateWait");
  const fp = await companyFingerprint(state_name());
  const { state: st, order, reason, expiresAt } = await verifyLicense();
  licensed = st === "licensed";
  if (licensed) {
    // 控えは注文番号ごとに1回だけ取る。以後は差し替えても上書きしない。
    const saved = loadSnap();
    if (saved && saved.order === order) {
      paidSnap = saved;
    } else {
      paidSnap = { order, fp, name: state_name(), input: JSON.parse(JSON.stringify(state)), at: Date.now() };
      saveSnap(paidSnap);
    }
    // 解錠の控えを取ったので、別タブ用に端末へ残した下書きは役目を終える
    try { localStorage.removeItem(DRAFT_KEY); } catch (err) { /* noop */ }
    try { render(); } catch (e) { /* 入力がまだ無いときは何もしない */ }
  } else {
    paidSnap = null;
  }
  showLicenseDiag(st, order, reason, expiresAt);
  // Square の注文（移行前）は通知の到着が遅れることがあるので、記録が無いときだけ確認し直す。
  // Stripe の注文は Worker が Stripe に直接たずねるため、待っても結果は変わらない
  if (st === "unlicensed" && order && reason === "not_found" && !isStripeOrder(order)) { pollLicense(order); return; }
  showGate(st === "licensed" ? "gateOk" : st === "offline" ? "gateOffline" : "gateBuy");
  if (st === "licensed" && window.gtag) gtag("event", "license_ok", { tool: "credit-pro" });
}

/**
 * 購入へ進む。入力内容を保存してから決済（Stripe）へ送る。
 * 決済後に戻ってきたとき、同じ内容のまま続けられるようにするため。
 */
/**
 * 貸借が合っていないあいだは、購入ボタンを押せない見た目にする。
 * 押せてしまうと、狂った判定にお金を払わせることになる。
 */
function syncBuyState() {
  const buy = $("btnBuy"); if (!buy) return;
  const off = balanceBad.length > 0;
  buy.classList.toggle("is-off", off);
  buy.setAttribute("aria-disabled", String(off));
  const note = $("buyNote");
  if (off && note) {
    note.textContent =
      `${balanceBad.join("・")}で、資産合計と負債・純資産合計が一致していません。` +
      `一致させてからお進みください（ずれたままの判定結果は、正しい数字になりません）。`;
  } else if (note && note.dataset.balance === "1") {
    note.textContent = ""; note.dataset.balance = "";
  }
  if (off && note) note.dataset.balance = "1";
}

async function onBuy(e) {
  e.preventDefault();
  // 貸借が合っていないときは、ここで止める。
  if (balanceBad.length) {
    $("buyNote").textContent =
      `${balanceBad.join("・")}で、資産合計と負債・純資産合計が一致していません。` +
      `この状態では判定結果が正しくならないため、購入できません。数値をご確認のうえ、一致させてからお進みください。`;
    const al = $("alertBalance");
    if (al) al.scrollIntoView({ behavior: "smooth", block: "center" });
    return;
  }
  // テスト環境では、決済へ行かずにその場で解錠する
  if (FREE_MODE) { freeUnlock(); return; }
  // 支払いリンクが未設定のまま押されたときは、遷移せずに理由を出す。
  // 黙って決済のエラーページへ飛ばすと、原因の切り分けができなくなる。
  // 会社名が空だと、決済しても「どの会社の分か」を確定できず解錠できない
  if (!state_name().trim()) {
    $("buyNote").textContent =
      "会社名をご入力ください。お支払いは1社分ごとのため、会社名が必要です（「会社の基本情報」欄）。";
    $("f_name")?.scrollIntoView({ behavior: "smooth", block: "center" });
    $("f_name")?.focus();
    return;
  }
  if (!payUrlReady()) {
    // 何が読み込まれているかを必ず表示する。
    // 「直したのに直らない」の大半は、ブラウザが古いlicense.jsを使っているだけなので、
    // 実際の値が見えれば一目で切り分けられる。
    $("buyNote").innerHTML =
      "支払いリンクが未設定です。credit-pro/license.js の PAY_URL に、Stripeで作成した支払いリンクを貼ってください。<br>" +
      "いま読み込まれている値：<code>" + esc(String(payUrl())) + "</code><br>" +
      "すでに貼り替えたのにこの表示が出る場合は、ブラウザが古いファイルを使っています。" +
      "Ctrl+Shift+R（Mac は Cmd+Shift+R）で読み込み直してください。";
    return;
  }
  // 会社名のハッシュを支払いリンクに付けて、決済と会社を結び付ける（1回の決済＝1社を、決済の時点で決める）
  const fp = await companyFingerprint(state_name());
  try { sessionStorage.setItem(DRAFT_KEY, JSON.stringify(state)); } catch (err) { /* noop */ }
  try { localStorage.setItem(DRAFT_KEY, JSON.stringify({ at: Date.now(), state })); } catch (err) { /* noop */ }
  if (window.gtag) gtag("event", "begin_checkout", { tool: "credit-pro", value: 500, currency: "JPY" });
  location.href = payUrl(fp);
}

/** 決済から戻ったとき、入力内容を復元する */
function restoreDraft() {
  try {
    let raw = sessionStorage.getItem(DRAFT_KEY);
    // 同じタブの控えが無く、決済から戻ってきたところなら、購入ボタンを押したときの控え（24時間以内）を使う
    if (!raw && hasReturnOrder()) {
      const bk = JSON.parse(localStorage.getItem(DRAFT_KEY) || "null");
      if (bk && bk.state && Date.now() - bk.at < 24 * 3600e3) {
        raw = JSON.stringify(bk.state);
        try { sessionStorage.setItem(DRAFT_KEY, raw); } catch (err) { /* noop */ }
      }
    }
    if (!raw) return false;
    const d = JSON.parse(raw);
    if (d && typeof d === "object") { state = { ...emptyInput(), ...d }; return true; }
  } catch (e) { /* noop */ }
  return false;
}

function onInput(e) {
  const t = e.target, k = t.dataset.k, i = t.dataset.i;
  // 返済額を手で触ったら、以後は自動見積で上書きしない
  if (k === "repayment") state.repayManual = true;
  if (k !== undefined && i !== undefined) {
    // 金額欄は表示単位で入力されるので、内部の百万円へ戻してから収める
    state[k][+i] = t.type === "number"
      ? (t.value === "" ? 0 : fromDisp(parseFloat(t.value)))
      : t.value;
  } else if (t.id && t.id.startsWith("f_")) {
    const f = t.id.slice(2);
    state[f] = t.type === "number" ? (t.value === "" ? 0 : parseFloat(t.value)) : t.value;
  } else return;
  render();
}

function paint() {
  ["name", "industry", "capitalTier", "listing", "founded", "baseDate", "employees",
   "capital", "ceoName", "ceoAge", "industryYears", "ceoYears", "ownHome",
   "disclosure", "successor", "memo"].forEach((f) => {
    const el = $("f_" + f); if (el) el.value = state[f] ?? "";
  });
  document.querySelectorAll("[data-k][data-i]").forEach((el) => {
    const v = state[el.dataset.k]?.[+el.dataset.i];
    // 金額欄は選ばれた単位に直して見せる。丸めは表示のときだけで、state は元の精度を保つ
    // 表示は整数に丸める。state は元の精度を保つので、単位を往復しても値は劣化しない。
    // 丸めないと百万円表示のときに 7426.9 のような小数が入力欄に出てしまう。
    el.value = el.type === "number"
      ? (typeof v === "number" ? Math.round(toDisp(v)) : "")
      : (v ?? "");
    // PDFから読み取れなかった項目は枠を赤くして、どこを埋めればよいか一目で分かるようにする
    const miss = missingCells[el.dataset.k];
    el.classList.toggle("is-missing", !!miss && miss.includes(+el.dataset.i));
  });
  syncFinAccordions();
}

/**
 * 貸借対照表・損益計算書の折りたたみを、読み取り結果に合わせて開閉する。
 * 全部読めていれば閉じて画面を短くし、抜けがあれば開いて赤い欄を見せる。
 */
function syncFinAccordions() {
  const BS = ["cash","receivables","inventory","otherCurrentAssets","tangible",
              "otherFixedAssets","deferred","payables","shortDebt","otherCurrentLiab",
              "longDebt","otherFixedLiab","equity"];
  const PL = ["sales","cogs","sga","nonOpInc","nonOpExp","extraInc","extraExp","tax","depreciation"];
  const apply = (id, keys, label) => {
    const acc = $(id), st = $(id + "-state");
    if (!acc) return;
    const n = keys.reduce((a, k) => a + (missingCells[k] ? missingCells[k].length : 0), 0);
    const touched = Object.keys(missingCells).length > 0 || anyValue(keys);
    if (!touched) { if (st) st.textContent = ""; return; }
    if (n > 0) {
      acc.open = true;
      if (st) { st.textContent = `${n}か所が未取得です`; st.className = "acc__state is-missing-tag"; }
    } else {
      acc.open = false;
      if (st) { st.textContent = "すべて読み取れました", st.className = "acc__state is-ok-tag"; }
    }
  };
  const anyValue = (keys) => keys.some((k) => Array.isArray(state[k]) && state[k].some((v) => v));
  apply("accBS", BS, "貸借対照表");
  apply("accPL", PL, "損益計算書");
}

/* ---------------------------------------------------------------- 再計算 */
/**
 * 借入金の返済計画を自動で見積もる。
 * Excel版 Pro と同じ考え方：長期借入金・社債 ÷ 残存平均返済年数（既定5年）を、
 * 1〜3年目とも同額の約定返済額とみなす。返済予定表がある場合は手で上書きできる。
 * 当座貸越・短期継続融資の折返し分は、返済していないので含めない（＝短期借入金は使わない）。
 */
function applyAutoRepay() {
  if (state.repayManual) return;
  const yrs = Math.max(1, Number(state.repayYears) || 5);
  const v = Math.round((Number(state.longDebt && state.longDebt[0]) || 0) / yrs);
  state.repayment = [v, v, v];
  [0, 1, 2].forEach((i) => {
    const el = document.querySelector(`input[data-k="repayment"][data-i="${i}"]`);
    if (el) el.value = toDisp(v);
  });
  const ys = $("f_repayYears"); if (ys && ys.value === "") ys.value = yrs;
  const st2 = $("repayState");
  if (st2) st2.textContent = v > 0
    ? `長期借入金・社債 ${yenU(state.longDebt[0])}${U_LABEL()} ÷ ${yrs}年 を初期値にしています。　`
    : "長期借入金・社債を入れると初期値が自動で入ります。　";
}

function render() {
  applyAutoRepay();
  const r = evaluate(state);
  document.querySelectorAll("[data-calc]").forEach((el) => {
    const v = r.fy[+el.dataset.i][el.dataset.calc];
    el.textContent = yenU(v);
    if (el.dataset.calc === "balanceCheck")
      el.style.background = isBalanced(v) ? "" : "#FBEDE6";
  });
  const bad = r.fy.map((p, i) => (isBalanced(p.balanceCheck) ? null : COLS[i])).filter(Boolean);
  balanceBad = bad;
  const al = $("alertBalance");
  al.className = "alert" + (bad.length ? " on" : "");
  al.textContent = bad.length
    ? `${bad.join("・")}で、資産合計と負債・純資産合計が一致していません。` +
      `このままでは自己資本比率も償還余力も正しく出ないため、判定結果はお売りできません。` +
      `「ここだけ入力してください」欄か、下の入力欄で数値をご確認ください。`
    : "";
  syncBuyState();
  // 判定結果とグラフは有料。未購入のあいだは中身を一切出さない。
  // 購入済みのときは、買ったときの内容（控え）からだけ作る。
  if (licensed && paidSnap) {
    const paidR = evaluate(paidSnap.input);
    const changed = !snapMatchesLive();
    $("resultCol").innerHTML = (changed ? changedNotice(paidSnap.name) : "") + report(paidR);
    attachTips($("resultCol"));
  } else {
    $("resultCol").innerHTML = lockedCard();
  }
}

function report(r) {
  const s = r.scores, cur = r.cur, bm = r.benchmark, red = r.redemption;
  const cls = s.rank === "C" ? " is-mid" : (s.rank === "D" || s.rank === "E") ? " is-low" : "";
  // 6軸の内訳は renderViz の「スコアの内訳」で図として出している

  const p = r.ratios.periods;
  const U = `<span class="unit-tag">${U_LABEL()}</span>`, PC = '<span class="unit-tag">％</span>';
  const figs = [
    ["売上高", U, yenU(cur.sales), yenU(r.prev.sales), yenU(r.prev2.sales)],
    ["経常利益", U, yenU(cur.ordinaryProfit), yenU(r.prev.ordinaryProfit), yenU(r.prev2.ordinaryProfit)],
    ["当期純利益", U, yenU(cur.netProfit), yenU(r.prev.netProfit), yenU(r.prev2.netProfit)],
    ["自己資本比率", PC, pct(p[0].equityRatio), pct(p[1].equityRatio), pct(p[2].equityRatio)],
    ["売上高経常利益率", PC, pct(p[0].ordinaryMargin), pct(p[1].ordinaryMargin), pct(p[2].ordinaryMargin)],
    ["有利子負債", U, yenU(cur.interestBearingDebt), yenU(r.prev.interestBearingDebt), yenU(r.prev2.interestBearingDebt)],
    ["簡易キャッシュフロー", U, yenU(cur.simpleCF), yenU(r.prev.simpleCF), yenU(r.prev2.simpleCF)],
  ].map(([n, u, a, b, c]) =>
    `<tr><th>${n}${u}</th><td>${a}</td><td>${b}</td><td>${c}</td></tr>`).join("");

  return `
  ${renderHead(r, { U_LABEL }, POLICY[s.rank])}

  <div class="calc">
    <h2>与信限度額の目安</h2>
    <p class="calc__hint">自己資本を基準にした金額と月商を基準にした金額のうち、小さいほうです。一次スクリーニングの出発点としてお使いください。</p>
    <div class="result${cls}" style="text-align:center;">
      <div class="result__score"><b>${yenU(r.creditLimit.value)}</b> ${U_LABEL()}<span class="unit-tag">単位：${U_LABEL()}</span></div>
    </div>
  </div>

  ${renderViz(r, { yenU, U_LABEL, pct })}

  <div class="calc">
    <h2>財務ハイライト（直近3期）</h2>
    <p class="calc__hint">金額の単位は${U_LABEL()}です。比率は％で表示しています。</p>
    <table class="figs"><thead><tr><th>項　目<span class="unit-tag">${U_LABEL()} ／ ％</span></th><th>今期</th><th>前期</th><th>前々期</th></tr></thead>
      <tbody>${figs}</tbody></table>
    <p class="calc__hint" style="margin-top:10px;">
      自己資本比率の業種基準は ${pct(bm.equityRatio)}、売上高経常利益率の業種基準は ${pct(bm.ordinaryMarginAvg3)} です
      （${esc(r.input.industry.trim())}／${esc(r.input.capitalTier)}）。
      債務償還年数は ${red.simpleCF <= 0 ? "算定不能（返済原資なし）" : red.years.toFixed(1) + "年"} です。</p>
  </div>

  <div class="calc">
    <h2>自動所見</h2>
    <p class="calc__hint">該当する項目だけが表示されます。稟議書の所見欄にそのままお使いいただけます。</p>
    <p style="font-weight:700;margin:14px 0 6px;">評価できる点</p>
    <ul class="remarks">${r.comments.strengths.map((t) => `<li>${esc(t)}</li>`).join("")}</ul>
    <p style="font-weight:700;margin:18px 0 6px;">留意すべき点</p>
    <ul class="remarks">${r.comments.concerns.map((t) => `<li>${esc(t)}</li>`).join("")}</ul>
  </div>`;
}

/* ------------------------------------------------- サンプルExcel（購入前）
 * 静的なファイルを置かず、その場で組み立てる。
 * 出力の書式を直したときにサンプルだけ古くなる、という事故が起きない。
 * 中身は記入例の架空データなので、購入前でも配れる。
 */
async function onSample() {
  // サンプルは、購入者に渡すのとまったく同じ経路で作る。
  // 静的ファイルを置くと本番と中身がずれるので、その場で組み立てる。
  const btn = $("btnSample"), note = $("sampleNote");
  const label = btn.textContent;
  btn.disabled = true; btn.textContent = "作成中…";
  note.textContent = "図を描いてExcelに貼っています。数秒かかります。";
  try {
    const d = demoKind === "bad" ? DEMO_BAD() : DEMO_GOOD();
    const yrs = Math.max(1, Number(d.repayYears) || 5);
    const v = Math.round((Number(d.longDebt[0]) || 0) / yrs);
    d.repayment = [v, v, v];
    const r = evaluate(d);
    // 画面で選んでいる単位（千円／百万円）をそのままサンプルにも反映する
    const f = { yenU, U_LABEL, pct };
    const figs = await renderFigures(r, f, 2);
    await downloadXlsx(r, "財務でポン_サンプル.xlsx", UNITS[dispUnit], figs, readingLines(r, f));
    note.textContent = `ダウンロードしました（単位：${U_LABEL()}）。⑥ダッシュボードのシートに図が入っています。`;
    if (window.gtag) gtag("event", "xlsx_sample", { tool: "credit-pro" });
  } catch (e) {
    note.textContent = "作成に失敗しました：" + e.message;
  } finally {
    btn.disabled = false; btn.textContent = label;
  }
}

/** 買ったあとに決算書を差し替えたときの断り書き */
function changedNotice(paidName) {
  return `
  <div class="calc changed">
    <b>いま表示しているのは、お支払い時に判定した「${esc(paidName || "（会社名なし）")}」の結果です。</b>
    <p>そのあと画面の内容が変わっています。別の決算書を判定するには、あらためてお求めください。
    1回のお支払いにつき1社分です。</p>
  </div>`;
}

/**
 * 未購入のときに判定結果の代わりに出すカード。
 * スコアもランクも出さない。何が見られるのかだけを書く。
 */
function lockedCard() {
  return `
  <div class="calc lock">
    <div class="lock__badge">判定は完了しました</div>
    <h2 class="lock__h">結果とグラフは、お支払い後にご覧いただけます</h2>
    <p class="lock__lead">読み取った内容はこの画面に残っています。決済後、そのまま結果が開きます。</p>
    <ul class="lock__list">
      <li>総合評点（100点満点）と信用程度 A〜E、取引方針の目安</li>
      <li>与信限度額の目安（自己資本基準・月商基準のいずれか小さい方）</li>
      <li>財務ハイライト（直近3期）と自動所見</li>
      <li>7つの図 — ①スコアの内訳／②財務指標のかたち／③貸借対照表のかたち／④売上と利益の推移／⑤返せるお金と、返す額／⑥借金を返し切るまでの年数／⑦グラフから読み取れること</li>
      <li>稟議に添付できるExcel（6シート・ダッシュボード付き）</li>
    </ul>
    <p class="lock__note">どんなものが出てくるかは、<a href="#step1">ページ上部の「評価の高い会社／低い会社」</a>で実物をご覧いただけます。</p>
  </div>`;
}

/* -------------------------------------------------------------- ダウンロード */
async function onDownload() {
  $("dlNote").textContent = FREE_MODE ? "" : "確認しています…";
  // ボタンの表示状態だけに頼らず、実行の直前にもう一度確認する。
  // このとき会社名の指紋を必ず一緒に送る。送り忘れると
  // 「別の会社に使い回そうとしている」と判定され、解錠が取り消されてしまう。
  const { state: st, order, reason, expiresAt } = await verifyLicense();
  if (st !== "licensed") {
    licensed = false; paidSnap = null;
    showLicenseDiag(st, order, reason, expiresAt);
    $("dlNote").textContent = "";
    showGate(st === "offline" ? "gateOffline" : "gateBuy");
    return;
  }
  $("dlNote").textContent = "";
  if (!paidSnap) { showGate("gateBuy"); return; }
  // 出力するのは、買ったときの内容。いま画面にある別の決算書ではない。
  const r = evaluate(paidSnap.input);
  if (r.fy.some((x) => !isBalanced(x.balanceCheck)) &&
      !confirm("貸借対照表の検算が0になっていません。このまま出力しますか？")) return;
  const btn = $("btnXlsx");
  btn.disabled = true;
  const label = btn.textContent;
  btn.textContent = "Excelを作成中…";
  $("dlNote").textContent = "図を描いてExcelに貼っています。数秒かかります。";
  try {
    const vf = { yenU, U_LABEL, pct };
    const figs = await renderFigures(r, vf, 2);
    await downloadXlsx(r, null, UNITS[dispUnit], figs, readingLines(r, vf));
    $("dlNote").textContent = "ダウンロードしました";
    if (window.gtag) gtag("event", "xlsx_download", { tool: "credit-pro" });
  } catch (e) {
    $("dlNote").textContent = "作成に失敗しました：" + e.message;
  } finally {
    btn.disabled = false;
    btn.textContent = label;
  }
}


/* ==========================================================================
 * 決算書PDFの読み取り
 * ・pdf.js はサイト内に置いてある（外部CDNを使わない）
 * ・読み取りはすべてブラウザ内。ファイルはどこにも送信しない
 * ・読み取った値は「そのまま反映」せず、必ず確認画面を挟む
 * ======================================================================== */
const PERIOD_LABEL = ["今期（直近）", "前期", "前々期"];
let pdfjsLib = null;
let pending = null;      // 確認待ちの読み取り結果
let accepted = null;     // 判定済みの読み取り結果（次のPDFを足せるよう保持する）
let metaCompany = null;  // PDFから読み取れた会社名
let metaCompanySeen = [];// 読み込んだPDFに出てきた会社名（取り違え検知用）
let missingCells = {};   // 読み取れなかった項目 { キー: [期のindex] }

async function getPdfjs() {
  if (pdfjsLib) return pdfjsLib;
  pdfjsLib = await import("../vendor/pdf.min.mjs");
  pdfjsLib.GlobalWorkerOptions.workerSrc = new URL("../vendor/pdf.worker.min.mjs", import.meta.url).href;
  return pdfjsLib;
}

function initUploader() {
  const zone = $("dropZone"), input = $("fileInput");
  zone.addEventListener("click", () => input.click());
  zone.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); input.click(); } });
  input.addEventListener("change", () => { if (input.files.length) readFiles([...input.files]); });
  ["dragenter", "dragover"].forEach((ev) =>
    zone.addEventListener(ev, (e) => { e.preventDefault(); zone.classList.add("is-over"); }));
  ["dragleave", "drop"].forEach((ev) =>
    zone.addEventListener(ev, (e) => { e.preventDefault(); zone.classList.remove("is-over"); }));
  zone.addEventListener("drop", (e) => {
    const fs = [...(e.dataTransfer?.files || [])].filter((f) => /\.pdf$/i.test(f.name));
    if (fs.length) readFiles(fs);
  });
  $("btnApply").addEventListener("click", applyRead);
  // 最初からやり直す：手入力した内容も含めて、まっさらな状態に戻す。
  // 前回の入力が残っていると、次の会社の判定に混ざって事故になる。
  const doReread = () => {
    if (!confirm("入力した内容をすべて消して、最初からやり直します。よろしいですか？")) return;
    try { sessionStorage.removeItem(DRAFT_KEY); } catch (err) { /* noop */ }
    try { localStorage.removeItem(DRAFT_KEY); } catch (err) { /* noop */ }
    location.reload();
  };
  $("btnReread").addEventListener("click", doReread);
  { const b2 = $("btnReread2"); if (b2) b2.addEventListener("click", doReread); }
  // PDFが無い場合の導線
  $("btnNoPdf").addEventListener("click", () => {
    showStep("step3", true);
    $("readBanner").innerHTML =
      '<div class="read-banner is-supp"><b>数値を直接ご入力ください</b>' +
      '<span>下の入力欄を上から順に埋めて、いちばん下の「この内容で判定する」を押してください。' +
      '決算書PDFをお持ちの場合は、上に置いていただくほうが早く済みます。</span></div>';
    $("readTable").innerHTML = "";
    $("readWarn").innerHTML = "";
    const fx = $("readFix"); if (fx) { fx.innerHTML = ""; fx.hidden = true; }
    $("btnApply").hidden = false;
    openManual();
    $("step3").scrollIntoView({ behavior: "smooth", block: "start" });
  });
  // 手入力欄の開閉
  $("btnManual").addEventListener("click", () => {
    const body = $("manualBody"), open = body.hidden;
    body.hidden = !open;
    $("btnManual").setAttribute("aria-expanded", String(open));
    $("btnManual").textContent = open ? "入力欄を閉じる" : "入力欄を開く";
  });
}

function setProgress(msg, ratio) {
  $("readState").hidden = false;
  $("readMsg").textContent = msg;
  $("progBar").style.width = Math.round((ratio || 0) * 100) + "%";
}

async function readFiles(files) {
  showStep("step3", false);
  try {
    const pdfjs = await getPdfjs();
    // 前回までに読み取れた期を引き継ぐ。
    // 「第10期、第9期…」と1件ずつ投げても積み上がるようにするため、毎回捨てない。
    const periods = (pending || accepted || []).slice();
    for (let fi = 0; fi < files.length; fi++) {
      const f = files[fi];
      setProgress(`${f.name} を読み取っています…`, 0);
      const buf = await f.arrayBuffer();
      const found = await scanPdf(pdfjs, buf, (pno, total) =>
        setProgress(`${f.name} を読み取っています…（${pno} / ${total} ページ）`, pno / total));
      // スキャン画像PDF（テキストレイヤーが無い）は、その旨を明示する
      if (found._noText) { periods.push({ file: f.name, failed: true, image: true }); continue; }
      // 対象外の業種、および四半期・中間の決算書は、読めても判定に進ませない。
      // 四半期の損益計算書は3か月ぶんなので、年商として扱うと回転率も償還年数も静かに狂う。
      const reject = found._outOfScope || found._interim;
      if (reject) { periods.push({ file: f.name, failed: true, scope: reject }); continue; }
      if (!found.BS && !found.PL) { periods.push({ file: f.name, failed: true }); continue; }
      // 有報・短信は1本で2期分（当期・前期）取れる
      const isTwoYear = (found.BS || found.PL).kind === "years";
      const yis = isTwoYear ? [-1, 0] : [-1];
      const per = found._period || {};
      const unit = found._unit || null;
      if (found._company) {
        if (!metaCompany) metaCompany = found._company;
        if (!metaCompanySeen.includes(found._company)) metaCompanySeen.push(found._company);
      }
      // 同じファイル、または同じ決算期をすでに読み取っていれば置き換える（重複防止）
      const sameIdx = periods.findIndex((q) =>
        q.file === f.name ||
        (q.period && per.end && q.period.end === per.end) ||
        (q.period && typeof per.no === "number" && q.period.no === per.no && !per.end && !q.period.end));
      if (sameIdx >= 0) periods.splice(sameIdx, 1);
      for (const yi of yis) {
        const raw = buildPeriod(found, yi);
        const values = scaleToMillion(raw.values, unit);
        const source = raw.source, warnings = raw.warnings.slice();
        if (!unit)
          warnings.push("金額の単位を読み取れませんでした。百万円として扱っています。単位が違う場合は、上の「金額の単位」で切り替えてご確認ください。");
        const { diff, messages } = validatePeriod(values);
        const grade = gradePeriod(values, diff, messages);
        // 1本で2期取れる様式（有報・短信）は、yi=0 が前期。並べ替えの鍵をずらしておく
        const key = yi === 0 ? shiftBack(per) : per;
        periods.push({ file: f.name, yi, values, source, warnings, diff, messages, grade,
                       period: key, unit,
                       pages: Object.entries(found).map(([k, v]) => `${k} ${v.page}ページ`).join(" / ") });
      }
    }
    setProgress("読み取りが終わりました。", 1);
    // 決算書の単位に表示を合わせる。
    // 千円・円で書かれた決算書を百万円で見せると、利用者が原本と突き合わせられない。
    // 読み取れた期のうち、最も細かい単位に寄せる（混在時に情報が落ちないようにするため）。
    const units = periods.filter((p) => !p.failed && p.unit).map((p) => p.unit.toMillion);
    if (units.length) setDispUnit(Math.min(...units) < 1 ? "thousand" : "million", "自動");
    sortPeriods(periods);
    // 3期を超えたぶんは、古いものから落とす（並べ替え済みなので後ろが古い）
    const live = periods.filter((p) => !p.failed);
    const kept = live.slice(0, 3);
    const dropped = live.length - kept.length;
    const failed = periods.filter((p) => p.failed);
    const merged = kept.concat(failed);
    if (dropped > 0) merged._dropped = dropped;
    showRead(merged);
  } catch (e) {
    setProgress("読み取りに失敗しました：" + e.message, 0);
  }
}






/**
 * 複数の会社名が混ざっていないかを知らせる。
 * 別会社の決算書を1つの判定に混ぜると、まったく意味のない数字が出るため、
 * 気づけるように必ず画面に出す。
 */
function warnCompany() {
  const box = $("companyWarn"), btn = $("btnApply");
  if (!box) return;
  const mixed = metaCompanySeen.length > 1;
  if (mixed) {
    // 別会社の決算書が混ざったままでは、出てくる数字に意味が無い。
    // 警告だけでは押されてしまうため、判定そのものを止める。
    box.hidden = false;
    box.classList.add("tip--warn");
    box.innerHTML = `<b>別の会社の決算書が混ざっています。このままでは判定できません</b><span>` +
      `読み取れた会社名：<b>${metaCompanySeen.map(esc).join("</b> と <b>")}</b>。` +
      `別々の会社の数字を1つにまとめても、結果は意味を持ちません。` +
      `下の「やり直す」で読み取りを消してから、1社分だけを置き直してください。</span>`;
    if (btn) { btn.disabled = true; btn.textContent = "会社が混ざっているため判定できません"; }
  } else {
    if (btn) { btn.disabled = false; btn.textContent = "この内容で判定する"; }
    if (metaCompany) {
      box.hidden = false;
      box.classList.remove("tip--warn");
      box.innerHTML = `<b>会社名を自動で入れました：${esc(metaCompany)}</b>` +
        `<span>PDFから読み取った名前です。誤っていれば、下の「会社の基本情報」で直してください。</span>`;
    } else {
      box.hidden = true;
    }
  }
}

/* --------------------------------------------------- Excel見本スライダー */
const SHOTS = [
  ["./assets/1-summary.jpg?v=2", "①判定サマリー",
   "総合評点・信用程度A〜E・6軸の評点内訳・与信限度額の目安・財務ハイライト・自動所見を1枚に。決裁欄つきで、そのまま回付できます。"],
  ["./assets/2-financial.jpg?v=2", "②財務分析",
   "損益計算書と貸借対照表の3期比較に、主要財務指標14種と運転資金分析。画面のレーダーは6指標ですが、ここでは14指標を業種基準と並べて見られます。"],
  ["./assets/3-repayment.jpg?v=2", "③資金償還表",
   "簡易キャッシュフローの作り方から債務償還年数、3年返済充足率まで。画面では結果だけをお見せしていますが、ここでは算定の過程が数字で追えます。"],
  ["./assets/4-scoring.jpg?v=2", "④配点内訳",
   "6軸それぞれの得点と、判定に用いた値。「なぜこの点数になったのか」を稟議の場で説明できます。"],
  ["./assets/5-input.jpg?v=2", "⑤入力データ",
   "判定に使った数値をそのまま記録。あとからの検証と、担当者が替わったときの引き継ぎに使えます。"],
  ["./assets/6-dashboard.jpg?v=2", "⑥ダッシュボード",
   "画面でご覧いただいた7つの図を、A3横1枚に。印刷してそのまま配れます。"],
];
let shotAt = 0;

/* ------------------------------------------------------------------------
 * 見本の自動切り替え
 *
 * タブを押せることに気づいてもらうため、①〜⑥を順番に自動で切り替える。
 * 選ばれているタブの下線が右まで伸びると、次のシートに移る。
 *
 *  ・枠が画面に半分以上入ってから動き出す。初めて見に来たときは①から始まる
 *  ・画面の外にあるとき、ブラウザの別タブを見ているときは止まる
 *  ・見本の上にマウスを乗せているあいだは止まる（離すと続きから）
 *  ・タブ／左右の矢印／「シート全体を見る」を押したら、自動切り替えはやめる
 *  ・SHOT_LOOPS 周したら①に戻って止まる。右上のボタンで、いつでも止める・再開できる
 *
 * 見本画像は縦の長さがまちまちで、そのまま切り替えるとページ全体が上下に跳ねる
 * （パソコンの幅で最大900pxほど）。そこで枠の縦横比を SHOT_RATIO に固定し、
 * 枠からはみ出す分は「シート全体を見る」で開けるようにしている。
 * SHOT_RATIO は、いちばん横長の⑥ダッシュボード（1400×1075）がちょうど収まる比率。
 * 見本画像を差し替えて縦横比が変わったら、ここを合わせる。
 *
 * 見た目（CSS）もこのファイルから差し込む。index.html を触らずに済むようにするため。
 * 自動切り替えの準備でつまずいても、タブの手動切り替えは今までどおり動く。
 * ---------------------------------------------------------------------- */
const SHOT_RATIO = [1400, 1075];   // 枠の［横, 縦］
const SHOT_LOOPS = 2;              // 何周したら止めるか
const SHOT_FADE_MS = 180;          // 切り替えのフェード（動きを減らす設定のときは使わない）
/** 1枚を見せる時間（ミリ秒）。説明文が長いシートほど長く、5〜9秒の範囲に収める */
const shotDwell = (i) => Math.min(9000, Math.max(5000, 2500 + SHOTS[i][2].length * 75));

const shotUI = {};                 // 画面の要素（initShots / initShotAuto で入れる）
const shotAuto = {
  on: false,                       // 自動切り替え中か
  holds: new Set(),                // 一時的に止めている理由（offscreen・hidden・hover・focus・loading・first）
  elapsed: 0,                      // いまのシートを見せ始めてからの経過ミリ秒
  loops: 0,                        // ⑥から①へ戻った回数
  raf: 0,
  last: 0,
};
const shotCache = new Map();       // 読み込んだ見本画像（src → Promise）
let shotSwap = 0;                  // 画像の差し替えの追い越しを防ぐための番号

const SHOT_ICON = {
  pause: '<svg viewBox="0 0 12 12" aria-hidden="true"><rect x="2.5" y="1.5" width="2.5" height="9" rx=".7"/><rect x="7" y="1.5" width="2.5" height="9" rx=".7"/></svg>',
  play: '<svg viewBox="0 0 12 12" aria-hidden="true"><path d="M3.5 1.8v8.4L10.3 6z"/></svg>',
  down: '<svg viewBox="0 0 10 10" aria-hidden="true"><path d="M1.6 3.4 5 6.8l3.4-3.4"/></svg>',
};

function initShots() {
  const tabs = document.querySelector(".shot__tabs");
  const img = $("shotImg");
  const cap = $("shotCap");
  if (!tabs || !img || !cap) return;
  shotUI.tabs = Array.from(tabs.querySelectorAll("[data-shot]"));
  shotUI.img = img;
  shotUI.cap = cap;

  // 手で選んだら、自動切り替えはやめる（見たいシートがある、という合図なので）
  const takeOver = () => { if (shotAuto.on) setShotAuto(false); };
  tabs.addEventListener("click", (e) => {
    const b = e.target.closest("[data-shot]");
    if (b) { takeOver(); showShot(+b.dataset.shot); }
  });
  $("shotPrev").addEventListener("click", () => { takeOver(); showShot(shotAt - 1); });
  $("shotNext").addEventListener("click", () => { takeOver(); showShot(shotAt + 1); });
  showShot(shotAt);

  // 自動切り替えは「あると親切」な機能。ここで例外を投げると init() の後ろ
  // （決済から戻ったときの復元など）まで止まってしまうので、自動切り替えだけを
  // あきらめて、理由はコンソールに残す。手動のタブはこの時点で動いている。
  try {
    initShotAuto(takeOver);
  } catch (err) {
    // 途中まで作った自動切り替えの部品は片づけて、手動のタブだけで動かす
    shotAuto.on = false;
    stopShotTick();
    if (shotUI.root) shotUI.root.classList.remove("is-auto");
    if (shotUI.play) shotUI.play.remove();
    console.warn("[財務でポン] 見本の自動切り替えを止めました（タブの手動切り替えは使えます）", err);
  }
}

function initShotAuto(takeOver) {
  const img = shotUI.img;
  const root = img.closest(".shot");
  const stage = img.closest(".shot__stage");
  const tabs = root && root.querySelector(".shot__tabs");
  if (!root || !stage || !tabs) return;

  injectShotStyle();

  // 画像を、縦横比を固定した枠に入れる（シートごとの高さの違いでページが跳ねないように）
  const frame = document.createElement("div");
  frame.className = "shot__frame";
  frame.id = "shotFrame";
  stage.insertBefore(frame, img);
  frame.appendChild(img);

  // 下が枠からはみ出すシートにだけ出す「シート全体を見る」
  const more = document.createElement("button");
  more.type = "button";
  more.className = "shot__more";
  more.setAttribute("aria-controls", "shotFrame");
  more.setAttribute("aria-expanded", "false");
  more.innerHTML = "<span>シート全体を見る</span>" + SHOT_ICON.down;
  stage.appendChild(more);

  // 止める／再開するボタン。見た目はタブの並びの右端、キーボードではタブの次に届く
  const play = document.createElement("button");
  play.type = "button";
  play.className = "shot__play";
  paintShotPlay(play, false);
  root.insertBefore(play, tabs.nextSibling);

  Object.assign(shotUI, { root, stage, more, play });

  img.addEventListener("load", () => markShotClip());
  if (!img.complete) {
    // 最初の1枚が届く前に、下線だけが進んでしまわないように
    holdShot("first");
    const done = () => releaseShot("first");
    img.addEventListener("load", done, { once: true });
    img.addEventListener("error", done, { once: true });
  }

  play.addEventListener("click", () => {
    if (shotAuto.on) { setShotAuto(false); return; }
    if (root.classList.contains("is-open")) setShotOpen(false, play);
    shotAuto.loops = 0;
    setShotAuto(true);
    loadShot((shotAt + 1) % SHOTS.length);
  });
  more.addEventListener("click", () => {
    takeOver();
    setShotOpen(!root.classList.contains("is-open"), more);
  });

  // 見本の上にマウスがあるあいだは止める。タッチ操作は対象外
  stage.addEventListener("pointerenter", (e) => { if (e.pointerType === "mouse") holdShot("hover"); });
  stage.addEventListener("pointerleave", (e) => { if (e.pointerType === "mouse") releaseShot("hover"); });

  // キーボードでタブや矢印に来たら止める。止める／再開するボタンに来たときは止めない
  const byKeyboard = (el) => { try { return el.matches(":focus-visible"); } catch (_) { return false; } };
  root.addEventListener("focusin", (e) => {
    if (e.target === play) releaseShot("focus");
    else if (byKeyboard(e.target)) holdShot("focus");
  });
  root.addEventListener("focusout", (e) => {
    if (!root.contains(e.relatedTarget)) releaseShot("focus");
  });

  // ブラウザの別タブを見ているあいだは止める
  if (document.hidden) holdShot("hidden");
  document.addEventListener("visibilitychange", () => {
    if (document.hidden) holdShot("hidden"); else releaseShot("hidden");
  });

  // 枠が画面に半分以上入ったら動かす
  if ("IntersectionObserver" in window) {
    holdShot("offscreen");
    new IntersectionObserver((entries) => {
      const en = entries[entries.length - 1];
      if (en.isIntersecting && en.intersectionRatio >= 0.49) {
        releaseShot("offscreen");
        if (shotAuto.on) loadShot((shotAt + 1) % SHOTS.length);
      } else {
        holdShot("offscreen");
      }
    }, { threshold: [0, 0.5] }).observe(stage);
  }

  shotUI.frame = frame;   // ここから先の切り替えは、フェードと先読みつきになる
  markShotClip();
  setShotAuto(true);
}

function injectShotStyle() {
  if (document.getElementById("shotAutoStyle")) return;
  const [w, h] = SHOT_RATIO;
  const style = document.createElement("style");
  style.id = "shotAutoStyle";
  style.textContent = `
.shot{position:relative;}
.shot__tabs{padding-right:56px;}
.shot__tabs button{position:relative;}
.shot.is-auto .shot__tabs button.is-on::before,
.shot.is-auto .shot__tabs button.is-on::after{content:"";position:absolute;left:12px;right:12px;bottom:4px;height:3px;border-radius:3px;}
.shot.is-auto .shot__tabs button.is-on::before{background:rgba(51,82,108,.16);}
.shot.is-auto .shot__tabs button.is-on::after{background:var(--sea-deep);transform:scaleX(var(--shot-p,0));transform-origin:0 50%;}
.shot__play{position:absolute;top:11px;right:12px;z-index:3;width:32px;height:32px;padding:0;display:grid;place-items:center;
  border:1.5px solid var(--line);border-radius:999px;background:var(--card);color:var(--ink);cursor:pointer;}
.shot__play:hover{border-color:var(--sea-deep);}
.shot__play svg{display:block;width:12px;height:12px;fill:currentColor;}
.shot__frame{position:relative;overflow:hidden;box-sizing:content-box;aspect-ratio:${w}/${h};
  border:1px solid var(--line);border-radius:8px;background:#fff;}
.shot.is-open .shot__frame{aspect-ratio:auto;}
.shot__stage .shot__frame img{border:0;border-radius:0;transition:opacity ${SHOT_FADE_MS}ms ease;}
.shot__frame.is-swapping img{opacity:.3;}
.shot__frame::after{content:"";position:absolute;left:0;right:0;bottom:0;height:88px;pointer-events:none;opacity:0;
  background:linear-gradient(rgba(255,255,255,0),rgba(255,255,255,.96) 78%);}
.shot.is-clipped .shot__frame::after{opacity:1;}
.shot.is-open .shot__frame::after{opacity:0;}
.shot__more{display:none;position:absolute;left:50%;bottom:26px;z-index:2;transform:translateX(-50%);
  align-items:center;gap:7px;white-space:nowrap;font-family:var(--sans);font-size:13px;font-weight:700;color:var(--ink);
  padding:8px 16px;border:1.5px solid var(--line);border-radius:999px;background:var(--card);box-shadow:var(--shadow);cursor:pointer;}
.shot.is-clipped .shot__more{display:inline-flex;}
.shot__more:hover{border-color:var(--sea-deep);}
.shot__more svg{display:block;width:10px;height:10px;fill:none;stroke:currentColor;stroke-width:1.8;stroke-linecap:round;stroke-linejoin:round;}
.shot.is-open .shot__more svg{transform:rotate(180deg);}
.shot.is-open.is-clipped .shot__stage{padding-bottom:64px;}
.shot.is-open .shot__more{bottom:16px;}
@media (max-width:820px){
  .shot__tabs{padding-right:48px;}
  .shot.is-auto .shot__tabs button.is-on::before,
  .shot.is-auto .shot__tabs button.is-on::after{left:10px;right:10px;bottom:3px;}
  .shot__play{top:9px;right:9px;width:30px;height:30px;}
  .shot__frame::after{height:56px;}
  .shot__more{bottom:22px;font-size:12px;padding:6px 12px;}
  .shot.is-open.is-clipped .shot__stage{padding-bottom:54px;}
  .shot.is-open .shot__more{bottom:12px;}
}`;
  document.head.appendChild(style);
}

/** i 番目のシートを見せる。手動・自動どちらの切り替えもここを通る */
function showShot(i) {
  shotAt = (i + SHOTS.length) % SHOTS.length;
  const [src, name] = SHOTS[shotAt];
  const { img, frame } = shotUI;
  const token = ++shotSwap;
  shotAuto.elapsed = 0;
  paintShotTabs();

  if (!frame) {
    // 自動切り替えの準備前（または準備に失敗したとき）は、今までどおり差し替えるだけ
    if (img.getAttribute("src") !== src) img.src = src;
    img.alt = name + "シートの見本";
    return;
  }
  if (img.getAttribute("src") === src) {
    frame.classList.remove("is-swapping");
    img.alt = name + "シートの見本";
    markShotClip();
    releaseShot("loading");
    return;
  }
  // 次の画像が届くまで古い画像を薄くして待つ。届くまで下線は進めない
  holdShot("loading");
  frame.classList.add("is-swapping");
  Promise.all([loadShot(shotAt), shotWait(shotReduced() ? 0 : SHOT_FADE_MS)]).then(([im]) => {
    if (token !== shotSwap) return;   // 待っているあいだに、別のシートが選ばれた
    img.src = src;
    img.alt = name + "シートの見本";
    frame.classList.remove("is-swapping");
    markShotClip(im.naturalWidth, im.naturalHeight);
    shotAuto.elapsed = 0;
    releaseShot("loading");
    if (shotAuto.on) loadShot((shotAt + 1) % SHOTS.length);   // 次の1枚を先に読んでおく
  });
}

function paintShotTabs() {
  shotUI.tabs.forEach((b, k) => {
    const on = k === shotAt;
    b.classList.toggle("is-on", on);
    b.setAttribute("aria-selected", String(on));
    b.style.setProperty("--shot-p", "0");
  });
  shotUI.cap.textContent = SHOTS[shotAt][2];
}

function setShotProgress(p) {
  const b = shotUI.tabs && shotUI.tabs[shotAt];
  if (b) b.style.setProperty("--shot-p", String(Math.round(p * 1000) / 1000));
}

/** 画像が枠より縦長（下が切れている）なら「シート全体を見る」を出す */
function markShotClip(w = shotUI.img.naturalWidth, h = shotUI.img.naturalHeight) {
  if (!shotUI.root || !(w > 0 && h > 0)) return;   // 読み込み前は測れない。load で測り直す
  shotUI.root.classList.toggle("is-clipped", h / w > SHOT_RATIO[1] / SHOT_RATIO[0] + 0.01);
}

/** 見本画像を読み込む。失敗しても解決する（切り替えが止まらないように） */
function loadShot(i) {
  const src = SHOTS[i][0];
  if (!shotCache.has(src)) {
    const im = new Image();
    im.decoding = "async";
    im.src = src;
    const ready = typeof im.decode === "function"
      ? im.decode()
      : new Promise((res, rej) => { im.onload = res; im.onerror = rej; });
    shotCache.set(src, ready.then(() => im, () => { shotCache.delete(src); return im; }));
  }
  return shotCache.get(src);
}

function shotWait(ms) {
  return new Promise((res) => setTimeout(res, ms));
}

function shotReduced() {
  return !!(window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches);
}

function setShotAuto(on) {
  if (!on) { shotAuto.on = false; stopShotTick(); }
  shotUI.root.classList.toggle("is-auto", on);
  paintShotPlay(shotUI.play, on);
  shotAuto.elapsed = 0;
  setShotProgress(0);
  if (on) { shotAuto.on = true; kickShot(); }
}

/** 止める／再開するボタンの見た目。on は「いま自動で切り替わっているか」 */
function paintShotPlay(play, on) {
  const label = on ? "シートの自動切り替えを止める" : "シートを自動で切り替える";
  play.innerHTML = on ? SHOT_ICON.pause : SHOT_ICON.play;
  play.setAttribute("aria-label", label);
  play.title = label;
}

/** 「シート全体を見る」の開け閉め。anchor は押されたボタン（閉じたときに指の下から逃がさない） */
function setShotOpen(open, anchor) {
  const { root, more } = shotUI;
  const pin = anchor && anchor.getClientRects().length ? anchor : null;
  const before = pin ? pin.getBoundingClientRect().top : 0;
  root.classList.toggle("is-open", open);
  more.setAttribute("aria-expanded", String(open));
  more.querySelector("span").textContent = open ? "たたむ" : "シート全体を見る";
  if (pin && !open) {
    const after = pin.getBoundingClientRect().top;
    if (Math.abs(after - before) > 1) window.scrollBy(0, after - before);
  }
}

function holdShot(why) {
  shotAuto.holds.add(why);
  stopShotTick();
}

function releaseShot(why) {
  if (shotAuto.holds.delete(why)) kickShot();
}

function stopShotTick() {
  if (shotAuto.raf) cancelAnimationFrame(shotAuto.raf);
  shotAuto.raf = 0;
  shotAuto.last = 0;
}

function kickShot() {
  if (!shotAuto.on || shotAuto.holds.size || shotAuto.raf) return;
  shotAuto.last = 0;
  shotAuto.raf = requestAnimationFrame(tickShot);
}

function tickShot(now) {
  shotAuto.raf = 0;
  if (!shotAuto.on || shotAuto.holds.size) return;
  // 処理が重くてコマの間が空いても、一気に進まないよう1コマ100ミリ秒までに抑える
  if (shotAuto.last) shotAuto.elapsed += Math.max(0, Math.min(now - shotAuto.last, 100));
  shotAuto.last = now;
  const dwell = shotDwell(shotAt);
  setShotProgress(Math.min(shotAuto.elapsed / dwell, 1));
  if (shotAuto.elapsed >= dwell) nextShotAuto();
  if (shotAuto.on && !shotAuto.holds.size && !shotAuto.raf) {
    shotAuto.raf = requestAnimationFrame(tickShot);
  }
}

function nextShotAuto() {
  const next = (shotAt + 1) % SHOTS.length;
  if (next === 0 && ++shotAuto.loops >= SHOT_LOOPS) {
    setShotAuto(false);   // 決めた周回を終えたら、①に戻して止める
    showShot(0);
    return;
  }
  showShot(next);
}


/* ============================================================ 比較デモ
 * 「評価の高い会社」「評価の低い会社」を並べて、出力される図を先に見てもらう。
 * 判定エンジンは本番とまったく同じものを使う。
 * ========================================================================= */
const DEMO_GOOD = () => Object.assign(emptyInput(), {
  name: "サンプル情報システム株式会社（評価の高い例）", industry: "　情報通信業",
  capitalTier: "1億円以上10億円未満", listing: "未上場",
  founded: "1998-04-01", baseDate: new Date().toISOString().slice(0, 10),
  employees: 320, capital: 150,
  terms: ["第28期（直近）", "第27期", "第26期"],
  sales: [10200, 9600, 9100], cogs: [7140, 6816, 6552], sga: [2150, 2080, 2030],
  nonOpInc: [20, 18, 16], nonOpExp: [30, 34, 38], extraInc: [0, 0, 0], extraExp: [0, 0, 30],
  tax: [290, 220, 150], depreciation: [260, 250, 240],
  cash: [2600, 2200, 1900], receivables: [1750, 1650, 1560], inventory: [320, 300, 290],
  otherCurrentAssets: [230, 210, 200], tangible: [1900, 1880, 1860],
  otherFixedAssets: [900, 850, 800], deferred: [0, 0, 0],
  payables: [780, 750, 720], shortDebt: [250, 280, 300], otherCurrentLiab: [850, 800, 780],
  longDebt: [700, 900, 1100], otherFixedLiab: [120, 120, 120], equity: [5000, 4240, 3590],
  ceoName: "見本　太郎", ceoAge: 54, industryYears: 26, ceoYears: 15,
  ownHome: "あり", disclosure: "あり", successor: "あり",
  repayYears: 5, repayManual: false,
});
const DEMO_BAD = () => Object.assign(emptyInput(), {
  name: "サンプル商事株式会社（評価の低い例）", industry: "　卸売業、小売業",
  capitalTier: "1,000万円以上1億円未満", listing: "未上場",
  founded: "2022-04-01", baseDate: new Date().toISOString().slice(0, 10),
  employees: 25, capital: 30,
  terms: ["第4期（直近）", "第3期", "第2期"],
  sales: [850, 920, 1000], cogs: [700, 745, 790], sga: [190, 190, 190],
  nonOpInc: [2, 2, 2], nonOpExp: [12, 10, 9], extraInc: [0, 0, 0], extraExp: [0, 0, 0],
  tax: [0, 0, 4], depreciation: [12, 12, 12],
  cash: [40, 70, 110], receivables: [180, 195, 210], inventory: [150, 160, 165],
  otherCurrentAssets: [20, 20, 20], tangible: [90, 100, 110],
  otherFixedAssets: [20, 20, 20], deferred: [0, 0, 0],
  payables: [120, 125, 130], shortDebt: [160, 150, 130], otherCurrentLiab: [50, 50, 50],
  longDebt: [130, 150, 170], otherFixedLiab: [10, 10, 10], equity: [30, 80, 105],
  ceoName: "見本　次郎", ceoAge: 41, industryYears: 6, ceoYears: 4,
  ownHome: "なし", disclosure: "なし", successor: "なし",
  repayYears: 5, repayManual: false,
});

let demoKind = "good";
function renderDemo(kind) {
  demoKind = kind;
  const col = $("demoCol"); if (!col) return;
  const d = kind === "bad" ? DEMO_BAD() : DEMO_GOOD();
  const yrs = Math.max(1, Number(d.repayYears) || 5);
  const v = Math.round((Number(d.longDebt[0]) || 0) / yrs);
  d.repayment = [v, v, v];
  const r = evaluate(d);
  // デモは百万円で固定して表示する（読み込んだ決算書の単位に引きずられないように）
  const f = { yenU, U_LABEL, pct };
  col.innerHTML = renderHead(r, f, POLICY[r.scores.rank]) + renderViz(r, f);
  attachTips(col);
  document.querySelectorAll("[data-demo]").forEach((b) => {
    const on = b.dataset.demo === kind;
    b.classList.toggle("is-on", on);
    b.setAttribute("aria-selected", String(on));
  });
}
/* ------------------------------------------------------- 金額単位の換算 */
// 本シートは百万円で計算する。決算書は千円単位が多く、換算しないと
// 規模の配点と与信限度額が1000倍ずれる。桁が大きいだけで数字は自然に見えるため気づけない。
const MONEY_KEYS = ["sales","cogs","sga","nonOpInc","nonOpExp","extraInc","extraExp","tax",
  "depreciation","cash","receivables","inventory","otherCurrentAssets","tangible",
  "otherFixedAssets","deferred","payables","shortDebt","otherCurrentLiab","longDebt",
  "otherFixedLiab","equity","currentAssets","fixedAssets","currentLiab","fixedLiab",
  "totalCapital","ordinary","operating","net"];

/** 読み取った1期分を百万円に揃える。単位が読めなければ触らない */
function scaleToMillion(values, unit) {
  if (!unit || unit.toMillion === 1) return values;
  const out = { ...values };
  // ここで四捨五入しない。
  // 千円単位の 4,767,955千円 を 4,768百万円 に丸めると、千円で表示し直したときに
  // 4,768,000 となり、決算書と下3桁が合わなくなる。
  // 4767.955 のまま持てば、百万円でも千円でも決算書どおりの数字を出せる。
  for (const k of MONEY_KEYS)
    if (typeof out[k] === "number") out[k] = out[k] * unit.toMillion;

  // 各項目を個別に四捨五入すると、内訳の合計が小計と1単位ずれることがある。
  // 画面の表は内訳を足して小計を出すため、そのままだと決算書では合っていた貸借が
  // 合わなくなり、赤い警告が出てしまう。端数は「その他」の欄に寄せて辻褄を合わせる。
  reconcile(out);
  return out;
}

/**
 * 換算で生じた端数を「その他◯◯」に吸収させ、内訳の合計＝小計＝貸借一致に揃える。
 * 元の決算書で合っていたものを、換算のせいで狂わせないための処理。
 */
function reconcile(v) {
  const n = (x) => (typeof x === "number" ? x : 0);
  const fit = (parts, sub, slack) => {
    // 小計が取れていなければ、内訳の合計をそのまま小計とする
    if (typeof v[sub] !== "number") { v[sub] = parts.reduce((a, k) => a + n(v[k]), 0); return; }
    const gap = v[sub] - parts.reduce((a, k) => a + n(v[k]), 0);
    // 丸めをやめたことで、値が小数を持つようになった。
    // 浮動小数点の計算誤差（1e-10 など）を「ズレ」と誤認しないよう、
    // 0.0005百万円（＝500円）未満は一致とみなす。
    if (Math.abs(gap) < 5e-4 || Math.abs(gap) > 3) return;   // 大きなズレは読み取り誤りなので触らない
    v[slack] = n(v[slack]) + gap;                  // 端数は「その他」で調整する
  };
  fit(["cash", "receivables", "inventory", "otherCurrentAssets"], "currentAssets", "otherCurrentAssets");
  fit(["tangible", "otherFixedAssets"], "fixedAssets", "otherFixedAssets");
  fit(["payables", "shortDebt", "otherCurrentLiab"], "currentLiab", "otherCurrentLiab");
  fit(["longDebt", "otherFixedLiab"], "fixedLiab", "otherFixedLiab");

  // 最後に資産側と負債・純資産側を突き合わせ、残った端数もその他流動資産に寄せる
  const assets = n(v.currentAssets) + n(v.fixedAssets) + n(v.deferred);
  const liabEq = n(v.currentLiab) + n(v.fixedLiab) + n(v.equity);
  const gap = assets - liabEq;
  if (Math.abs(gap) >= 5e-4 && Math.abs(gap) <= 3) {
    v.otherCurrentAssets = n(v.otherCurrentAssets) - gap;
    v.currentAssets = n(v.currentAssets) - gap;
  }
}

/* ----------------------------------------------------------- 画面の段 */
function showStep(id, on) { const el = $(id); if (el) el.hidden = !on; }

/** 手入力欄を開き、判定と決済の段も見えるようにする */
function openManual() {
  const body = $("manualBody");
  if (body && body.hidden) {
    body.hidden = false;
    $("btnManual").setAttribute("aria-expanded", "true");
    $("btnManual").textContent = "入力欄を閉じる";
  }
}

/* -------------------------------------------------- 決算期の並べ替え */
/** 有報の「前期」列を、決算期の鍵として1年ぶん戻す */
function shiftBack(per) {
  const out = { end: null, no: null, label: "" };
  if (per.end) {
    const d = per.end.split("-");
    out.end = `${+d[0] - 1}-${d[1]}-${d[2]}`;
    out.label = `${+d[0] - 1}年${+d[1]}月期`;
  }
  if (typeof per.no === "number") {
    out.no = per.no - 1;
    out.label = out.end ? `第${out.no}期（${out.label}）` : `第${out.no}期`;
  }
  return out;
}

/**
 * 読み取れた期を新しい順（今期→前期→前々期）に並べ替える。
 * 決算日が取れたものを優先し、無ければ期数で比べる。
 * どちらも読めない期が混ざっている場合は、取り違えを避けるため並べ替えを行わず投入順のままにする。
 */
function sortPeriods(periods) {
  const live = periods.filter((p) => !p.failed);
  if (live.length < 2) return;
  const keyed = live.every((p) => p.period && (p.period.end || typeof p.period.no === "number"));
  if (!keyed) return;
  const val = (p) => (p.period.end ? p.period.end : "") ;
  const sorted = live.slice().sort((a, b) => {
    const ea = val(a), eb = val(b);
    if (ea && eb && ea !== eb) return eb.localeCompare(ea);
    const na = a.period.no ?? -Infinity, nb = b.period.no ?? -Infinity;
    if (na !== nb) return nb - na;
    return 0;
  });
  // 元配列の live 部分だけを並べ替え後の順に置き換える
  let i = 0;
  for (let k = 0; k < periods.length; k++)
    if (!periods[k].failed) periods[k] = sorted[i++];
}

/* ------------------------------------------------------------ 判定と手入力 */
// 抽出の成否を左右する中核項目。これが欠けたら「要確認」とする
const CORE_KEYS = ["sales", "cash", "currentAssets", "fixedAssets", "currentLiab", "fixedLiab", "equity"];

/**
 * 1期分の読み取り結果を3段階で評価する。
 *   "ok"    … 貸借が合い、中核項目も減価償却費も揃っている（そのまま使える）
 *   "supp"  … 抽出は成功。ただし決算書に単独で載っていない項目（減価償却費など）だけ手入力が要る
 *   "check" … 貸借不一致、経常利益の不整合、または中核項目が取れていない（数値の確認・入力が要る）
 */
function gradePeriod(values, diff, messages) {
  const coreMiss = CORE_KEYS.some((k) => values[k] === null || values[k] === undefined);
  // 検算メッセージは種類が増えた（営業利益・当期純利益の積み上がり、
  // 有利子負債と区分合計の整合）。特定の文言だけを見ていると、
  // 新しく検出した不整合を「そのまま使えます」と表示してしまう。
  const anyNg = messages.length > 0;
  if (diff !== 0 || anyNg || coreMiss) return "check";
  if (values.depreciation === null || values.depreciation === undefined) return "supp";
  return "ok";
}

// 手入力欄に出す候補。
// 出す基準は「未取得だと判定結果が狂う項目」に限る。
// 決算書に載っていない/合計に吸収済みの内訳（売上原価の区分がない様式など）まで
// 欄にすると、「そのまま使えます」と言いながら大量の入力を求める矛盾した画面になる。
//   depreciation … ⑥償還余力（30点）の算定に必須
//   CORE_KEYS     … 貸借・規模・資本構成の土台。欠けると判定そのものが立たない
//   shortDebt / longDebt … 有利子負債。0で確定できないため、欠けたら必ず確認を求める
const MANUAL_LABELS = {
  depreciation: ["減価償却費", "販管費明細・製造原価報告書の合計。決算書に単独の行が無いことが多い項目です。"],
  sales: ["売上高", ""], cash: ["現金・預金", ""],
  currentAssets: ["流動資産合計", ""], fixedAssets: ["固定資産合計", ""],
  currentLiab: ["流動負債合計", ""], fixedLiab: ["固定負債合計", ""],
  equity: ["純資産合計", ""],
  shortDebt: ["短期借入金（1年内返済分を含む）", ""], longDebt: ["長期借入金・社債", ""],
};
// 表示順（重要な順）
const MANUAL_ORDER = ["depreciation", "sales", "cash", "currentAssets", "fixedAssets",
                      "currentLiab", "fixedLiab", "equity", "shortDebt", "longDebt"];

/** その期で実際に入力を求めるべきキーを返す */
function manualKeysFor(p) {
  const nil = (k) => p.values[k] === null || p.values[k] === undefined;
  return MANUAL_ORDER.filter((k) => {
    if (k === "shortDebt" || k === "longDebt") {
      // buildPeriod は見つからなければ0を入れ、警告を出す。
      // 「無借金なのか読み落としなのか」を利用者に確かめてもらう
      return p.warnings.some((w) => w.includes("有利子負債"));
    }
    return nil(k);
  });
}

const ROWS_SHOW = [
  ["売上高", "sales"], ["売上原価", "cogs"], ["販売費及び一般管理費", "sga"],
  ["営業外収益", "nonOpInc"], ["営業外費用", "nonOpExp"], ["法人税等", "tax"],
  ["減価償却費", "depreciation"], ["現金・預金", "cash"], ["受取手形・売掛金", "receivables"],
  ["棚卸資産", "inventory"], ["流動資産合計", "currentAssets"], ["固定資産合計", "fixedAssets"],
  ["支払手形・買掛金", "payables"], ["短期借入金", "shortDebt"], ["流動負債合計", "currentLiab"],
  ["長期借入金・社債", "longDebt"], ["固定負債合計", "fixedLiab"], ["純資産合計", "equity"],
];

function bannerFor(worst, hasFields) {
  if (worst === "ok")
    return hasFields
      ? ["good", "読み取れました", "貸借は一致しています。下の確認欄だけ目を通してから「反映する」を押してください。"]
      : ["good", "✓ そのまま使えます", "読み取った数値をご確認のうえ、下の「反映する」を押してください。"];
  if (worst === "supp")
    return ["supp", "あと少しで完成します", "決算書に単独で載っていない項目だけ、下の欄にご入力ください。入力すると判定とExcelが完成します。"];
  return ["check", "数値の確認・入力をお願いします", "読み取れなかった項目を下の欄に入力すると完成します。貸借がずれている場合は、各項目の値もあわせてご確認ください。"];
}

let lastRead = null;   // 単位を切り替えたとき、読み取り結果の表も作り直すために覚えておく

function showRead(periods) {
  lastRead = periods;
  const live = periods.filter((p) => !p.failed);
  const setFix = (html) => {
    const el = $("readFix");
    if (!el) return;
    el.innerHTML = html;
    el.hidden = !html;
  };

  showStep("step3", true);
  if (!live.length) {
    // 全滅：画像PDFかどうかで文言を変え、手入力へ誘導する
    const anyImage = periods.some((p) => p.image);
    const scope = periods.find((p) => p.scope);
    $("readTable").innerHTML = "";
    const SCOPE_MSG = {
      bank: ["銀行の決算書のようです。このツールの対象外です",
        "銀行の貸借対照表には流動・固定の区分が無く、損益計算書も売上高ではなく経常収益で構成されるため、本ツールの判定モデルには載りません。判定に用いる業界基準の統計も金融業・保険業を対象外としています。一般事業会社の決算書でお試しください。"],
      insurance: ["保険会社の決算書のようです。このツールの対象外です",
        "保険会社の貸借対照表には流動・固定の区分が無く、損益計算書も売上高ではなく経常収益で構成されるため、本ツールの判定モデルには載りません。判定に用いる業界基準の統計も金融業・保険業を対象外としています。一般事業会社の決算書でお試しください。"],
      securities: ["証券会社（金融商品取引業）の決算書のようです。このツールの対象外です",
        "証券会社の損益計算書は営業収益と受入手数料で構成され、貸借対照表にも流動・固定の区分がありません。本ツールの判定モデルには載らず、業界基準の統計も金融業を対象外としています。一般事業会社の決算書でお試しください。"],
      interim: ["四半期・中間の決算書のようです。このツールの対象外です",
        "四半期の損益計算書に載っている売上高や利益は3か月ぶんの金額です。これを年間の実績として扱うと、総資産回転率も債務償還年数も与信限度額も実態からずれた数字になります。しかも一見それらしい数字が出るため、誤りに気づけません。通期（1年分）の決算書をご用意ください。"],
    };
    const msg = scope ? SCOPE_MSG[scope.scope] : null;
    $("readBanner").innerHTML = msg
      ? `<div class="read-banner is-ng"><b>${msg[0]}</b><span>${msg[1]}</span></div>`
      : `<div class="read-banner is-ng"><b>${anyImage ? "この決算書からは文字を取り出せませんでした" : "この様式は読み取れませんでした"}</b>` +
        `<span>下の入力欄に直接ご入力いただければ、判定もExcelの作成も問題なく行えます。</span></div>`;
    setFix("");
    $("readWarn").innerHTML = anyImage
      ? "<li>スキャンされた画像PDFのため、文字を読み取れませんでした（このツールは画像の文字起こし＝OCRは行いません）。お手数ですが、下の入力欄に直接ご入力ください。判定とExcelの作成は問題なく行えます。</li>"
      : "<li>この決算書は自動読み取りに対応していませんでした。お手数ですが、下の入力欄に直接ご入力ください。判定とExcelの作成は問題なく行えます。</li>";
    if (scope && scope.scope === "interim") {
      // 通期の決算書でない以上、手入力に誘導しても判定は成り立たない
      $("readWarn").innerHTML = "<li>四半期・中間の決算書では判定を行いません。通期（1年分）の決算書をご用意ください。</li>";
      return;
    }
    $("btnApply").hidden = false;   // 手入力だけで判定できるようにする
    openManual();
    return;
  }
  $("btnApply").hidden = false;
  pending = live;
  if (metaCompany) { state.name = metaCompany; paint(); }
  warnCompany();

  // ---- 判定バナー（最も注意の要る期に合わせる） ----
  const rank = { ok: 0, supp: 1, check: 2 };
  const worst = live.reduce((w, p) => (rank[p.grade] > rank[w] ? p.grade : w), "ok");
  // 入力欄が1つでも出るかを先に調べ、バナーの文言と矛盾しないようにする
  const anyFields = live.some((p) => manualKeysFor(p).length || p.diff !== 0);
  const [cls, title, lead] = bannerFor(worst, anyFields);
  $("readBanner").innerHTML =
    `<div class="read-banner is-${cls}"><b>${title}</b><span>${lead}</span></div>`;

  // ---- 読み取り結果の表 ----
  let h = `<thead><tr><th>科　目<span class="unit-tag">単位：${U_LABEL()}</span></th>`;
  live.forEach((p, i) => {
    const lab = p.period && p.period.label ? `<br><span class="th-sub">${esc(p.period.label)}</span>` : "";
    // 元の決算書が何円単位だったかも出す。換算したことを隠さない
    const u = p.unit ? `<br><span class="th-sub">原本：${esc(p.unit.label)}</span>` : "";
    h += `<th>${PERIOD_LABEL[i]}${lab}${u}</th>`;
  });
  h += "</tr></thead><tbody>";
  h += "<tr><th>読み取り元</th>" + live.map((p) =>
    `<td style="font-size:12.5px">${esc(p.file)}<br>${esc(p.pages || "")}</td>`).join("") + "</tr>";
  for (const [label, key] of ROWS_SHOW) {
    h += `<tr><th>${label}</th>` + live.map((p) => {
      const v = p.values[key];
      return v === null || v === undefined
        ? '<td class="miss">未取得</td>'
        : `<td class="ok">${yenU(v)}</td>`;
    }).join("") + "</tr>";
  }
  h += "<tr><th>貸借の検算</th>" + live.map((p) =>
    `<td class="${isBalanced(p.diff) ? "ok" : "miss"}">${isBalanced(p.diff) ? "一致" : yenU(p.diff) + " のズレ"}</td>`).join("") + "</tr>";
  $("readTable").innerHTML = h + "</tbody>";

  // ---- ここだけ入力してください（未取得のエンジン項目だけを欄にする） ----
  let fix = "";
  let depHint = false;
  live.forEach((p, i) => {
    const miss = manualKeysFor(p);
    const unbalanced = p.diff !== 0;
    if (!miss.length && !unbalanced) return;
    if (miss.includes("depreciation")) depHint = true;
    fix += live.length > 1 ? `<p class="fix-period">${PERIOD_LABEL[i]}</p>` : "";
    if (unbalanced)
      fix += `<p class="fix-warn">資産合計と負債・純資産合計が <b>${yenU(p.diff)}</b> ずれています。下の項目、または反映後の入力欄で各数値をご確認ください。</p>`;
    if (miss.length) {
      fix += '<div class="pro-2 fix-grid">';
      for (const k of miss) {
        const label = MANUAL_LABELS[k][0];
        fix += `<label class="pro-field"><span>${label}</span>` +
               `<input type="number" step="1" inputmode="numeric" placeholder="未入力" ` +
               `data-fixkey="${k}" data-fixidx="${i}"></label>`;
      }
      fix += "</div>";
    }
  });
  if (fix && depHint)
    fix += '<p class="note-s" style="margin-top:8px;">減価償却費は、損益計算書に単独の行が無い決算書が普通です。' +
           'キャッシュ・フロー計算書、製造原価報告書、販売費及び一般管理費の明細に載っています。無ければ0のままでも作成できますが、償還余力（30点）の判定には必要です。</p>';
  const fixHead = worst === "ok" ? "ここだけご確認ください" : "ここだけ入力してください";
  setFix(fix ? `<h3 class="fix-head">${fixHead}</h3>${fix}` : "");
  // 読み取った時点で入力欄へ流し込む。判定ボタンを押すまで0が並ぶのは分かりにくいため。
  // 不足項目の欄を作り直した後に呼ぶ（前回の入力値を拾わないようにする）。
  applyValues();

  // ---- 補足メッセージ ----
  const msgs = [];
  live.forEach((p, i) => {
    [...new Set([...p.warnings, ...p.messages])].forEach((w) =>
      msgs.push(`${PERIOD_LABEL[i]}：${w}`));
  });
  // 決算期が連続していない場合に知らせる。
  // 例：2026年3月期と2024年3月期だけが読めたとき、間の2025年3月期が抜けたまま
  // 「前期」の欄に前々期の数字が入る。黙って詰めると気づけないため必ず伝える。
  {
    const dated = live.filter((p) => p.period && p.period.end);
    if (dated.length >= 2) {
      const gaps = [];
      for (let i = 0; i < dated.length - 1; i++) {
        const y0 = +dated[i].period.end.slice(0, 4), y1 = +dated[i + 1].period.end.slice(0, 4);
        if (y0 - y1 > 1) gaps.push(`${dated[i + 1].period.label}と${dated[i].period.label}`);
      }
      if (gaps.length)
        msgs.push(`決算期が連続していません（${gaps.join("、")}のあいだが抜けています）。` +
          `間の期のPDFを追加で読み込むか、順番が意図どおりかご確認ください。`);
    }
  }
  if (periods._dropped)
    msgs.push(`4期以上を読み取ったため、新しい3期分だけを残しました（${periods._dropped}期分を除きました）。`);
  if (live.length < 3) {
    const dated = live.filter((p) => p.period && p.period.label).map((p) => p.period.label);
    const got = dated.length === live.length ? `（${dated.join("・")}）` : "";
    msgs.push(`${live.length}期分を読み取りました${got}。3期分そろうと、損益の推移と償還余力まで判定できます。` +
      `別の期の決算書PDFを、この画面にもう一度ドロップしてください。1件ずつでも、まとめてでもかまいません。` +
      `決算期は自動で読み取り、新しい順に並べ替えます。`);
  }
  $("readWarn").innerHTML = msgs.length
    ? msgs.map((m) => `<li>${esc(m)}</li>`).join("")
    : "<li>特に注意すべき点は検出されませんでした。念のため数値をご確認ください。</li>";
}

/**
 * 読み取った数値を入力欄へ反映する。
 * 判定ボタンを押す前に呼ぶので、赤い枠や3期分の数値がその場で見える。
 * 会社名・代表者情報・返済計画など、利用者が自分で入れた項目は消さない。
 */
function applyValues() {
  if (!pending) return;
  // 前回の読み取り値が残っていると、今回読めなかった項目に古い数字が居座る。
  // 3期分すべてを一度0に戻してから入れ直す。
  const blank = emptyInput();
  const KEEP = ["repayment"];   // 返済計画は決算書に載っておらず、利用者が入れたもの
  for (const k of Object.keys(blank))
    if (Array.isArray(blank[k]) && blank[k].length === 3 && !KEEP.includes(k)) state[k] = [0, 0, 0];
  state.terms = ["", "", ""];
  missingCells = {};   // 赤く出す対象を作り直す

  pending.forEach((p, i) => {
    const v = toEngineFields(p.values);
    // 読み取れなかった項目を覚えておき、入力欄を赤くする
    for (const [k, val] of Object.entries(v))
      if (val === null || val === undefined) (missingCells[k] ||= []).push(i);
    for (const [k, val] of Object.entries(v)) {
      if (!Array.isArray(state[k])) state[k] = [0, 0, 0];
      state[k][i] = val ?? 0;
    }
    // 決算期が読み取れていれば、決算期欄にも入れる（判定には使わないが、Excelの見出しになる）
    if (p.period && p.period.label) {
      if (!Array.isArray(state.terms)) state.terms = ["", "", ""];
      state.terms[i] = p.period.label;
    }
  });
  // 会社名。PDFから読めたものを優先して入れる。
  // 別の会社の決算書を続けて読ませたときに、前の会社名が残ったままにならないようにする。
  if (metaCompany) state.name = metaCompany;
  warnCompany();
  paint(); render();
}

/**
 * 「この内容で判定する」を押したとき。
 * 値の流し込みは読み取り直後に済ませてあるので、ここでは上書きしない。
 * 上書きすると、利用者が表で直した数値が読み取り値に戻ってしまう。
 * ここで取り込むのは「ここだけ入力してください」の欄だけ。
 */
function applyRead() {
  document.querySelectorAll("#readFix input[data-fixkey]").forEach((inp) => {
    if (inp.value === "") return;
    const num = parseFloat(inp.value);
    if (Number.isNaN(num)) return;
    const k = inp.dataset.fixkey, i = +inp.dataset.fixidx;
    if (!Array.isArray(state[k])) state[k] = [0, 0, 0];
    state[k][i] = num;
    // 埋まったので赤い表示を解除する
    if (missingCells[k]) missingCells[k] = missingCells[k].filter((x) => x !== i);
  });
  paint(); render();
  if (window.gtag && pending) gtag("event", "pdf_applied", { tool: "credit-pro", periods: pending.length });
  accepted = pending || accepted;   // 追加のPDFを置いたときに積み上げられるよう残す
  $("readState").hidden = true;
  showStep("step4", true);
  showStep("step5", true);
  refreshLicense();
  $("step4").scrollIntoView({ behavior: "smooth", block: "start" });
}

/* ------------------------------------------------------------------ 記入例 */
function demo() {
  const d = emptyInput();
  Object.assign(d, {
    name: "サンプル情報システム株式会社（記入例）", industry: "　情報通信業",
    capitalTier: "1億円以上10億円未満", listing: "未上場",
    founded: "1998-04-01", baseDate: new Date().toISOString().slice(0, 10),
    employees: 320, capital: 150,
    terms: ["2026年3月期", "2025年3月期", "2024年3月期"],
    sales: [10200, 9600, 9100], cogs: [7140, 6816, 6552], sga: [2150, 2080, 2030],
    nonOpInc: [20, 18, 16], nonOpExp: [30, 34, 38],
    extraInc: [0, 0, 0], extraExp: [0, 0, 30], tax: [290, 220, 150],
    depreciation: [260, 250, 240],
    cash: [2600, 2200, 1900], receivables: [1750, 1650, 1560],
    inventory: [320, 300, 290], otherCurrentAssets: [230, 210, 200],
    tangible: [1900, 1880, 1860], otherFixedAssets: [900, 850, 800], deferred: [0, 0, 0],
    payables: [780, 750, 720], shortDebt: [250, 280, 300], otherCurrentLiab: [850, 800, 780],
    longDebt: [700, 900, 1100], otherFixedLiab: [120, 120, 120], equity: [5000, 4240, 3590],
    ceoName: "見本　太郎", ceoAge: 54, industryYears: 26, ceoYears: 15,
    ownHome: "あり", disclosure: "あり", successor: "あり",
    repayment: [250, 230, 220], repayManual: true,
    memo: "主力は金融機関向け業務システムの受託開発。上位5社で売上の約6割を占めるが、いずれも長期契約で取引関係は安定。3期連続の増収増益で、実質無借金。",
  });
  return d;
}

init();
