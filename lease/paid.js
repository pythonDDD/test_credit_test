/* ============================================================================
 * /lease/paid.js — 有料版（グラフ＋Excel）の案内と、購入後の画面
 *
 * 購入前：この見積で「まだ分かっていないこと」を5つ見せ、見本のグラフはぼかして置く
 * 購入後：購入した見積（物件価額・期間・月額）の数字で、動かせるグラフとExcelを出す
 *
 * 1回の購入は見積1件。購入した瞬間の数字を控え（スナップショット）、
 * そのあと入力を変えても、グラフとExcelは控えた見積から作る。
 * ========================================================================== */
import { analyze, defaultLife, LIFE_MIN, LIFE_MAX } from "./engine.js?v=3";
import { donutMarkup, lineCompare, areaRemaining, barsExpense, attachTips, toPng, METHODS } from "./viz.js?v=3";
import { payUrl, payUrlReady, quoteFingerprint, verifyOrder, readReturnOrder, cleanReturnUrl } from "./license.js?v=2";
import { downloadLeaseXlsx } from "./xlsx-export.js?v=4";

/* テスト環境（test_credit_test）では true にする。本番は必ず false。
   true でも kazumono.com の上では無料にならない（取り違えて上げたときの歯止め） */
export const FREE_BUILD = true;
const FREE_MODE = FREE_BUILD && !/(^|\.)kazumono\.com$/i.test(location.hostname);

const $ = (id) => document.getElementById(id);
const yen = (n) => Math.round(n).toLocaleString("ja-JP");
const SNAP_KEY = "kazumono.lease.paid";        // 開いている見積の控え（注文番号と期限つき）
const PENDING_KEY = "kazumono.lease.pending";  // 購入ボタンを押した時点の見積。Stripe から戻ったときに使う
const ORDER_KEY = "kazumono.lease.order";      // 戻ってきたが、まだ見積と結び付けられていない注文番号
const TEST_HOST = !/(^|\.)kazumono\.com$/i.test(location.hostname);   // 本番以外（テスト環境）
const CONTACT = "stats.okinawa@gmail.com";
/* 見本のグラフとシート見本に使う見積。「例を入れる」と同じ数字にすると、例を試した人に
   その見積の答えを無料で見せてしまうので、わざと別の数字にしている */
const SAMPLE = { price: 3000000, months: 60, monthly: 57000, residual: 0, life: 5 };

const P = { quote: null, life: null, lifeTouched: false, snap: null, over: {}, cancelAt: 12,
  show: { A: true, D: true, C: true, B: true } };

/* ============================================================ Excelのシート見本（自動で流れる）
 * 財務でポン！で検証済みの仕組みを、そのまま使っている */
const SHOTS = [
  ["./assets/lease-1-dashboard.jpg?v=2", "①ダッシュボード",
   "月額・金利・リース会社の利益と、4つのグラフを1枚に。A4縦1枚で印刷して、そのまま稟議に添付できます。"],
  ["./assets/lease-2-cost.jpg?v=2", "②原価内訳",
   "物件代金・保険料・税金・資金の金利、そしてリース会社の利益。月額の中身を上から順に分けています。"],
  ["./assets/lease-3-schedule.jpg?v=2", "③支払予定表",
   "毎回の支払を元本と利息に分け、その時点で残っている支払も並べました。途中で解約するときの目安になります。"],
  ["./assets/lease-4-tax.jpg?v=2", "④償却資産税",
   "リース会社が毎年納める税金を年度ごとに。買った場合に自社で払う税金の目安にもなります。"],
  ["./assets/lease-5-compare.jpg?v=2", "⑤現金・借入との比較",
   "リース・現金・銀行借入を、税金の効果まで含めた実質負担で並べます。いちばん負担の小さい買い方が一目で分かります。"],
  ["./assets/lease-6-kappu.jpg?v=2", "⑥割賦との比較",
   "分割払いで買う割賦と、リースを比べます。お金の差に加えて、持ち主や途中でやめるときの違いも。"],
  ["./assets/lease-7-depreciation.jpg?v=2", "⑦減価償却費",
   "買った場合に、毎年いくら経費になるか。帳簿価額の移り変わりと合わせて年度別に。"],
];
let shotAt = 0;

/* ------------------------------------------------------------------------
 * 見本の自動切り替え
 *
 * タブを押せることに気づいてもらうため、①〜⑦を順番に自動で切り替える。
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
 * SHOT_RATIO は、見本画像の横1400pxに対して縦1050px（短いシートは余白を足してこの比率にしてある）。
 * 見本画像を差し替えて縦横比が変わったら、ここを合わせる。
 *
 * 見た目（CSS）もこのファイルから差し込む。index.html を触らずに済むようにするため。
 * 自動切り替えの準備でつまずいても、タブの手動切り替えは今までどおり動く。
 * ---------------------------------------------------------------------- */
const SHOT_RATIO = [1400, 1050];   // 枠の［横, 縦］。短いシートの見本は、この比率まで下に余白を足してある
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
    console.warn("[リース見積診断] 見本の自動切り替えを止めました（タブの手動切り替えは使えます）", err);
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



/* ============================================================ 購入の控え */
function readSnap() {
  try {
    const s = JSON.parse(localStorage.getItem(SNAP_KEY) || "null");
    return s && s.price > 0 && s.months > 0 && s.monthly > 0 ? s : null;
  } catch (_) { return null; }
}
function writeSnap(s) {
  try { localStorage.setItem(SNAP_KEY, JSON.stringify(s)); } catch (_) { /* 保存できなくても、この画面では開いておく */ }
  P.snap = s;
}
function forgetSnap() {
  try { localStorage.removeItem(SNAP_KEY); } catch (_) { /* noop */ }
  P.snap = null;
}
function readJSON(key) {
  try { return JSON.parse(localStorage.getItem(key) || "null"); } catch (_) { return null; }
}
function writeJSON(key, v) {
  try { localStorage.setItem(key, JSON.stringify(v)); } catch (_) { /* noop */ }
}
function removeKey(key) {
  try { localStorage.removeItem(key); } catch (_) { /* noop */ }
}

/* ============================================================ 決済から戻ったとき */
/* 画面の上に出す案内。[文, 「この見積で開く」ボタンを出すか] */
const NOTES = {
  checking: ["お支払いを確認しています…", false],
  askQuote: ["お支払いの記録は届いています。購入した見積の数字（物件価額・期間・月額）を下のツールに入れてから、「この見積で開く」を押してください。", true],
  other_quote: ["購入した見積と数字が違います。購入したときと同じ物件価額・期間・月額を入れて、もう一度「この見積で開く」を押してください。", true],
  not_paid: ["お支払いがまだ完了していないようです。完了していれば、少し待ってからページを再読み込みしてください。", false],
  expired: ["最初に開いてから24時間が過ぎたため、購入した見積を閉じました。", false],
  network: ["確認の通信に失敗しました。時間をおいて、ページを再読み込みしてください。", false],
  other: [`この注文では開けませんでした。お支払い済みの場合は、領収メールを添えて ${CONTACT} までご連絡ください。`, false],
};
function showNote(key) {
  const [text, withButton] = NOTES[key] || NOTES.other;
  $("payNoteMsg").textContent = text;
  $("payNoteBtn").hidden = !withButton;
  $("payNote").hidden = false;
}
function hideNote() { $("payNote").hidden = true; }

/** ページを開いたときに一度だけ呼ぶ：Stripe から戻ってきたか、控えがまだ有効かを確かめる */
async function resumePurchase() {
  const ret = readReturnOrder();
  if (ret) {
    cleanReturnUrl();
    writeJSON(ORDER_KEY, { order: ret, at: Date.now() });
    const pending = readJSON(PENDING_KEY);
    if (pending && pending.fp && pending.snap) {
      await openWith(ret, pending.fp, pending.snap);
    } else {
      // 別のブラウザで戻ってきた（PayPayのアプリを経由した場合など）。見積を入れ直してもらう
      showNote("askQuote");
      $("payNote").scrollIntoView({ behavior: "smooth", block: "center" });
    }
    return;
  }
  if (P.snap && P.snap.order) { await recheck(); return; }
  const waiting = readJSON(ORDER_KEY);
  if (waiting && waiting.order && Date.now() - waiting.at < 7 * 24 * 3600e3) {
    // 前回、確認の途中で止まった（通信の失敗など）。購入ボタンを押したときの見積が残っていれば、それで確かめ直す
    const pending = readJSON(PENDING_KEY);
    if (pending && pending.fp && pending.snap) await openWith(waiting.order, pending.fp, pending.snap);
    else showNote("askQuote");
  }
}

/** 注文番号と見積で Worker に確かめ、よければ購入後の画面を開く */
async function openWith(order, fp, base) {
  showNote("checking");
  let res;
  try {
    res = await verifyOrder(order, fp);
  } catch (e) {
    showNote("network");
    track("lease_verify_failed", { reason: "network" });
    return false;
  }
  if (!res || res.valid !== true) {
    const reason = res && res.reason;
    showNote(NOTES[reason] ? reason : "other");
    track("lease_verify_failed", { reason: String(reason) });
    return false;
  }
  const expiresAt = Date.parse(res.expiresAt) || Date.now() + 24 * 3600e3;
  const snap = { price: base.price, months: base.months, monthly: base.monthly, residual: base.residual || 0,
    life: base.life ?? defaultLife(base.months), order, expiresAt, at: Date.now() };
  writeSnap(snap);
  removeKey(PENDING_KEY);
  removeKey(ORDER_KEY);
  hideNote();
  P.over = {};
  paintAll();
  $("paidOpen").scrollIntoView({ behavior: "smooth", block: "start" });
  track("lease_purchase_verified");
  return true;
}

/** 控えが残っているとき：まだ開いてよいかを Worker に確かめ直す */
async function recheck() {
  const s = P.snap;
  try {
    const res = await verifyOrder(s.order, await quoteFingerprint(s));
    if (res.valid === true) {
      writeSnap({ ...s, expiresAt: Date.parse(res.expiresAt) || s.expiresAt });
      paintAll();
    } else if (["expired", "other_quote", "inactive", "wrong_product", "wrong_link", "not_found", "bad_order", "no_fingerprint"].includes(res.reason)) {
      forgetSnap();
      paintAll();
      showNote(res.reason === "expired" ? "expired" : "other");
    }
    // Stripe や Worker 側の一時的な不具合（stripe_error など）では、お客さまの控えは消さない
  } catch (_) {
    // 通信できないときは、手元の期限までそのまま開いておく
  }
}

/** 別のブラウザで戻ってきた人が、見積を入れ直して押すボタン */
async function onOpenWithQuote() {
  const waiting = readJSON(ORDER_KEY);
  if (!waiting || !waiting.order) { hideNote(); return; }
  const q = P.quote;
  if (!q) {
    showNote("askQuote");
    $("tool").scrollIntoView({ behavior: "smooth", block: "start" });
    return;
  }
  const base = { price: q.price, months: q.months, monthly: q.monthly, residual: q.residual || 0, life: P.life ?? defaultLife(q.months) };
  await openWith(waiting.order, await quoteFingerprint(base), base);
}

function leftText(t) {
  const ms = t - Date.now();
  if (!(ms > 0)) return "";
  const h = Math.floor(ms / 3600e3), m = Math.floor((ms % 3600e3) / 60e3);
  return `あと ${h}時間${String(m).padStart(2, "0")}分 開けます（最初に開いてから24時間）`;
}

/* ============================================================ はじめに */
export function initPaid() {
  P.snap = readSnap();
  renderSamples();
  try { initShots(); } catch (e) { console.warn("[リース見積診断] シート見本を動かせませんでした", e); }

  $("lifeChips").addEventListener("click", (e) => {
    const b = e.target.closest("[data-life]"); if (!b) return;
    setLife(+b.dataset.life);
  });
  $("lifeInput").addEventListener("input", () => {
    const v = parseInt($("lifeInput").value, 10);
    if (v >= LIFE_MIN && v <= LIFE_MAX) setLife(v, true);
  });
  $("buy").addEventListener("click", onBuy);
  $("dl").addEventListener("click", onDownload);
  $("relock").addEventListener("click", () => {
    forgetSnap(); removeKey(PENDING_KEY); removeKey(ORDER_KEY); hideNote(); paintAll();
    $("tool").scrollIntoView({ behavior: "smooth", block: "start" });
  });
  $("payNoteBtn").addEventListener("click", onOpenWithQuote);
  // 残り時間の表示を30秒ごとに更新し、期限が来たら閉じる
  setInterval(() => {
    if (!P.snap || !P.snap.expiresAt) return;
    if (Date.now() > P.snap.expiresAt) paintAll();
    else $("openLeft").textContent = leftText(P.snap.expiresAt);
  }, 30000);
  $("openLife").addEventListener("change", () => {
    const v = parseInt($("openLife").value, 10);
    if (P.snap && v >= LIFE_MIN && v <= LIFE_MAX) { writeSnap({ ...P.snap, life: v }); paintOpen(); }
  });
  document.querySelectorAll("[data-q]").forEach((card) => card.addEventListener("click", () => {
    $("buyBox").scrollIntoView({ behavior: "smooth", block: "center" });
  }));

  // 購入後のグラフを動かすつまみ
  const sliders = [["sFund", "fundRate", 0.001], ["sLoan", "loanRate", 0.001], ["sKappu", "kappuRate", 0.001], ["sTax", "corpTax", 0.001]];
  for (const [id, key, unit] of sliders) {
    $(id).addEventListener("input", () => { P.over[key] = +$(id).value * unit; paintOpen(); });
  }
  $("sCancel").addEventListener("input", () => { P.cancelAt = +$("sCancel").value; paintCancel(); });
  $("legend").addEventListener("click", (e) => {
    const b = e.target.closest("[data-m]"); if (!b) return;
    const k = b.dataset.m;
    if (Object.values(P.show).filter(Boolean).length === 1 && P.show[k]) return;   // 全部は消さない
    P.show[k] = !P.show[k];
    paintOpen();
  });
  $("resetOver").addEventListener("click", () => { P.over = {}; paintOpen(true); });
  ["pDonut", "pCompare", "pCancel", "pExpense"].forEach((id) => attachTips($(id)));
  // 画面の向きや幅が変わったら、グラフの細身・通常を描き分け直す
  let wasCompact = compact().compact, t = 0;
  window.addEventListener("resize", () => {
    clearTimeout(t);
    t = setTimeout(() => { const c = compact().compact; if (c !== wasCompact && P.snap) { wasCompact = c; paintOpen(); } }, 150);
  });
  paintAll();
  resumePurchase().catch((e) => console.warn("[リース見積診断] 購入の確認でつまずきました", e));
}

/** 無料ツールで計算するたびに呼ばれる。quote は計算できないとき null */
export function updatePaid(quote) {
  P.quote = quote ? { ...quote, monthly: Math.round(quote.monthly) } : null;
  if (P.quote && !P.lifeTouched) P.life = defaultLife(P.quote.months);
  paintAll();
}

function setLife(v, fromInput = false) {
  P.life = v; P.lifeTouched = true;
  if (!fromInput) $("lifeInput").value = "";
  paintBuy();
}

function paintAll() {
  if (P.snap && P.snap.expiresAt && Date.now() > P.snap.expiresAt) {
    forgetSnap();
    showNote("expired");
  }
  $("paid").hidden = !(P.quote && !P.snap);
  $("paidOpen").hidden = !P.snap;
  if (P.quote && !P.snap) { paintTotal(); paintBuy(); }
  if (P.snap) paintOpen(true);
}

/* ============================================================ 購入前 */
/** 見出しの文中に、この見積の支払総額を入れる */
function paintTotal() {
  const q = P.quote;
  $("paidTotal").textContent = `${yen(q.monthly * q.months)}円`;
}

function paintBuy() {
  const q = P.quote;
  const life = P.life ?? (q ? defaultLife(q.months) : 5);
  document.querySelectorAll("#lifeChips [data-life]").forEach((b) => b.classList.toggle("is-on", +b.dataset.life === life));
  $("lifeNow").textContent = `${life}年`;
  const btn = $("buy"), msg = $("buyMsg");
  if (FREE_MODE) {
    btn.disabled = false; btn.textContent = "テスト用：無料で開く";
    msg.textContent = "テスト環境です。支払いなしで、購入後の画面を確かめられます。";
  } else if (payUrlReady()) {
    btn.disabled = false; btn.textContent = "1,000円で購入する";
    msg.textContent = "購入するのは、この画面の見積（物件価額・期間・月額）1件分です。お支払いは Stripe の画面で、カード・Apple Pay・Google Pay・PayPay から選べます。";
  } else {
    btn.disabled = true; btn.textContent = "近日公開";
    msg.textContent = "ただいま販売の準備中です。";
  }
}

async function onBuy() {
  const q = P.quote; if (!q) return;
  const snap = { price: q.price, months: q.months, monthly: q.monthly, residual: q.residual || 0,
    life: P.life ?? defaultLife(q.months), at: Date.now() };
  if (FREE_MODE) {
    writeSnap(snap); P.over = {};
    paintAll();
    $("paidOpen").scrollIntoView({ behavior: "smooth", block: "start" });
    track("lease_unlock_test");
    return;
  }
  if (payUrlReady()) {
    // 見積のハッシュを支払いリンクに付けて、決済と見積を結び付ける。
    // 戻ってきたとき用に、押した時点の見積を控えておく（別のタブや再読み込みでも残るよう localStorage に置く）
    const btn = $("buy");
    btn.disabled = true;
    try {
      const fp = await quoteFingerprint(snap);
      writeJSON(PENDING_KEY, { fp, snap, at: Date.now() });
      track("lease_buy_click");
      location.href = payUrl(fp);
    } catch (e) {
      btn.disabled = false;
      $("buyMsg").textContent = "決済の画面を開けませんでした。ページを再読み込みしてお試しください。";
    }
  }
}

/** 購入前に置く見本のグラフ（ぼかす）。見本の見積の数字なので、お客さまの見積は含まない */
function renderSamples() {
  const r = analyze(SAMPLE);
  $("qfig1").innerHTML = donutMarkup(r);
  $("qfig2").innerHTML = lineCompare(r, { A: true, D: true });
  $("qfig3").innerHTML = lineCompare(r, { A: true, B: true, C: true });
  $("qfig4").innerHTML = areaRemaining(r, 24);
  $("qfig5").innerHTML = barsExpense(r);
}

/* ============================================================ 購入後 */
function result() {
  return analyze({ ...P.snap, ...P.over });
}
/** スマホの幅では、字が小さくならないよう細身のグラフに描き替える */
const compact = () => ({ compact: window.matchMedia("(max-width: 560px)").matches });

function paintOpen(syncSliders = false) {
  const s = P.snap; if (!s) return;
  const r = result();
  $("openQuote").textContent = `物件価額 ${yen(s.price)}円 ／ ${s.months}か月 ／ 月額 ${yen(s.monthly)}円${s.residual > 0 ? ` ／ 残価 ${yen(s.residual)}円` : ""}`;
  $("openLife").value = s.life;
  const q = P.quote;
  const differs = q && (q.price !== s.price || q.months !== s.months || Math.round(q.monthly) !== s.monthly);
  $("openDiff").hidden = !differs;
  $("openLeft").textContent = s.expiresAt ? leftText(s.expiresAt) : "";
  $("relock").hidden = !(FREE_MODE || TEST_HOST);

  if (syncSliders) {
    const a = r.input;
    $("sFund").value = Math.round(a.fundRate * 1000);
    $("sLoan").value = Math.round(a.loanRate * 1000);
    $("sKappu").value = Math.round(r.kappuRate * 1000);
    $("sTax").value = Math.round(a.corpTax * 1000);
    $("sCancel").max = s.months;
    P.cancelAt = Math.min(P.cancelAt, s.months);
    $("sCancel").value = P.cancelAt;
  }
  $("vFund").textContent = `${(r.input.fundRate * 100).toFixed(1)}%`;
  $("vLoan").textContent = `${(r.input.loanRate * 100).toFixed(1)}%`;
  $("vKappu").textContent = `${(r.kappuRate * 100).toFixed(1)}%`;
  $("vTax").textContent = `${(r.input.corpTax * 100).toFixed(1)}%`;

  // 1. 利益
  keepTip("pDonut", donutMarkup(r, compact()));
  $("profitLine").innerHTML = r.lease.profit >= 0
    ? `リース会社の利益は <b>${yen(r.lease.profit)}円</b>（リース料総額の${(r.margin * 100).toFixed(1)}%）と推計されます。`
    : `この月額では、リース会社の利益は <b>${yen(r.lease.profit)}円</b> と、マイナスの計算になります。調達金利の前提を下げて確かめてください。`;
  $("profitNote").textContent = `損益分岐の月額は ${yen(r.lease.breakEven)}円。見積の月額との差 ${yen(s.monthly - r.lease.breakEven)}円 が、1回あたりの利益の目安です。`;

  // 2. 買い方の比較
  $("legend").innerHTML = Object.entries(METHODS).map(([k, m]) =>
    `<button type="button" data-m="${k}" aria-pressed="${P.show[k]}" style="--c:${m.color}">${m.label}</button>`).join("");
  keepTip("pCompare", lineCompare(r, P.show, compact()));
  const order = Object.keys(METHODS).sort((x, y) => r.cmpNpv[x] - r.cmpNpv[y]);
  $("cmpTable").innerHTML = `<tr><th>買い方</th><th>実質負担の合計</th><th>現在価値（小さい順）</th></tr>` + order.map((k, i) =>
    `<tr class="${i === 0 ? "is-best" : ""}"><td><i class="dot" style="background:${METHODS[k].color}"></i>${METHODS[k].label}${i === 0 ? "<em>現在価値でいちばん小さい</em>" : ""}</td><td>${yen(r.cmpTotal[k])}円</td><td>${yen(r.cmpNpv[k])}円</td></tr>`).join("");
  $("cmpNote").textContent = `現在価値は、先の支払ほど軽く数えて（年${(r.input.discRate * 100).toFixed(1)}%で割り引いて）足した金額です。払う時期の違いまで含めて比べられます。`;

  // 3. 解約
  paintCancel(r);

  // 4. 経費と税金
  keepTip("pExpense", barsExpense(r, compact()));
  const yrs = Math.min(5, r.dep.rows.length);
  const depSum = r.dep.rows.slice(0, yrs).reduce((t, d) => t + d.amount, 0);
  $("expenseLine").textContent = `買った場合、最初の${yrs}年で減価償却費は合計 ${yen(depSum)}円、償却資産税は合計 ${yen(r.cost.taxTotal)}円 です。`;
}

function paintCancel(r = result()) {
  const n = r.schedule.length, k = Math.min(Math.max(P.cancelAt, 0), n);
  keepTip("pCancel", areaRemaining(r, k, compact()));
  const rem = k >= n ? 0 : (k <= 0 ? r.lease.total : r.schedule[k - 1].remaining);
  $("vCancel").textContent = `${k}か月`;
  $("cancelLine").innerHTML = k >= n
    ? "満了まで払い終えると、残りの支払はありません。"
    : `${k}回払った時点で解約すると、残りの支払は <b>${yen(rem)}円</b> です。規定損害金は、この金額を基に契約ごとに決まります。`;
}

/** グラフを書き換えても、ツールチップの入れ物は残す */
function keepTip(id, markup) {
  const box = $(id), tip = box.querySelector(".vtip");
  box.innerHTML = markup;
  if (tip) box.appendChild(tip);
}

async function onDownload() {
  const btn = $("dl"); if (!P.snap) return;
  if (P.snap.expiresAt && Date.now() > P.snap.expiresAt) { paintAll(); return; }
  const label = btn.textContent;
  btn.disabled = true; btn.textContent = "Excelを作っています…";
  try {
    const r = result();
    const figs = [];
    const X = { excel: true };   // Excel には、そろった大きさ（520×300）で描いたグラフを貼る
    for (const [key, markup] of [["donut", donutMarkup(r, X)], ["compare", lineCompare(r, undefined, X)], ["remaining", areaRemaining(r, P.cancelAt, X)], ["expense", barsExpense(r, X)]]) {
      try { figs.push({ key, ...(await toPng(markup)) }); } catch (e) { console.warn("[リース見積診断] グラフを画像にできませんでした", key, e); }
    }
    await downloadLeaseXlsx(r, figs);
    track("lease_download");
    $("dlMsg").textContent = figs.length === 4 ? "ダウンロードしました。" : "ダウンロードしました（一部のグラフは画像にできませんでした）。";
  } catch (e) {
    $("dlMsg").textContent = e.message || "Excelを作れませんでした。ページを再読み込みしてお試しください。";
  } finally {
    btn.disabled = false; btn.textContent = label;
  }
}

function track(name, params = {}) {
  try { if (typeof gtag === "function") gtag("event", name, params); } catch (_) { /* noop */ }
}
