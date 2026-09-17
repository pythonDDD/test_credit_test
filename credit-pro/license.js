/* ============================================================================
 * license.js — 決済確認とロック解除
 *
 * 仕組み（2026-09 に Square から Stripe へ移行）:
 *   1. 購入ボタンで、会社名のハッシュを client_reference_id に付けて Stripe の支払いリンクへ移る。
 *      決済の時点で「どの会社の分か」が Stripe 側に残る
 *   2. 決済すると ?session_id=cs_live_… を付けて戻ってくる
 *   3. その注文番号を localStorage に保存し、Worker の /verify?svc=credit-pro&order=…&fp=… に問い合わせる。
 *      Worker は Stripe に直接たずねるので、通知の到着を待たずにその場で確かめられる
 *   ※ 移行前に Square で買った注文（transactionId）は、24時間の期限までこれまでの確認口で確かめる
 *
 * 方針:
 *   ・画面上の判定は常に無料。課金の対象はExcelファイルの受け取りのみ
 *   ・購入前に判定結果を最後まで見せる。何を買うのか分からないまま払わせない
 *   ・Workerに障害があっても、画面上の判定は動き続ける（決済確認と計算を分離）
 * ========================================================================== */

const WORKER = "https://square-license.stats-okinawa.workers.dev";

/* ★★ Stripe の支払いリンク（財務でポン！・500円）★★
   支払い完了後の戻り先は、Stripe の管理画面（支払いリンク → 支払い完了ページ →「確認ページを表示しない」）で設定する。
     テスト中 : https://pythonddd.github.io/test_credit_test/credit-pro/?session_id={CHECKOUT_SESSION_ID}
     公開後   : https://kazumono.com/credit-pro/?session_id={CHECKOUT_SESSION_ID}
   【注意】このファイルを新しい版で上書きすると、下の PAY_URL も一緒に置き換わります。 */
const PAY_URL = "https://buy.stripe.com/9B67sMb9AgC73nH7IwaR200";

/** 支払いリンクが未設定かどうか。未設定のまま黙って遷移させないための判定 */
export function payUrlReady() {
  return typeof PAY_URL === "string" && /^https:\/\/buy\.stripe\.com\/[A-Za-z0-9_]+$/.test(PAY_URL);
}

const KEY = "kazumono.credit-pro.order";

/** 決済から戻ってきたところか（URLに注文番号が付いているか）。Stripe は session_id、移行前の Square は transactionId */
export function hasReturnOrder() {
  const q = new URLSearchParams(location.search);
  return !!(q.get("session_id") || q.get("transactionId") || q.get("orderId"));
}

/** Stripe の注文番号か。Square の注文と確認口を分けるために使う */
export function isStripeOrder(order) {
  return /^cs_(live|test)_/.test(String(order || ""));
}

/** URLに注文番号が付いていれば保存し、URLからは消す（リロードで消えないように） */
function captureOrder() {
  const q = new URLSearchParams(location.search);
  const oid = q.get("session_id") || q.get("orderId") || q.get("transactionId");
  if (oid) {
    try { localStorage.setItem(KEY, oid); } catch (e) { /* プライベートモード等 */ }
    q.delete("session_id"); q.delete("orderId"); q.delete("transactionId");
    const rest = q.toString();
    history.replaceState(null, "", location.pathname + (rest ? "?" + rest : ""));
    return oid;
  }
  try { return localStorage.getItem(KEY); } catch (e) { return null; }
}

/**
 * ライセンス状態を返す。
 * @returns {Promise<{state:"licensed"|"unlicensed"|"offline", order:string|null}>}
 *   offline は「確認できなかった」状態。決済済みの人を締め出さないため区別する。
 */
/**
 * 会社名を、そのままでは送らずにハッシュ化する。
 * 1回の決済を1社分に限るための目印として使う。
 * 取引先の名前をこちらのサーバに残さないため、必ずブラウザ内で変換してから送る。
 */
/* このページが何のサービスかを表す名前。サービスを増やすときは必ず別の値にする。
   同じ500円のサービスが2つあると、金額だけでは区別できず、
   片方の支払いでもう片方が解錠できてしまうため、指紋に混ぜて切り分ける。 */
export const SERVICE_ID = "credit-pro";

export async function companyFingerprint(name) {
  const norm = String(name || "")
    .normalize("NFKC")
    .replace(/[\s　]/g, "")
    .replace(/[（(].*?[）)]/g, "")
    .toLowerCase();
  const buf = await crypto.subtle.digest("SHA-256",
    new TextEncoder().encode("svc:" + SERVICE_ID + "|co:" + norm));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("").slice(0, 32);
}

/**
 * 解錠してよいかをWorkerに尋ねる。
 * 判定はすべてWorker側で行う。画面のJSは誰でも書き換えられるため、
 * ここで独自に判断しても意味がないので、返ってきた答えに従うだけにする。
 */
export async function checkLicense(fp) {
  const order = captureOrder();
  if (!order) return { state: "unlicensed", order: null, reason: null };
  try {
    const ctl = new AbortController();
    const timer = setTimeout(() => ctl.abort(), 8000);
    // Stripe の注文はサービス名を付けて Stripe の確認口へ。Square の注文（移行前）は、これまでの確認口へ
    const q = (isStripeOrder(order) ? `svc=${SERVICE_ID}&` : "") +
      `order=${encodeURIComponent(order)}` + (fp ? `&fp=${encodeURIComponent(fp)}` : "");
    const res = await fetch(`${WORKER}/verify?${q}`, { signal: ctl.signal, cache: "no-store" });
    clearTimeout(timer);
    if (!res.ok) return { state: "offline", order, reason: null };
    const data = await res.json();
    return {
      state: data.valid === true ? "licensed" : "unlicensed",
      order,
      reason: data.reason || null,
      expiresAt: data.expiresAt || null,
    };
  } catch (e) {
    return { state: "offline", order, reason: null };
  }
}

/** 購入ボタンの行き先。会社名のハッシュを付けて、決済と会社を結び付ける（引数なしならリンクそのもの） */
export function payUrl(fp) {
  return fp ? `${PAY_URL}?client_reference_id=${encodeURIComponent(fp)}` : PAY_URL;
}

export function forgetOrder() {
  try { localStorage.removeItem(KEY); } catch (e) { /* noop */ }
}
