/* ============================================================================
 * /lease/license.js — 決済確認とロック解除（リース見積診断）
 *
 * 仕組み:
 *   1. Squareの支払いリンクで決済すると ?transactionId=… を付けて戻ってくる
 *   2. その注文番号を localStorage に保存する
 *   3. Worker の /verify?order=… に問い合わせ、valid:true ならロックを外す
 *
 * 方針:
 *   ・リース料と金利の計算は常に無料。課金の対象は、利益の内訳・比較のグラフとExcel
 *   ・Workerに障害があっても、無料の計算は動き続ける（決済確認と計算を分離）
 * ========================================================================== */

const WORKER = "https://square-license.stats-okinawa.workers.dev";

/* ★★ ここだけ書き換えてください ★★
   【注意】このファイルを新しい版で上書きすると、下のPAY_URLも一緒に置き換わります。
   ZIPを差し替えたあとは、必ずこの行が自分のリンクになっているか確認してください。
   Squareの支払いリンクURL。Square管理画面で作成して貼り替えます。
   作成時の設定：
     金額 1,000円 ／ Frequency: One-time（Monthlyにすると定期課金になります）／
     Redirect to a website after checkout: ON
     リダイレクト先は「今そのページを公開しているURL」にします。
       本番 : https://kazumono.com/lease/
   未設定（XXXXXXXXのまま）だと購入ボタンは押せず、「近日公開」と表示します。
   【公開前に必須】Worker 側で、この注文が1,000円（リース版）の支払いかを確かめる改修が要ります。
   それまでは、財務でポン！（500円）の注文でも解錠できてしまう恐れがあります。 */
const PAY_URL = "https://square.link/u/XXXXXXXX";

/** 支払いリンクが未設定かどうか。未設定のまま黙って遷移させないための判定 */
export function payUrlReady() {
  return typeof PAY_URL === "string" && /^https:\/\/square\.link\/u\/[A-Za-z0-9]+$/.test(PAY_URL)
    && !PAY_URL.includes("XXXXXXXX");
}

const KEY = "kazumono.lease.order";

/** URLに注文番号が付いていれば保存し、URLからは消す（リロードで消えないように） */
function captureOrder() {
  const q = new URLSearchParams(location.search);
  const oid = q.get("orderId") || q.get("transactionId");
  if (oid) {
    try { localStorage.setItem(KEY, oid); } catch (e) { /* プライベートモード等 */ }
    q.delete("orderId"); q.delete("transactionId");
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
/* このページが何のサービスかを表す名前。サービスを増やすときは必ず別の値にする。
   同じ500円のサービスが2つあると、金額だけでは区別できず、
   片方の支払いでもう片方が解錠できてしまうため、指紋に混ぜて切り分ける。 */
export const SERVICE_ID = "lease";

export async function quoteFingerprint(q) {
  // 1回の支払いを見積1件に限る目印。数字だけを混ぜ、社名などは送らない
  const norm = [q.price, q.months, Math.round(q.monthly)].map((x) => String(Math.round(+x || 0))).join("/");
  const buf = await crypto.subtle.digest("SHA-256",
    new TextEncoder().encode("svc:" + SERVICE_ID + "|q:" + norm));
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
    const q = `order=${encodeURIComponent(order)}` + (fp ? `&fp=${encodeURIComponent(fp)}` : "");
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

export function payUrl() { return PAY_URL; }

export function forgetOrder() {
  try { localStorage.removeItem(KEY); } catch (e) { /* noop */ }
}
