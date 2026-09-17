/* ============================================================================
 * /lease/license.js — 決済と解錠の確認（リース見積診断・Stripe）
 *
 * 流れ
 *   1. 購入ボタン：見積（物件価額・期間・月額）のハッシュを client_reference_id に付けて、
 *      Stripe の支払いリンクへ移る。決済の時点で「どの見積の購入か」が Stripe 側に残る
 *   2. 支払いが終わると、Stripe が ?session_id=cs_live_… を付けてこのページへ戻す
 *   3. Worker に「この注文番号とこの見積で開いてよいか」を聞く。
 *      Worker は Stripe に直接問い合わせて、支払い済み・1,000円・この支払いリンク・同じ見積かを確かめる
 *
 * 判断はすべて Worker が行う。画面の JS は誰でも読めるので、ブラウザ側だけの制限は当てにしない。
 * 見積の数字そのものは送らない。ブラウザの中でハッシュにしてから送る。
 * ========================================================================== */

/* Stripe の支払いリンク（本番）。支払い完了ページの戻り先は Stripe の管理画面で設定する
     テスト中 : https://pythonddd.github.io/test_credit_test/lease/?session_id={CHECKOUT_SESSION_ID}
     公開後   : https://kazumono.com/lease/?session_id={CHECKOUT_SESSION_ID} */
const PAY_URL = "https://buy.stripe.com/14A7sM6Tk85B8I16EsaR201";
const WORKER = "https://square-license.stats-okinawa.workers.dev";

export const SERVICE_ID = "lease";

export function payUrlReady() {
  return /^https:\/\/buy\.stripe\.com\/[A-Za-z0-9_]+$/.test(PAY_URL);
}

/** 購入ボタンの行き先。見積のハッシュを付けて、決済と見積を結び付ける */
export function payUrl(fp) {
  return `${PAY_URL}?client_reference_id=${encodeURIComponent(fp)}`;
}

/** 見積の目印（SHA-256）。1回の支払いを見積1件に限るために使う。数字だけを混ぜ、社名などは入れない */
export async function quoteFingerprint(q) {
  const norm = [q.price, q.months, Math.round(q.monthly)].map((x) => String(Math.round(+x || 0))).join("/");
  const buf = await crypto.subtle.digest("SHA-256",
    new TextEncoder().encode("svc:" + SERVICE_ID + "|q:" + norm));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

/** Stripe から戻ってきたときの注文番号（無ければ null） */
export function readReturnOrder() {
  const id = new URLSearchParams(location.search).get("session_id");
  return id && /^cs_(live|test)_[A-Za-z0-9]{10,}$/.test(id) ? id : null;
}

/** 注文番号をアドレス欄から消す（ブックマークや画面の共有に残さないため） */
export function cleanReturnUrl() {
  const u = new URL(location.href);
  if (!u.searchParams.has("session_id")) return;
  u.searchParams.delete("session_id");
  history.replaceState(null, "", u.pathname + (u.search ? u.search : "") + u.hash);
}

/**
 * Worker に解錠してよいかを聞く。
 * @returns {Promise<{valid:boolean, reason?:string, expiresAt?:string}>}
 * 通信に失敗したときは例外を投げる（呼び出し側で「確認できなかった」と扱う）
 */
export async function verifyOrder(order, fp) {
  const url = `${WORKER}/verify?svc=${SERVICE_ID}&order=${encodeURIComponent(order)}&fp=${encodeURIComponent(fp)}`;
  const res = await fetch(url, { cache: "no-store" });
  if (!res.ok) throw new Error(`verify ${res.status}`);
  return res.json();
}
