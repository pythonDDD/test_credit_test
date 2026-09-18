/* ============================================================================
 * /lease/engine.js — リース料の計算エンジン（画面とExcelの両方で使う）
 *
 * 約束ごと
 *   ・金額は円、率は小数（3% → 0.03）
 *   ・支払は毎月の期末払い（1回目の支払は開始から1か月後）で統一する
 *   ・残価は「満了時に物件の価値として残す額」。0 ならフルペイアウト
 *
 * 画面（無料）で使うのは forward（条件→月額）と reverse（月額→金利）。
 * analyze は有料のExcel用で、原価内訳・支払予定表・償却資産税・減価償却・
 * 調達方法の比較をまとめて計算する。
 * ========================================================================== */

/* 減価償却資産の耐用年数等に関する省令 別表第十（200%定率法）
 * 耐用年数: [償却率, 改定償却率, 保証率] */
const DB_TABLE = {2:[1.0,1.0,0.0], 3:[0.667,1.0,0.11089], 4:[0.5,1.0,0.12499], 5:[0.4,0.5,0.108], 6:[0.333,0.334,0.09911], 7:[0.286,0.334,0.0868], 8:[0.25,0.334,0.07909], 9:[0.222,0.25,0.07126], 10:[0.2,0.25,0.06552], 11:[0.182,0.2,0.05992], 12:[0.167,0.2,0.05566], 13:[0.154,0.167,0.0518], 14:[0.143,0.167,0.04854], 15:[0.133,0.143,0.04565], 16:[0.125,0.143,0.04294], 17:[0.118,0.125,0.04038], 18:[0.111,0.112,0.03884], 19:[0.105,0.112,0.03693], 20:[0.1,0.112,0.03486], 21:[0.095,0.1,0.03335], 22:[0.091,0.1,0.03182], 23:[0.087,0.091,0.03052], 24:[0.083,0.084,0.02969], 25:[0.08,0.084,0.02841], 26:[0.077,0.084,0.02716], 27:[0.074,0.077,0.02624], 28:[0.071,0.072,0.02568], 29:[0.069,0.072,0.02463], 30:[0.067,0.072,0.02366], 31:[0.065,0.067,0.02286], 32:[0.063,0.067,0.02216], 33:[0.061,0.063,0.02161], 34:[0.059,0.063,0.02097], 35:[0.057,0.059,0.02051], 36:[0.056,0.059,0.01974], 37:[0.054,0.056,0.0195], 38:[0.053,0.056,0.01882], 39:[0.051,0.053,0.0186], 40:[0.05,0.053,0.01791], 41:[0.049,0.05,0.01741], 42:[0.048,0.05,0.01694], 43:[0.047,0.048,0.01664], 44:[0.045,0.046,0.01664], 45:[0.044,0.046,0.01634], 46:[0.043,0.044,0.01601], 47:[0.043,0.044,0.01532], 48:[0.042,0.044,0.01499], 49:[0.041,0.042,0.01475], 50:[0.04,0.042,0.0144]};
/* 償却資産（固定資産税）の減価率 r（旧定率法の償却率）
 * 前年中に取得した年は 1−r/2、それ以降は 1−r を掛ける */
const OLD_DB_RATE = {2:0.684, 3:0.536, 4:0.438, 5:0.369, 6:0.319, 7:0.28, 8:0.25, 9:0.226, 10:0.206, 11:0.189, 12:0.175, 13:0.162, 14:0.152, 15:0.142, 16:0.134, 17:0.127, 18:0.12, 19:0.114, 20:0.109, 21:0.104, 22:0.099, 23:0.095, 24:0.092, 25:0.088, 26:0.085, 27:0.082, 28:0.079, 29:0.076, 30:0.074, 31:0.072, 32:0.069, 33:0.067, 34:0.066, 35:0.064, 36:0.062, 37:0.06, 38:0.059, 39:0.057, 40:0.056, 41:0.055, 42:0.053, 43:0.052, 44:0.051, 45:0.05, 46:0.049, 47:0.048, 48:0.047, 49:0.046, 50:0.045};

export const DEFAULTS = Object.freeze({
  fundRate: 0.02,     // リース会社の資金調達金利（年）
  insRate: 0.01,      // 動産総合保険料（物件価額に対する率・期間合計）
  taxRate: 0.014,     // 償却資産税率（標準税率）
  initCost: 0,        // その他の初期費用（円）
  corpTax: 0.3062,    // 法人実効税率
  loanRate: 0.015,    // 銀行借入金利（年）
  discRate: 0.03,     // 比較に使う割引率（年）
  selfIns: 0,         // 自社で掛ける保険料（年額・購入の場合）
  kappuRate: null,    // 割賦の金利（年）。null ならリースと同じ条件から決める
});

/* ------------------------------------------------------------ 丸め */
const EPS = 1e-6;
/** Excel の ROUND と同じ（0.5 は0から遠い方へ） */
export function round(x, digits = 0) {
  const m = Math.pow(10, digits);
  return Math.sign(x) * Math.round(Math.abs(x) * m + EPS * 1e-3) / m;
}
/** Excel の ROUNDDOWN（0に近い方へ切り捨て） */
export function roundDown(x, digits = 0) {
  const m = Math.pow(10, digits);
  return Math.sign(x) * Math.floor(Math.abs(x) * m + EPS) / m;
}
/** Excel の ROUNDUP（0から遠い方へ切り上げ） */
export function roundUp(x, digits = 0) {
  const m = Math.pow(10, digits);
  return Math.sign(x) * Math.ceil(Math.abs(x) * m - EPS) / m;
}

/* ------------------------------------------------------------ 金利の基本 */
/** 元利均等の毎回の支払額（正の数）。fv は満了時に残す額、type は 0=期末 1=期首 */
export function pmt(i, n, pv, fv = 0, type = 0) {
  if (!(n > 0)) return 0;
  if (Math.abs(i) < 1e-12) return (pv - fv) / n;
  const v = Math.pow(1 + i, -n);
  return (pv - fv * v) * i / ((1 - v) * (type ? 1 + i : 1));
}

/** 毎回 p を n 回払い、満了時に fv を返すとき、今いくら分に当たるか */
function presentValue(i, n, p, fv = 0, type = 0) {
  if (Math.abs(i) < 1e-12) return p * n + fv;
  const v = Math.pow(1 + i, -n);
  return p * (1 - v) / i * (type ? 1 + i : 1) + fv * v;
}

/** 1回あたりの利率を逆算する（Excel の RATE と同じ意味）。解けなければ NaN */
export function rate(n, p, pv, fv = 0, type = 0) {
  if (!(n > 0) || !(p > 0) || !(pv > 0)) return NaN;
  let lo = -0.05, hi = 1;
  if (presentValue(lo, n, p, fv, type) < pv || presentValue(hi, n, p, fv, type) > pv) return NaN;
  for (let k = 0; k < 200; k++) {
    const mid = (lo + hi) / 2;
    if (presentValue(mid, n, p, fv, type) > pv) lo = mid; else hi = mid;
  }
  return (lo + hi) / 2;
}

/* ------------------------------------------------------------ 無料：条件→月額 */
/**
 * 物件価額・期間・年利・残価・その他費用率から、月額リース料を出す。
 * @returns 月額、リース料率、支払総額、内訳（元本・金利相当額・その他費用）、実質年率（金利）
 */
export function forward({ price, months, annualRate, residual = 0, miscRate = 0 }) {
  const i = annualRate / 12;
  const base = pmt(i, months, price, residual);
  const misc = price * miscRate / 12;

  /* 見積書に載る月額は1円単位の数字なので、ここで先に確定させる。
     支払総額・内訳・実質年率は、すべてこの確定した月額から導く。
     （丸める前の値で総額を出すと「月額×回数」と合わなくなる） */
  const monthly = round(base + misc);
  const total = monthly * months;
  const principal = round(Math.abs(i) < 1e-12 ? price - residual : price - residual * Math.pow(1 + i, -months));
  const miscTotal = round(misc * months);

  return {
    monthly,
    ratio: monthly / price,
    total,
    principal,                        // 物件価額の回収（元本）
    interest: total - principal - miscTotal,   // 残り＝金利相当額。3つ足すと必ず支払総額になる
    miscTotal,                        // その他費用（税・保険・管理費）
    residual,
    multiple: total / price,
    effective: rate(months, monthly, price, residual) * 12,   // 実質年率（金利）
  };
}

/* ------------------------------------------------------------ 無料：月額→金利 */
/** 見積の月額（または料率）から、実質年率（金利）を逆算する */
export function reverse({ price, months, monthly: raw, residual = 0 }) {
  const monthly = round(raw);       // 料率から入れた場合も、月額は1円単位に揃える
  const total = monthly * months;
  return {
    monthly,
    ratio: monthly / price,
    total,
    extra: total - price,             // 物件価額より多く払う額
    multiple: total / price,
    residual,
    effective: rate(months, monthly, price, residual) * 12,
  };
}

/** 残価を 0%, 5%, … と置いたときの実質年率（金利）を並べる */
export function residualScenarios({ price, months, monthly }, steps = [0, 0.05, 0.1, 0.15, 0.2, 0.25, 0.3]) {
  return steps.map((r) => ({ ratio: r, residual: price * r, effective: rate(months, monthly, price, price * r) * 12 }));
}

/** 残価を 0%, 5%, … と置いたときの月額リース料を並べる */
export function residualPayments(args, steps = [0, 0.05, 0.1, 0.15, 0.2, 0.25, 0.3]) {
  return steps.map((r) => ({ ratio: r, residual: args.price * r, monthly: forward({ ...args, residual: args.price * r }).monthly }));
}

/* ------------------------------------------------------------ 償却資産税 */
/** 年度ごとの評価額と税額（課税標準は千円未満、税額は百円未満を切り捨て） */
export function propertyTax(price, life, taxRate, years) {
  const r = OLD_DB_RATE[life];
  if (r === undefined) throw new Error("耐用年数は2〜50年で入れてください");
  const rows = [];
  let prev = price, cum = 0;
  for (let y = 1; y <= Math.max(years, 0); y++) {
    const keep = y === 1 ? 1 - r / 2 : 1 - r;
    const value = roundDown(prev * keep);
    const floor = price * 0.05;
    const base = Math.max(value, floor);
    const tax = roundDown(roundDown(base, -3) * taxRate, -2);
    cum += tax;
    rows.push({ year: y, keep, prev, value, floor, base, tax, cum });
    prev = value;
  }
  return rows;
}

/* ------------------------------------------------------------ 減価償却（200%定率法） */
export function depreciation(price, life) {
  const t = DB_TABLE[life];
  if (!t) throw new Error("耐用年数は2〜50年で入れてください");
  const [dbRate, revRate, guarantee] = t;
  const guaranteeAmt = roundDown(price * guarantee);
  const rows = [];
  let book = price, cum = 0, revisedBase = null;
  for (let y = 1; book > 1 && y <= 60; y++) {
    const adjusted = roundDown(book * dbRate);
    if (revisedBase === null && adjusted < guaranteeAmt) revisedBase = book;
    const amount = revisedBase !== null
      ? Math.min(roundDown(revisedBase * revRate), book - 1)
      : Math.min(adjusted, book - 1);
    cum += amount;
    rows.push({ year: y, open: book, adjusted, revised: revisedBase !== null, amount, close: book - amount, cum });
    book -= amount;
  }
  return { rows, dbRate, revRate, guarantee, guaranteeAmt };
}

/* ------------------------------------------------------------ 有料：Excel用の全体計算 */
/**
 * 見積（物件価額・期間・月額）を、原価とリース会社の利益に分ける。
 * 数字の決め方は note 版「リース料計算書 Pro」と同じにしてある（値の一致を検証済み）。
 */
export function analyze(input) {
  const a = { ...DEFAULTS, residual: 0, type: 0, ...input };
  const { price, months, monthly, life, residual, type } = a;
  const taxYears = a.taxYears ?? Math.ceil(months / 12);

  // A. 元本原価（リース会社が立て替えるお金）
  const ins = round(price * a.insRate);
  const tax = propertyTax(price, life, a.taxRate, Math.min(taxYears, 12));
  const taxTotal = tax.reduce((s, r) => s + r.tax, 0);
  const principalCost = price + ins + taxTotal + a.initCost;

  // B. 資金原価（立て替えている間の利息）
  const fundMonthlyRate = a.fundRate / 12;
  const fundPmt = pmt(fundMonthlyRate, months, principalCost, residual, type);
  const fundInterest = round(fundPmt * months + residual - principalCost);
  const leaseCost = principalCost + fundInterest;

  // C. 利益
  const total = monthly * months;
  const profit = total + residual - leaseCost;
  const breakEven = roundUp((leaseCost - residual) / months, -2);

  // D. 指標
  const lessorMonthly = rate(months, monthly, principalCost, residual, type);
  const lesseeRate = rate(months, monthly, price, residual, type) * 12;

  // 支払予定表（毎回の支払を、元本相当と利息相当に分ける）
  const schedule = [];
  let balance = principalCost, cumPaid = 0;
  for (let k = 1; k <= months; k++) {
    let interestPart = type === 1 && k === 1 ? 0 : round(balance * lessorMonthly);
    let principalPart = monthly - interestPart;
    // 毎回1円単位に丸めると最後に数円ずれる。最終回で吸収して、残高が残価ちょうどで終わるようにする。
    if (k === months && isFinite(balance)) {
      principalPart = balance - residual;
      interestPart = monthly - principalPart;
    }
    balance -= principalPart;
    cumPaid += monthly;
    schedule.push({ no: k, year: Math.ceil(k / 12), payment: monthly, principal: principalPart,
      interest: interestPart, balance, cumPaid, remaining: total - cumPaid });
  }

  // 減価償却（現金・借入で買った場合）
  const dep = depreciation(price, life);

  // 調達方法の比較（12年分）
  const Y = 12;
  const loanPmtExact = pmt(a.loanRate / 12, months, price);
  const loanPmt = round(loanPmtExact);
  const loanInterestByYear = yearlyLoanInterest(price, a.loanRate / 12, months, loanPmtExact, Y);
  // 割賦：同じリース会社から、税金と保険を除いた同じ条件で分割払いにした場合の金利を既定にする
  const kappuRate = a.kappuRate ?? kappuDefaultRate({ price, months, monthly, residual, taxTotal, ins });
  const kappuPmtExact = pmt(kappuRate / 12, months, price);
  const kappuPmt = round(kappuPmtExact);
  const kappuInterestByYear = yearlyLoanInterest(price, kappuRate / 12, months, kappuPmtExact, Y);
  const cmp = [];
  for (let y = 1; y <= Y; y++) {
    const payCount = Math.min(12, Math.max(0, months - 12 * (y - 1)));
    const taxY = tax[y - 1] ? tax[y - 1].tax : 0;
    const depY = dep.rows[y - 1] && y <= 20 ? dep.rows[y - 1].amount : 0;
    const taxIns = taxY + (payCount > 0 ? a.selfIns : 0);
    const leasePay = monthly * payCount;
    const saveA = -round(leasePay * a.corpTax);
    const cash = y === 1 ? price : 0;
    const saveB = -round((depY + taxIns) * a.corpTax);
    const repay = round(loanPmt * payCount);
    const loanInt = payCount === 0 ? 0 : round(loanInterestByYear[y - 1]);
    const saveC = -round((depY + loanInt + taxIns) * a.corpTax);
    const kRepay = round(kappuPmt * payCount);
    const kInt = payCount === 0 ? 0 : round(kappuInterestByYear[y - 1]);
    const saveD = -round((depY + kInt + taxIns) * a.corpTax);
    cmp.push({ year: y, payCount, tax: taxY, dep: depY,
      A: { pay: leasePay, save: saveA, net: leasePay + saveA },
      B: { pay: cash, taxIns, save: saveB, net: cash + taxIns + saveB },
      C: { pay: repay, interest: loanInt, taxIns, save: saveC, net: repay + taxIns + saveC },
      D: { pay: kRepay, interest: kInt, taxIns, save: saveD, net: kRepay + taxIns + saveD } });
  }
  const sum = (f) => cmp.reduce((s, r) => s + f(r), 0);
  const npv = (f) => cmp.reduce((s, r, k) => s + f(r) / Math.pow(1 + a.discRate, k + 1), 0);

  return {
    input: a, taxYears,
    cost: { price, ins, taxTotal, init: a.initCost, principalCost, fundMonthlyRate, fundPmt, fundInterest, leaseCost },
    lease: { monthly, months, total, profit, breakEven, residual },
    ratio: monthly / price,
    margin: total ? profit / total : 0,
    lessorMonthly, lessorYield: lessorMonthly * 12, lesseeRate,
    diffCash: total - price,
    schedule, tax, dep,
    cmp, loanPmt, kappuRate, kappuPmt,
    cmpTotal: { A: sum((r) => r.A.net), B: sum((r) => r.B.net), C: sum((r) => r.C.net), D: sum((r) => r.D.net) },
    cmpNpv: { A: npv((r) => r.A.net), B: npv((r) => r.B.net), C: npv((r) => r.C.net), D: npv((r) => r.D.net) },
    cmpCum: ["A", "B", "C", "D"].reduce((o, k) => {
      let c = 0; o[k] = cmp.map((r) => (c += r[k].net)); return o;
    }, {}),
  };
}

/** 借入の年ごとの支払利息（Excel の CUMIPMT を1年ずつ足したもの・期末払い） */
function yearlyLoanInterest(pv, i, n, p, years) {
  const out = new Array(years).fill(0);
  let bal = pv;
  for (let k = 1; k <= n && k <= years * 12; k++) {
    const it = bal * i;
    out[Math.ceil(k / 12) - 1] += it;
    bal -= p - it;
  }
  return out;
}

/** 割賦の金利の既定値：リースの月額から、リース会社が払う税金と保険の分を除いて金利に直す */
export function kappuDefaultRate({ price, months, monthly, residual = 0, taxTotal, ins }) {
  const m = monthly - (taxTotal + ins) / months;
  const r = rate(months, m, price, residual) * 12;
  return isFinite(r) ? Math.max(r, 0) : 0;
}

/** 実質年率（金利）が delta 下がったとき、支払総額がいくら減るか（交渉の目安） */
export function negotiationValue({ price, months, monthly, residual = 0 }, delta = 0.001) {
  const r = rate(months, monthly, price, residual) * 12;
  if (!isFinite(r) || r - delta <= -0.5) return NaN;
  const m2 = pmt((r - delta) / 12, months, price, residual);
  return (monthly - m2) * months;
}

/** 耐用年数の初期値：リース期間に近い年数（2〜50年） */
export function defaultLife(months) {
  return Math.min(LIFE_MAX, Math.max(LIFE_MIN, Math.round(months / 12)));
}

export const LIFE_MIN = 2, LIFE_MAX = 50;
