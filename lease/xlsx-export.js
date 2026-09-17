/* ============================================================================
 * /lease/xlsx-export.js — Excel版の出力（ExcelJS）
 *
 * 方針
 *   ・数式を一切書き込まない（値のみ）。数式入りの note 版と役割を分けるため
 *   ・参照表（償却率・減価残存率）や、入力シートは入れない
 *   ・見た目は Web 版（リース料金ポン！）の色に揃える
 *   ・ブラウザの中で作ってダウンロードする。サーバーには何も送らない
 * ========================================================================== */
let _ExcelJS = null;
async function getExcelJS() {
  if (_ExcelJS) return _ExcelJS;
  if (typeof window !== "undefined" && window.ExcelJS) return (_ExcelJS = window.ExcelJS);
  throw new Error("Excelを作る部品を読み込めませんでした。ページを再読み込みしてお試しください。");
}

const C = {
  deep: "FF33526C", sea: "FF527695", mist: "FFDFF6F1", mistL: "FFF2FAF8", paper: "FFFAF9F6",
  line: "FFD3DDE2", ink: "FF0F1A22", soft: "FF4A5B69", plum: "FFB67AB4", plumL: "FFF7EBF6",
  white: "FFFFFFFF", sub: "FFCFE0EA",
};
let F = "Meiryo UI";   // 見本画像を作るときだけ、描画環境にある字体へ差し替える
const YEN = '#,##0"円";"▲"#,##0"円";"-"';
const NUM = '#,##0;"▲"#,##0;"-"';
const PCT2 = '0.00%;"▲"0.00%;"-"';
const PCT3 = '0.000%';
const METHOD = { A: "リース", B: "現金で購入", C: "銀行借入で購入", D: "割賦" };

const thin = { style: "thin", color: { argb: C.line } };
const BOX = { top: thin, left: thin, bottom: thin, right: thin };
const fill = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
const font = (size = 10, bold = false, color = C.ink) => ({ name: F, size, bold, color: { argb: color } });
const AL = {
  l: { vertical: "middle", horizontal: "left", indent: 1 },
  c: { vertical: "middle", horizontal: "center" },
  r: { vertical: "middle", horizontal: "right", indent: 1 },
  w: { vertical: "middle", horizontal: "left", indent: 1, wrapText: true },
};
const colL = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

function put(ws, addr, value, o = {}) {
  const c = ws.getCell(addr);
  c.value = value === undefined ? null : value;
  c.font = o.font || font();
  if (o.fill) c.fill = fill(o.fill);
  if (o.numFmt) c.numFmt = o.numFmt;
  c.alignment = o.align || AL.l;
  if (o.border !== false) c.border = BOX;
  return c;
}
function band(ws, row, fromCol, toCol, value, o = {}) {
  ws.mergeCells(`${colL(fromCol)}${row}:${colL(toCol)}${row}`);
  return put(ws, `${colL(fromCol)}${row}`, value, o);
}

/** どのシートにも付ける見出し（2行）と、シートの設定 */
function sheet(wb, name, title, r, widths, { landscape = false, onePage = false } = {}) {
  const ws = wb.addWorksheet(name, {
    views: [{ showGridLines: false }],
    pageSetup: { paperSize: 9, orientation: landscape ? "landscape" : "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: onePage ? 1 : 0,
      margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } },
    properties: { defaultRowHeight: 18 },
  });
  widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  const last = widths.length;
  ws.getRow(1).height = 8;
  ws.getRow(2).height = 30;
  band(ws, 2, 2, last, `リース料金ポン！　${title}`, { font: font(15, true, C.white), fill: C.deep, border: false });
  ws.getRow(3).height = 20;
  band(ws, 3, 2, last, quoteLine(r), { font: font(9.5, false, C.sub), fill: C.deep, border: false });
  ws.getRow(4).height = 10;
  return ws;
}
function quoteLine(r) {
  const q = r.input;
  const rv = q.residual > 0 ? ` ／ 残価 ${q.residual.toLocaleString("ja-JP")}円` : "";
  return `物件価額 ${q.price.toLocaleString("ja-JP")}円 ／ ${q.months}か月 ／ 月額 ${Math.round(q.monthly).toLocaleString("ja-JP")}円${rv} ／ 法定耐用年数 ${q.life}年 ／ 作成 ${today()}`;
}
const today = () => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`; };

function section(ws, row, fromCol, toCol, label) {
  ws.getRow(row).height = 22;
  band(ws, row, fromCol, toCol, label, { font: font(11, true, C.deep), fill: C.mist, border: false });
}
function note(ws, row, fromCol, toCol, s, size = 9) {
  ws.mergeCells(`${colL(fromCol)}${row}:${colL(toCol)}${row}`);
  put(ws, `${colL(fromCol)}${row}`, s, { font: font(size, false, C.soft), align: AL.w, border: false });
  ws.getRow(row).height = Math.max(18, Math.ceil(s.length / 60) * 15);
}
const FOOT = "本ファイルは計算結果のみを収録しています（数式は入っていません）。一般的な計算方法による試算であり、実際の契約条件・税務の扱いは、リース会社と顧問税理士にご確認ください。";

/* ------------------------------------------------------------ ① ダッシュボード */
function dashboard(wb, r, figs) {
  const ws = sheet(wb, "ダッシュボード", "見積の分析", r, [2, ...Array(12).fill(10.5), 2], { onePage: true });
  const tiles = [
    ["月額リース料", r.lease.monthly, YEN, "見積の月額（税抜）"],
    ["支払総額", r.lease.total, YEN, "月額 × 支払回数"],
    ["リース料率", r.ratio, PCT3, "月額 ÷ 物件価額"],
    ["実質年率（金利）", r.lesseeRate, PCT2, "借入の金利に直した値"],
    ["リース会社の利益", r.lease.profit, YEN, "推計。前提は下の表"],
    ["物件価額より多く払う額", r.lease.total + r.input.residual - r.input.price, YEN, "支払総額 − 物件価額"],
    ["損益分岐の月額", r.lease.breakEven, YEN, "リース会社の利益がゼロになる月額"],
    ["リース会社の利回り", r.lessorYield, PCT2, "立て替えたお金に対する年利"],
  ];
  tiles.forEach((t, i) => {
    const row = 5 + Math.floor(i / 4) * 4, col = 2 + (i % 4) * 3;
    const hot = i === 4;
    const bg = hot ? C.plumL : C.mistL;
    band(ws, row, col, col + 2, t[0], { font: font(9, true, hot ? C.plum : C.soft), fill: bg, border: false });
    ws.mergeCells(`${colL(col)}${row + 1}:${colL(col + 2)}${row + 1}`);
    put(ws, `${colL(col)}${row + 1}`, t[1], { font: font(17, true, hot ? C.plum : C.ink), fill: bg, numFmt: t[2], align: AL.l, border: false });
    band(ws, row + 2, col, col + 2, t[3], { font: font(8.5, false, C.soft), fill: bg, border: false });
    ws.getRow(row + 1).height = 30;
  });
  ws.getRow(8).height = 8; ws.getRow(12).height = 10;

  const place = (key, col, row, widthPx) => {
    const g = figs.find((f) => f.key === key);
    if (!g) return;
    const id = wb.addImage({ base64: g.dataUrl.split(",")[1], extension: "png" });
    ws.addImage(id, { tl: { col, row }, ext: { width: widthPx, height: Math.round(widthPx * g.h / g.w) } });
  };
  section(ws, 13, 2, 7, "物件価額より多く払う分の中身");
  section(ws, 13, 8, 13, "4つの買い方の比較（税金の効果を含む累計）");
  place("donut", 1.1, 13.3, 420);
  place("compare", 7.1, 13.3, 450);
  section(ws, 31, 2, 7, "途中で解約したときの残りの支払");
  section(ws, 31, 8, 13, "買った場合の、毎年の経費と税金");
  place("remaining", 1.1, 31.3, 430);
  place("expense", 7.1, 31.3, 430);

  section(ws, 49, 2, 13, "計算の前提（Webで動かした値のまま）");
  const a = r.input;
  const pre = [
    ["リース会社の資金調達金利", a.fundRate, PCT2], ["保険料（物件価額に対して・期間合計）", a.insRate, PCT2],
    ["償却資産税率", a.taxRate, PCT2], ["法人実効税率", a.corpTax, PCT2],
    ["銀行借入の金利", a.loanRate, PCT2], ["割賦の金利", r.kappuRate, PCT2],
    ["比較に使う割引率", a.discRate, PCT2], ["法定耐用年数", a.life, '0"年"'],
  ];
  pre.forEach((p, i) => {
    const row = 50 + Math.floor(i / 2), col = i % 2 === 0 ? 2 : 8;
    ws.mergeCells(`${colL(col)}${row}:${colL(col + 3)}${row}`);
    put(ws, `${colL(col)}${row}`, p[0], { font: font(9.5), fill: C.paper });
    ws.mergeCells(`${colL(col + 4)}${row}:${colL(col + 5)}${row}`);
    put(ws, `${colL(col + 4)}${row}`, p[1], { font: font(10, true), numFmt: p[2], align: AL.r });
  });
  note(ws, 55, 2, 13, "リース会社の利益は、見積の月額から、物件代金・保険料・償却資産税・資金の金利を差し引いた推計です。割賦の金利は、同じリース会社から税金と保険を除いた同じ条件で分割払いにした場合を置いています。");
  note(ws, 56, 2, 13, FOOT);
}

/* ------------------------------------------------------------ ② 原価内訳 */
function costSheet(wb, r) {
  const ws = sheet(wb, "原価内訳", "原価内訳（リース会社の利益）", r, [2, 5, 34, 17, 52, 2]);
  const c = r.cost, L = r.lease, a = r.input;
  let row = 5;
  const line = (no, label, v, fmt, memo, o = {}) => {
    put(ws, `B${row}`, no, { align: AL.c, font: font(10, true, C.deep), fill: o.bg });
    put(ws, `C${row}`, label, { font: font(10, !!o.bold), fill: o.bg });
    put(ws, `D${row}`, v, { numFmt: fmt, align: AL.r, font: font(o.big ? 12 : 10, true, o.color || C.ink), fill: o.bg });
    put(ws, `E${row}`, memo, { font: font(9, false, C.soft), align: AL.w, fill: o.bg });
    ws.getRow(row).height = o.big ? 26 : 21;
    row++;
  };
  section(ws, row++, 2, 5, "A　リース会社が立て替えるお金");
  line("①", "物件代金", c.price, YEN, "リース会社がメーカー・販売店に払う金額です。");
  line("②", "保険料", c.ins, YEN, `物件価額の${(a.insRate * 100).toFixed(2)}%を期間合計として置いています。`);
  line("③", "償却資産税（期間合計）", c.taxTotal, YEN, "所有者のリース会社が毎年納めます。年度別は「償却資産税」シートにあります。");
  line("④", "その他の初期費用", c.init, YEN, "登録費用・事務手数料など。");
  line("", "小計", c.principalCost, YEN, "この金額を、期間中に回収します。", { bold: true, bg: C.mistL });
  row++;
  section(ws, row++, 2, 5, "B　立て替えている間の金利");
  line("⑤", "資金の金利", c.fundInterest, YEN, `リース会社が年${(a.fundRate * 100).toFixed(2)}%で資金を調達したと置いた場合の利息です。`);
  line("", "リース会社の原価合計", c.leaseCost, YEN, "これを下回る月額では、リース会社は赤字です。", { bold: true, bg: C.mistL });
  row++;
  section(ws, row++, 2, 5, "C　リース会社の利益");
  line("", "リース料の総額", L.total, YEN, `月額 ${Math.round(L.monthly).toLocaleString("ja-JP")}円 × ${L.months}回`);
  if (L.residual > 0) line("", "満了時の残価", L.residual, YEN, "満了時に物件の価値として回収する分です。");
  line("⑥", "リース会社の利益（推計）", L.profit, YEN, "総額（＋残価）− 原価合計。交渉の余地があるのは、主にこの部分です。", { bold: true, big: true, color: C.plum, bg: C.plumL });
  line("", "利益率", r.margin, PCT2, "リース料の総額に対する利益の割合です。");
  row++;
  section(ws, row++, 2, 5, "D　見積の見方");
  line("", "リース料率", r.ratio, PCT3, "月額 ÷ 物件価額。見積書で使われる指標です。");
  line("", "実質年率（金利）", r.lesseeRate, PCT2, "税金・保険・利益も含めて、借入の金利に直した値です。");
  line("", "リース会社の利回り", r.lessorYield, PCT2, "立て替えたお金（小計）に対する年利です。");
  line("", "損益分岐の月額", L.breakEven, YEN, "リース会社の利益がゼロになる月額（100円単位に切り上げ）です。");
  line("", "物件価額より多く払う額", L.total + L.residual - c.price, YEN, "②〜⑥の合計と同じです。");
  row++;
  note(ws, row++, 2, 5, FOOT);
  ws.views = [{ showGridLines: false }];
}

/* ------------------------------------------------------------ ③ 支払予定表 */
function scheduleSheet(wb, r) {
  const ws = sheet(wb, "支払予定表", "支払予定表", r, [2, 8, 7, 14, 14, 13, 15, 15, 15, 2]);
  const head = ["回数", "年目", "支払額", "うち元本相当", "うち利息相当", "元本残高", "支払累計", "残りの支払"];
  ws.getRow(5).height = 24;
  head.forEach((h, i) => put(ws, `${colL(i + 2)}5`, h, { font: font(9.5, true, C.white), fill: C.deep, align: AL.c }));
  r.schedule.forEach((s, i) => {
    const row = 6 + i, bg = s.year % 2 === 0 ? C.mistL : C.white;
    const vals = [s.no, s.year, s.payment, s.principal, s.interest, Math.max(s.balance, 0), s.cumPaid, s.remaining];
    vals.forEach((v, k) => put(ws, `${colL(k + 2)}${row}`, v, {
      numFmt: k < 2 ? "0" : NUM, align: k < 2 ? AL.c : AL.r, fill: bg,
      font: font(9.5, k === 7, k === 7 ? C.deep : C.ink) }));
  });
  const end = 6 + r.schedule.length;
  note(ws, end + 1, 2, 9, "「残りの支払」は、その回まで払い終えた時点で残っているリース料の合計です。途中で解約するときの規定損害金は、これを基に契約ごとに決まります（割引の有無は契約書で確認してください）。");
  note(ws, end + 2, 2, 9, "元本相当・利息相当は、リース会社が立て替えたお金（原価内訳の小計）を、リース会社の利回りで割り振ったものです。");
  note(ws, end + 3, 2, 9, FOOT);
  ws.views = [{ state: "frozen", ySplit: 5, showGridLines: false }];
  ws.pageSetup.printTitlesRow = "5:5";
}

/* ------------------------------------------------------------ ④ 償却資産税 */
function taxSheet(wb, r) {
  const ws = sheet(wb, "償却資産税", "償却資産税（リース会社が毎年納める税金）", r, [2, 10, 18, 18, 16, 16, 2]);
  const head = ["年度", "評価額", "課税標準額", "税額", "税額の累計"];
  ws.getRow(5).height = 24;
  head.forEach((h, i) => put(ws, `${colL(i + 2)}5`, h, { font: font(9.5, true, C.white), fill: C.deep, align: AL.c }));
  r.tax.forEach((t, i) => {
    const row = 6 + i, bg = i % 2 ? C.mistL : C.white;
    [`${t.year}年目`, t.value, t.base, t.tax, t.cum].forEach((v, k) =>
      put(ws, `${colL(k + 2)}${row}`, v, { numFmt: k ? NUM : undefined, align: k ? AL.r : AL.c, fill: bg, font: font(10, k === 3) }));
  });
  const end = 6 + r.tax.length;
  put(ws, `B${end}`, "合計", { font: font(10, true), fill: C.mist, align: AL.c });
  ["C", "D"].forEach((c) => put(ws, `${c}${end}`, null, { fill: C.mist }));
  put(ws, `E${end}`, r.cost.taxTotal, { numFmt: NUM, align: AL.r, font: font(11, true, C.deep), fill: C.mist });
  put(ws, `F${end}`, null, { fill: C.mist });
  note(ws, end + 2, 2, 6, `毎年1月1日時点の所有者が納めます。リース期間中の所有者はリース会社なので、この税金はリース料に含まれています。税率は${(r.input.taxRate * 100).toFixed(1)}%、課税年数は${r.taxYears}年で置いています。`);
  note(ws, end + 3, 2, 6, "買った場合は自社で納めます。会社全体の償却資産が少なく、課税標準額の合計が150万円未満なら、免税点により課税されないことがあります（このシートでは考慮していません）。");
  note(ws, end + 4, 2, 6, FOOT);
}

/* ------------------------------------------------------------ ⑤⑥ 比べ方の表 */
function compareSheet(wb, r, name, title, keys, lead, diffs) {
  const years = Math.min(12, Math.max(1, ...r.cmp.filter((x) => x.payCount > 0 || x.dep > 0 || x.tax > 0).map((x) => x.year)));
  const widths = [2, 22, ...Array(years).fill(11.5), 13, 13, 2];
  const ws = sheet(wb, name, title, r, widths, { landscape: true });
  const lastCol = widths.length - 1;
  note(ws, 5, 2, lastCol, lead, 9.5);
  let row = 7;
  ws.getRow(row).height = 22;
  put(ws, `B${row}`, "区分", { font: font(9.5, true, C.white), fill: C.deep, align: AL.c });
  for (let y = 1; y <= years; y++) put(ws, `${colL(2 + y)}${row}`, `${y}年目`, { font: font(9.5, true, C.white), fill: C.deep, align: AL.c });
  put(ws, `${colL(3 + years)}${row}`, "合計", { font: font(9.5, true, C.white), fill: C.deep, align: AL.c });
  put(ws, `${colL(4 + years)}${row}`, "現在価値", { font: font(9.5, true, C.white), fill: C.deep, align: AL.c });
  row++;
  const rowsOf = {
    A: [["リース料の支払", (x) => x.A.pay], ["節税の効果", (x) => x.A.save], ["実質負担", (x) => x.A.net, true]],
    B: [["物件代金の支払", (x) => x.B.pay], ["償却資産税・保険", (x) => x.B.taxIns], ["節税の効果", (x) => x.B.save], ["実質負担", (x) => x.B.net, true]],
    C: [["返済（元金＋利息）", (x) => x.C.pay], ["うち支払利息", (x) => x.C.interest], ["償却資産税・保険", (x) => x.C.taxIns], ["節税の効果", (x) => x.C.save], ["実質負担", (x) => x.C.net, true]],
    D: [["割賦の支払（元金＋手数料）", (x) => x.D.pay], ["うち割賦手数料", (x) => x.D.interest], ["償却資産税・保険", (x) => x.D.taxIns], ["節税の効果", (x) => x.D.save], ["実質負担", (x) => x.D.net, true]],
  };
  for (const k of keys) {
    band(ws, row, 2, lastCol, `${METHOD[k]}`, { font: font(10.5, true, C.deep), fill: C.mist, border: false });
    row++;
    for (const [label, f, total] of rowsOf[k]) {
      const bg = total ? C.mistL : C.white;
      put(ws, `B${row}`, label, { font: font(9.5, !!total), fill: bg });
      let sum = 0;
      for (let y = 1; y <= years; y++) {
        const v = f(r.cmp[y - 1]); sum += v;
        put(ws, `${colL(2 + y)}${row}`, v, { numFmt: NUM, align: AL.r, fill: bg, font: font(9.5, !!total) });
      }
      put(ws, `${colL(3 + years)}${row}`, sum, { numFmt: NUM, align: AL.r, fill: bg, font: font(10, true, total ? C.deep : C.ink) });
      put(ws, `${colL(4 + years)}${row}`, total ? Math.round(r.cmpNpv[k]) : null, { numFmt: NUM, align: AL.r, fill: bg, font: font(10, true, C.deep) });
      row++;
    }
    row++;
  }
  const best = keys.reduce((m, k) => (r.cmpTotal[k] < r.cmpTotal[m] ? k : m), keys[0]);
  const bestNpv = keys.reduce((m, k) => (r.cmpNpv[k] < r.cmpNpv[m] ? k : m), keys[0]);
  band(ws, row, 2, lastCol, `実質負担の合計がいちばん小さいのは「${METHOD[best]}」、お金の時間価値まで含めた現在価値でいちばん小さいのは「${METHOD[bestNpv]}」です。`,
    { font: font(10.5, true, C.plum), fill: C.plumL, border: false });
  ws.getRow(row).height = 26;
  row += 2;
  const a = r.input;
  note(ws, row++, 2, lastCol, `前提：法人実効税率 ${(a.corpTax * 100).toFixed(2)}%、銀行借入の金利 ${(a.loanRate * 100).toFixed(2)}%、割賦の金利 ${(r.kappuRate * 100).toFixed(2)}%、割引率 ${(a.discRate * 100).toFixed(2)}%。節税の効果はマイナスで表示しています。満了後の物件の価値（売却や継続利用）は含めていません。`);
  if (diffs) {
    row++;
    section(ws, row++, 2, lastCol, "リースと割賦の、お金以外の違い（一般的な例。契約により異なります）");
    for (const [label, lease, kappu] of diffs) {
      const span = Math.floor((lastCol - 2) / 2);
      put(ws, `B${row}`, label, { font: font(9.5, true), fill: C.paper, align: AL.w });
      ws.mergeCells(`${colL(3)}${row}:${colL(2 + span)}${row}`);
      put(ws, `C${row}`, `リース：${lease}`, { font: font(9.5), align: AL.w });
      ws.mergeCells(`${colL(3 + span)}${row}:${colL(lastCol)}${row}`);
      put(ws, `${colL(3 + span)}${row}`, `割賦：${kappu}`, { font: font(9.5), align: AL.w });
      ws.getRow(row).height = 34;
      row++;
    }
  }
  row++;
  note(ws, row, 2, lastCol, FOOT);
}

/* ------------------------------------------------------------ ⑦ 減価償却費 */
function depSheet(wb, r) {
  const ws = sheet(wb, "減価償却費", "減価償却費（買った場合）", r, [2, 10, 18, 16, 18, 18, 2]);
  const head = ["年度", "期首の帳簿価額", "減価償却費", "期末の帳簿価額", "償却費の累計"];
  ws.getRow(5).height = 24;
  head.forEach((h, i) => put(ws, `${colL(i + 2)}5`, h, { font: font(9.5, true, C.white), fill: C.deep, align: AL.c }));
  r.dep.rows.forEach((d, i) => {
    const row = 6 + i, bg = i % 2 ? C.mistL : C.white;
    [`${d.year}年目`, d.open, d.amount, d.close, d.cum].forEach((v, k) =>
      put(ws, `${colL(k + 2)}${row}`, v, { numFmt: k ? NUM : undefined, align: k ? AL.r : AL.c, fill: bg, font: font(10, k === 2) }));
  });
  const end = 6 + r.dep.rows.length;
  note(ws, end + 1, 2, 6, "現金や銀行借入・割賦で買った場合の、税務上の減価償却費です。200%定率法で、取得した年も1年分として計算しています（月割りはしていません）。最後は帳簿に1円を残します。");
  note(ws, end + 2, 2, 6, FOOT);
}

/* ------------------------------------------------------------ まとめて作る */
export async function buildWorkbook(r, figs = [], opts = {}) {
  F = opts.font || "Meiryo UI";
  const ExcelJS = await getExcelJS();
  const wb = new ExcelJS.Workbook();
  wb.creator = "数字のものさし";
  wb.title = "リース料金ポン！ 見積の分析";
  wb.created = new Date();
  dashboard(wb, r, figs);
  costSheet(wb, r);
  scheduleSheet(wb, r);
  taxSheet(wb, r);
  compareSheet(wb, r, "現金・借入との比較", "現金・銀行借入との比較", ["A", "B", "C"],
    "同じ物件を、リースで使う場合と、現金で買う場合・銀行から借りて買う場合を、税金の効果まで含めた「実質負担」で比べます。表面の支払総額だけで決めると、答えを間違えることがあります。");
  compareSheet(wb, r, "割賦との比較", "割賦との比較", ["A", "D"],
    "割賦は、分割払いで物件を買う方法です。自分の物になるので、償却資産税や保険は自社で払い、減価償却と割賦手数料が経費になります。リースとの違いを、実質負担で比べます。",
    [["物件の持ち主", "リース会社", "自社（代金を払い終えるまで、所有権は売り手に残すことが多い）"],
     ["税金・保険", "リース会社が払い、月額に含まれる", "自社で払う"],
     ["満了・完済のあと", "返却・再リース・買い取りから選ぶ", "そのまま自社の物"],
     ["途中でやめるとき", "原則として解約できず、残りのリース料に相当する額を払う", "残金を一括で払うのが一般的"]]);
  depSheet(wb, r);
  return wb;
}

export async function downloadLeaseXlsx(r, figs, filename) {
  const wb = await buildWorkbook(r, figs);
  const buf = await wb.xlsx.writeBuffer();
  const name = filename || `リース料金ポン_見積の分析_${r.input.price}円_${r.input.months}か月.xlsx`;
  const url = URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  const a = document.createElement("a");
  a.href = url; a.download = name; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}
