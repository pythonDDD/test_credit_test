/* ============================================================================
 * /credit-pro/viz.js — 判定結果の見える化
 *
 * engine.js の evaluate() が返す結果オブジェクトだけを読んで SVG を組み立てる。
 * 計算は一切やり直さない（Excel版 Pro と同じ数字がそのまま図になる）。
 * 外部ライブラリなし。app.js から renderViz(r, f) を呼ぶ。
 * ========================================================================== */

/* サイトのパレット（index.html の :root と同じ値） */
const C = {
  seaDeep: "#33526C", sea: "#527695", sky: "#84D2F5", aqua: "#B0F1F0", mist: "#DFF6F1",
  ink: "#0F1A22", soft: "#4A5A66", faint: "#7A8A95",
  warn: "#B5623F", amber: "#E0A94A", peridot: "#82B33A", green: "#4E8F2E",
  rose: "#F4C3D3", bloom: "#F8A3BF", deep: "#8C3B22",
  track: "#E1E9ED", grid: "#E2E9EC", axis: "#9FB0BC",
};

/* ------------------------------------------------------------ 小道具 */
const esc = (s) => String(s ?? "").replace(/[&<>"]/g,
  (m) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[m]));
const n1 = (x) => (isFinite(x) ? x.toFixed(1) : "—");
const n2 = (x) => (isFinite(x) ? x.toFixed(2) : "—");
const p1 = (x) => (isFinite(x) ? (x * 100).toFixed(1) + "%" : "—");
const red3 = (r) => r.redemption.required <= 0 ? "実質無借金"
  : r.redemption.simpleCF <= 0 ? "返済原資なし" : n1(r.redemption.years) + "年";
const clip = (t, n) => (String(t).length > n ? String(t).slice(0, n - 1) + "…" : String(t));

function T(x, y, t, o = {}) {
  const halo = o.halo
    ? `<text x="${x}" y="${y}" font-size="${o.s || 11}" ${o.a ? `text-anchor="${o.a}"` : ""}
        ${o.w ? `font-weight="${o.w}"` : ""} fill="none" stroke="#fff" stroke-width="3.4"
        stroke-linejoin="round">${esc(t)}</text>` : "";
  return halo + `<text x="${x}" y="${y}" font-size="${o.s || 11}" fill="${o.c || C.soft}"
    ${o.a ? `text-anchor="${o.a}"` : ""} ${o.w ? `font-weight="${o.w}"` : ""}>${esc(t)}</text>`;
}
const R = (x, y, w, h, f, o = {}) =>
  (w > 0 && h > 0)
    ? `<rect x="${x}" y="${y}" width="${w}" height="${h}" fill="${f}"${o.rx ? ` rx="${o.rx}"` : ""}${o.st ? ` stroke="${o.st}" stroke-width="${o.sw || 1}"` : ""}/>`
    : "";
const L = (x1, y1, x2, y2, c, w = 1, d) =>
  `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${c}" stroke-width="${w}"${d ? ` stroke-dasharray="${d}"` : ""}/>`;

/** 目盛りをきりのいい数字にする */
function niceStep(x) {
  if (!(x > 0)) return 1;
  const e = Math.pow(10, Math.floor(Math.log10(x))), m = x / e;
  for (const c of [1, 1.2, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10]) if (m <= c + 1e-9) return c * e;
  return 10 * e;
}
/** lo〜hi を4区間のきりのいい目盛りに収める（0は必ず目盛り線に乗る） */
function axis4(lo, hi) {
  if (hi <= lo) hi = lo + 1;
  let step = niceStep((hi - lo) / 4), a, b, guard = 0;
  do {
    a = Math.floor(lo / step) * step; b = a + step * 4;
    if (b >= hi - 1e-9) break;
    step = niceStep(step * 1.05);
  } while (++guard < 20);
  return { lo: a, hi: b, ticks: [0, 1, 2, 3, 4].map((i) => a + step * i) };
}
/** ツールチップ付きの当たり判定。title|行|行… の形で渡す */
const tip = (inner, lines) =>
  `<g class="viz__hot" data-tip="${esc(lines.filter(Boolean).join("|"))}">${inner}</g>`;

const svg = (vb, body) =>
  `<svg class="viz__svg" viewBox="${vb}" role="img" xmlns="http://www.w3.org/2000/svg">${body}</svg>`;

/* ======================================================= ① スコアの内訳 */
function scoreBars(r) {
  const s = r.scores, kd = s.kibo_detail, sd = s.soneki_detail, ke = s.keiei_detail;
  const rows = [
    ["① 業歴", s.gyoreki, 10, `創業から${r.businessYears}年`, "長く続いていること自体が、環境変化を乗り越えてきた証拠になります。"],
    ["② 資本構成", s.shihon, 12, `自己資本比率 ${p1(r.equityRatio)}／業種基準 ${p1(r.benchmark.equityRatio)}（倍率 ${n2(r.equityMultiple)}）`,
      "業種と資本金規模で補正した基準の何倍かで採点します。"],
    ["③ 規模", s.kibo, 18, `業容${kd.gyoyo}点＋年商${kd.nensho}点＋上場区分${kd.listing}点＋従業員${kd.employees}点`,
      "規模そのものが信用の裏づけになるため、配点は18点と重めです。"],
    ["④ 損益", s.soneki, 10, `${sd.pattern}（基礎${sd.base}点）＋加点${sd.bonus}点`,
      "黒字がどれだけ続いているかを見ます。利益額による加点もあります。"],
    ["⑤ 経営者", s.keiei, 20, `業界歴${ke.industry}点＋経営者歴${ke.ceo}点＋持ち家${ke.home}点＋開示姿勢${ke.disclosure}点`,
      "数字に出ない部分。決算書の開示姿勢だけで14点を置いています。"],
    ["⑥ 償還余力", s.shokan, 30, `債務償還年数${red3(r)}→${r.redemption.scoreA}点／3年返済充足率${n2(r.redemption.ratio)}倍→${r.redemption.scoreB}点`,
      "返す力。100点のうち最も重い30点を割り当てています。"],
  ];
  const MAX = 30, bx = 108, bw = 250, rowH = 38, top = 14;
  let h = "";
  rows.forEach(([name, got, max, basis, why], i) => {
    const y = top + i * rowH, track = bw * max / MAX, ratio = max ? got / max : 0;
    const col = ratio >= 0.8 ? C.green : ratio >= 0.6 ? C.peridot
      : ratio >= 0.4 ? C.sea : ratio >= 0.2 ? C.warn : C.deep;
    const body = T(0, y + 18, name, { s: 12.5, w: "700", c: C.ink })
      + R(bx, y + 5, track, 17, C.track, { rx: 5 })
      + R(bx, y + 5, Math.max(track * ratio, ratio > 0 ? 2 : 0), 17, col, { rx: 5 })
      + T(bx + track + 9, y + 18, `${got} / ${max}`, { s: 11.5, w: "700", c: C.ink })
      + T(0, y + 31, `${Math.round(ratio * 100)}%を取得`, { s: 9.5, c: C.faint })
      + R(0, y, 420, rowH - 2, "transparent");
    h += tip(body, [`${name}　${got} / ${max}点`, basis, why]);
  });
  const y = top + rows.length * rowH + 4;
  h += L(0, y, 420, y, "#C9D3D8", 1);
  h += T(0, y + 19, "合計", { s: 13.5, w: "700", c: C.ink });
  h += T(bx + bw + 9, y + 19, `${s.total} / 100`, { s: 13.5, w: "700", c: C.ink });
  h += T(0, y + 33, "バーの長さは配点の大きさ（30点満点＝いちばん長い）", { s: 9.5, c: C.faint });
  return svg("0 0 420 300", h);
}

/* ================================================= ② 財務指標のかたち */
const RADAR = [
  { tip: "総資本のうち、返さなくてよいお金の割合。高いほど財務は安全です。", n: "自己資本比率", u: "%", dec: 1, dir: 1, star: true,
    get: (r) => r.ratios.periods[0].equityRatio * 100, bm: (r) => r.benchmark.equityRatio * 100 },
  { tip: "売上100円あたり、いくら経常利益が残るか。本業と財務を合わせた稼ぐ力です。", n: "経常利益率", u: "%", dec: 1, dir: 1, star: true,
    get: (r) => r.ratios.periods[0].ordinaryMargin * 100, bm: (r) => r.benchmark.ordinaryMarginAvg3 * 100 },
  { tip: "本業だけで売上100円あたりいくら稼げるか。原価と販管費の効率を映します。", n: "営業利益率", u: "%", dec: 1, dir: 1, star: true,
    get: (r) => r.ratios.periods[0].operatingMargin * 100, bm: (r) => r.benchmark.operatingMarginLatest * 100 },
  { tip: "投じた資本を1年で何回売上に変えたか。高いほど資産を効率よく使っています。", n: "総資本回転率", u: "回", dec: 2, dir: 1, star: false,
    get: (r) => r.ratios.periods[0].assetTurnover, bm: () => 1.0 },
  { tip: "1年以内に返す負債に対し、1年以内に現金化できる資産がどれだけあるか。", n: "流動比率", u: "%", dec: 0, dir: 1, star: false,
    get: (r) => r.ratios.periods[0].currentRatio * 100, bm: () => 120 },
  { tip: "月商の何か月分の借入があるか。小さいほど身軽です。", n: "借入金月商倍率", u: "か月", dec: 2, dir: -1, star: false,
    get: (r) => r.ratios.periods[0].gearingMonths, bm: () => 6.0 },
];
const norm = (v, b, dir) => {
  if (!isFinite(v)) return 3;
  if (dir > 0) return b > 0 ? Math.max(3, Math.min(100, v / b * 50)) : 3;
  if (v <= 0) return 100;
  return b > 0 ? Math.max(3, Math.min(100, b / v * 50)) : 3;
};

function radar(r) {
  if (!(r.cur.totalCapital > 0) || !(r.cur.sales > 0))
    return svg("0 0 420 330", R(0, 0, 420, 330, "#F3F6F7", { rx: 10 })
      + T(210, 168, "決算書の数字を入れるとチャートが出ます", { a: "middle", s: 12, c: C.faint }));
  const cx = 210, cy = 148, RR = 95;
  const vals = RADAR.map((ax) => {
    const v = ax.get(r), b = ax.bm(r);
    return { ax, v, b, s: norm(v, b, ax.dir) };
  });
  let h = "";
  [25, 50, 75, 100].forEach((k) => {
    const p = vals.map((_, i) => {
      const a = (i * 60 - 90) * Math.PI / 180;
      return `${cx + Math.cos(a) * RR * k / 100},${cy + Math.sin(a) * RR * k / 100}`;
    });
    h += `<polygon points="${p.join(" ")}" fill="none" stroke="${k === 100 ? "#B4C4CC" : "#DCE5E9"}" stroke-width="1"/>`;
  });
  for (let i = 0; i < 6; i++) {
    const a = (i * 60 - 90) * Math.PI / 180;
    h += L(cx, cy, cx + Math.cos(a) * RR, cy + Math.sin(a) * RR, "#DCE5E9", 1);
  }
  const base = vals.map((_, i) => {
    const a = (i * 60 - 90) * Math.PI / 180;
    return `${cx + Math.cos(a) * RR * 0.5},${cy + Math.sin(a) * RR * 0.5}`;
  });
  const mine = vals.map((o, i) => {
    const a = (i * 60 - 90) * Math.PI / 180;
    return `${cx + Math.cos(a) * RR * o.s / 100},${cy + Math.sin(a) * RR * o.s / 100}`;
  });
  h += `<polygon points="${base.join(" ")}" fill="none" stroke="#5A6B76" stroke-width="1.6" stroke-dasharray="5 4"/>`;
  h += `<polygon points="${mine.join(" ")}" fill="${C.sea}" fill-opacity="0.26" stroke="${C.sea}" stroke-width="2.4"/>`;
  mine.forEach((p, i) => {
    const [x, y] = p.split(","), o = vals[i];
    const v = isFinite(o.v) ? o.v.toFixed(o.ax.dec) + o.ax.u : "—";
    h += tip(`<circle cx="${x}" cy="${y}" r="3.4" fill="${C.seaDeep}"/><circle cx="${x}" cy="${y}" r="13" fill="transparent"/>`,
      [o.ax.n, `実績 ${v}　／　基準 ${o.b.toFixed(o.ax.dec)}${o.ax.u}`,
       o.ax.dir > 0 ? (o.v >= o.b ? "基準を上回っています。" : "基準を下回っています。")
                    : (o.v <= o.b ? "目安の範囲に収まっています。" : "目安を超えています。"),
       o.ax.tip]);
  });
  vals.forEach((o, i) => {
    const a = (i * 60 - 90) * Math.PI / 180;
    const lx = cx + Math.cos(a) * (RR + 24), ly = cy + Math.sin(a) * (RR + 24);
    const an = Math.cos(a) > 0.2 ? "start" : Math.cos(a) < -0.2 ? "end" : "middle";
    h += T(lx, ly, o.ax.n + (o.ax.star ? " ★" : ""), { a: an, s: 10.5, w: "700", c: C.ink });
    h += T(lx, ly + 13, (isFinite(o.v) ? o.v.toFixed(o.ax.dec) : "—") + o.ax.u, { a: an, s: 10.5, w: "700", c: C.sea });
    if (o.ax.dir < 0) h += T(lx, ly + 24, "小さいほど良い", { a: an, s: 8.5, c: C.faint });
  });
  h += T(210, 318, "点線＝基準。外にはみ出していれば基準より良い", { a: "middle", s: 10, c: C.faint });
  return svg("0 0 420 330", h);
}

function radarTable(r, f) {
  const rows = RADAR.map((ax) => {
    const v = ax.get(r), b = ax.bm(r);
    const ok = ax.dir > 0 ? v >= b : v <= b;
    const mark = ok ? "○" : "△";
    const word = ax.star ? (ok ? "基準以上" : "基準未満") : (ok ? (ax.dir > 0 ? "目安以上" : "目安以内") : (ax.dir > 0 ? "目安未満" : "目安超"));
    return `<tr><th>${esc(ax.n)}${ax.star ? ' <span class="viz__star">★</span>' : ""}</th>
      <td>${isFinite(v) ? v.toFixed(ax.dec) + ax.u : "—"}</td>
      <td class="viz__bm">${b.toFixed(ax.dec)}${ax.u}</td>
      <td class="${ok ? "viz__ok" : "viz__ng"}">${mark} ${word}</td></tr>`;
  }).join("");
  return `<table class="viz__table"><thead><tr><th>指　標</th><th>実　績</th><th>基準・目安</th><th>判定</th></tr></thead>
    <tbody>${rows}</tbody></table>
    <p class="viz__note">★＝財務省 法人企業統計の該当業種の値（${esc(r.input.industry.trim())}／${esc(r.input.capitalTier)}）。
    ★のない3つは実務上の一般的な目安（総資本回転率1.0回・流動比率120%・借入金月商倍率6か月）です。</p>`;
}

/* ============================================= ③ 貸借対照表のかたち */
function bsBlock(r, f) {
  const cur = r.cur;
  const assets = cur.totalAssets, liabEq = cur.totalCapital;
  if (!(assets > 0))
    return svg("0 0 420 300", R(0, 0, 420, 300, "#F3F6F7", { rx: 10 })
      + T(210, 154, "貸借対照表を入力すると図が出ます", { a: "middle", s: 12, c: C.faint }));
  const top = 44, bot = 272, H = bot - top, lx = 46, rx = 214, cw = 118;
  const scale = Math.max(assets, liabEq);
  const px = (a) => H * a / scale;
  let h = T(lx + cw / 2, 30, "資産（持ち物）", { a: "middle", s: 11.5, w: "700", c: C.ink })
        + T(rx + cw / 2, 30, "負債・純資産（お金の出どころ）", { a: "middle", s: 11.5, w: "700", c: C.ink });
  const stack = (x, items) => {
    let y = top, out = "";
    items.forEach((it) => {
      const hh = px(it.a);
      if (hh <= 0) return;
      let blk = R(x, y, cw, hh, it.c, { st: "#FFFFFF", sw: 1 });
      if (hh >= 19) {
        blk += T(x + 8, y + hh / 2 - 2, it.n, { s: 10, w: "700", c: it.light ? "#FFFFFF" : C.ink });
        blk += T(x + cw - 8, y + hh / 2 + 11, `${f.yenU(it.a)}（${Math.round(it.a / scale * 100)}%）`,
          { a: "end", s: 9.5, c: it.light ? "#FFFFFF" : C.soft });
      }
      out += tip(blk, [it.n, `${f.yenU(it.a)} ${f.U_LABEL()}　／　総資産の ${Math.round(it.a / scale * 100)}%`, it.tip]);
      y += hh;
    });
    return out;
  };
  h += stack(lx, [
    { n: "現金・預金", a: cur.cash, c: C.sky, tip: "すぐ使えるお金。月商の何か月分あるかが「手元流動性」です。" },
    { n: "その他の流動資産", a: Math.max(cur.currentAssets - cur.cash, 0), c: C.aqua,
      tip: "売掛金・受取手形・棚卸資産など、1年以内に現金化する見込みの資産です。" },
    { n: "固定資産", a: cur.fixedAssets + cur.deferred, c: C.mist,
      tip: "建物・機械・投資有価証券など、長く使う資産。多いほど資本が寝ています。" },
  ]);
  h += stack(rx, [
    { n: "有利子負債", a: cur.interestBearingDebt, c: C.sea, light: true,
      tip: "利息を払って借りているお金（短期借入金＋長期借入金・社債）。債務償還年数の分子になります。" },
    { n: "その他の負債", a: Math.max(cur.totalLiab - cur.interestBearingDebt, 0), c: "#CBDEE7",
      tip: "買掛金・未払金など、利息のつかない負債です。" },
    { n: "純資産", a: Math.max(cur.equity, 0), c: C.peridot, light: true,
      tip: "返さなくていいお金。総資産に占めるこの厚みが自己資本比率です。" },
  ]);
  if (cur.equity < 0) {
    const yA = top + px(assets), yL = top + px(cur.totalLiab);
    h += L(lx, yA, rx + cw, yA, C.deep, 1.6, "6 4");
    h += R(rx, yA, cw, yL - yA, C.deep, { rx: 0 });
    h += T(rx + cw / 2, (yA + yL) / 2 + 4, "債務超過", { a: "middle", s: 11, w: "700", c: "#FFFFFF" });
  }
  h += T(210, 292, "左右の高さは必ず同じ。右下の緑が厚いほど、返さなくていいお金で買っている",
    { a: "middle", s: 9.5, c: C.faint });
  return svg("0 0 420 300", h);
}

/* ============================================== ④ 売上と利益の推移 */
function trend(r, f) {
  const S = [r.prev2.sales, r.prev.sales, r.cur.sales];
  const O = [r.prev2.ordinaryProfit, r.prev.ordinaryProfit, r.cur.ordinaryProfit];
  const labels = [r.input.terms[2] || "前々期", r.input.terms[1] || "前期", r.input.terms[0] || "今期"];
  if (S.every((x) => x <= 0))
    return svg("0 0 420 280", R(0, 0, 420, 280, "#F3F6F7", { rx: 10 })
      + T(210, 144, "売上高を入れると推移が出ます", { a: "middle", s: 12, c: C.faint }));
  const Lm = 54, Rm = 54, Tm = 26, Bm = 46, W = 420, Hh = 280;
  const x0 = Lm, x1 = W - Rm, y0 = Tm, y1 = Hh - Bm, pw = x1 - x0, ph = y1 - y0;
  const sA = axis4(0, Math.max(...S)), oA = axis4(Math.min(0, ...O), Math.max(0, ...O));
  const YS = (v) => y1 - ph * (v - sA.lo) / (sA.hi - sA.lo);
  const YO = (v) => y1 - ph * (v - oA.lo) / (oA.hi - oA.lo);
  let h = "";
  for (let g = 0; g <= 4; g++) {
    const gy = y0 + ph * g / 4, k = 4 - g;
    h += L(x0, gy, x1, gy, C.grid, 1);
    h += T(x0 - 7, gy + 4, f.yenU(sA.ticks[k]), { a: "end", s: 9, c: C.faint });
    h += T(x1 + 7, gy + 4, f.yenU(oA.ticks[k]), { s: 9, c: C.warn });
  }
  if (oA.lo < 0) h += L(x0, YO(0), x1, YO(0), C.warn, 1.4, "4 3");
  const bw = pw / 3 * 0.40, pts = [];
  S.forEach((s, i) => {
    const cx = x0 + pw * (i + 0.5) / 3;
    if (s > 0) {
      h += tip(R(cx - bw / 2, YS(s), bw, y1 - YS(s), C.sky, { rx: 3 }),
        [`${labels[i]}　売上高`, `${f.yenU(s)} ${f.U_LABEL()}`,
         i > 0 ? (S[i] >= S[i - 1] ? `前期比 +${f.yenU(s - S[i - 1])}（増収）` : `前期比 ${f.yenU(s - S[i - 1])}（減収）`) : "",
         "棒の高さより、折れ線の向きを先に見てください。"]);
      // 経常利益の点が棒より上にあるときは、棒のラベルを棒の中に入れて衝突を避ける
      const inside = YO(O[i]) < YS(s) + 10;
      h += T(cx, YS(s) + (inside ? 16 : -6), f.yenU(s),
        { a: "middle", s: 9.5, w: "700", c: inside ? C.seaDeep : C.seaDeep, halo: !inside });
    }
    pts.push(`${cx},${YO(O[i])}`);
    h += T(cx, y1 + 16, clip(labels[i], 9), { a: "middle", s: 10, w: "700", c: C.ink });
  });
  h += `<polyline points="${pts.join(" ")}" fill="none" stroke="${C.warn}" stroke-width="2.6"/>`;
  O.forEach((o, i) => {
    const cx = x0 + pw * (i + 0.5) / 3, yy = YO(o);
    h += tip(`<circle cx="${cx}" cy="${yy}" r="4.2" fill="${C.warn}" stroke="#fff" stroke-width="1.4"/>`
      + `<circle cx="${cx}" cy="${yy}" r="13" fill="transparent"/>`,
      [`${labels[i]}　経常利益`, `${f.yenU(o)} ${f.U_LABEL()}`,
       S[i] > 0 ? `売上高経常利益率 ${(o / S[i] * 100).toFixed(1)}%` : "",
       "本業の利益に、受取利息や支払利息などを加減したもの。会社の総合的な稼ぐ力です。"]);
    const onBar = S[i] > 0 && yy > YS(S[i]);
    h += T(cx, yy + (onBar ? 18 : -10), f.yenU(o), { a: "middle", s: 9.5, w: "700", c: C.warn, halo: 1 });
  });
  h += T(x0 - 7, y0 - 9, "売上高", { a: "end", s: 9.5, w: "700", c: C.seaDeep });
  h += T(x1 + 7, y0 - 9, "経常利益", { s: 9.5, w: "700", c: C.warn });
  h += T(210, Hh - 8, `単位：${f.U_LABEL()}`, { a: "middle", s: 9.5, c: C.faint });
  return svg("0 0 420 280", h);
}

/* ========================================== ⑤ 返せるお金と、返す額 */
function repay(r, f) {
  const cf = r.redemption.simpleCF;
  const rp = [0, 1, 2].map((i) => Number(r.input.repayment[i]) || 0);
  if (cf === 0 && rp.every((x) => x === 0))
    return svg("0 0 420 280", R(0, 0, 420, 280, "#F3F6F7", { rx: 10 })
      + T(210, 144, "返済計画を入れると図が出ます", { a: "middle", s: 12, c: C.faint }));
  const Lm = 58, Rm = 18, Tm = 26, Bm = 52, W = 420, Hh = 280;
  const x0 = Lm, x1 = W - Rm, y0 = Tm, y1 = Hh - Bm, pw = x1 - x0, ph = y1 - y0;
  const all = [cf, ...rp];
  const A = axis4(Math.min(0, ...all), Math.max(0, ...all));
  const Y = (v) => y1 - ph * (v - A.lo) / (A.hi - A.lo);
  let h = "";
  for (let g = 0; g <= 4; g++) {
    const gy = y0 + ph * g / 4;
    h += L(x0, gy, x1, gy, C.grid, 1);
    h += T(x0 - 7, gy + 4, f.yenU(A.ticks[4 - g]), { a: "end", s: 9, c: C.faint });
  }
  if (A.lo < 0) h += L(x0, Y(0), x1, Y(0), C.deep, 1.4, "4 3");
  const gw = pw / 3, bw = gw * 0.28;
  for (let i = 0; i < 3; i++) {
    const gx = x0 + gw * i + gw / 2, a = cf, b = rp[i];
    h += tip(R(gx - bw - 3, Math.min(Y(a), Y(0)), bw, Math.abs(Y(a) - Y(0)), C.sea, { rx: 3 }),
      [`${i + 1}年目　簡易キャッシュフロー`, `${f.yenU(a)} ${f.U_LABEL()}`,
       "当期純利益＋減価償却費。1年で手元に残るお金の目安です。",
       "減価償却費は、費用として引かれているのにお金が出ていかないため足し戻します。"]);
    h += tip(R(gx + 3, Math.min(Y(b), Y(0)), bw, Math.abs(Y(b) - Y(0)), C.rose, { rx: 3 }),
      [`${i + 1}年目　約定返済額`, `${f.yenU(b)} ${f.U_LABEL()}`,
       a - b >= 0 ? `返済後に ${f.yenU(a - b)} 残る見込みです。` : `${f.yenU(b - a)} 不足します。借換えか手元資金の取り崩しが前提になります。`]);
    if (a !== 0) h += T(gx - bw / 2 - 3, Y(a) + (a < 0 ? 12 : -6), f.yenU(a), { a: "middle", s: 9, w: "700", c: C.seaDeep });
    if (b !== 0) h += T(gx + bw / 2 + 3, Y(b) - 6, f.yenU(b), { a: "middle", s: 9, w: "700", c: C.warn });
    h += T(gx, y1 + 16, `${i + 1}年目`, { a: "middle", s: 10, w: "700", c: C.ink });
    const sur = a - b;
    if (a !== 0 || b !== 0)
      h += T(gx, y1 + 30, sur >= 0 ? `余力 +${f.yenU(sur)}` : `不足 ${f.yenU(-sur)}`,
        { a: "middle", s: 9.5, w: "700", c: sur >= 0 ? C.green : C.deep });
  }
  h += T(210, Hh - 5, `単位：${f.U_LABEL()}（簡易CF＝当期純利益＋減価償却費。3年とも直近期と同水準で見込む）`,
    { a: "middle", s: 9, c: C.faint });
  return svg("0 0 420 280", h);
}

/* ====================================== ⑥ 借金を返し切るまでの年数 */
function gauge(r) {
  const red = r.redemption;
  const x0 = 34, x1 = 386, y = 64, bh = 24, MAX = 25;
  const X = (yr) => x0 + (x1 - x0) * Math.max(0, Math.min(MAX, yr)) / MAX;
  const WHY = "⑤は「今年の返済に、今年の稼ぎが足りるか」。⑥は「いまの借金を、いまの稼ぎで何年かかって返し切れるか」。別々に採点します。";
  const WHY2 = "⑤が足りていても、借金の絶対額が大きければ⑥は長くなります。逆に⑥が短くても、今年の返済が集中していれば⑤は不足します。";
  let h = tip(R(x0, y, X(10) - x0, bh, C.peridot, { rx: 0 }),
      ["健全（10年以内）", "有利子負債の返済能力に問題は見られない水準です。",
       "金融機関が「正常先」と見る一つの目安が10年以内です。", WHY])
    + tip(R(X(10), y, X(20) - X(10), bh, C.amber),
      ["要注意（10〜20年）", "返し切るのに時間がかかりすぎている水準です。",
       "設備投資が重い業種では長くなりやすいので、同業と比べて判断します。", WHY])
    + tip(R(X(20), y, x1 - X(20), bh, C.warn),
      ["厳しい（20年超）", "いまの稼ぐペースでは返し切るのが難しい水準です。",
       "借換えが続くことが前提になり、金融環境が変わると資金繰りに直結します。", WHY]);
  [0, 5, 10, 15, 20, 25].forEach((t) => {
    h += L(X(t), y + bh, X(t), y + bh + 5, C.faint, 1);
    h += T(X(t), y + bh + 17, t === 25 ? "25年〜" : `${t}年`, { a: "middle", s: 9.5, c: C.faint });
  });
  h += T(x0 + (X(10) - x0) / 2, y + 16, "健全", { a: "middle", s: 11, w: "700", c: "#fff" });
  h += T((X(10) + X(20)) / 2, y + 16, "要注意", { a: "middle", s: 11, w: "700", c: "#fff" });
  h += T((X(20) + x1) / 2, y + 16, "厳しい", { a: "middle", s: 11, w: "700", c: "#fff" });
  if (red.required <= 0) {
    h += tip(T(210, 26, "実質無借金", { a: "middle", s: 18, w: "700", c: C.green })
      + `<polygon points="${x0},${y - 2} ${x0 - 8},${y - 15} ${x0 + 8},${y - 15}" fill="${C.green}"/>`
      + R(120, 8, 180, 26, "transparent"),
      ["実質無借金", "要償還債務がゼロです。",
       "有利子負債から、現預金と正常運転資金（売掛金＋棚卸資産−買掛金）を引いた残りがマイナスという意味です。",
       "手元の現金と、商売に必要な運転資金で借金をまかなえている状態です。", WHY2]);
  } else if (red.simpleCF <= 0) {
    h += tip(T(210, 26, "返済原資なし（簡易CFがマイナス）", { a: "middle", s: 14, w: "700", c: C.deep })
      + R(60, 8, 300, 26, "transparent"),
      ["返済原資なし", "簡易キャッシュフロー（当期純利益＋減価償却費）がマイナスです。",
       "1年間の営業の結果として手元にお金が残っていないため、年数は計算できません。",
       "返済は借換えか資産の取り崩しに頼ることになります。⑥の配点15点は0点になります。", WHY2]);
  } else {
    const px = X(red.years);
    h += tip(`<polygon points="${px},${y - 2} ${px - 8},${y - 15} ${px + 8},${y - 15}" fill="${C.ink}"/>`
      + `<rect x="${px - 16}" y="${y - 18}" width="32" height="${bh + 20}" fill="transparent"/>`,
      ["債務償還年数", `${n1(red.years)} 年`,
       red.years <= 10 ? "10年以内で、健全とされる水準です。"
         : red.years <= 20 ? "10年を超えており、銀行がまず気にする水準です。" : "20年超。いまの稼ぐペースでは返し切るのが難しい水準です。",
       "業種で適正な長さは変わります。設備が重い業種は長く、サービス業は短くなりやすい指標です。"]);
    h += T(210, 26, `${n1(red.years)} 年`, { a: "middle", s: 20, w: "700", c: C.ink });
  }
  h += tip(T(210, y + bh + 34, "要償還債務（有利子負債−現預金−正常運転資金）÷ 簡易キャッシュフロー",
    { a: "middle", s: 9.5, c: C.faint }) + R(60, y + bh + 24, 300, 16, "transparent"),
    ["この年数の作り方",
     `要償還債務 ＝ 有利子負債 − 現預金 − 正常運転資金`,
     "正常運転資金（売掛金＋棚卸資産−買掛金）は、商売を回すのに必ず要るお金なので、返済に充てられない前提で差し引きます。",
     "分母の簡易キャッシュフローは、当期純利益に減価償却費を足し戻したもの。帳簿では引かれているのに、お金は出ていっていないからです。",
     WHY2]);
  return svg("0 0 420 150", h);
}

/* ================================================ グラフの読み取り */
function readings(r, f) {
  const s = r.scores, cur = r.cur, red = r.redemption, bm = r.benchmark, p = r.ratios.periods[0];
  const out = [];
  const axes = [["① 業歴", s.gyoreki, 10], ["② 資本構成", s.shihon, 12], ["③ 規模", s.kibo, 18],
                ["④ 損益", s.soneki, 10], ["⑤ 経営者", s.keiei, 20], ["⑥ 償還余力", s.shokan, 30]];
  const worst = axes.slice().sort((a, b) => a[1] / a[2] - b[1] / b[2])[0];
  out.push(`評点100点のうち、いちばん取りこぼしているのは「${worst[0].replace(/^[①-⑥]\s*/, "")}」です。`
    + `${worst[2]}点満点中${worst[1]}点、達成率${Math.round(worst[1] / worst[2] * 100)}%。①の棒グラフで、薄い部分がいちばん長い行です。`);

  if (cur.totalCapital > 0) {
    const er = p.equityRatio;
    out.push(cur.equity < 0
      ? "純資産がマイナスです。債務超過の状態にあり、与信判断では最も重い事実として扱われます。"
      : `自己資本比率は${p1(er)}。業種×資本金階層で補正した基準${p1(bm.equityRatio)}を`
        + (er >= bm.equityRatio ? "上回っており、②のレーダーでは外側に出ます。" : `${p1(bm.equityRatio - er)}下回っており、②のレーダーでは内側にへこみます。`)
        + `　③のB/S図では、右の柱の緑の厚みが自己資本比率、いちばん上の濃い青が有利子負債（${f.yenU(cur.interestBearingDebt)}${f.U_LABEL()}）です。`);
  }
  if (r.prev.sales > 0 && cur.sales > 0) {
    const up = cur.sales >= r.prev.sales, pu = cur.ordinaryProfit >= r.prev.ordinaryProfit;
    const lab = (up ? "増収" : "減収") + (pu ? "増益" : "減益");
    const note = up
      ? (pu ? "売上も利益も伸びています。" : "売上は伸びているのに、利益は減っています。原価と販管費のどちらが動いたのかを確認したいところです。")
      : (pu ? "売上は減りましたが、利益は増えています。狙って縮めたのか、たまたまかで意味が変わります。" : "売上も利益も減っています。与信判断では最も警戒する形です。");
    out.push(`④の推移：直近期は${lab}です。${note}棒の高さより、折れ線の向きを先に見てください。`);
  }
  if (red.required <= 0) {
    out.push(`⑥の償還余力：要償還債務はゼロ（実質無借金）です。有利子負債${f.yenU(red.interestBearingDebt)}${f.U_LABEL()}に対し、`
      + `現預金${f.yenU(red.cash)}と正常運転資金${f.yenU(red.workingCapital)}の合計が上回っています。`);
  } else if (red.simpleCF <= 0) {
    out.push("⑥の償還余力：簡易キャッシュフロー（当期純利益＋減価償却費）がマイナスです。"
      + "1年間の営業の結果として手元にお金が残っていないため、債務償還年数は計算できません。返済は借換えか資産の取り崩しに頼ることになります。");
  } else {
    out.push(`⑥の償還余力：債務償還年数は${n1(red.years)}年です。`
      + (red.years <= 10 ? "10年以内で、健全とされる水準に収まっています。"
        : red.years <= 15 ? "10年を超えており、注意を要する水準です。" : "15年を超えており、過大とみられる水準です。")
      + `　DSCR（簡易CF÷1年目返済額）は${n2(red.dscr)}倍。1.0倍未満は当年の返済原資が不足します。`);
  }
  if (red.repay3 > 0) {
    out.push(`⑤の返済：今後3年の約定返済額${f.yenU(red.repay3)}${f.U_LABEL()}に対し、簡易キャッシュフローの3年累計見込は${f.yenU(red.cf3)}、`
      + `充足率は${n2(red.ratio)}倍。`
      + (red.ratio >= 1 ? "青が桃色を上回っており、返済原資は自力で賄える見込みです。" : "桃色が青を上回っています。借換えか手元資金の取り崩しが前提になります。"));
  }
  out.push("グラフは入力された数字をそのまま描いています。粉飾、保証債務、簿外債務、経営者の資質は図には映りません。");
  return `<ul class="remarks">${out.map((t) => `<li>${esc(t)}</li>`).join("")}</ul>`;
}

/* ==================================================== 組み立て（公開） */
export function renderViz(r, f) {
  const card = (title, hint, body, extra = "") => `
    <div class="viz__card">
      <h3 class="viz__h">${title}</h3>
      <p class="viz__hint">${hint}</p>
      ${body}${extra}
    </div>`;
  return `
  <div class="calc">
    <h2>見える化</h2>
    <p class="calc__hint">上と同じ判定結果を、図にしたものです。数字の表のままでは伝わらない相手に、形と色で見せるために使ってください。</p>
    <div class="viz">
      ${card("① スコアの内訳",
        "100点をどこで積み、どこで落としたか。バーの長さが配点、濃い部分が取れた点数です。",
        scoreBars(r))}
      ${card("② 財務指標のかたち",
        "点線の六角形が基準です。実線がその外にあれば基準より良い、内側なら基準より悪い。へこんでいる方向が弱点です。",
        radar(r), radarTable(r, f))}
      ${card("③ 貸借対照表のかたち",
        "左が持ち物、右がそのお金の出どころ。右下の緑が厚いほど、返さなくていいお金で持ち物を買っていることになります。",
        bsBlock(r, f))}
      ${card("④ 売上と利益の推移",
        "棒が売上高、折れ線が経常利益です。棒が伸びて線が下がっていれば増収減益、両方下がっていれば減収減益です。",
        trend(r, f))}
      ${card("⑤ 返せるお金と、返す額",
        "青が1年で手元に残るお金の目安（簡易キャッシュフロー）、桃色がその年の約定返済額。青が桃色より高ければ、返済は回っています。",
        repay(r, f))}
      ${card("⑥ 借金を返し切るまでの年数",
        "いまの稼ぐペースのまま、要償還債務を返し切るのに何年かかるか。⑤が「今年の返済に足りるか」なのに対し、⑥は「借金の絶対量」を見ます。配点は⑤と⑥で15点ずつ、合わせて30点。100点のうち最も重い項目です。",
        gauge(r))}
      <div class="viz__card viz__card--wide">
        <h3 class="viz__h">⑦ グラフから読み取れること</h3>
        <p class="viz__hint">入力された数字から自動で書き出した観察です。判断そのものではありません。</p>
        ${readings(r, f)}
      </div>
    </div>
  </div>`;
}

/* ============================================ 信用程度バロメーター */
const RANKS = [
  ["E", 0, 35, "#8C3B22", "原則として新規与信は見合わせ、既存与信は回収・保全を優先"],
  ["D", 36, 50, "#B5623F", "慎重な対応が必要。保全策の検討を推奨"],
  ["C", 51, 65, "#527695", "概ね可としつつ、定期的なモニタリングが必要"],
  ["B", 66, 85, "#82B33A", "通常の与信取引に支障はない水準"],
  ["A", 86, 100, "#4E8F2E", "積極的に取り組んで差し支えない水準"],
];
function rankBar(total) {
  const W = 640, x0 = 10, x1 = W - 10, y = 40, bh = 30;
  const X = (v) => x0 + (x1 - x0) * Math.max(0, Math.min(100, v)) / 100;
  let h = "";
  RANKS.forEach(([g, lo, hi, col, pol]) => {
    const a = X(lo), b = X(hi + (g === "A" ? 0 : 1));
    h += tip(R(a, y, b - a - 2, bh, col, { rx: 3 })
      + T((a + b) / 2, y + 20, g, { a: "middle", s: 15, w: "700", c: "#fff" }),
      [`信用程度 ${g}`, `${lo}〜${hi}点`, pol]);
  });
  const px = X(total);
  h += `<polygon points="${px},${y - 3} ${px - 7},${y - 15} ${px + 7},${y - 15}" fill="#0F1A22"/>`;
  h += L(px, y - 3, px, y + bh + 3, "#0F1A22", 2.4);
  h += T(Math.max(26, Math.min(W - 26, px)), y + bh + 20, `${total}点`, { a: "middle", s: 13, w: "700", c: C.ink });
  h += T(x0, y + bh + 20, "0", { s: 11, c: C.faint });
  h += T(x1, y + bh + 20, "100", { a: "end", s: 11, c: C.faint });
  return svg("0 0 640 90", h);
}

/** 判定結果ヘッダ（総合スコア・信用程度・バロメーター） */
export function renderHead(r, f, policy) {
  const s = r.scores;
  const cls = s.rank === "C" ? " is-mid" : (s.rank === "D" || s.rank === "E") ? " is-low" : "";
  return `
  <div class="calc">
    <h2>判定結果</h2>
    <div class="vizhead${cls}">
      <p class="vizhead__co">${esc(r.input.name || "（会社名が未入力です）")}</p>
      <p class="vizhead__meta">${esc(r.input.industry.trim())}　／　${esc(r.input.capitalTier)}　／　${esc(r.input.listing)}　／　単位：${f.U_LABEL()}</p>
      <div class="vizhead__row">
        <div class="vizhead__num">
          <span class="vizhead__lab">総合スコア</span>
          <p class="vizhead__score"><b>${s.total}</b> / 100</p>
        </div>
        <div class="vizhead__num">
          <span class="vizhead__lab">信用程度</span>
          <p class="vizhead__grade">${s.rank}</p>
        </div>
        <div class="vizhead__bar">${rankBar(s.total)}</div>
      </div>
      <p class="vizhead__policy">${esc(policy)}</p>
    </div>
  </div>`;
}

/* ================================================== ツールチップの配線 */
let tipEl = null;
/** 図の上にマウス（または指）を乗せると、その数字の意味を出す */
export function attachTips(root) {
  if (!root) return;
  if (!tipEl) {
    tipEl = document.createElement("div");
    tipEl.className = "viz__tip";
    tipEl.setAttribute("role", "tooltip");
    document.body.appendChild(tipEl);
  }
  const hide = () => { tipEl.classList.remove("is-on"); };
  const show = (el, x, y) => {
    const raw = el.getAttribute("data-tip");
    if (!raw) return;
    const [head, ...rest] = raw.split("|");
    tipEl.innerHTML = `<b>${head}</b>` + rest.map((t) => `<span>${t}</span>`).join("");
    tipEl.classList.add("is-on");
    const w = tipEl.offsetWidth, h = tipEl.offsetHeight;
    let left = x + 14, top = y + 16;
    if (left + w > window.innerWidth - 8) left = x - w - 14;
    if (top + h > window.innerHeight - 8) top = y - h - 14;
    tipEl.style.left = Math.max(8, left) + "px";
    tipEl.style.top = Math.max(8, top) + "px";
  };
  root.addEventListener("pointermove", (e) => {
    const el = e.target.closest("[data-tip]");
    if (el) show(el, e.clientX, e.clientY); else hide();
  });
  root.addEventListener("pointerleave", hide);
  root.addEventListener("pointerdown", (e) => {
    const el = e.target.closest("[data-tip]");
    if (el) show(el, e.clientX, e.clientY);
  });
  window.addEventListener("scroll", hide, { passive: true });
}

/* ====================================================== Excel用の図の書き出し
 * 画面と同じ図を、そのまま画像にしてExcelに貼るためのもの。
 * 画像なので、Excelのグラフ機能も、参照先のデータも、数式も付いてこない。
 * ＝ 配点表やしきい値といったロジックは一切外に出ない。
 * ========================================================================== */

/** ダッシュボードに載せる6枚。番号は画面の①〜⑥と同じ。 */
export function figures(r, f) {
  return [
    { no: "①", title: "評点の内訳", svg: scoreBars(r) },
    { no: "②", title: "財務指標のかたち", svg: radar(r) },
    { no: "③", title: "貸借対照表のかたち", svg: bsBlock(r, f) },
    { no: "④", title: "売上と利益の推移", svg: trend(r, f) },
    { no: "⑤", title: "返せるお金と、返す額", svg: repay(r, f) },
    { no: "⑥", title: "借金を返し切るまでの年数", svg: gauge(r) },
  ];
}

/** ⑦の所見は文章なので、そのまま配列で渡す */
export function readingLines(r, f) {
  const html = readings(r, f);
  return [...html.matchAll(/<li>(.*?)<\/li>/g)].map((m) =>
    m[1].replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&amp;/g, "&"));
}

/** SVGをPNGのdataURLにする。ブラウザの描画をそのまま使う。 */
const FONT_STACK = "'Hiragino Sans','Hiragino Kaku Gothic ProN','Yu Gothic UI','Yu Gothic',YuGothic," +
  "'Noto Sans JP','Noto Sans CJK JP','Meiryo','MS PGothic',sans-serif";
export function toPng(svgMarkup, scale = 2) {
  return new Promise((resolve, reject) => {
    const vb = /viewBox="0 0 (\d+) (\d+)"/.exec(svgMarkup);
    if (!vb) return reject(new Error("viewBoxが読み取れません"));
    const W = +vb[1], H = +vb[2];
    // svg() が出力する時点で xmlns は入っている。ここで足すと属性が重複し、
    // XMLとして壊れて画像化に失敗するので、足さないこと。
    let m = svgMarkup
      .replace(/ data-tip="[^"]*"/g, "")
      .replace(/ class="viz__(?:svg|hot)"/g, "")
      .replace(/^<svg/, `<svg width="${W}" height="${H}" font-family="${FONT_STACK}"`);
    m = m.replace(">", `><rect x="0" y="0" width="${W}" height="${H}" fill="#FFFFFF"/>`);
    const img = new Image();
    img.onload = () => {
      try {
        const cv = document.createElement("canvas");
        cv.width = Math.round(W * scale); cv.height = Math.round(H * scale);
        const cx = cv.getContext("2d");
        cx.fillStyle = "#FFFFFF"; cx.fillRect(0, 0, cv.width, cv.height);
        cx.setTransform(scale, 0, 0, scale, 0, 0);
        cx.drawImage(img, 0, 0);
        resolve({ dataUrl: cv.toDataURL("image/png"), w: W, h: H });
      } catch (e) { reject(e); }
    };
    img.onerror = () => reject(new Error("図を画像に変換できませんでした"));
    img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(m);
  });
}

/** 6枚まとめてPNGにする。1枚失敗しても残りは出す。 */
export async function renderFigures(r, f, scale = 2) {
  const out = []; const failed = [];
  for (const g of figures(r, f)) {
    try {
      const png = await toPng(g.svg, scale);
      out.push({ no: g.no, title: g.title, ...png });
    } catch (e) { failed.push(g.no + " " + g.title); }
  }
  // 1枚も作れなかったときは黙って落とさない。原因が分からなくなるため。
  if (!out.length) throw new Error("図を画像にできませんでした（" + failed.join("／") + "）");
  if (failed.length && typeof console !== "undefined") {
    console.warn("[財務でポン] 画像化できなかった図:", failed.join("／"));
  }
  return out;
}
