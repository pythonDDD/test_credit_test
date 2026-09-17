/* ============================================================================
 * /lease/viz.js — 有料版のグラフ（SVG）
 * 画面では data-tip を持つ要素にマウスや指を乗せると、数字が出る。
 * Excelには、同じSVGをPNGにして貼る（toPng）。
 * ========================================================================== */
const COL = {
  deep: "#33526C", sea: "#527695", sky: "#84D2F5", lime: "#C3D64A", bloom: "#F8A3BF",
  plum: "#B67AB4", ink: "#0F1A22", soft: "#465A6C", grid: "rgba(24,38,47,.10)",
};
export const METHODS = {
  A: { label: "リース", color: "#33526C" },
  D: { label: "割賦", color: "#B67AB4" },
  C: { label: "銀行借入で購入", color: "#3E93C9" },
  B: { label: "現金で購入", color: "#8FB339" },
};
const yen = (n) => Math.round(n).toLocaleString("ja-JP");
const man = (n) => `${(n / 10000).toLocaleString("ja-JP", { maximumFractionDigits: 1 })}万`;
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");
/** 目盛りをキリのよい数（1・2・2.5・5 の10倍ごと）にそろえる */
function niceScale(minV, maxV, steps = 4) {
  const span = Math.max(maxV - minV, 1);
  const raw = span / steps, mag = Math.pow(10, Math.floor(Math.log10(raw))), norm = raw / mag;
  const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10) * mag;
  const lo = Math.floor(minV / step) * step, hi = Math.ceil(maxV / step) * step;
  const ticks = [];
  for (let v = lo; v <= hi + step / 2; v += step) ticks.push(v);
  return { lo, hi, ticks };
}
const svg = (w, h, body, label) =>
  `<svg class="viz" viewBox="0 0 ${w} ${h}" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="${esc(label)}">${body}</svg>`;
const text = (x, y, s, o = {}) =>
  `<text x="${x}" y="${y}" font-size="${o.size || 12}" fill="${o.fill || COL.soft}" text-anchor="${o.anchor || "start"}"${o.bold ? ' font-weight="700"' : ""}>${esc(s)}</text>`;

/* ------------------------------------------------------------ 1. 多く払う分の中身（ドーナツ） */
export function donutMarkup(r, o = {}) {
  const cp = !!o.compact;
  const parts = [
    { label: "資金の金利（リース会社が払う）", short: "資金の金利", v: r.cost.fundInterest, c: COL.bloom },
    { label: "税金（償却資産税）", short: "税金", v: r.cost.taxTotal, c: COL.lime },
    { label: "保険料", short: "保険料", v: r.cost.ins, c: COL.sky },
    { label: "その他の初期費用", short: "初期費用", v: r.cost.init, c: COL.sea },
    { label: "リース会社の利益", short: "リース会社の利益", v: r.lease.profit, c: COL.plum },
  ].filter((p) => p.v > 0);
  const sum = parts.reduce((s, p) => s + p.v, 0) || 1;
  const W = cp ? 360 : 520, cx = cp ? 180 : 140, cy = cp ? 138 : 140, R = cp ? 118 : 112, r0 = cp ? 76 : 70;
  const H = cp ? 290 + Math.ceil(parts.length / 2) * 46 : 280;
  let a0 = -Math.PI / 2, body = "";
  for (const p of parts) {
    const a1 = a0 + (p.v / sum) * Math.PI * 2;
    const large = a1 - a0 > Math.PI ? 1 : 0;
    const pt = (ang, rad) => `${(cx + rad * Math.cos(ang)).toFixed(2)},${(cy + rad * Math.sin(ang)).toFixed(2)}`;
    const d = parts.length === 1
      ? `M${cx - R},${cy} a${R},${R} 0 1,0 ${2 * R},0 a${R},${R} 0 1,0 ${-2 * R},0 M${cx - r0},${cy} a${r0},${r0} 0 1,1 ${2 * r0},0 a${r0},${r0} 0 1,1 ${-2 * r0},0`
      : `M${pt(a0, R)} A${R},${R} 0 ${large},1 ${pt(a1, R)} L${pt(a1, r0)} A${r0},${r0} 0 ${large},0 ${pt(a0, r0)} Z`;
    body += `<path d="${d}" fill="${p.c}" stroke="#fff" stroke-width="2" data-tip="${esc(`${p.label}｜${yen(p.v)}円（${(p.v / sum * 100).toFixed(1)}%）`)}"/>`;
    a0 = a1;
  }
  const profit = r.lease.profit;
  body += text(cx, cy - 8, "リース会社の利益", { anchor: "middle", size: cp ? 13 : 12 });
  body += text(cx, cy + 16, `${yen(profit)}円`, { anchor: "middle", size: cp ? 20 : 18, bold: true, fill: profit < 0 ? "#B5623F" : COL.ink });
  // 残価がある見積では、支払総額より「回収する額（残価を含む）」のほうが正しい言い方になる
  const share = (Math.max(profit, 0) / sum * 100).toFixed(1);
  body += text(cx, cy + 36, r.input.residual > 0 ? `回収する分の${share}%` : `多く払う分の${share}%`, { anchor: "middle", size: cp ? 12 : 11 });
  parts.forEach((p, i) => {
    const x = cp ? 16 + (i % 2) * 176 : 300, y = cp ? 292 + Math.floor(i / 2) * 46 : 50 + i * 42;
    body += `<rect x="${x}" y="${y - 11}" width="14" height="14" rx="4" fill="${p.c}"/>`;
    body += text(x + 22, y, cp ? p.short : p.label, { size: cp ? 13 : 12.5, fill: COL.ink });
    body += text(x + 22, y + 19, `${yen(p.v)}円`, { size: cp ? 14 : 13, bold: true, fill: COL.ink });
  });
  return svg(W, H, body, r.input.residual > 0 ? "物件価額を超えて回収する分の中身（残価を含む）" : "物件価額より多く払う分の中身");
}

/* ------------------------------------------------------------ 2. 4つの買い方の比較（累計の実質負担） */
export function lineCompare(r, show = { A: true, D: true, C: true, B: true }, o = {}) {
  const cp = !!o.compact, fs = cp ? 13 : 11;
  const lastYear = Math.max(1, ...r.cmp.filter((x) => x.payCount > 0 || x.dep > 0 || x.tax > 0).map((x) => x.year));
  const years = Math.min(12, lastYear);
  const keys = Object.keys(METHODS).filter((k) => show[k]);
  const vals = keys.flatMap((k) => r.cmpCum[k].slice(0, years));
  const sc = niceScale(Math.min(0, ...vals), Math.max(1, ...vals));
  const W = cp ? 360 : 720, H = cp ? 280 : 300, L = cp ? 52 : 62, Rr = cp ? 12 : 20, T = 16, B = 38;
  const X = (i) => L + (W - L - Rr) * (years === 1 ? 0.5 : i / (years - 1));
  const Y = (v) => T + (H - T - B) * (1 - (v - sc.lo) / (sc.hi - sc.lo || 1));
  let body = "";
  for (const v of sc.ticks) {
    const y = Y(v);
    body += `<line x1="${L}" y1="${y}" x2="${W - Rr}" y2="${y}" stroke="${COL.grid}"/>`;
    body += text(L - 6, y + 4, man(v), { anchor: "end", size: fs });
  }
  for (let i = 0; i < years; i++) body += text(X(i), H - 14, cp ? `${i + 1}` : `${i + 1}年目`, { anchor: "middle", size: fs });
  if (cp) body += text(W - Rr, H - 1, "年目", { anchor: "end", size: 11 });
  for (const k of keys) {
    const pts = r.cmpCum[k].slice(0, years).map((v, i) => `${X(i).toFixed(1)},${Y(v).toFixed(1)}`).join(" ");
    body += `<polyline points="${pts}" fill="none" stroke="${METHODS[k].color}" stroke-width="3" stroke-linejoin="round" stroke-linecap="round"/>`;
    const lastI = years - 1;
    body += `<circle cx="${X(lastI)}" cy="${Y(r.cmpCum[k][lastI])}" r="4.5" fill="${METHODS[k].color}"/>`;
  }
  // 年ごとの当たり判定（縦の帯）。乗せると、その年までの累計が出る
  const colW = (W - L - Rr) / Math.max(years - 1, 1);
  for (let i = 0; i < years; i++) {
    const tip = `${i + 1}年目までの実質負担｜` + keys.map((k) => `${METHODS[k].label} ${man(r.cmpCum[k][i])}円`).join("｜");
    body += `<rect x="${(X(i) - colW / 2).toFixed(1)}" y="${T}" width="${colW.toFixed(1)}" height="${H - T - B}" fill="transparent" data-tip="${esc(tip)}"/>`;
  }
  return svg(W, H, body, "4つの買い方の、年ごとの実質負担（累計）");
}

/* ------------------------------------------------------------ 3. 途中で解約したら（残りの支払） */
export function areaRemaining(r, k, o = {}) {
  const cp = !!o.compact, fs = cp ? 13 : 11;
  const n = r.schedule.length, total = r.lease.total;
  const W = cp ? 360 : 720, H = cp ? 240 : 260, L = cp ? 52 : 62, Rr = cp ? 12 : 20, T = 16, B = 34;
  const sc = niceScale(0, total);
  const X = (i) => L + (W - L - Rr) * (i / n);
  const Y = (v) => T + (H - T - B) * (1 - v / (sc.hi || 1));
  let body = "";
  for (const v of sc.ticks) {
    const y = Y(v);
    body += `<line x1="${L}" y1="${y}" x2="${W - Rr}" y2="${y}" stroke="${COL.grid}"/>`;
    body += text(L - 6, y + 4, man(v), { anchor: "end", size: fs });
  }
  const step = Math.max(12, Math.ceil(n / (cp ? 36 : 60)) * 12);
  for (let m = 0; m <= n; m += step) body += text(X(m), H - 12, cp ? `${m}` : `${m}か月`, { anchor: "middle", size: fs });
  if (cp) body += text(W - Rr, H - 1, "か月", { anchor: "end", size: 11 });
  let d = `M${X(0)},${Y(total)}`;
  r.schedule.forEach((s, i) => { d += ` L${X(i + 1).toFixed(1)},${Y(s.remaining).toFixed(1)}`; });
  body += `<path d="${d} L${X(n)},${Y(0)} L${X(0)},${Y(0)} Z" fill="rgba(132,210,245,.35)"/>`;
  body += `<path d="${d}" fill="none" stroke="${COL.sea}" stroke-width="2.5"/>`;
  const rem = k >= n ? 0 : (k <= 0 ? total : r.schedule[k - 1].remaining);
  body += `<line x1="${X(k)}" y1="${T}" x2="${X(k)}" y2="${H - B}" stroke="${COL.plum}" stroke-width="2" stroke-dasharray="5 4"/>`;
  body += `<circle cx="${X(k)}" cy="${Y(rem)}" r="6" fill="${COL.plum}" stroke="#fff" stroke-width="2"/>`;
  return svg(W, H, body, "途中で解約したときの、残りの支払");
}

/* ------------------------------------------------------------ 4. 毎年の経費と税金（買った場合） */
export function barsExpense(r, o = {}) {
  const cp = !!o.compact, fs = cp ? 13 : 11;
  const years = Math.min(12, Math.max(r.dep.rows.length, r.tax.length));
  const rows = Array.from({ length: years }, (_, i) => ({
    y: i + 1, dep: r.dep.rows[i] ? r.dep.rows[i].amount : 0, tax: r.tax[i] ? r.tax[i].tax : 0 }));
  const sc = niceScale(0, Math.max(1, ...rows.map((x) => x.dep)));
  const W = cp ? 360 : 720, H = cp ? 250 : 280, L = cp ? 52 : 62, Rr = cp ? 12 : 20, T = 16, B = 34;
  const bw = (W - L - Rr) / years;
  const Y = (v) => T + (H - T - B) * (1 - v / sc.hi);
  let body = "";
  for (const v of sc.ticks) {
    const y = Y(v);
    body += `<line x1="${L}" y1="${y}" x2="${W - Rr}" y2="${y}" stroke="${COL.grid}"/>`;
    body += text(L - 6, y + 4, man(v), { anchor: "end", size: fs });
  }
  rows.forEach((x, i) => {
    const x0 = L + bw * i + bw * 0.16, w = bw * 0.68;
    body += `<rect x="${x0}" y="${Y(x.dep)}" width="${w * 0.62}" height="${H - B - Y(x.dep)}" rx="3" fill="${COL.sea}" data-tip="${esc(`${x.y}年目｜減価償却費 ${yen(x.dep)}円`)}"/>`;
    body += `<rect x="${x0 + w * 0.66}" y="${Y(x.tax)}" width="${w * 0.34}" height="${H - B - Y(x.tax)}" rx="3" fill="${COL.lime}" data-tip="${esc(`${x.y}年目｜償却資産税 ${yen(x.tax)}円`)}"/>`;
    body += text(L + bw * i + bw / 2, H - 12, cp ? `${x.y}` : `${x.y}年目`, { anchor: "middle", size: fs });
  });
  if (cp) body += text(W - Rr, H - 1, "年目", { anchor: "end", size: 11 });
  return svg(W, H, body, "買った場合の、年ごとの減価償却費と償却資産税");
}

/* ------------------------------------------------------------ 画面のツールチップ */
export function attachTips(box) {
  if (!box || box.dataset.tips) return;
  box.dataset.tips = "1";
  const tip = document.createElement("div");
  tip.className = "vtip";
  tip.hidden = true;
  box.appendChild(tip);
  const show = (e) => {
    const t = e.target.closest("[data-tip]");
    if (!t || !box.contains(t)) { tip.hidden = true; return; }
    tip.innerHTML = t.getAttribute("data-tip").split("｜").map((s, i) => i ? `<span>${esc(s)}</span>` : `<b>${esc(s)}</b>`).join("");
    tip.hidden = false;
    const rb = box.getBoundingClientRect();
    const px = (e.touches ? e.touches[0].clientX : e.clientX) - rb.left;
    const py = (e.touches ? e.touches[0].clientY : e.clientY) - rb.top;
    const w = tip.offsetWidth;
    tip.style.left = `${Math.min(Math.max(px - w / 2, 4), rb.width - w - 4)}px`;
    tip.style.top = `${Math.max(py - tip.offsetHeight - 14, 4)}px`;
  };
  box.addEventListener("pointermove", show);
  box.addEventListener("pointerdown", show);
  box.addEventListener("pointerleave", () => { tip.hidden = true; });
}

/* ------------------------------------------------------------ Excel用のPNG */
const FONT_STACK = "'Hiragino Sans','Hiragino Kaku Gothic ProN','Yu Gothic UI','Yu Gothic',YuGothic,'Noto Sans JP','Noto Sans CJK JP','Meiryo',sans-serif";
export function toPng(markup, scale = 2) {
  return new Promise((resolve, reject) => {
    const vb = /viewBox="0 0 (\d+) (\d+)"/.exec(markup);
    if (!vb) return reject(new Error("viewBoxが読み取れません"));
    const W = +vb[1], H = +vb[2];
    // xmlns は svg() で1回だけ入れている。ここで足すと重複してXMLが壊れるので足さない
    const m = markup.replace(/ data-tip="[^"]*"/g, "").replace(/ class="viz"/, "")
      .replace(/^<svg/, `<svg width="${W}" height="${H}" font-family="${FONT_STACK}"`);
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
    img.onerror = () => reject(new Error("グラフを画像にできませんでした"));
    img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(m);
  });
}
