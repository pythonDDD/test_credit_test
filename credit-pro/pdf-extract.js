/* ============================================================================
 * pdf-extract.js — 決算書PDFから数値を取り出す
 *
 * pdf.js のテキストレイヤーだけを使う。OCRもサーバー送信も行わない。
 * Python版（検証済み）と同じアルゴリズムを移植したもの。
 *
 * 対応している様式:
 *   ・計算書類（会社計算規則）  … 資産の部／負債の部 の左右2段組み・単年
 *   ・有価証券報告書 / 決算短信 … 前期／当期 の2年並記・単段
 *
 * 気をつけている点:
 *   ・分散配置（「現 金 及 び 預 金」）→ 空白を全除去してから照合する
 *   ・注記番号（※1,※2）→ 科目名から除去する
 *   ・2段組みには「科目が2列」と「年度が2列」の2種類がある。
 *     取り違えると前期の数字を当期として静かに取り込むため、必ず判別する
 *   ・有利子負債は「社債＋長期借入金」のように合算が要る。選択ではなく加算
 * ========================================================================== */

/* ------------------------------------------------------------------ 正規化 */
/** 分散配置・全角・注記番号を吸収して比較可能な形にする */
export function norm(s) {
  return String(s)
    .normalize("NFKC")
    .replace(/[\s\u3000]+/g, "")
    .replace(/[※＊*]\s*[\d,、，]*/g, "")
    .replace(/[（(]注\d*[）)]/g, "")
    // 小計ラベルを囲む隅付き括弧・かぎ括弧を除去する。
    // 例:「【流動資産】」「【売上高】」を辞書の「流動資産」「売上高」と一致させるため。
    // 単位表記（単位：千円）や（連結）は丸括弧なので影響しない。
    .replace(/[【】「」『』〔〕［］]/g, "")
    // IFRSの有価証券報告書は、科目名の直後に注記番号を置く。
    // 例:「営業債権及びその他の債権7、27」「リース負債12、27」
    // 末尾の数字（と区切りの読点）だけを落とす。先頭の数字（1年内返済予定…）は残す。
    .replace(/\d+(?:[、，,]\d+)*$/, "")
    // 「法人税、住民税および事業税」のように接続詞を仮名で書く決算書がある。
    // 辞書は漢字なので一致せず、その科目が丸ごと欠落する。
    // 実例：法人税等が調整額だけになり、当期純利益が積み上がらなくなっていた。
    .replace(/および/g, "及び").replace(/ならびに/g, "並びに")
    .replace(/または/g, "又は").replace(/もしくは/g, "若しくは");
}

const NUM_RE = /^[△▲\-−(（]?[\d,]+[)）]?$/;

/**
 * pdf.js は「△ 63,414」のように、符号と数字のあいだに空白を含んだまま
 * 1つの item として返すことがある（決算公告に多い）。
 * 空白を落とさずに数値判定すると、その金額は数値と認められず科目名の側に流れ、
 * さらに norm() の「末尾の注記番号を落とす」処理に削られて、金額ごと消える。
 * 実例：「法人税等調整額△63,414」が消え、法人税等が合計だけになっていた。
 */
const deSpace = (s) => String(s).normalize("NFKC").replace(/[\s\u3000]+/g, "");

export function toNum(tok) {
  const t0 = deSpace(tok).replace(/,/g, "");
  const neg = /^[△▲\-−(（]/.test(t0);
  const t = t0.replace(/^[△▲\-−(（]/, "").replace(/[)）]$/, "");
  if (!/^\d+$/.test(t)) return null;
  const v = parseInt(t, 10);
  return neg ? -v : v;
}

const isNum = (tok) => NUM_RE.test(deSpace(tok)) && toNum(tok) !== null;

/**
 * 「数字の一部になりうる文字」かどうか。
 * PDFによっては金額を1文字ずつ別のitemで出す（例: "4" "3" "2" "," "4" "6" "4"）。
 * isNum は "," 単体を数値と見なさないため、そこで語の結合が切れ、
 * カンマだけが科目名の側に取り残されて「流動資産合計,,」のようになり辞書と一致しなくなる。
 * 結合の判定にはこちらを使い、数値としての判定には従来どおり isNum を使う。
 */
const isNumish = (tok) => /^[\d,.\u25b3\u25b2\-\u2212()（）]+$/.test(deSpace(tok));

/* -------------------------------------------------- pdf.js の item → 語 */
/**
 * pdf.js は文字単位でitemを返すことがある（分散配置の決算書は特に）。
 * y座標で行にまとめ、x方向に近いものを1語として連結する。
 */
function itemsToWords(items, viewportHeight) {
  const raw = [];
  for (const it of items) {
    const s = it.str;
    if (!s || !s.trim()) continue;
    const x0 = it.transform[4];
    const y = it.transform[5];
    raw.push({ text: s, x0, x1: x0 + (it.width || 0), top: viewportHeight - y });
  }
  raw.sort((a, b) => a.top - b.top || a.x0 - b.x0);

  // 行にまとめる
  const lines = [];
  for (const r of raw) {
    const ln = lines.find((l) => Math.abs(l.top - r.top) <= 3);
    if (ln) { ln.items.push(r); ln.top = (ln.top + r.top) / 2; }
    else lines.push({ top: r.top, items: [r] });
  }

  // 行内でx方向に近いものを連結して語にする
  const words = [];
  for (const ln of lines) {
    ln.items.sort((a, b) => a.x0 - b.x0);
    let cur = null;
    for (const it of ln.items) {
      const h = 10; // 想定文字高
      // 決算公告は金額欄を【】で囲む。左段の「】」と右段の「【流動負債】」は
      // 数ピクセルしか離れておらず、そのまま結合すると左右の段が1つの語になり、
      // 段組みの分割が効かなくなる。隅付き括弧はまたいで結合しない。
      const bracketBreak = /】$/.test(cur ? cur.text : "") || /^【/.test(it.text);
      if (cur && !bracketBreak && it.x0 - cur.x1 <= h * 0.9 && isNumish(it.text) === isNumish(cur.text)) {
        cur.text += it.text; cur.x1 = it.x1;
      } else {
        if (cur) words.push(cur);
        cur = { text: it.text, x0: it.x0, x1: it.x1, top: ln.top };
      }
    }
    if (cur) words.push(cur);
  }
  return words;
}

/**
 * 科目名を突き合わせるための鍵。括弧の類をすべて落とす。
 * 決算公告は小計を「【流動資産】」と囲むが、PDFによっては半角で「流動資産[]」、
 * 「有形固定資産(」のように括弧が科目名に貼りついた形で出てくる。
 * norm では全角の隅付き括弧しか落としていなかったため、半角のものが残って辞書と一致せず、
 * 区分の合計が丸ごと欠けて貸借が合わなくなっていた。
 * ラベル側と辞書側の両方に同じ処理をかけるので、「(純額)」付きの辞書項目も壊れない。
 */
const keyOf = (s) => String(s)
  .replace(/[【】「」『』〔〕［］\[\]（）()]/g, "")
  // 「当期純損失(△)」「税金等調整前当期純損失(△)」のように、
  // 損失であることを示す△が科目名に付く様式がある。括弧を外しても△が残り、
  // 辞書の「当期純損失」と一致せず当期純利益が丸ごと欠測していた。
  .replace(/[△▲]/g, "");

/* ------------------------------------------------------------ レイアウト判定 */
/** 右揃えで3回以上現れるx1＝金額列。見出しの年号やページ番号は頻度で除外する */
function amountBands(words, pageWidth) {
  const nums = words.filter((w) => isNum(w.text));
  if (nums.length < 4) return [];
  // 注記番号の列（「10」「15」など）を金額列と誤認しないようにする。
  // 注記番号は桁が短く符号も付かないため、列ごとの「最大桁数」で見分ける。
  const cnt = new Map(), width = new Map();
  for (const w of nums) {
    const k = Math.round(w.x1);
    cnt.set(k, (cnt.get(k) || 0) + 1);
    const digits = w.text.replace(/[^\d]/g, "").length;
    const signed = /^[△▲\-−(（]/.test(w.text.normalize("NFKC"));
    width.set(k, Math.max(width.get(k) || 0, digits + (signed ? 10 : 0)));
  }
  let xs = [...cnt.entries()].filter(([, c]) => c >= 3).map(([x]) => x).sort((a, b) => a - b);
  if (xs.length > 1) {
    // 最も「金額らしい」列の桁数を基準に、明らかに桁の小さい列（注記番号）を落とす
    const maxW = Math.max(...xs.map((x) => width.get(x)));
    const kept = xs.filter((x) => width.get(x) >= Math.min(4, maxW) || width.get(x) >= maxW - 2);
    if (kept.length) xs = kept;
  }
  if (!xs.length) return [];
  const bands = [];
  let cur = [xs[0]];
  for (const x of xs.slice(1)) {
    if (x - cur[cur.length - 1] < pageWidth * 0.10) cur.push(x);
    else { bands.push(cur); cur = [x]; }
  }
  bands.push(cur);
  return bands.map((b) => Math.max(...b));
}

/**
 * 段組みを判定する。
 *  "single"   … 単段
 *  "accounts" … 科目が左右2列（計算書類）→ splitX で分割して読む
 *  "years"    … 年度が2列（有報・短信）→ 分割せず、行内の数値配列で持つ
 */
function classify(words, pageWidth, pageText) {
  const bands = amountBands(words, pageWidth);
  if (bands.length < 2) return { kind: "single", splitX: null };
  if (/(前連結会計年度|前事業年度|前期).{0,40}(当連結会計年度|当事業年度|当期)/.test(pageText))
    return { kind: "years", splitX: null };
  // 左段の金額と右段の金額の“あいだ”に科目名があれば「科目が2列」、なければ「年度が2列」。
  // 判定を誤ると当期と前期を取り違えるため、次の2点を必ず確認する。
  //   ・あいだの文字が、左段の金額のすぐ右から始まっているか（表の途中の注釈を拾わない）
  //   ・その文字が複数行にわたって存在するか（1行だけの脚注を科目列と誤認しない）
  const lo = bands[0], hi = bands[1];
  const mids = words.filter((w) => !isNum(w.text) && w.x0 > lo + 2 && w.x0 < hi - 20);
  const midRows = new Set(mids.map((w) => Math.round(w.top / 3)));
  const numRows = new Set(words.filter((w) => isNum(w.text) && Math.abs(w.x1 - hi) <= 4)
                               .map((w) => Math.round(w.top / 3)));
  // 右段の金額がある行のうち、あいだに科目名も存在する行の割合
  let share = 0;
  if (numRows.size) {
    let n = 0;
    for (const r of numRows) if (midRows.has(r)) n++;
    share = n / numRows.size;
  }
  if (mids.length >= 5 && midRows.size >= 3 && share >= 0.6)
    return { kind: "accounts", splitX: lo + 6 };
  return { kind: "years", splitX: null };
}

/* ---------------------------------------------------------------- 行の抽出 */
function rowsOf(words, pageWidth, pageText) {
  const { kind, splitX } = classify(words, pageWidth, pageText);
  const buckets = new Map();
  for (const w of words) {
    const col = splitX === null || w.x0 < splitX ? 0 : 1;
    const key = col + ":" + Math.round(w.top / 3);
    if (!buckets.has(key)) buckets.set(key, []);
    buckets.get(key).push(w);
  }
  // 金額列のx1位置。ここに右端が揃っている数値だけを金額として採用し、
  // 注記番号（左寄り・少桁）を金額として拾ってしまうのを防ぐ。
  const bands = amountBands(words, pageWidth);
  const onBand = (w) => bands.some((b) => Math.abs(w.x1 - b) <= 4);

  // 「１年以内返済／長期借入金」のように科目名が2行に折り返すと、
  // 金額だけの行と、科目名だけの行に分かれてしまう。
  // 同じ段の隣接する行どうしに限って、科目名を金額の行へ寄せる。
  {
    const keys = [...buckets.keys()];
    const info = new Map();
    for (const k of keys) {
      const ws = buckets.get(k);
      const [col, top] = k.split(":").map(Number);
      info.set(k, { col, top, hasNum: ws.some((w) => isNum(w.text)),
                    hasLabel: ws.some((w) => !isNum(w.text)) });
    }
    for (const k of keys) {
      const me = info.get(k);
      if (!me) continue;                               // 既に他の行へ寄せた行
      if (!me.hasNum || me.hasLabel) continue;         // 金額だけの行が対象
      // 折り返しは上下2行にまたがることがある（1行目「１年以内返済」/2行目「長期借入金」）。
      // 上下とも科目名だけの行なら、上の行から順に金額の行へ寄せる。
      let moved = 0;
      for (const d of [-2, -1, 1, 2]) {
        if (moved >= 2) break;
        const nk = me.col + ":" + (me.top + d);
        const nb = info.get(nk);
        if (!nb || nb.hasNum || !nb.hasLabel) continue;
        buckets.get(k).push(...buckets.get(nk));       // 科目名を金額の行へ移す
        buckets.delete(nk);
        info.delete(nk);
        moved++;
      }
    }
  }

  const rows = [];
  for (const ws of buckets.values()) {
    // 「販売費及び一般管／理費」のように2行へ折り返した科目名は、
    // x0だけで並べると下の行が先に来て「理費販売費及び一般管」になる。
    // 上下（top）を先に見てから左右（x0）で並べる。
    ws.sort((a, b) => (Math.round(a.top / 3) - Math.round(b.top / 3)) || (a.x0 - b.x0));
    const nums = ws.filter((w) => isNum(w.text));
    const labels = ws.filter((w) => !isNum(w.text)).map((w) => w.text);
    // 決算公告には、マイナス記号を金額から離れた「符号欄」に置く様式がある。
    //   法人税、住民税および事業税        △          627
    // 「△」が単独の語になると数値と認められず科目名の側に流れ、
    // 金額は正の数として読まれる。税額の還付が支払として入り、当期純利益が合わなくなる。
    // ただし、隣の欄から紛れ込んだ「△」を巻き込むと符号を壊すので、
    // 単独の△の個数が金額の個数と一致するときにだけ符号を反転する。
    const signs = ws.filter((w) => { const t = deSpace(w.text); return t === "△" || t === "▲"; }).length;
    let amountWords = bands.length ? nums.filter(onBand) : nums;
    if (!amountWords.length) amountWords = nums;        // バンドを検出できない表は従来どおり
    let amounts = amountWords.map((w) => toNum(w.text));
    if (signs > 0 && signs === amounts.length && amounts.every((v) => v >= 0)) {
      amounts = amounts.map((v) => -v);
    }
    // その金額が何列目にあるかを覚えておく。
    // 「内訳の右隣に区分合計を置く」様式では、内訳と合計が別の列に並ぶ。
    // 列を無視して右端を採ると、内訳を足すときに合計まで足してしまう。
    const cols = amountWords.map((w) => {
      if (!bands.length) return -1;
      let bi = 0, bd = Infinity;
      for (let i = 0; i < bands.length; i++) {
        const d = Math.abs(w.x1 - bands[i]);
        if (d < bd) { bd = d; bi = i; }
      }
      return bi;
    });
    if (!labels.length || !amounts.length) continue;
    rows.push({ label: norm(labels.join("")), amounts, cols });
  }
  return { rows, kind };
}

/* ------------------------------------------------------------------ 科目辞書 */
// agg   … 集計済みの科目。1つでも見つかればそれを採用
// parts … 内訳科目。aggが無いときに全部足す
const BS_DICT = {
  cash: { agg: ["現金及び預金", "現金預金", "現金・預金", "現金及び現金同等物"], parts: ["現金", "預金"] },
  receivables: { agg: ["受取手形及び売掛金", "売上債権", "受取手形、売掛金及び契約資産", "売掛金及びその他の短期債権", "受取手形及び売掛金(純額)", "営業債権及びその他の債権", "営業債権及び契約資産", "受取手形、売掛金及び契約資産(純額)"], parts: ["受取手形", "売掛金", "契約資産", "電気事業未収金", "電子記録債権", "売掛金及び契約資産", "完成工事未収入金", "受取手形及び電子記録債権"] },
  inventory: { agg: ["棚卸資産"], parts: ["商品及び製品", "商品", "製品", "仕掛品", "原材料及び貯蔵品", "原材料", "貯蔵品", "未成工事支出金", "販売用不動産", "仕掛販売用不動産", "未成業務支出金"] },
  tangible: { agg: ["有形固定資産合計", "有形固定資産"], parts: [] },
  currentAssets: { agg: ["流動資産合計", "流動資産"], parts: [] },
  // IFRSは「非流動資産合計」。日本基準の「固定資産合計」と同じ位置づけ
  fixedAssets: { agg: ["固定資産合計", "非流動資産合計", "固定資産"], parts: [] },
  deferred: { agg: ["繰延資産合計", "繰延資産"], parts: [] },
  payables: { agg: ["支払手形及び買掛金", "仕入債務", "買掛金及びその他の短期債務", "営業債務及びその他の債務"], parts: ["支払手形", "買掛金", "電気事業未払金", "電子記録債務", "工事未払金", "支払手形及び電子記録債務"] },
  shortDebt: { agg: [], parts: ["短期借入金", "1年内返済予定の長期借入金", "1年以内に期限到来の固定負債",
                                "1年内償還予定の社債", "コマーシャル・ペーパー", "リース債務",
                                "その他の短期金融負債", "リース負債", "短期借入債務", "リース債務流動", "1年以内返済長期借入金", "1年内返済予定長期借入金",
                                // 流動負債の節に載る「長期借入金」「社債」は1年内返済・償還分。
                                // 節を限定して拾うので、固定負債側の同名科目とは混ざらない。
                                "長期借入金", "社債"] },
  currentLiab: { agg: ["流動負債合計", "流動負債"], parts: [] },
  longDebt: { agg: [], parts: ["長期借入金", "社債", "リース債務", "長期金融負債", "リース負債", "長期借入債務", "借入金", "社債及び借入金", "リース債務固定"] },
  fixedLiab: { agg: ["固定負債合計", "非流動負債合計", "固定負債"], parts: [] },
  // 決算公告には合計行を置かず、区分の見出し行に金額を書く様式がある（「純資産の部 42,301,401」）。
  // 「純資産の部」を拾えないと、貸借が純資産のぶんだけ合わなくなる。
  equity: { agg: ["純資産合計", "純資産の部合計", "資本合計", "純資産の部", "純資産"], parts: [] },
  totalCapital: { agg: ["負債純資産合計", "負債及び純資産合計", "負債及び資本合計", "負債・純資産合計", "負債・純資産の部合計", "負債純資産の部合計", "負債及び純資産の部合計",
                       // 総資本は総資産と同額。上記が無い様式では資産側の合計で代用する
                       "資産合計", "資産の部合計", "資産の部計"], parts: [] },
};
const PL_DICT = {
  // 事業別に売上を並べる様式（建設・不動産）では、合計行を先に見ないと
  // 最初の事業の売上だけを全社の売上として取り込んでしまう。
  // 実例：大東建託で 1,842,357 のところを 540,975（完成工事高）にしていた。
  sales: { agg: ["売上高合計", "営業収益合計", "売上収益合計", "売上高", "営業収益", "純営業収益",
                 "売上収益", "完成工事高", "事業収益", "営業収益及び営業外収益"], parts: [] },
  cogs: { agg: ["売上原価合計", "売上原価", "完成工事原価", "営業原価"], parts: [] },
  sga: { agg: ["販売費及び一般管理費合計", "販売費及び一般管理費計", "販売費及び一般管理費",
               "販売費及び一般管理費等", "一般管理費", "販売費"], parts: [] },
  // IFRSには「営業利益」を置かず「コア営業利益」「事業利益」とする会社がある。
  // 本来の営業利益が見つかったときはそちらが優先されるよう、末尾に置く。
  operating: { agg: ["営業利益", "営業損失", "営業利益又は営業損失", "営業損失(△)", "コア営業利益", "事業利益"], parts: [] },
  // IFRSには営業外収益/費用が無い。金融収益・その他収益・持分法投資利益を合算して同じ位置に置く
  nonOpInc: { agg: ["営業外収益", "営業外収益合計"], parts: ["金融収益", "その他収益", "持分法による投資利益"] },
  nonOpExp: { agg: ["営業外費用", "営業外費用合計"], parts: ["金融費用", "その他費用", "持分法による投資損失"] },
  ordinary: { agg: ["経常利益", "経常損失", "当期経常利益", "経常利益又は経常損失"], parts: [] },
  extraInc: { agg: ["特別利益", "特別利益合計"], parts: [] },
  extraExp: { agg: ["特別損失", "特別損失合計"], parts: [] },
  // 「法人税等」は合計とはかぎらない。「法人税等」と「法人税等調整額」を別の行に置く様式があり、
  // 「法人税等」を合計として採ると調整額のぶんだけ税額を取り違える。
  // 実例：東京生活館で 147,774,199 のところを 219,353,700 にしていた。
  // 明示的な合計行だけを agg に置き、それ以外は内訳として足し合わせる。
  tax: { agg: ["法人税等合計", "法人所得税費用", "法人税、住民税及び事業税等合計"],
         parts: ["法人税、住民税及び事業税", "法人税住民税及び事業税", "法人税等", "法人税等調整額"] },
  net: { agg: ["当期純利益", "当期純損失", "当期利益", "中間利益", "中間(当期)利益", "親会社株主に帰属する当期純利益", "親会社の所有者に帰属する当期利益", "当期純利益又は当期純損失"], parts: [] },
};
const GROSS = { agg: ["売上総利益", "売上総利益合計", "営業総利益"], parts: [] };
// 売上高のほかに「営業収入」を並べる様式（小売業）。営業収益はこの合算になる。
const OPINCOME = { agg: ["営業収入合計", "営業収入", "その他の営業収入"], parts: [] };
const DEP = { agg: ["減価償却費", "減価償却費及びその他の償却費", "減価償却費及び償却費", "減価償却費及びのれん償却額", "減価償却及び償却費"], parts: [] };
const OPEX = ["営業費用", "営業費用合計", "営業費", "営業費合計", "経常費用", "経常費用合計"];

/** yi: 0=左（有報なら前期）／-1=右（当期） */
/**
 * yi: -1=当期（右端）／0=前期（右から2番目）
 * 有報・短信には「注記」列があり、注記番号が数値として拾われることがある。
 * 先頭から数えると注記番号を金額として取り込むため、必ず右端から数える。
 */
function pull(rows, spec, yi) {
  const val = (r) => {
    const a = r.amounts;
    if (yi === -1) return a[a.length - 1];
    return a.length >= 2 ? a[a.length - 2] : a[a.length - 1];
  };
  for (const name of spec.agg) {
    const k = keyOf(name);
    const r = rows.find((x) => keyOf(x.label) === k);
    if (r) return { value: val(r), source: name };
  }
  // 内訳の合算では、行の右端が「区分合計」であることがある。
  //   法人税、住民税及び事業税   477,659            ← 内訳の列
  //   法人税等調整額     29,373  507,032            ← 右端は区分合計の列
  // 右端をそのまま足すと、内訳と合計を両方足して二重計上になる。
  // 実例：テレビ大阪で法人税等が 507,032 → 984,691（＋94%）になっていた。
  //
  // かといって「金額が多い行は末尾を捨てる」では逆に壊れる。
  // 有価証券報告書には、隣の表の数値が1つだけ紛れ込んで
  // 前期・当期の左側に余分な列ができる行があり、そこでは捨てるべきは左端だからである。
  // 実例：キオクシアの「社債及び借入金」。
  //
  // そこで列の位置そのものを見る。足し合わせる科目すべてに共通して存在する列を選ぶ。
  // 内訳どうしは必ず同じ列に並ぶので、区分合計の列は自然に外れる。
  const partKeys = spec.parts.map(keyOf);
  const hits = rows.filter((x) => partKeys.includes(keyOf(x.label)));
  if (hits.length) {
    const col = commonCol(hits, yi);
    const valPart = (r) => {
      if (col !== null) {
        const i = r.cols ? r.cols.indexOf(col) : -1;
        if (i >= 0) return r.amounts[i];
      }
      return val(r);
    };
    return { value: hits.reduce((a, r) => a + valPart(r), 0), source: hits.map((h) => h.label).join("＋") };
  }
  return { value: null, source: null };
}

/**
 * 合算する行すべてに共通して存在する金額列のうち、当期にあたるものを返す。
 * yi=-1 なら共通列の右端、yi=0 ならその1つ左。
 * 列を特定できないときは null を返し、呼び出し側は従来どおり行の右端を使う。
 */
function commonCol(hits, yi) {
  if (!hits.length || !hits.every((r) => Array.isArray(r.cols) && r.cols.length === r.amounts.length)) return null;
  let common = null;
  for (const r of hits) {
    const s = new Set(r.cols.filter((c) => c >= 0));
    if (!s.size) return null;
    common = common === null ? s : new Set([...common].filter((c) => s.has(c)));
    if (!common.size) return null;
  }
  const sorted = [...common].sort((a, b) => a - b);
  if (yi === -1) return sorted[sorted.length - 1];
  return sorted.length >= 2 ? sorted[sorted.length - 2] : sorted[sorted.length - 1];
}

/* ------------------------------------------------------------ 財務諸表の特定 */
/**
 * 見出しの検出。
 * pdf.js はテキストを content stream の順で返すため、見出しがページ末尾に来ることがある。
 * したがって「先頭◯文字」ではなくページ全文から探す。
 */
// 「要約中間連結損益計算書」「連結財政状態計算書」「貸借対照表」など語順・修飾語の差を吸収する
const HEAD_RE = /(?:【([^】]{0,12}?)】?|([^\n]{0,12}?))(財政状態計算書|貸借対照表|損益計算書|キャッシュ・?フロー計算書)/g;

function headingsOf(pageText) {
  const out = [];
  HEAD_RE.lastIndex = 0;
  let m;
  while ((m = HEAD_RE.exec(pageText)) !== null) {
    // 直前12文字のうち、見出しに直接つながる修飾語だけを見る。
    // 「…包括利益計算書】【要約中間連結損益計算書」のように直前に別の見出しが
    // 来ることがあるため、最後の「】」より後ろだけを修飾語として扱う。
    let mod = (m[1] || m[2] || "");
    const cut = mod.lastIndexOf("】");
    if (cut >= 0) mod = mod.slice(cut + 1);
    mod = mod.replace(/^[^ぁ-んァ-ヶ一-龥]+/, "");     // 先頭の数字・記号を落とす
    // 「２．四半期連結財務諸表及び主な注記（１）四半期連結貸借対照表」のように、
    // 直前に別の見出しの語が来ることがある。「注記」「財務諸表」より後ろだけを修飾語とする。
    for (const cutWord of ["注記", "財務諸表等", "財務諸表"]) {
      const j = mod.lastIndexOf(cutWord);
      if (j >= 0) { mod = mod.slice(j + cutWord.length).replace(/^[^ぁ-んァ-ヶ一-龥]+/, ""); break; }
    }
    if (/包括利益$/.test(mod)) continue;               // 包括利益計算書はPLではない
    if (/計上額|について|に関する|における|場合|とき|より|ため/.test(mod)) continue;  // 本文中の言及を除く
    const kind = /財政状態|貸借/.test(m[3]) ? "BS" : /損益/.test(m[3]) ? "PL" : "CF";
    out.push({ kind, text: mod + m[3], consolidated: mod.includes("連結") });
  }
  return out;
}

/** 見出しではなく中身で財務諸表と判断できるか（見出しが別ページにある場合の保険） */
function bodyLooksLike(kind, labels) {
  const keys = new Set([...labels].map(keyOf));
  const has = (...ks) => ks.some((k) => keys.has(keyOf(k)));
  // 電気事業・ガス事業の貸借対照表は合計行を「合計」としか書かず、
  // 「流動資産合計」「資産合計」のいずれも存在しない。
  // 区分の見出し行だけで構成されるため、それを手掛かりに認める。
  // これが無いと、見出しは正しく取れているのに本文で弾かれて丸ごと欠測になる。
  if (kind === "BS" && has("流動資産") && has("固定資産", "非流動資産")) return true;
  if (kind === "BS") return has("流動資産合計", "非流動資産合計", "固定資産合計", "資産合計",
                                "流動負債合計", "非流動負債合計", "負債合計", "資本合計", "純資産合計",
                                // 決算公告は「合計」ではなく「〜の部合計」と書く様式がある
                                "資産の部合計", "負債の部合計", "純資産の部合計",
                                "負債・純資産合計", "負債純資産合計", "負債及び純資産合計");
  // 利益科目が無いものは損益計算書とみなさない。
  // 「売上高」だけを条件にすると、経営指標の推移表などを誤って拾う。
  if (kind === "PL") return has("営業利益", "営業損失", "営業利益又は営業損失", "経常利益", "当期経常利益",
                                "税引前利益", "税引前中間利益", "税引前当期純利益", "税金等調整前当期純利益",
                                "当期純利益", "当期純損失", "中間利益", "売上総利益");
  if (kind === "CF") return has("減価償却費", "減価償却費及びその他の償却費");
  return false;
}

/**
 * PDF全体を走査して BS/PL/CF を特定する。
 *
 * 重要な設計:
 *   ・連結を単体より優先する。混ざると貸借が合わなくなり、しかも数字は自然に見えるため気づけない
 *   ・BSは資産の部と資本の部で2ページに分かれることがある（IFRS・有報で頻出）。
 *     見出しページに続く「同じ表の続き」を結合する
 */
/* ------------------------------------------------- 会社名・金額単位の検出 */
const CO_KINDS = ["株式会社", "有限会社", "合同会社", "合資会社", "合名会社"];
// 会計用語や住所の断片が混ざっていたら、それは会社名ではなく本文の切れ端
const CO_NG = /資産|負債|利益|損失|費用|収益|合計|計算書|報告書|剰余金|余金|原価|税引|除却|配当|事項|営業外|注記|明細|科目|金額|株主資本|変動|附属|監査|決算|貸借|損益|現在|まで|から|支社|支店|本社|営業所|御中|様|[0-9]\s*[年月日]/;
// 提出会社ではないのに有価証券報告書の表紙に必ず載る名前。これを社名として拾わない。
const CO_EXCLUDE = /証券取引所|取引所|信託銀行|証券代行|監査法人|会計事務所|印刷|EDINET|縦覧/;
const CO_OK = /^[ぁ-んァ-ヶ一-龥ａ-ｚＡ-Ｚa-zA-Z0-9・ー]{2,20}$/;

/**
 * 決算書の金額単位を返す。決算書は必ず「（単位：千円）」等を明記している。
 * 本シートは百万円で計算するため、換算に使う。取り違えると規模と与信限度額が桁違いになる。
 */
export function detectUnit(pageTexts) {
  const t = pageTexts.join("");
  // ①「（単位：千円）」が最も確実。決算書・計算書類はほぼこの形
  let m = t.match(/単位[：:]?\s*([百千]?万?円)/);
  // ②有価証券報告書は「売上高(百万円)」「営業収益(百万円)」のように項目名に単位を付ける
  if (!m) m = t.match(/(?:売上高|営業収益|経常収益|純資産額|総資産額)[（(]([百千]?万?円)[）)]/);
  // 決算公告には「（単位：〜）」を書かず、金額欄の見出しに単位だけを置く様式がある。
  // 例:「科目 金額 科目 金額 円 円 流動資産 …」
  if (!m) m = t.match(/金[\s　]*額[）)]?[\s　]*[（(]?([百千]?万?円)/);
  // 単位の表示が一切なく、表の下の注記だけが単位を語っている決算公告がある。
  // 例:「記載金額は、千円未満の端数を切り捨てて表示しております。」
  // ここを読み落とすと桁が3つも6つもずれるため、必ず拾う。
  if (!m) m = t.match(/([百千]?万?円)未満(?:の端数)?[をは]?切(?:り)?捨/);
  if (!m) return null;
  const u = m[1];
  if (u === "百万円") return { label: "百万円", toMillion: 1 };
  if (u === "千円") return { label: "千円", toMillion: 1 / 1000 };
  if (u === "万円") return { label: "万円", toMillion: 1 / 100 };
  if (u === "円") return { label: "円", toMillion: 1 / 1000000 };
  return null;
}

/**
 * 会社名を返す。誤った名前を自動入力すると手入力より危険なので、
 * 確度が低い候補は採用せず null を返す（空欄のままにする）。
 * 採用条件は「2回以上出てくる」または「1ページ目の冒頭＝表題部にある」こと。
 */
export function detectCompany(pageTexts) {
  const t = pageTexts.join("");

  // ① 有価証券報告書の表紙は「【会社名】株式会社◯◯」と明記している。
  //    ここが読めれば推測は要らないので、最優先で採用する。
  // 本文の正規化で【】は既に外れているため、括弧なしの形で照合する。
  // 例：「…事業年度第13期(…)会社名株式会社マネーフォワード英訳名MoneyForward…」
  const KIND_RE = "株式会社|有限会社|合同会社|合資会社|合名会社";
  const labeled = (pageTexts[0] || "").match(new RegExp(
    "(?:会社名|商号|名称)\\s*((?:" + KIND_RE + ")[ぁ-んァ-ヶ一-龥ａ-ｚＡ-Ｚa-zA-Z0-9・ー]{1,20}" +
    "|[ぁ-んァ-ヶ一-龥ａ-ｚＡ-Ｚa-zA-Z0-9・ー]{1,20}(?:" + KIND_RE + "))"));
  if (labeled) {
    // 直後の見出し語（英訳名 など）を巻き込んでいたら落とす
    const n = labeled[1].replace(/(英訳名|代表者|本店|電話番号|事務連絡者|最寄り|縦覧|上場取引所|コード番号|URL|証券コード|決算期|問合せ先).*$/, "")
      // 有価証券報告書の表紙は社名の直後にEDINETコード（E04678）を置く。
      // ページの区切りで「E」だけが残ることがあるため、日本語のあとに続くEは落とす。
      .replace(/([ぁ-んァ-ヶ一-龥])E\d{0,6}$/, "$1");
    if (n.length >= 4) return n;
  }

  const cands = new Map();
  // 有価証券報告書の表紙は社名の直後にEDINETコード（E04678）を置く。
  // ページの区切りで「E」だけが残ることがあるため、日本語のあとに続くEは落とす。
  const add = (n0) => {
    const n = String(n0 || "").replace(/([ぁ-んァ-ヶ一-龥])E\d{0,6}$/, "$1");
    if (n) cands.set(n, (cands.get(n) || 0) + 1);
  };
  for (const kind of CO_KINDS) {
    let i = -1;
    while ((i = t.indexOf(kind, i + 1)) !== -1) {
      // 後置形（株式会社○○）：直後を名前として読む。次の会社種別・記号・数字で打ち切る
      let after = t.slice(i + kind.length, i + kind.length + 14)
                   .split(/[（(【］\]、。：:0-9０-９]/)[0];
      for (const k2 of CO_KINDS) { const p2 = after.indexOf(k2); if (p2 >= 0) after = after.slice(0, p2); }
      if (CO_OK.test(after) && !CO_NG.test(after) && !CO_EXCLUDE.test(kind + after)) add(kind + after);
      // 前置形（○○株式会社）：直前を名前として読む。長い候補から順に見る
      const before = t.slice(Math.max(0, i - 20), i);
      for (let s = 0; s < before.length; s++) {
        let cand = before.slice(s);
        // 「…でOTNet株式会社」のように直前の助詞を巻き込むのを防ぐ
        cand = cand.replace(/^[ぁ-ん]{1,2}(?=[ァ-ヶ一-龥A-Za-zＡ-Ｚ0-9])/, "");
        // 「…3月31日沖縄電力株式会社」のように日付の末尾文字を巻き込むのを防ぐ
        cand = cand.replace(/^[年月日期至自現在]+/, "");
        // 「有価証券報告書GMO…」のように直前の見出し語を巻き込むのを防ぐ
        cand = cand.replace(/^(?:有価証券報告書|報告書|書類|表紙|提出会社|会社名|商号|名称|当社|同社)+/, "");
        // 「…個別注記表浜銀ファイナンス」のように、直前の書類名を巻き込むのを防ぐ
        cand = cand.replace(/^.*(?:個別注記表|注記表|計算書類|貸借対照表|損益計算書|決算公告|決算報告書)/, "");
        // 「…牧港五丁目2番1号FRT」のように、直前の住所を巻き込むのを防ぐ
        cand = cand.replace(/^.*(?:丁目|番地|[0-9０-９]+番[0-9０-９]*号?|[0-9０-９]+号)/, "");
        // 住所やページ番号の名残りで先頭に数字が付くことがある（「…2番1号1東京海上ミレア…」）
        cand = cand.replace(/^[0-9０-９]+/, "");
        // 数字を落とすと日付の助数詞が先頭に出てくる（「…3月31日東京生活館」→「日東京生活館」）
        cand = cand.replace(/^[年月日期至自現在]+/, "");
        // 切り出した先頭が語の途中だと、別の語の断片を社名にしてしまう。
        // 例：「…個別注記表株式会社東京ドーム」から「記表株式会社」を作ってしまっていた。
        // 直前の文字が日本語なら、語の境界ではないので採用しない。
        {
          const at = i - (before.length - s);
          const prev = at > 0 ? t[at - 1] : "";
          if (prev && /[ぁ-んァ-ヶ一-龥]/.test(prev) && cand === before.slice(s)) continue;
        }
        if (CO_OK.test(cand) && !CO_NG.test(cand) && !CO_EXCLUDE.test(cand + kind)) { add(cand + kind); break; }
      }
      i += kind.length;
    }
  }
  // 1ページ目の表題部にある社名が最も信頼できる。有価証券報告書は必ずここに提出会社名を書く。
  // 本文には「株式会社を設立」のような文章の断片が何度も出るため、頻度だけで選ぶと負ける。
  const head = (pageTexts[0] || "").slice(0, 400);
  const inHead = [...cands.keys()].filter((k) => head.includes(k));
  if (inHead.length) {
    // 表題部に複数あるときは、
    //  ① 表題部での出現位置が最も後ろ（提出会社名は見出しの最後に書かれる）
    //  ② 同じ位置なら短いほう（余計な語を巻き込んでいない）
    // の順で選ぶ。長いほうを選ぶと「役員…縦覧に供する場所株式会社」のような塊を掴む。
    return inHead.sort((a, b) => {
      const d = head.indexOf(a) - head.indexOf(b);
      return d !== 0 ? d : a.length - b.length;
    })[0];
  }
  // 表題部で見つからないときだけ、2回以上出てくるものを採る
  let best = 0, company = null;
  for (const [k, v] of cands) if (v >= 2 && v > best) { best = v; company = k; }
  return company;
}

/**
 * 本ツールの判定モデルに載らない業態かどうかを返す。
 * 銀行・保険の決算書は、貸借対照表に流動／固定の区分が無く、
 * 損益計算書も売上高ではなく経常収益で構成されるため、そもそも様式が違う。
 * また業界基準に用いる統計も金融業・保険業を対象外としている。
 * 「読み取れませんでした」ではなく、対象外であることを伝えるために使う。
 */
export function detectOutOfScope(pageTexts) {
  const t = pageTexts.join("");
  const bank = /経常収益/.test(t) && /(?:預金|貸出金|コールローン|資金運用収益)/.test(t);
  if (bank) return "bank";
  // 「責任準備金」だけで保険業と決めてはいけない。
  // 退職給付の簡便法を説明する定型文
  //   「年金財政計算上の責任準備金を退職給付債務とする方法を用いた簡便法」
  // が中小企業の個別注記表に頻繁に載っており、正常な会社を対象外として
  // 門前払いしていた（実例：鉄道会社の計算書類）。
  // 保険業に固有の科目か、保険業であることが本文から明らかな場合に限る。
  const insHard = /(?:保険料等収入|支払備金|保険引受収益|保険引受費用|支払保険金)/.test(t);
  const insSoft = /責任準備金/.test(t) &&
                  /(?:保険業法|少額短期保険|生命保険業|損害保険業|保険会社|共済事業)/.test(t);
  if (insHard || insSoft) return "insurance";
  // 証券会社・金融商品取引業。損益計算書が営業収益と受入手数料で構成され、
  // 貸借対照表にも流動／固定の区分が無いため、本ツールの判定モデルに載らない。
  const sec = /(?:受入手数料|トレーディング損益|信用取引資産|信用取引負債|金融商品取引業)/.test(t);
  if (sec) return "securities";
  return null;
}

/**
 * 四半期・中間の決算書かどうかを、見つかった財務諸表の「見出し」で判定する。
 * 本文には有価証券報告書でも「四半期」の語が出るため、本文検索では誤判定する。
 * 四半期の損益計算書は3か月ぶんなので、年商として扱うと
 * 回転率も償還年数も与信限度額も静かに狂う。読めても判定してはいけない。
 */
export function detectInterim(found) {
  const heads = [(found && found.BS && found.BS.heading) || "",
                 (found && found.PL && found.PL.heading) || ""].join(" ");
  return /四半期|中間/.test(heads) ? "interim" : null;
}

/* ------------------------------------------------------------------ 決算期 */
/**
 * 決算期を特定する。決算書の様式によって書き方が違うため、次の順で探す。
 *   ① 会計期間「自2025年4月1日 至2026年3月31日」… 決算日が確実に分かる
 *   ② 貸借対照表日「2026年3月31日現在」
 *   ③ 期数「第47期」… 決算日は分からないが、期の前後関係は分かる
 * 戻り値の end は並べ替えに使う（新しいものが後ろ）。label は画面表示用。
 */
export function detectPeriod(pageTexts) {
  const t = pageTexts.join("");
  const out = { end: null, no: null, label: "" };

  // 決算書には前期と当期の2列があり、前期のほうが先に書かれている。
  // 最初の一致を採ると必ず1期古い日付になるため、見つかった中で最も新しいものを採る。
  const ends = [];
  const pad = (x) => String(x).padStart(2, "0");
  for (const mm of t.matchAll(/自\d{4}年\d{1,2}月\d{1,2}日至(\d{4})年(\d{1,2})月(\d{1,2})日/g))
    ends.push(`${mm[1]}-${pad(mm[2])}-${pad(mm[3])}`);
  if (!ends.length)
    for (const mm of t.matchAll(/(\d{4})年(\d{1,2})月(\d{1,2})日現在/g))
      ends.push(`${mm[1]}-${pad(mm[2])}-${pad(mm[3])}`);
  // 決算公告には和暦で書かれたものがある（令和7年12月31日現在）
  if (!ends.length)
    for (const mm of t.matchAll(/(令和|平成)(\d{1,2})年(\d{1,2})月(\d{1,2})日現在/g)) {
      const base = mm[1] === "令和" ? 2018 : 1988;
      ends.push(`${base + Number(mm[2])}-${pad(mm[3])}-${pad(mm[4])}`);
    }
  if (ends.length) out.end = ends.sort()[ends.length - 1];

  let m = t.match(/第(\d{1,3})期/);
  if (m) out.no = parseInt(m[1], 10);

  if (out.end) {
    const [y, mo] = out.end.split("-");
    out.label = `${y}年${parseInt(mo, 10)}月期`;
    if (out.no) out.label = `第${out.no}期（${out.label}）`;
  } else if (out.no) {
    out.label = `第${out.no}期`;
  }
  return out;
}

export async function scanPdf(pdfjs, buf, onProgress) {
  // cMap（文字コード変換表）を必ず渡す。
  // 有価証券報告書は MS-Gothic 等のCIDフォントを埋め込みつつ ToUnicode を持たないものが多く、
  // これが無いと文字が1文字も取り出せず「画像PDF」と誤判定してしまう。
  const doc = await pdfjs.getDocument({
    data: new Uint8Array(buf),
    useWorkerFetch: false,
    isEvalSupported: false,
    cMapUrl: new URL("../vendor/cmaps/", import.meta.url).href,
    cMapPacked: true,
  }).promise;
  const pages = [];
  const pageTexts = [];
  let totalChars = 0;

  for (let i = 1; i <= doc.numPages; i++) {
    if (onProgress) onProgress(i, doc.numPages);
    const page = await doc.getPage(i);
    const vp = page.getViewport({ scale: 1 });
    const tc = await page.getTextContent();
    const words = itemsToWords(tc.items, vp.height);
    const pageText = norm(words.map((w) => w.text).join(""));
    totalChars += pageText.length;
    if (pageTexts.length < 12) pageTexts.push(pageText);
    const { rows, kind } = words.length ? rowsOf(words, vp.width, pageText) : { rows: [], kind: "single" };
    pages.push({ no: i, rows, kind, text: pageText, labels: new Set(rows.map((r) => r.label)),
                 headings: headingsOf(pageText) });
  }
  await doc.destroy();

  // テキストレイヤーが無いPDF（スキャン画像）は、その旨を明示して返す
  if (totalChars < 200) return { _noText: true };

  // 決算期を拾う。複数のPDFを投入されたとき、新しい順に並べ替えるために使う。
  // 読めなくても処理は続ける（並べ替えを諦めて投入順にするだけ）。
  const period = detectPeriod(pageTexts);
  const outOfScope = detectOutOfScope(pageTexts);
  const company = detectCompany(pageTexts);

  const found = {};
  // 減価償却費を全ページから探すために保持する。列挙対象に混ざらないよう非列挙にする。
  Object.defineProperty(found, "_pages", { value: pages, enumerable: false });
  Object.defineProperty(found, "_period", { value: period, enumerable: false });
  Object.defineProperty(found, "_company", { value: company, enumerable: false });
  Object.defineProperty(found, "_outOfScope", { value: outOfScope, enumerable: false });
  for (const wantConsolidated of [true, false]) {          // 連結を先に探す
    for (const kind of ["BS", "PL", "CF"]) {
      if (found[kind]) continue;
      for (const pg of pages) {
        const h = pg.headings.find((x) => x.kind === kind && x.consolidated === wantConsolidated);
        if (!h || !bodyLooksLike(kind, pg.labels)) continue;
        found[kind] = { page: pg.no, rows: pg.rows.slice(), kind: pg.kind,
                        consolidated: h.consolidated, heading: h.text };
        // 損益計算書が「特別損失合計」で切れ、税金と当期純利益だけが次ページに載る様式がある。
        // 無条件に結合すると次ページの包括利益計算書を飲み込むため、
        // ①見出しが無い ②包括利益の科目を含まない ③税金か純利益の科目を含む
        // の3つを満たすページに限って、1ページだけ結合する。
        if (kind === "PL") {
          const taxRows = ["法人税等合計", "法人税等", "税金等調整前当期純利益", "税引前当期純利益",
                           "当期純利益", "親会社株主に帰属する当期純利益"];
          const hasTail = (labels) => taxRows.some((k) => labels.has(k));
          if (!hasTail(pg.labels)) {
            const nx = pages[pg.no];               // 次のページ（pagesは0起点）
            const isComprehensive = (labels) =>
              [...labels].some((k) => /包括利益/.test(k));
            if (nx && !nx.headings.length && nx.rows.length &&
                !isComprehensive(nx.labels) && hasTail(nx.labels)) {
              found.PL.rows.push(...nx.rows);
              found.PL.continued = [nx.no];
            }
          }
        }
        // 続きのページを結合する。
        // 貸借対照表だけが「資産の部」と「資本の部」で2ページに割れるため、BSに限定する。
        if (kind === "BS") {
          const need = ["資本合計", "純資産合計", "負債及び資本合計", "負債純資産合計", "負債及び純資産合計"];
          const hasEquity = (labels) => need.some((k) => labels.has(k));
          for (let n = pg.no; n < pages.length && !hasEquity(new Set(found.BS.rows.map((r) => r.label))); n++) {
            const nx = pages[n];
            if (!nx || nx.headings.length || !nx.rows.length) break;
            if (!bodyLooksLike("BS", nx.labels)) break;
            found.BS.rows.push(...nx.rows);
            found.BS.continued = (found.BS.continued || []).concat(nx.no);
          }
        }
        break;
      }
    }
    if (found.BS && found.PL) break;   // 連結で揃ったら単体は見ない
  }

  // 金額の単位は「決算書そのもののページ」から採る。
  // 全文の先頭から探すと、為替レート表の「（単位：円）」のような
  // 決算書と無関係な記載を拾い、百万円を円と取り違える。桁が6つずれる。
  const unitPages = [];
  for (const k of ["BS", "PL"]) {
    if (!found[k]) continue;
    const nos = [found[k].page, ...(found[k].continued || [])];
    for (const n of nos) { const pg = pages.find((x) => x.no === n); if (pg) unitPages.push(pg.text); }
  }
  const unit = detectUnit(unitPages.length ? unitPages : pageTexts) || detectUnit(pageTexts);
  Object.defineProperty(found, "_unit", { value: unit, enumerable: false });
  Object.defineProperty(found, "_interim", { value: detectInterim(found), enumerable: false });

  return found;
}

/** 負債の部を流動セクション／固定（非流動）セクションに切り分ける */
function sliceSection(rows, which) {
  const find = (re) => rows.findIndex((r) => re.test(keyOf(r.label)));

  // ① 合計行がある様式（有価証券報告書など）。合計行が節の末尾になる。
  const endCur = find(/^流動負債合計$/);
  const endFix = find(/^(固定負債合計|非流動負債合計)$/);
  if (endCur >= 0 && endFix > endCur) {
    const i = find(/^流動負債$/);
    const startCur = Math.max(0, i >= 0 && i < endCur ? i : endCur - 20);
    return which === "current"
      ? rows.slice(startCur, endCur + 1)
      : rows.slice(endCur + 1, endFix + 1);
  }

  // ② 合計行が無く、「流動負債」「固定負債」の見出し行だけがある様式（決算公告に多い）。
  //    区切らないまま全体を渡すと、流動と固定の両方に同じ名前で載る「リース債務」等を
  //    二重に拾い、有利子負債が実際の倍近くになる。数字は自然に見えるので気づけない。
  const hCur = find(/^流動負債$/);
  const hFix = find(/^(固定負債|非流動負債)$/);
  if (hCur >= 0 && hFix > hCur) {
    const hTot = find(/^(負債合計|負債の部合計|負債の部計|負債計)$/);
    return which === "current"
      ? rows.slice(hCur, hFix)
      : rows.slice(hFix, hTot > hFix ? hTot : rows.length);
  }

  // ③ 固定負債を先に書く様式（電気事業・ガス事業の貸借対照表）。
  //    「固定負債 … 流動負債 … 負債合計」の順で、合計行を持たない。
  //    ②は流動が先にある前提なので判別できず、全体を返してしまい、
  //    固定側の社債・長期借入金を短期有利子負債にも計上していた。
  //    実例：沖縄電力で短期有利子負債 320,460（流動負債 86,246）という値になっていた。
  if (hFix >= 0 && hCur > hFix) {
    const hTot = find(/^(負債合計|負債の部合計|負債の部計|負債計)$/);
    return which === "fixed"
      ? rows.slice(hFix, hCur)
      : rows.slice(hCur, hTot > hCur ? hTot : rows.length);
  }

  return rows;   // 判別できなければ全体を対象
}

/* -------------------------------------------------------- 1期分の組み立て */

/**
 * 損益計算書の「区分合計」を、内訳の右隣にある1つ多い金額から拾う。
 *
 * 中小企業の計算書類や決算公告には、こういう様式がある。
 *
 *     営業外収益
 *       受取利息及び配当金   1,162
 *       その他                  21     1,265   ← 区分合計は内訳の右隣
 *     営業外費用
 *       支払利息                58
 *       その他                   4        62
 *
 * 「営業外収益」という見出し行そのものに金額が無いため、辞書では拾えない。
 * そこで、内訳より金額が1つ多い行を区分合計とみなす。
 *
 * ただし推測なので、必ず算術で検算してから採用する。
 *   営業利益 ＋ 営業外収益 − 営業外費用 ＝ 経常利益
 *   経常利益 ＋ 特別利益   − 特別損失   ＝ 税引前当期純利益
 * 合わなければ何も入れない。間違った値を静かに入れるくらいなら、未取得のままがよい。
 */
function inferPlSections(rows, out, source, yi) {
  if (!rows.length) return;
  const val = (r) => (yi === -1 ? r.amounts[r.amounts.length - 1]
                                : (r.amounts.length >= 2 ? r.amounts[r.amounts.length - 2]
                                                         : r.amounts[r.amounts.length - 1]));
  // 内訳行がいくつ金額を持つか（最頻値）。区分合計はそれより1つ多い。
  const freq = new Map();
  for (const r of rows) freq.set(r.amounts.length, (freq.get(r.amounts.length) || 0) + 1);
  let base = 1, best = -1;
  for (const [n, c] of freq) if (c > best) { best = c; base = n; }

  const idx = (names) => rows.findIndex((r) => names.includes(keyOf(r.label)));
  const iOp  = idx(["営業利益", "営業損失"]);
  const iOrd = idx(["経常利益", "経常損失"]);
  const iPre = idx(["税引前当期純利益", "税引前当期純損失", "税金等調整前当期純利益", "税引前利益"]);

  const subtotalsBetween = (a, b) => {
    if (a < 0 || b < 0 || b <= a) return [];
    return rows.slice(a + 1, b).filter((r) => r.amounts.length > base);
  };
  const near = (x, y) => Math.abs(x - y) <= Math.max(2, Math.abs(y) * 0.005);

  /**
   * 区間内の区分合計から、収益側と費用側を決める。
   * 2つあれば「前が収益・後が費用」。
   * 1つしか無い様式（特別利益が無く特別損失だけ、など）もあるため、
   * その場合は収益とみなす場合・費用とみなす場合の両方を検算し、合うほうを採る。
   * どちらも合わなければ何も入れない。
   */
  const solve = (from, to, target, incKey, expKey, incLabel, expLabel) => {
    if (out[incKey] !== null || out[expKey] !== null) return;
    if (target === null || target === undefined) return;
    const sub = subtotalsBetween(from, to);
    const start = from === iOp ? out.operating : out.ordinary;
    if (start === null || start === undefined) return;
    if (sub.length === 2) {
      const inc = val(sub[0]), exp = val(sub[1]);
      if (near(start + inc - exp, target)) {
        out[incKey] = inc; source[incKey] = incLabel + "（区分合計より）";
        out[expKey] = exp; source[expKey] = expLabel + "（区分合計より）";
      }
      return;
    }
    if (sub.length === 1) {
      const x = val(sub[0]);
      if (near(start - x, target)) {          // 費用だけがある
        out[expKey] = x; source[expKey] = expLabel + "（区分合計より）";
        out[incKey] = 0; source[incKey] = incLabel + "（記載なし）";
      } else if (near(start + x, target)) {   // 収益だけがある
        out[incKey] = x; source[incKey] = incLabel + "（区分合計より）";
        out[expKey] = 0; source[expKey] = expLabel + "（記載なし）";
      }
      return;
    }
  };

  // ① 営業外収益・営業外費用
  solve(iOp, iOrd, out.ordinary, "nonOpInc", "nonOpExp", "営業外収益", "営業外費用");

  // ② 特別利益・特別損失
  if (iPre >= 0) {
    solve(iOrd, iPre, val(rows[iPre]), "extraInc", "extraExp", "特別利益", "特別損失");
  }

  // ③ 売上高・売上原価。
  // リース業などには、売上を内訳だけで並べ「売上高」の行を置かない様式がある。
  //     リース売上高   29,079
  //     割賦売上高        629
  //     その他の売上高      8   31,116   ← 区分合計は内訳の右隣
  // 売上高が取れないと回転率も月商倍率も出せないため、区分合計から推定する。
  // 売上総利益（無ければ営業利益）と突き合わせて合う場合にだけ採用する。
  if (out.sales === null && out.cogs === null) {
    const iGross = idx(["売上総利益", "売上総利益合計", "営業総利益"]);
    const end = iGross >= 0 ? iGross : iOp;
    const target = iGross >= 0 ? val(rows[iGross]) : out.operating;
    if (end > 0 && target !== null && target !== undefined) {
      const sub = rows.slice(0, end).filter((r) => r.amounts.length > base);
      if (sub.length === 2) {
        const inc = val(sub[0]), exp = val(sub[1]);
        if (near(inc - exp, target)) {
          out.sales = inc; source.sales = "営業収益の区分合計より";
          out.cogs = exp; source.cogs = "営業費用の区分合計より";
          // 営業利益で突き合わせた場合、費用側の合計に販管費まで含まれている。
          // そのまま販管費を別に引くと二重計上になるので、費用側から取り除く。
          if (iGross < 0) {
            if (out.sga !== null && out.sga !== undefined) {
              out.cogs = exp - out.sga;
              source.cogs = "営業費用の区分合計−販売費及び一般管理費";
            } else {
              out.sga = 0; source.sga = "営業費用に含まれるため区分なし";
            }
          }
        }
      }
    }
  }
}

/**
 * 事業別に損益を並べる様式で、全事業の合計に差し替える。
 * 合算した収益と費用の差が「全事業営業利益」と一致する場合にだけ採用する。
 * 合わなければ何も変えない。
 */
function mergeSegmentPl(rows, out, source, yi, warnings) {
  const val = (r) => (yi === -1 ? r.amounts[r.amounts.length - 1]
                                : (r.amounts.length >= 2 ? r.amounts[r.amounts.length - 2]
                                                         : r.amounts[r.amounts.length - 1]));
  const total = rows.find((r) => /^全事業(営業利益|営業損失)$/.test(keyOf(r.label)));
  if (!total) return;
  const totalOp = val(total);
  const incRows = rows.filter((r) => ["営業収益", "売上高"].includes(keyOf(r.label)));
  const expRows = rows.filter((r) => ["営業費", "営業費用", "売上原価"].includes(keyOf(r.label)));
  if (incRows.length < 2 || !expRows.length) return;
  const inc = incRows.reduce((a, r) => a + val(r), 0);
  const exp = expRows.reduce((a, r) => a + val(r), 0);
  if (Math.abs(inc - exp - totalOp) > Math.max(2, Math.abs(inc) * 0.0005)) return;
  out.sales = inc;      source.sales = `全事業の営業収益（${incRows.length}事業の合計）`;
  out.sga = exp;        source.sga = `全事業の営業費（${expRows.length}事業の合計）`;
  out.cogs = 0;
  out.operating = totalOp; source.operating = "全事業営業利益";
  warnings.push("事業別に区分された損益計算書のため、全事業を合算しました。この様式には売上原価と販管費の区分がないため、営業費の全額を販管費として扱っています。");
}

export function buildPeriod(found, yi) {
  const out = {}, source = {}, warnings = [];
  const bs = found.BS?.rows || [], pl = found.PL?.rows || [], cf = found.CF?.rows || [];

  for (const [k, spec] of Object.entries(BS_DICT)) {
    // 有利子負債は「リース負債」のように流動・非流動で同名の科目があるため、
    // 区間を限定して拾わないと二重計上になる
    const scope = k === "shortDebt" ? sliceSection(bs, "current")
                : k === "longDebt" ? sliceSection(bs, "fixed") : bs;
    const { value, source: s } = pull(scope, spec, yi);
    out[k] = value; if (s) source[k] = s;
  }
  for (const [k, spec] of Object.entries(PL_DICT)) {
    const { value, source: s } = pull(pl, spec, yi);
    out[k] = value; if (s) source[k] = s;
  }

  // 事業別に損益を積む様式（鉄道・不動産など）。
  //   鉄道事業  営業収益 6,791,578 / 営業費 6,346,123 / 営業利益 445,454
  //   不動産事業 営業収益   814,538 / 営業費   545,008 / 営業利益 269,529
  //   全事業   営業利益   714,984
  // 同じ科目名が繰り返されるため、そのままでは最初の事業の数字だけを
  // 全社の売上高・営業利益として取り込んでしまう。
  // 実例：売上高が 7,606,116 のところ 6,791,578（−10.7%）、
  //       営業利益が 714,984 のところ 445,454（−37.7%）になっていた。
  // 「全事業営業利益」がある場合に限り、各事業を合算して差し替える。
  mergeSegmentPl(pl, out, source, yi, warnings);

  // 区分の見出し行に金額が無い様式では、内訳の右隣にある区分合計を推測する（検算つき）
  inferPlSections(pl, out, source, yi);

  // 「流動負債合計」「固定負債合計」を置かず、負債合計だけを書く様式への手当て。
  // 区分が両方とも空のままだと貸借が合わず、判定にも進めない。
  // 負債合計から固定負債（判明していれば）を引いた残りを流動負債とみなす。
  // 流動負債を多めに見ることになるが、流動比率は低く出る方向なので与信では安全側。
  if (out.currentLiab === null) {
    const liab = pull(bs, { agg: ["負債合計", "負債の部合計", "負債の部計", "負債計"], parts: [] }, yi);
    if (liab.value !== null) {
      out.currentLiab = liab.value - (out.fixedLiab || 0);
      source.currentLiab = "負債合計−固定負債（推定）";
      warnings.push("流動負債と固定負債の区分がこの決算書には記載されていないため、負債合計から推定しました。区分の内訳が必要な場合は手入力で補ってください。");
    }
  }

  // 売上原価の内訳（期首棚卸＋仕入−期末棚卸）だけを並べ、「売上原価」の行を置かない様式。
  // 売上総利益が読めていれば差額で確定できる。推測ではなく恒等式なので安全。
  if (out.cogs === null && out.sales !== null) {
    const g = pull(pl, GROSS, yi);
    if (g.value !== null) {
      out.cogs = out.sales - g.value;
      source.cogs = "売上高−売上総利益";
    }
  }

  // 小売業には「売上高」と「営業収入」を並べ、その合算を営業収益とする様式がある。
  // 売上高だけを採ると、営業利益が積み上がらない。
  // 実例：サンエーで 245,547 のところを 225,485 にしていた（営業収入 20,062 が抜けていた）。
  // 足したほうが営業利益の計算に合う場合にだけ採用する。
  if (out.sales !== null && out.cogs !== null && out.sga !== null && out.operating !== null) {
    const gap = out.sales - out.cogs - out.sga - out.operating;
    if (Math.abs(gap) > Math.max(2, Math.abs(out.sales) * 0.0005)) {
      const oi = pull(pl, OPINCOME, yi);
      if (oi.value !== null && Math.abs(out.sales + oi.value - out.cogs - out.sga - out.operating)
                               <= Math.max(2, Math.abs(out.sales) * 0.0005)) {
        out.sales += oi.value;
        source.sales = (source.sales || "売上高") + "＋" + oi.source;
      }
    }
  }

  // 総資本の行が「合計」としか書かれない様式（電気事業など）。
  // 資産側の区分が揃っていれば、その和が総資本と一致する。
  if (out.totalCapital === null && out.currentAssets !== null && out.fixedAssets !== null) {
    out.totalCapital = out.currentAssets + out.fixedAssets + (out.deferred || 0);
    source.totalCapital = "流動資産＋固定資産（合計行なし）";
  }

  // 減価償却費：CF計算書 →（無ければ）損益計算書 →（無ければ）全ページ の順に探す。
  // 中小企業の計算書類にはCF計算書が無く、販売費及び一般管理費明細書や
  // 製造原価報告書に載っていることが多いため、最後は全ページを走査する。
  let dep = pull(cf, DEP, yi);
  if (dep.value === null) {
    dep = pull(pl, DEP, yi);
    if (dep.value !== null) dep.source += "（損益計算書より）";
  }
  if (dep.value === null && Array.isArray(found._pages)) {
    // 販管費明細と製造原価報告書の両方にある場合は合算する（両方が費用計上分）
    const hits = [];
    for (const pg of found._pages) {
      if (found.CF && pg.no === found.CF.page) continue;
      if (found.PL && pg.no === found.PL.page) continue;
      const r = pg.rows.find((x) => DEP.agg.includes(x.label));
      if (r) {
        const a = r.amounts;
        const v = yi === -1 ? a[a.length - 1] : (a.length >= 2 ? a[a.length - 2] : a[a.length - 1]);
        if (typeof v === "number") hits.push({ page: pg.no, value: v });
      }
    }
    if (hits.length) {
      dep = { value: hits.reduce((a, h) => a + h.value, 0),
              source: hits.map((h) => `${h.page}ページ`).join("＋") + "（明細より）" };
      if (hits.length > 1) {
        warnings.push("減価償却費を複数ページから合算しました（" +
          hits.map((h) => `${h.page}ページ ${h.value.toLocaleString()}`).join(" ＋ ") +
          "）。二重計上でないかご確認ください。");
      }
    }
  }
  // それでも見つからないときは、注記の本文から拾う。
  // 中小企業の計算書類は「損益計算書に関する注記」に
  //   （3）減価償却費 2,467,046千円
  // と文章で書くことがあり、表になっていないため行としては取れない。
  // 単位付きで書かれているものだけを採り、決算書本体の単位へ換算する。
  if (dep.value === null && Array.isArray(found._pages) && found._unit) {
    const RE = /減価償却費(?:及び[^0-9]{0,10})?[^0-9△▲]{0,4}([△▲]?[\d,]+)(円|千円|百万円)/;
    const scale = { "円": 0.000001, "千円": 0.001, "百万円": 1 };
    for (const pg of found._pages) {
      const m = RE.exec(pg.text);
      if (!m) continue;
      const raw = toNum(m[1]);
      if (raw === null || raw === 0) continue;
      const v = Math.round(raw * scale[m[2]] / found._unit.toMillion);
      dep = { value: v, source: `${pg.no}ページの注記より` };
      warnings.push("減価償却費は注記の本文から読み取りました。金額をご確認ください。");
      break;
    }
  }
  out.depreciation = dep.value;
  if (dep.source) source.depreciation = dep.source;
  if (dep.value === null)
    warnings.push("減価償却費が見つかりませんでした。償還余力の算定に必要です。製造原価明細・販管費明細からご入力ください。");

  // 売上原価／販管費の区分が無い様式（電気事業・通信事業など）
  if (out.cogs === null && out.sga === null) {
    for (const name of OPEX) {
      const r = pl.find((x) => keyOf(x.label) === keyOf(name));
      if (r) {
        // 前期を取るときに左端から数えてはいけない。
        // 有価証券報告書には、隣の表の数値が左に1つ紛れ込む行があり、
        // そこを前期の金額として拾ってしまう。
        // 実例：東京電力の営業費用が 6,575,938 のところ 2 になっていた。
        const a = r.amounts;
        out.sga = yi === -1 ? a[a.length - 1] : (a.length >= 2 ? a[a.length - 2] : a[a.length - 1]);
        out.cogs = 0;
        source.sga = name + "（売上原価との区分なし）";
        warnings.push("この様式には売上原価と販管費の区分がありません。営業費用の全額を販管費として扱っています。売上総利益率は実態を表しません。");
        break;
      }
    }
  }

  // IFRSの損益計算書は費用を「△835,371」のようにマイナス表記する。
  // 本シートは費用を正の数で扱うため、符号を揃える。
  // 反転する前に「費用を負で書く様式かどうか」を覚えておく。
  const negStyle = ["cogs", "sga", "nonOpExp", "extraExp"]
    .some((k) => typeof out[k] === "number" && out[k] < 0);
  for (const k of ["cogs", "sga", "nonOpExp", "extraExp"]) {
    if (typeof out[k] === "number" && out[k] < 0) out[k] = -out[k];
  }
  // 法人税等だけは別扱いにする。
  // 税額は還付超過で本当にマイナスになることがあり、それを正に反転すると
  // 当期純利益が積み上がらなくなる（実例：法人税等 △2,868 を ＋2,868 にしていた）。
  // 費用を負で書く様式のときにだけ反転する。
  if (typeof out.tax === "number" && out.tax < 0 && negStyle) out.tax = -out.tax;
  // 営業外収益がマイナスで返るケース（その他費用を含めて相殺された場合）も正に寄せる
  if (typeof out.nonOpInc === "number" && out.nonOpInc < 0) {
    out.nonOpExp = (out.nonOpExp || 0) + Math.abs(out.nonOpInc);
    out.nonOpInc = 0;
  }

  // 「その他」は小計からの差額で埋める。こうすると小計が必ず一致する
  const d = (total, ...items) =>
    out[total] === null || out[total] === undefined
      ? null
      : out[total] - items.reduce((a, k) => a + (out[k] || 0), 0);
  out.otherCurrentAssets = d("currentAssets", "cash", "receivables", "inventory");
  out.otherFixedAssets = d("fixedAssets", "tangible");
  out.otherCurrentLiab = d("currentLiab", "payables", "shortDebt");
  out.otherFixedLiab = d("fixedLiab", "longDebt");

  // 決算書自体の端数処理で、資産側と負債側が1〜数単位ずれることがある。
  // 重要性のない差は「その他流動資産」で吸収して検算を0にする。
  // 吸収しないと、正しく読み取れているのに赤い警告が出て利用者を不安にさせる。
  {
    const a = (out.currentAssets || 0) + (out.fixedAssets || 0) + (out.deferred || 0);
    const l = (out.currentLiab || 0) + (out.fixedLiab || 0) + (out.equity || 0);
    const gap = a - l;
    const tol = Math.max(2, Math.abs(a) * 0.0001);   // 総資産の0.01%まで
    if (gap !== 0 && Math.abs(gap) <= tol && out.otherCurrentAssets !== null) {
      out.otherCurrentAssets -= gap;
      out.currentAssets -= gap;
      source.roundingAdjust = `端数調整 ${gap > 0 ? "−" : "＋"}${Math.abs(gap)}`;
    }
  }

  // 固定負債が無い会社（決算公告の小規模会社に多い）は、固定負債の行そのものが存在しない。
  // 総資本・流動負債・純資産が揃っていれば差額で埋められる。埋めないと区分が欠けたまま判定に入る。
  if (out.fixedLiab === null && out.totalCapital !== null &&
      out.currentLiab !== null && out.equity !== null) {
    const rest = out.totalCapital - out.currentLiab - out.equity;
    if (rest >= 0 && rest <= Math.abs(out.totalCapital) * 0.02) {
      out.fixedLiab = rest;
      source.fixedLiab = "差額（総資本−流動負債−純資産）";
    }
  }
  // 電気事業・鉄道事業の貸借対照表は「有形固定資産」という行を置かず、
  // 「電気事業固定資産」「鉄軌道事業固定資産」のように事業別に並べる。
  // 内訳を足し合わせると、その下位項目まで拾って二重計上になる（例：その他の電気事業固定資産）。
  // 固定資産から投資その他の資産と無形固定資産を引く恒等式のほうが確実。
  if (out.tangible === null && out.fixedAssets !== null) {
    const hasBiz = bs.some((r) => /事業固定資産$/.test(keyOf(r.label)));
    const inv = pull(bs, { agg: ["投資その他の資産", "投資その他の資産合計", "投資その他資産"], parts: [] }, yi);
    const intan = pull(bs, { agg: ["無形固定資産", "無形固定資産合計"], parts: [] }, yi);
    if (hasBiz && inv.value !== null) {
      out.tangible = out.fixedAssets - inv.value - (intan.value || 0);
      source.tangible = "固定資産−投資その他の資産" + (intan.value ? "−無形固定資産" : "");
    }
  }

  if (out.tangible === null && out.fixedAssets !== null)
    warnings.push("有形固定資産の内訳が特定できませんでした。固定資産合計は正しいですが、有形と無形の内訳はご確認ください。");
  for (const k of ["shortDebt", "longDebt"]) {
    if (out[k] === null) {
      out[k] = 0;
      warnings.push("有利子負債に該当する科目が見つかりませんでした。無借金でない場合は必ずご入力ください。");
      break;
    }
  }
  return { values: out, source, warnings };
}

/** 独立に抽出した資産側小計と負債側小計を突き合わせる */
export function validatePeriod(v) {
  const a = (v.currentAssets || 0) + (v.fixedAssets || 0) + (v.deferred || 0);
  const l = (v.currentLiab || 0) + (v.fixedLiab || 0) + (v.equity || 0);
  const msgs = [];
  const has = (...ks) => ks.every((k) => v[k] !== null && v[k] !== undefined);
  const jp = (n) => Number(n).toLocaleString();

  if (Math.abs(a - l) > Math.max(2, Math.abs(a) * 0.0001))
    msgs.push(`貸借が一致しません（資産計 ${jp(a)} と 負債純資産計 ${jp(l)} の差 ${jp(a - l)}）`);

  // 経常利益の積み上がり
  if (has("ordinary", "operating", "nonOpInc", "nonOpExp")) {
    const calc = v.operating + v.nonOpInc - v.nonOpExp;
    if (Math.abs(calc - v.ordinary) > tol(v.ordinary))
      msgs.push(`経常利益が積み上がりません（計算 ${jp(calc)} と 記載 ${jp(v.ordinary)}）`);
  }

  // 営業利益の積み上がり。
  // 売上原価と販管費の区分が無い様式では cogs=0 として扱っているため、そこも通る。
  if (has("sales", "cogs", "sga", "operating", "ordinary")) {   // IFRSは経常利益が無く、営業利益にその他収益/費用を含むため対象外
    const calc = v.sales - v.cogs - v.sga;
    if (Math.abs(calc - v.operating) > tol(v.operating, v.sales))
      msgs.push(`営業利益が積み上がりません（売上高−売上原価−販管費 ${jp(calc)} と 記載 ${jp(v.operating)}）`);
  }

  // 当期純利益までの積み上がり。
  // ここに検算が無かったため、法人税等の二重計上や事業別損益の取りこぼしが
  // 貸借の一致だけをすり抜けて、そのまま判定に流れ込んでいた。
  if (has("ordinary", "extraInc", "extraExp", "tax", "net")) {
    const calc = v.ordinary + v.extraInc - v.extraExp - v.tax;
    if (Math.abs(calc - v.net) > tol(v.net, v.ordinary))
      msgs.push(`当期純利益が積み上がりません（経常利益＋特別損益−法人税等 ${jp(calc)} と 記載 ${jp(v.net)}）`);
  }

  // 区分をまたいだ二重計上の検出。
  // 有利子負債は流動負債・固定負債の内数なので、区分の合計を超えることはない。
  // 超えていれば、同名科目を流動と固定の両方で拾っている。
  if (has("shortDebt", "currentLiab") && v.shortDebt > v.currentLiab * 1.001 + 2)
    msgs.push(`短期有利子負債 ${jp(v.shortDebt)} が流動負債 ${jp(v.currentLiab)} を超えています（区分の重複の可能性）`);
  if (has("longDebt", "fixedLiab") && v.longDebt > v.fixedLiab * 1.001 + 2)
    msgs.push(`長期有利子負債 ${jp(v.longDebt)} が固定負債 ${jp(v.fixedLiab)} を超えています（区分の重複の可能性）`);

  return { diff: a - l, messages: msgs };
}

/** 決算書側の端数処理を吸収する許容差。基準額の0.5%か、規模の0.05%の大きいほう */
function tol(base, scale) {
  return Math.max(2, Math.abs(base || 0) * 0.005, Math.abs(scale || 0) * 0.0005);
}

/** エンジン入力のキー名に合わせて1期分を返す */
export function toEngineFields(v) {
  return {
    sales: v.sales, cogs: v.cogs, sga: v.sga,
    nonOpInc: v.nonOpInc, nonOpExp: v.nonOpExp,
    extraInc: v.extraInc, extraExp: v.extraExp, tax: v.tax,
    depreciation: v.depreciation,
    cash: v.cash, receivables: v.receivables, inventory: v.inventory,
    otherCurrentAssets: v.otherCurrentAssets, tangible: v.tangible,
    otherFixedAssets: v.otherFixedAssets, deferred: v.deferred,
    payables: v.payables, shortDebt: v.shortDebt,
    otherCurrentLiab: v.otherCurrentLiab, longDebt: v.longDebt,
    otherFixedLiab: v.otherFixedLiab, equity: v.equity,
  };
}
