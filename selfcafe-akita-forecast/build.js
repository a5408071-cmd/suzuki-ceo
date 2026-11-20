/**
 * 未来屋書店 秋田店　開業後3か年の売上見込み（30坪・業務委託型）
 * A4横 3ページ／濃緑バナー＋ブロンズの新意匠
 */
const pptxgen = require("pptxgenjs");
const path = require("path");

/* ---------------- palette / type ---------------- */
const C = {
  bg: "FFFFFF",
  deep: "0B4A2F",      // ヘッダー帯
  green: "16764B",     // 主グリーン
  mint: "C9E0D4",      // 淡グリーン
  tint: "EEF5F1",
  bronze: "B5812C",    // アクセント
  bronzeT: "F8F1E2",
  ink: "16201B",
  body: "3A423E",
  muted: "7C8580",
  rule: "D9DFDB",
  soft: "F4F7F5",
  white: "FFFFFF",
};
const F = { jp: "Yu Gothic", num: "Arial" };
const HAS_JP = /[ぁ-んァ-ヶ一-龥々〜ー％円杯年月日]/;
const nf = (s) => (HAS_JP.test(String(s)) ? F.jp : F.num);

const W = 11.69, H = 8.27, M = 0.5, CW = W - M * 2, R = M + CW;
const BAND = 0.92;

const pres = new pptxgen();
pres.defineLayout({ name: "A4L", width: W, height: H });
pres.layout = "A4L";
pres.author = "セルフカフェ株式会社";
pres.title = "秋田店 開業後3か年の売上見込み";

let PAGE = 0;
const TOTAL = 3;

/* ---------------- chrome ---------------- */
function head(s, title, sub) {
  PAGE++;
  s.background = { color: C.bg };
  s.addShape("rect", { x: 0, y: 0, w: W, h: BAND, fill: { color: C.deep } });
  s.addShape("rect", { x: 0, y: BAND, w: W, h: 0.045, fill: { color: C.bronze } });
  s.addShape("rect", { x: M, y: 0.28, w: 0.075, h: 0.38, fill: { color: C.bronze } });
  s.addText(title, {
    x: M + 0.2, y: 0.24, w: 7.6, h: 0.3,
    fontFace: F.jp, fontSize: 15, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  if (sub)
    s.addText(sub, {
      x: M + 0.2, y: 0.55, w: 7.6, h: 0.2,
      fontFace: F.jp, fontSize: 8.5, color: "AFC6B9", margin: 0, valign: "middle",
    });
  s.addText("未来屋書店株式会社　御中", {
    x: R - 4.0, y: 0.26, w: 4.0, h: 0.2,
    fontFace: F.jp, fontSize: 8.5, color: "AFC6B9", align: "right", margin: 0, valign: "middle",
  });
  s.addText(PAGE + " / " + TOTAL, {
    x: R - 4.0, y: 0.50, w: 4.0, h: 0.22,
    fontFace: F.num, fontSize: 10, bold: true, color: C.bronze, align: "right", margin: 0, valign: "middle",
  });
}
const foot = (s, notes) => {
  s.addShape("rect", { x: M, y: H - 0.80, w: CW, h: 0.01, fill: { color: C.rule } });
  (notes || []).forEach((t, i) =>
    s.addText(t, {
      x: M, y: H - 0.72 + i * 0.145, w: CW - 3.6, h: 0.14,
      fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
    })
  );
  s.addText("セルフカフェ株式会社", {
    x: R - 3.5, y: H - 0.34, w: 3.5, h: 0.16,
    fontFace: F.jp, fontSize: 7.5, color: C.muted, align: "right", margin: 0, valign: "middle",
  });
};
const sec = (s, x, y, w, t, note) => {
  s.addShape("rect", { x, y: y + 0.02, w: 0.06, h: 0.19, fill: { color: C.bronze } });
  s.addText(t, {
    x: x + 0.15, y, w: w - 0.15, h: 0.23,
    fontFace: F.jp, fontSize: 10.5, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
  if (note)
    s.addText(note, {
      x: x + w - 3.4, y, w: 3.4, h: 0.23,
      fontFace: F.jp, fontSize: 7.5, color: C.muted, align: "right", margin: 0, valign: "middle",
    });
};

/* ---------------- table cells ---------------- */
const th = (t, al) => ({
  text: t,
  options: { fill: { color: C.deep }, color: C.white, bold: true, fontSize: 8.5, fontFace: F.jp, align: al || "left", valign: "middle", margin: [0.03, 0.1, 0.03, 0.1] },
});
const td = (t, o = {}) => ({
  text: t,
  options: {
    fill: { color: o.fill || C.white }, color: o.color || C.body, bold: !!o.bold,
    fontSize: o.size || 8.5, fontFace: o.face || nf(t), align: o.align || "left",
    valign: "middle", margin: [0.03, 0.1, 0.03, 0.1],
  },
});

/* ---------------- 棒グラフ ---------------- */
function bars(s, o) {
  const { x, y, w, h, data, max, unit } = o;
  const n = data.length;
  const gap = o.gap === undefined ? 0.16 : o.gap;
  const bw = o.bw || (w - gap * (n - 1)) / n;
  const base = y + h;
  // 目盛線
  (o.grid || []).forEach((g) => {
    const gy = base - (g.v / max) * h;
    s.addShape("rect", { x, y: gy, w, h: 0.008, fill: { color: g.strong ? C.bronze : C.rule } });
    if (g.label)
      s.addText(g.label, {
        x: x - 1.35, y: gy - 0.09, w: 1.25, h: 0.18,
        fontFace: F.jp, fontSize: 7.5, bold: !!g.strong, color: g.strong ? C.bronze : C.muted, align: "right", margin: 0, valign: "middle",
      });
  });
  data.forEach((d, i) => {
    const bx = x + i * (bw + gap);
    const bh = Math.max(0.03, (d.v / max) * h);
    s.addShape("rect", { x: bx, y: base - bh, w: bw, h: bh, fill: { color: d.color || C.green } });
    s.addText(d.t || String(d.v), {
      x: bx - 0.25, y: base - bh - 0.26, w: bw + 0.5, h: 0.22,
      fontFace: F.jp, fontSize: d.fs || 9.5, bold: true, color: d.lc || C.ink, align: "center", margin: 0, valign: "middle",
    });
    s.addText(d.label, {
      x: bx - 0.25, y: base + 0.05, w: bw + 0.5, h: 0.2,
      fontFace: F.jp, fontSize: 8, color: C.body, align: "center", margin: 0, valign: "middle",
    });
    if (d.sub)
      s.addText(d.sub, {
        x: bx - 0.25, y: base + 0.24, w: bw + 0.5, h: 0.18,
        fontFace: F.jp, fontSize: 7, color: C.muted, align: "center", margin: 0, valign: "middle",
      });
  });
  s.addShape("rect", { x, y: base, w, h: 0.012, fill: { color: C.ink } });
  if (unit)
    s.addText(unit, {
      x, y: y - 0.22, w: 2.0, h: 0.18,
      fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
    });
}

/* ================================================= P1 サマリー */
{
  const s = pres.addSlide();
  head(s, "秋田店　開業後3か年の売上見込み", "30坪（区画確定）／業務委託型・3年 特別条件（契約確定）／2027年3月上旬 開業");

  // KPI
  const K = [
    { l: "1年目", p: "2027年3月〜2028年2月", v: "2,943,600", u: "円", s2: "月平均 245,300円" },
    { l: "2年目", p: "2028年3月〜2029年2月", v: "3,435,000", u: "円", s2: "月平均 286,250円" },
    { l: "3年目", p: "2029年3月〜2030年2月", v: "3,510,600", u: "円", s2: "月平均 292,550円" },
    { l: "3年間 累計", p: "契約期間と同じ3年間", v: "9,889,200", u: "円", s2: "約 989万円", hot: true },
  ];
  const kw = (CW - 0.18 * 3) / 4;
  K.forEach((k, i) => {
    const x = M + i * (kw + 0.18), y = 1.22, h = 1.14;
    s.addShape("rect", { x, y, w: kw, h, fill: { color: k.hot ? C.bronzeT : C.soft } });
    s.addShape("rect", { x, y, w: kw, h: 0.045, fill: { color: k.hot ? C.bronze : C.green } });
    s.addText(k.l, {
      x: x + 0.16, y: y + 0.12, w: kw - 0.32, h: 0.2,
      fontFace: F.jp, fontSize: 9, bold: true, color: k.hot ? C.bronze : C.green, margin: 0, valign: "middle",
    });
    s.addText(k.p, {
      x: x + 0.16, y: y + 0.31, w: kw - 0.32, h: 0.16,
      fontFace: F.jp, fontSize: 6.8, color: C.muted, margin: 0, valign: "middle",
    });
    s.addText(
      [
        { text: k.v, options: { fontFace: F.num, fontSize: 19, bold: true, color: k.hot ? C.bronze : C.ink } },
        { text: " " + k.u, options: { fontFace: F.jp, fontSize: 9, bold: true, color: C.body } },
      ],
      { x: x + 0.16, y: y + 0.50, w: kw - 0.32, h: 0.34, margin: 0, valign: "middle" }
    );
    s.addText(k.s2, {
      x: x + 0.16, y: y + 0.86, w: kw - 0.32, h: 0.18,
      fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
    });
  });
  s.addText("上記は貴社のお受取額（セルフカフェ売上の25％ ＋ 月額固定50,000円＝貴社のみの特別条件）でございます。すべて税抜。", {
    x: M, y: 2.42, w: CW, h: 0.18,
    fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
  });

  // 表
  const lw = 6.35, rx = M + lw + 0.38, rw = R - rx;
  sec(s, M, 2.72, lw, "3か年の内訳");
  const Y = [
    { c: 62, cm: "1,860", sm: "781,200", sy: "9,374,400", q: "195,300", fm: "245,300", fy: "2,943,600" },
    { c: 75, cm: "2,250", sm: "945,000", sy: "11,340,000", q: "236,250", fm: "286,250", fy: "3,435,000" },
    { c: 77, cm: "2,310", sm: "970,200", sy: "11,642,400", q: "242,550", fm: "292,550", fy: "3,510,600" },
  ];
  const rc = (t, o) => td(t, Object.assign({ align: "right" }, o));
  s.addTable(
    [
      [th("項目"), th("1年目", "right"), th("2年目", "right"), th("3年目", "right")],
      [td("1日あたりの販売杯数（平均）"), ...Y.map((y) => rc(y.c + "杯", { bold: true, color: C.ink }))],
      [td("月間販売杯数（30日）"), ...Y.map((y) => rc(y.cm + "杯"))],
      [td("セルフカフェ売上（月）"), ...Y.map((y) => rc(y.sm))],
      [td("セルフカフェ売上（年）", { fill: C.tint }), ...Y.map((y) => rc(y.sy, { fill: C.tint, bold: true, color: C.green }))],
      [td("　うち 売上の25％（月）"), ...Y.map((y) => rc(y.q))],
      [td("　月額固定（月）"), ...Y.map(() => rc("50,000"))],
      [td("貴社お受取額（月）", { fill: C.bronzeT, bold: true, color: C.bronze }), ...Y.map((y) => rc(y.fm, { fill: C.bronzeT, bold: true, color: C.bronze, size: 10 }))],
      [td("貴社お受取額（年）", { fill: C.bronzeT, bold: true, color: C.bronze }), ...Y.map((y) => rc(y.fy, { fill: C.bronzeT, bold: true, color: C.bronze, size: 10 }))],
    ],
    { x: M, y: 3.02, w: lw, colW: [2.33, 1.34, 1.34, 1.34], rowH: [0.22, 0.26, 0.24, 0.24, 0.26, 0.24, 0.24, 0.28, 0.28], border: { type: "solid", color: C.rule, pt: 0.75 }, autoPage: false }
  );

  // グラフ
  sec(s, rx, 2.72, rw, "貴社お受取額の推移");
  bars(s, {
    x: rx + 0.1, y: 3.26, w: rw - 0.2, h: 1.62, max: 400, unit: "万円／年", gap: 0.4,
    data: [
      { v: 294.4, t: "294.4", label: "1年目", sub: "立ち上がり", color: C.mint, lc: C.green },
      { v: 343.5, t: "343.5", label: "2年目", sub: "巡航", color: C.green },
      { v: 351.1, t: "351.1", label: "3年目", sub: "巡航＋α", color: C.deep },
    ],
  });
  s.addShape("rect", { x: rx, y: 5.26, w: rw, h: 0.72, fill: { color: C.tint } });
  s.addText("初期費用（設置工事費 30〜50万円）は、開業後 1〜2か月で回収できる見込みでございます。", {
    x: rx + 0.16, y: 5.34, w: rw - 0.32, h: 0.56,
    fontFace: F.jp, fontSize: 8.5, color: C.ink, margin: 0, valign: "middle", lineSpacingMultiple: 1.25,
  });

  sec(s, M, 6.18, CW, "この見込みの読み方");
  const RD = [
    ["1年目は立ち上がり期", "開業直後は館内での認知が十分でないため、巡航値の約83％（1日62杯）で置いております。"],
    ["2年目に巡航へ", "1日75杯。当初ご提案の30坪想定と同じ水準で、以降の基準となる数値でございます。"],
    ["3年目は微増", "リピート利用の定着を見込み、巡航値から約3％の増加（1日77杯）としております。"],
  ];
  const cw3 = (CW - 0.2 * 2) / 3;
  RD.forEach((d, i) => {
    const x = M + i * (cw3 + 0.2);
    s.addShape("rect", { x, y: 6.48, w: 0.045, h: 0.72, fill: { color: C.bronze } });
    s.addText(d[0], {
      x: x + 0.14, y: 6.48, w: cw3 - 0.14, h: 0.2,
      fontFace: F.jp, fontSize: 9, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(d[1], {
      x: x + 0.14, y: 6.70, w: cw3 - 0.14, h: 0.52,
      fontFace: F.jp, fontSize: 8, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
    });
  });

  foot(s, [
    "※　想定値であり、売上およびお受取額を保証するものではございません。金額はすべて税抜でございます。",
    "※　水道光熱費・通信費・消耗品費は貴社ご負担のため、本試算には含めておりません。",
  ]);
}

/* ================================================= P2 杯数の立ち上がり */
{
  const s = pres.addSlide();
  head(s, "販売杯数の立ち上がりと、その根拠", "1日あたりの販売杯数の推移（1年目は四半期ごと）");

  sec(s, M, 1.18, CW, "1日あたりの販売杯数の推移", "巡航値 75杯／日（破線）");
  bars(s, {
    x: M + 1.45, y: 1.86, w: CW - 1.75, h: 2.20, max: 90, unit: "杯／日", gap: 0.34,
    grid: [{ v: 75, strong: true, label: "巡航値 75杯" }],
    data: [
      { v: 45, t: "45", label: "1年目 1Q", sub: "開業〜3か月", color: C.mint, lc: C.green },
      { v: 60, t: "60", label: "1年目 2Q", sub: "4〜6か月", color: C.mint, lc: C.green },
      { v: 68, t: "68", label: "1年目 3Q", sub: "7〜9か月", color: "8FBFA6", lc: C.green },
      { v: 75, t: "75", label: "1年目 4Q", sub: "10〜12か月", color: C.green },
      { v: 75, t: "75", label: "2年目", sub: "通年平均", color: C.green },
      { v: 77, t: "77", label: "3年目", sub: "通年平均", color: C.deep },
    ],
  });
  s.addText("1年目の通年平均＝62杯／日", {
    x: M, y: 4.58, w: 4.4, h: 0.2,
    fontFace: F.jp, fontSize: 8.5, bold: true, color: C.bronze, margin: 0, valign: "middle",
  });

  const lw = 6.35, rx = M + lw + 0.38, rw = R - rx;
  sec(s, M, 4.84, lw, "巡航値 75杯／日 の置き方");
  s.addTable(
    [
      [th("参照した数値"), th("1日あたりの販売杯数", "right"), th("位置づけ")],
      [td("書店併設の既存2店　実績平均"), td("42.9杯", { align: "right", bold: true, color: C.ink }), td("下限の目安", { size: 8 })],
      [td("商業施設内の大型店　実績"), td("122.7杯", { align: "right", bold: true, color: C.ink }), td("上限の目安", { size: 8 })],
      [td("秋田店　30坪・約59席", { fill: C.bronzeT, bold: true, color: C.bronze }), td("75杯", { align: "right", bold: true, color: C.bronze, size: 10.5, fill: C.bronzeT }), td("本試算の巡航値", { size: 8, fill: C.bronzeT, color: C.bronze })],
    ],
    { x: M, y: 5.14, w: lw, colW: [2.95, 1.8, 1.6], rowH: [0.22, 0.26, 0.26, 0.3], border: { type: "solid", color: C.rule, pt: 0.75 }, autoPage: false }
  );
  s.addText("館内の来客数と席数（約59席）から、既存2店の実績と商業施設内大型店の実績の間に置いております。販売杯数を左右するのは館内の来客数と席の回転でございます。確定いたしました区画（30坪）のレイアウトを前提としております。", {
    x: M, y: 6.26, w: lw, h: 0.6,
    fontFace: F.jp, fontSize: 8, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
  });

  sec(s, rx, 4.84, rw, "立ち上がりの考え方");
  const ST = [
    ["1Q　45杯", "開業告知とメニュー認知の形成期。書店ご利用のお客様への浸透が中心となります。"],
    ["2Q　60杯", "リピート利用が生まれ始める時期。席の回転と滞在の型が定まってまいります。"],
    ["3Q　68杯", "館内での定着期。平日の稼働が底上げされてまいります。"],
    ["4Q　75杯", "巡航値に到達。以降はこの水準を基準といたします。"],
  ];
  ST.forEach((t, i) => {
    const y = 5.14 + i * 0.48;
    s.addShape("rect", { x: rx, y, w: rw, h: 0.44, fill: { color: i % 2 ? C.soft : C.white }, line: { color: C.rule, width: 0.5 } });
    s.addText(t[0], {
      x: rx + 0.12, y: y + 0.03, w: rw - 0.24, h: 0.18,
      fontFace: F.jp, fontSize: 8.5, bold: true, color: C.green, margin: 0, valign: "middle",
    });
    s.addText(t[1], {
      x: rx + 0.12, y: y + 0.21, w: rw - 0.24, h: 0.2,
      fontFace: F.jp, fontSize: 7.5, color: C.body, margin: 0, valign: "middle",
    });
  });

  foot(s, [
    "※　既存店の実績は社内管理台帳（2026年5月〜8月の月間実杯数の平均）によるものでございます。",
    "※　販売杯数は立地・季節・館内の動線により変動いたします。想定値であり保証するものではございません。",
  ]);
}

/* ================================================= P3 ケース別・前提 */
{
  const s = pres.addSlide();
  head(s, "ケース別の見込みと、試算の前提", "巡航値を±10％動かした場合の3か年");

  const lw = 7.1, rx = M + lw + 0.38, rw = R - rx;
  sec(s, M, 1.18, lw, "3つのケース");
  const CASE = [
    { n: "保守ケース", d: "巡航 68杯／日（標準の▲10％）", y: ["2,716,800", "3,170,400", "3,246,000"], t: "9,133,200", c: C.muted, f: C.soft },
    { n: "標準ケース", d: "巡航 75杯／日（本試算の基準）", y: ["2,943,600", "3,435,000", "3,510,600"], t: "9,889,200", c: C.bronze, f: C.bronzeT, hot: true },
    { n: "上振れケース", d: "巡航 83杯／日（標準の＋10％）", y: ["3,208,200", "3,737,400", "3,813,000"], t: "10,758,600", c: C.green, f: C.tint },
  ];
  s.addTable(
    [
      [th("ケース"), th("1年目", "right"), th("2年目", "right"), th("3年目", "right"), th("3年累計", "right")],
      ...CASE.map((k) => [
        { text: [{ text: k.n + "\n", options: { fontFace: F.jp, fontSize: 9, bold: true, color: k.c } }, { text: k.d, options: { fontFace: F.jp, fontSize: 7.2, color: C.muted } }],
          options: { fill: { color: k.f }, valign: "middle", margin: [0.03, 0.1, 0.03, 0.1] } },
        ...k.y.map((v) => td(v, { align: "right", fill: k.f, size: 9, bold: !!k.hot, color: k.hot ? k.c : C.body })),
        td(k.t, { align: "right", fill: k.f, size: 10.5, bold: true, color: k.c }),
      ]),
    ],
    { x: M, y: 1.48, w: lw, colW: [2.3, 1.14, 1.14, 1.14, 1.38], rowH: [0.22, 0.46, 0.46, 0.46], border: { type: "solid", color: C.rule, pt: 0.75 }, autoPage: false }
  );
  s.addText("いずれのケースでも、3年間の累計お受取額は900万円を超える見込みでございます。上振れケースは、館内の来客数が想定を上回って推移した場合を見込んだものでございます。", {
    x: M, y: 3.18, w: lw, h: 0.42,
    fontFace: F.jp, fontSize: 8.5, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
  });

  sec(s, rx, 1.18, rw, "3年累計の比較");
  bars(s, {
    x: rx + 0.12, y: 1.80, w: rw - 0.24, h: 1.32, max: 1200, unit: "万円（3年累計）", gap: 0.34,
    data: [
      { v: 913.3, t: "913", label: "保守", color: "B9C4BD", lc: C.muted, fs: 9 },
      { v: 988.9, t: "989", label: "標準", color: C.bronze, fs: 9 },
      { v: 1075.9, t: "1,076", label: "上振れ", color: C.green, fs: 9 },
    ],
  });

  sec(s, M, 3.78, lw, "試算の前提");
  const PRE = [
    ["契約形態", "業務委託型・3年　特別条件（確定）"],
    ["区画", "30坪（図面確定）／約59席を想定"],
    ["お支払い", "セルフカフェ売上の25％ ＋ 月額固定 50,000円（貴社のみの特別条件）"],
    ["販売単価", "1杯 420円（税抜・定価ベース）"],
    ["営業日数", "月30日・年12か月で算出"],
    ["開業", "2027年3月上旬"],
    ["貴社の初期費用", "設置工事費 30〜50万円（税抜）のみ。お見積りのご提出はございません。"],
    ["貴社のご負担費用", "水道光熱費・通信費・消耗品費（本試算には含めておりません）"],
  ];
  s.addTable(
    [...PRE.map((p, i) => [
      td(p[0], { bold: true, color: C.ink, size: 8.5, fill: i % 2 ? C.soft : C.white }),
      td(p[1], { size: 8.5, fill: i % 2 ? C.soft : C.white }),
    ])],
    { x: M, y: 4.08, w: lw, colW: [1.85, lw - 1.85], rowH: PRE.map(() => 0.32), border: { type: "solid", color: C.rule, pt: 0.75 }, autoPage: false }
  );

  sec(s, rx, 3.78, rw, "ご留意事項");
  const NT = [
    "本試算は想定値であり、売上およびお受取額を保証するものではございません。",
    "販売杯数は、館内の来客数・季節・動線により変動いたします。",
    "館の営業時間が変更となった場合、稼働時間に応じて増減いたします。",
    "水道光熱費・通信費・消耗品費を差し引いた実質手取りベースの試算も、ご要望に応じてご提出いたします。",
    "開業後は月次の実績をご報告し、半期ごとに見込みを見直してまいります。",
  ];
  NT.forEach((t, i) => {
    const y = 4.08 + i * 0.52;
    s.addText("・", {
      x: rx, y, w: 0.18, h: 0.48,
      fontFace: F.jp, fontSize: 8, color: C.bronze, margin: 0, valign: "top",
    });
    s.addText(t, {
      x: rx + 0.18, y, w: rw - 0.18, h: 0.48,
      fontFace: F.jp, fontSize: 8, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
    });
  });

  foot(s, [
    "※　金額はすべて税抜でございます。想定値であり、売上およびお受取額を保証するものではございません。",
    "※　1年目＝2027年3月〜2028年2月、2年目＝2028年3月〜2029年2月、3年目＝2029年3月〜2030年2月といたしました。",
  ]);
}

pres.writeFile({ fileName: path.join(__dirname, "akita-forecast.pptx") }).then((f) => console.log("written:", f));
