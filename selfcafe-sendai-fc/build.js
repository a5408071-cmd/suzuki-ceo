/**
 * 未来屋書店 仙台泉（新規出店）FCプランのご提案
 * A4横 3ページ／白地・単色・見出し明朝の報告書体裁
 */
const pptxgen = require("pptxgenjs");
const path = require("path");

const C = {
  bg: "FFFFFF",
  ink: "1A1A1A",
  body: "333333",
  muted: "767676",
  ruleS: "1A1A1A",   // 太罫
  ruleT: "808080",   // 表罫
  hair: "C8C8C8",
  headFill: "E4E4E4",
  band: "F2F2F2",
  soft: "F7F8F9",
  bandD: "E9F1EC",
  barD: "106838",
  gr: "106838",
  grPale: "E9F1EC",
  barL: "BFBFBF",
  white: "FFFFFF",
};
const FM = "Yu Mincho";   // 見出し
const FG = "Yu Gothic";   // 本文・表・数字

const W = 11.69, H = 8.27, M = 0.58, CW = W - M * 2, R = M + CW;

const pres = new pptxgen();
pres.defineLayout({ name: "A4L", width: W, height: H });
pres.layout = "A4L";
pres.author = "セルフカフェ株式会社";
pres.title = "仙台泉 FCプランのご提案";

let PAGE = 0;
const TOTAL = 4;

/* ---------------- 体裁 ---------------- */
function head(s, title, sub) {
  PAGE++;
  s.background = { color: C.bg };
  s.addText(title, {
    x: M, y: 0.36, w: CW - 3.3, h: 0.36,
    fontFace: FM, fontSize: 17, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
  if (sub)
    s.addText(sub, {
      x: M, y: 0.72, w: CW - 3.3, h: 0.2,
      fontFace: FG, fontSize: 8.5, color: C.body, margin: 0, valign: "middle",
    });
  s.addText("未来屋書店株式会社　御中", {
    x: R - 3.3, y: 0.34, w: 3.3, h: 0.2,
    fontFace: FG, fontSize: 9, color: C.ink, align: "right", margin: 0, valign: "middle",
  });
  s.addText("2026年11月26日　セルフカフェ株式会社", {
    x: R - 3.3, y: 0.54, w: 3.3, h: 0.18,
    fontFace: FG, fontSize: 8, color: C.muted, align: "right", margin: 0, valign: "middle",
  });
  s.addText("― " + PAGE + " / " + TOTAL + " ―", {
    x: R - 3.3, y: 0.73, w: 3.3, h: 0.18,
    fontFace: FG, fontSize: 8, color: C.muted, align: "right", margin: 0, valign: "middle",
  });
  s.addShape("rect", { x: M, y: 0.98, w: CW, h: 0.019, fill: { color: C.gr } });
  s.addShape("rect", { x: M, y: 1.025, w: CW, h: 0.006, fill: { color: C.ruleS } });
}
const foot = (s, notes) => {
  s.addShape("rect", { x: M, y: H - 0.78, w: CW, h: 0.006, fill: { color: C.hair } });
  (notes || []).forEach((t, i) =>
    s.addText(t, {
      x: M, y: H - 0.70 + i * 0.15, w: CW - 2.6, h: 0.15,
      fontFace: FG, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
    })
  );
  s.addText("セルフカフェ株式会社", {
    x: R - 2.4, y: H - 0.3, w: 2.4, h: 0.16,
    fontFace: FG, fontSize: 7.5, color: C.muted, align: "right", margin: 0, valign: "middle",
  });
};
let SEC = 0;
const sec = (s, x, y, w, t, note) => {
  SEC++;
  s.addText(SEC + "．" + t, {
    x, y, w, h: 0.24,
    fontFace: FM, fontSize: 11.5, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
  if (note)
    s.addText(note, {
      x: x + w - 3.6, y, w: 3.6, h: 0.24,
      fontFace: FG, fontSize: 8, color: C.muted, align: "right", margin: 0, valign: "middle",
    });
  s.addShape("rect", { x, y: y + 0.25, w, h: 0.008, fill: { color: C.gr } });
};

const th = (t, al) => ({
  text: t,
  options: { fill: { color: C.grPale }, color: C.gr, bold: true, fontSize: 9, fontFace: FG, align: al || "left", valign: "middle", margin: [0.03, 0.1, 0.03, 0.1] },
});
const td = (t, o = {}) => ({
  text: t,
  options: {
    fill: { color: o.fill || C.white }, color: o.color || C.body, bold: !!o.bold,
    fontSize: o.size || 9, fontFace: FG, align: o.align || "left",
    valign: "middle", margin: [0.03, 0.1, 0.03, 0.1],
  },
});
const TB = { type: "solid", color: C.ruleT, pt: 0.5 };

/* ---------------- 棒グラフ（単色） ---------------- */
function bars(s, o) {
  const { x, y, w, h, data, max, unit } = o;
  const n = data.length, gap = o.gap === undefined ? 0.16 : o.gap;
  const bw = (w - gap * (n - 1)) / n, base = y + h;
  data.forEach((d, i) => {
    const bx = x + i * (bw + gap), bh = Math.max(0.03, (d.v / max) * h);
    s.addShape("rect", { x: bx, y: base - bh, w: bw, h: bh, fill: { color: d.color || C.barL } });
    s.addText(d.t, {
      x: bx - 0.25, y: base - bh - 0.25, w: bw + 0.5, h: 0.22,
      fontFace: FG, fontSize: d.fs || 9, bold: true, color: C.ink, align: "center", margin: 0, valign: "middle",
    });
    s.addText(d.label, {
      x: bx - 0.25, y: base + 0.06, w: bw + 0.5, h: 0.2,
      fontFace: FG, fontSize: 8, color: C.body, align: "center", margin: 0, valign: "middle",
    });
    if (d.sub)
      s.addText(d.sub, {
        x: bx - 0.25, y: base + 0.24, w: bw + 0.5, h: 0.18,
        fontFace: FG, fontSize: 7, color: C.muted, align: "center", margin: 0, valign: "middle",
      });
  });
  s.addShape("rect", { x, y: base, w, h: 0.008, fill: { color: C.ink } });
  if (unit)
    s.addText(unit, {
      x, y: y - 0.22, w: 2.2, h: 0.18,
      fontFace: FG, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
    });
}


/* ---------------- 試算データ ---------------- */
const UP = 420, DAYS = 30, FIXFEE = 50000;
const MACH = 76000, TERM = 10000, CLOUD = 5000, SUP = 15000, ROY = 50000;
const UTL = { 30: 60000, 60: 70000 }, ETC = { 30: 20000, 60: 30000 };
const CLN = 30000, SECF = 20000;
const CASE = [
  { t: "30坪", pos: "店頭配置", cups: 86, tb: 30, hot: true },
  { t: "30坪", pos: "店奥配置", cups: 60, tb: 30 },
  { t: "60坪", pos: "店頭配置", cups: 98, tb: 60, ref: true },
  { t: "60坪", pos: "店奥配置", cups: 69, tb: 60, ref: true },
].map((c) => {
  const sales = c.cups * UP * DAYS;
  const mat = Math.floor(sales * 0.2);
  const fx = UTL[c.tb] + MACH + TERM + CLOUD + SUP + ROY + ETC[c.tb];
  const fPf = sales - mat - fx;
  const iIn = Math.floor(sales * 0.25) + FIXFEE;
  const iCost = UTL[c.tb] + ETC[c.tb];
  const iPf = iIn - iCost;
  return Object.assign({}, c, { sales, mat, fx, fCost: mat + fx, fPf, iIn, iCost, iPf, d: fPf - iPf });
});
const B = CASE[0];
const yen = (v) => v.toLocaleString();
const neg = (v) => "▲ " + yen(v);

/* ================================================= P1 違い */
{
  const s = pres.addSlide();
  head(s, "業務委託型とFCプランの違い", "仙台泉（新規出店）　30坪・1日86杯（店頭配置）／2028年度内 開業予定");

  s.addShape("rect", { x: M, y: 1.14, w: CW, h: 0.62, fill: { color: C.white }, line: { color: C.ruleS, width: 1 } });
  s.addText("基本工事（床・壁・天井・照明）は、いずれの契約形態でも貴社のご負担でございます。", {
    x: M + 0.18, y: 1.18, w: CW - 0.36, h: 0.24,
    fontFace: FG, fontSize: 9.5, bold: true, color: C.gr, margin: 0, valign: "middle",
  });
  s.addText("他の書店区画と同等の仕上がりが必要となるためでございます。本資料で申し上げる「ご投資」は、それに加わる部分（業務委託型であれば弊社が負担しておりました内装造作・什器等）を指しております。", {
    x: M + 0.18, y: 1.42, w: CW - 0.36, h: 0.3,
    fontFace: FG, fontSize: 8.5, color: C.body, margin: 0, valign: "middle",
  });

  sec(s, M, 1.92, CW, "仕組みの違い", "運営はいずれも貴社スタッフ。違いは「売上の帰属」と「工事の負担」の2点でございます");
  const pw = (CW - 0.3) / 2;
  const PLAN = [
    { t: "業務委託型", sub: "弊社が事業主体となります", f: C.soft, bd: C.ruleT, c: C.ink,
      rows: [["運　営", "未来屋書店様（弊社からの業務委託）"], ["売上の帰属", "セルフカフェ"], ["内装工事・什器", "セルフカフェが負担"], ["貴社のご負担", "設置工事費 30〜50万円"], ["貴社の収入", "売上の25％ ＋ 月額固定 50,000円"]],
      pf: yen(B.iPf) },
    { t: "FCプラン", sub: "貴社が事業主体となられます", f: C.grPale, bd: C.gr, c: C.gr, hot: true,
      rows: [["運　営", "未来屋書店様"], ["売上の帰属", "未来屋書店様"], ["内装工事・什器", "未来屋書店様が負担（ご投資に含む）"], ["貴社のご投資", "内装造作・什器等（別途お見積り）"], ["貴社の収入", "売上高（経費を差し引いた残りが利益）"]],
      pf: yen(B.fPf) },
  ];
  PLAN.forEach((p, i) => {
    const x = M + i * (pw + 0.3);
    s.addShape("rect", { x, y: 2.28, w: pw, h: 2.52, fill: { color: C.white }, line: { color: p.bd, width: p.hot ? 1.25 : 0.75 } });
    s.addShape("rect", { x, y: 2.28, w: pw, h: 0.42, fill: { color: p.f } });
    s.addText(p.t, { x: x + 0.2, y: 2.28, w: pw * 0.42, h: 0.42, fontFace: FM, fontSize: 13, bold: true, color: p.c, margin: 0, valign: "middle" });
    s.addText(p.sub, { x: x + pw * 0.42, y: 2.28, w: pw * 0.58 - 0.2, h: 0.42, fontFace: FG, fontSize: 8.5, color: C.body, align: "right", margin: 0, valign: "middle" });
    p.rows.forEach((r, j) => {
      const y = 2.78 + j * 0.3;
      s.addText(r[0], { x: x + 0.2, y, w: 1.45, h: 0.28, fontFace: FG, fontSize: 8.8, bold: true, color: C.ink, margin: 0, valign: "middle" });
      s.addText(r[1], { x: x + 1.65, y, w: pw - 1.85, h: 0.28, fontFace: FG, fontSize: 8.8, color: C.body, margin: 0, valign: "middle" });
    });
    s.addShape("rect", { x: x + 0.2, y: 4.32, w: pw - 0.4, h: 0.006, fill: { color: C.hair } });
    s.addText("30坪・店頭配置の営業利益", { x: x + 0.2, y: 4.38, w: 2.3, h: 0.34, fontFace: FG, fontSize: 8.5, color: C.body, margin: 0, valign: "middle" });
    s.addText(
      [{ text: p.pf, options: { fontFace: FG, fontSize: 20, bold: true, color: p.c } },
       { text: " 円／月", options: { fontFace: FG, fontSize: 9.5, bold: true, color: C.body } }],
      { x: x + pw - 3.0, y: 4.36, w: 2.8, h: 0.38, align: "right", margin: 0, valign: "middle" }
    );
  });
  s.addText("FCプランは 2.62倍 でございます。", {
    x: M, y: 4.88, w: CW, h: 0.22,
    fontFace: FG, fontSize: 9.5, bold: true, color: C.gr, align: "center", margin: 0, valign: "middle",
  });

  sec(s, M, 5.22, CW, "条件の比較");
  const r3 = (a, b, c) => [td(a, { size: 8.8 }), td(b, { align: "center", size: 8.8 }), td(c, { align: "center", size: 8.8, fill: C.grPale })];
  s.addTable(
    [
      [th("観点"), th("業務委託型", "center"), th("FCプラン", "center")],
      r3("基本工事（床・壁・天井・照明）", "貴社がご負担", "貴社がご負担"),
      r3("貴社がご負担になる経費", "水道光熱費・雑費（通信費・消耗品費等）", "原料・水道光熱費・設備利用料・ロイヤリティ・雑費 ほか"),
      r3("日常業務（清掃・補充・発注）", "貴社スタッフ", "貴社スタッフ"),
      r3("売上変動のリスク", "小（月額固定5万円が下支え）", "貴社がご負担"),
      r3("食品衛生責任者", "不要", "1名のご選任が必要"),
    ],
    { x: M, y: 5.58, w: CW, colW: [3.0, 3.4, CW - 6.4], rowH: [0.24, 0.26, 0.3, 0.24, 0.26, 0.24], border: TB, autoPage: false }
  );

  foot(s, [
    "※　本資料は開業前の机上試算でございます。想定値であり、売上および利益を保証するものではございません。金額はすべて税抜。",
    "※　日常業務はいずれの方式でも貴社スタッフによる運用を想定。人件費は両案とも未計上、家賃は0円としております。",
  ]);
}

/* ================================================= P2 月次損益 */
{
  const s = pres.addSlide();
  head(s, "月次損益の比較", "30坪・店頭配置（1日86杯）　同じ条件・同じ経費区分で並べております　単位：円（税抜）");

  const hw = (CW - 0.3 * 2) / 3;
  [
    { l: "業務委託型　営業利益", v: yen(B.iPf), f: C.soft, bd: C.ruleT, c: C.ink },
    { l: "FCプラン　営業利益", v: yen(B.fPf), f: C.white, bd: C.gr, c: C.gr },
    { l: "差　額", v: "＋" + yen(B.d), f: C.grPale, bd: C.gr, c: C.gr, hot: true },
  ].forEach((h, i) => {
    const x = M + i * (hw + 0.3);
    s.addShape("rect", { x, y: 1.14, w: hw, h: 0.92, fill: { color: h.f }, line: { color: h.bd, width: h.hot ? 1.25 : 0.75 } });
    s.addText(h.l, { x: x + 0.18, y: 1.19, w: hw - 0.36, h: 0.22, fontFace: FG, fontSize: 9, bold: true, color: h.c, margin: 0, valign: "middle" });
    s.addText(
      [{ text: h.v, options: { fontFace: FG, fontSize: h.hot ? 25 : 22, bold: true, color: h.c } },
       { text: " 円／月", options: { fontFace: FG, fontSize: 9.5, bold: true, color: C.body } }],
      { x: x + 0.18, y: 1.42, w: hw - 0.36, h: 0.56, margin: 0, valign: "middle" }
    );
  });
  [0, 1].forEach((i) => s.addText(i === 0 ? "→" : "＝", {
    x: M + (i + 1) * hw + i * 0.3, y: 1.40, w: 0.3, h: 0.4,
    fontFace: FG, fontSize: 14, bold: true, color: C.muted, align: "center", margin: 0, valign: "middle",
  }));
  s.addText("年間では ＋" + yen(B.d * 12) + "円、販売1杯あたりに換算いたしますと 151円 の差でございます。", {
    x: M, y: 2.12, w: CW, h: 0.22,
    fontFace: FG, fontSize: 9.5, bold: true, color: C.gr, align: "center", margin: 0, valign: "middle",
  });

  sec(s, M, 2.46, CW, "月次損益の内訳");
  const c1 = (v, o = {}) => td(v, Object.assign({ align: "right", size: 9 }, o));
  const c2 = (v, o = {}) => td(v, Object.assign({ align: "right", size: 9, fill: C.grPale }, o));
  const c3 = (v, o = {}) => td(v, Object.assign({ align: "right", size: 9, bold: true }, o));
  const ROWS = [
    ["セルフカフェ売上（月）", yen(B.sales), yen(B.sales), "－", {}],
    ["貴社の収入", yen(B.iIn), yen(B.sales), "＋" + yen(B.sales - B.iIn), { b: true }],
    ["　ドリンク原料（約20％）", "－", neg(B.mat), neg(B.mat), {}],
    ["　水道光熱費（書店営業時間ベース）", neg(UTL[30]), neg(UTL[30]), "－", {}],
    ["　マシン使用料（38,000円×2台）", "－", neg(MACH), neg(MACH), {}],
    ["　決済端末利用料（5,000円×2台）", "－", neg(TERM), neg(TERM), {}],
    ["　クラウド利用料", "－", neg(CLOUD), neg(CLOUD), {}],
    ["　営業サポート（ミニマムサポート）", "－", neg(SUP), neg(SUP), {}],
    ["　ロイヤリティ（一律・売上連動なし）", "－", neg(ROY), neg(ROY), {}],
    ["　雑費（通信費・消耗品費等）", neg(ETC[30]), neg(ETC[30]), "－", {}],
    ["　清掃業務委託費　※有人のため不要", "－", "（" + yen(CLN) + "）", "－", { g: true }],
    ["　セキュリティ費　※有人のため不要", "－", "（" + yen(SECF) + "）", "－", { g: true }],
    ["ランニングコスト計", neg(B.iCost), neg(B.fCost), neg(B.fCost - B.iCost), { b: true, cost: true }],
    ["営業利益（月）", yen(B.iPf), yen(B.fPf), "＋" + yen(B.d), { b: true, f: C.band, big: true }],
  ];
  s.addTable(
    [
      [th("項目"), th("業務委託型", "right"), th("FCプラン", "right"), th("差　額", "right")],
      ...ROWS.map((x) => {
        const g = x[4].g, sz = x[4].big ? 11 : 9;
        return [
          td(x[0], { size: 9, bold: !!x[4].b, color: g ? C.muted : (x[4].b ? C.ink : C.body), fill: x[4].f }),
          c1(x[1], { bold: !!x[4].b, color: g ? C.muted : (x[4].b ? C.ink : C.body), fill: x[4].f, size: sz }),
          c2(x[2], { bold: !!x[4].b, color: g ? C.muted : (x[4].b && !x[4].cost ? C.gr : C.body), size: sz }),
          c3(x[3], { fill: x[4].f, size: sz, color: g || x[3] === "－" ? C.muted : (x[3][0] === "▲" ? C.ink : C.gr) }),
        ];
      }),
    ],
    { x: M, y: 2.82, w: CW, colW: [4.05, 2.16, 2.16, CW - 8.37], rowH: [0.26, 0.26, 0.3, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.28, 0.34], border: TB, autoPage: false }
  );
  s.addText("収入が ＋" + yen(B.sales - B.iIn) + "円 増え、経費が ▲" + yen(B.fCost - B.iCost) + "円 増えるため、差引 ＋" + yen(B.d) + "円 となります。平均荒利率は 80.0％ でございます。", {
    x: M, y: 6.76, w: CW, h: 0.24,
    fontFace: FG, fontSize: 9.5, bold: true, color: C.gr, align: "center", margin: 0, valign: "middle",
  });

  foot(s, [
    "※　業務委託型の「貴社の収入」は、セルフカフェ売上の25％ ＋ 月額固定50,000円でございます。FCプランは売上高がそのまま貴社に帰属いたします。",
    "※　清掃業務委託費・セキュリティ費は、貴社書店内への併設（有人）のため不要でございます。参考として括弧内に標準額を記載しております。",
  ]);
}

/* ================================================= P3 坪数・配置別 */
{
  const s = pres.addSlide();
  head(s, "坪数・配置別の試算", "販売杯数は、既にご提出済みの売上試算と同一でございます　単位：円（税抜）");

  sec(s, M, 1.14, CW, "4パターンの比較");
  const cI = (v, o = {}) => td(v, Object.assign({ align: "right", size: 8.8 }, o));
  const cF = (v, o = {}) => td(v, Object.assign({ align: "right", size: 8.8, fill: C.grPale }, o));
  const ln = (label, fi, ff, o = {}) => [
    td(label, { size: 8.8, bold: !!o.b, color: o.b ? C.ink : C.body, fill: o.lf }),
    ...CASE.flatMap((c) => [cI(fi(c), { bold: !!o.b, color: o.b ? C.ink : C.body }), cF(ff(c), { bold: !!o.b, color: o.b && !o.cost ? C.gr : C.body })]),
  ];
  s.addTable(
    [
      [th("項目"), ...CASE.flatMap((c) => [th(c.t + (c.ref ? "（参考）" : "") + "\n業務委託型", "right"), th(c.t + (c.ref ? "（参考）" : "") + "\nFCプラン", "right")])],
      [td("配置", { size: 8.8 }), ...CASE.flatMap((c) => [cI(c.pos), cF(c.pos)])],
      ln("販売杯数（1日）", (c) => c.cups + "杯", (c) => c.cups + "杯"),
      ln("セルフカフェ売上（月）", (c) => yen(c.sales), (c) => yen(c.sales)),
      ln("貴社の収入", (c) => yen(c.iIn), (c) => yen(c.sales), { b: true }),
      ln("ランニングコスト計", (c) => neg(c.iCost), (c) => neg(c.fCost), { cost: true }),
      ln("営業利益（月）", (c) => yen(c.iPf), (c) => yen(c.fPf), { b: true, lf: C.band }),
      ln("営業利益（年）", (c) => yen(c.iPf * 12), (c) => yen(c.fPf * 12), { b: true, lf: C.band }),
    ],
    { x: M, y: 1.50, w: CW, colW: [2.13, 1.05, 1.05, 1.05, 1.05, 1.05, 1.05, 1.05, 1.05], rowH: [0.36, 0.24, 0.24, 0.26, 0.3, 0.28, 0.32, 0.32], border: TB, autoPage: false }
  );

  sec(s, M, 4.06, CW, "営業利益（月）の比較");
  bars(s, {
    x: M + 1.0, y: 4.60, w: CW - 2.0, h: 1.12, max: 80, unit: "単位：万円／月", gap: 0.16,
    data: CASE.flatMap((c) => [
      { v: c.iPf / 10000, t: (c.iPf / 10000).toFixed(1), label: c.t, sub: c.pos.replace("配置", "") + "・委託", color: C.barL, fs: 8 },
      { v: c.fPf / 10000, t: (c.fPf / 10000).toFixed(1), label: c.t, sub: c.pos.replace("配置", "") + "・FC", color: C.gr, fs: 8 },
    ]),
  });

  sec(s, M, 6.26, CW, "損益分岐点");
  s.addTable(
    [
      [th("区分"), th("固定費（月）", "right"), th("損益分岐点", "right"), th("想定販売杯数（店頭配置）", "right")],
      [td("30坪", { size: 8.8 }), td(yen(236000), { align: "right", size: 8.8 }), td("1日 約24杯", { align: "right", size: 8.8, bold: true, color: C.gr }), td("1日 86杯", { align: "right", size: 8.8 })],
      [td("60坪", { size: 8.8 }), td(yen(256000), { align: "right", size: 8.8 }), td("1日 約26杯", { align: "right", size: 8.8, bold: true, color: C.gr }), td("1日 98杯", { align: "right", size: 8.8 })],
    ],
    { x: M, y: 6.62, w: CW, colW: [2.2, 2.6, 2.6, CW - 7.4], rowH: [0.24, 0.26, 0.26], border: TB, autoPage: false }
  );

  foot(s, [
    "※　平均荒利率はいずれも 80.0％ でございます（変動費はドリンク原料のみ）。想定値であり、売上および利益を保証するものではございません。",
  ]);
}

/* ================================================= P4 投資と利点・留意点 */
{
  const s = pres.addSlide();
  head(s, "ご投資の考え方と、ご判断の材料", "お見積りは、区画と仕様の確定後にあらためてご提出いたします");

  sec(s, M, 1.14, CW, "ご投資の区分");
  const SEG = [
    { t: "Ａ", h: "業務委託型では弊社が負担していた部分", d: "内装造作工事（カウンター・造作・二次側の配線配管）／家具什器／サイン工事", n: "FCプランでは貴社のご投資となります。貴社よりご指摘いただいた「プラスになる分」は、こちらでございます。", hot: true },
    { t: "Ｂ", h: "FCプランに固有の費用", d: "加盟金／保証金／開業準備金（研修・オープン販促・初期在庫）／設計・申請・諸経費", n: "保証金は契約終了時にご返還いたします。" },
    { t: "Ｃ", h: "基本工事（床・壁・天井・照明）", d: "他の書店区画と同等の仕上がりが必要となる部分", n: "いずれの契約形態でも貴社のご負担のため、上記の差には含まれません。" },
  ];
  const sw = (CW - 0.24 * 2) / 3;
  SEG.forEach((g, i) => {
    const x = M + i * (sw + 0.24);
    s.addShape("rect", { x, y: 1.50, w: sw, h: 1.46, fill: { color: g.hot ? C.grPale : C.soft }, line: { color: g.hot ? C.gr : C.ruleT, width: g.hot ? 1 : 0.75 } });
    s.addText(g.t, { x: x + 0.16, y: 1.56, w: 0.3, h: 0.24, fontFace: FM, fontSize: 12, bold: true, color: g.hot ? C.gr : C.ink, margin: 0, valign: "middle" });
    s.addText(g.h, { x: x + 0.46, y: 1.56, w: sw - 0.62, h: 0.24, fontFace: FG, fontSize: 8.8, bold: true, color: g.hot ? C.gr : C.ink, margin: 0, valign: "middle" });
    s.addText(g.d, { x: x + 0.16, y: 1.84, w: sw - 0.32, h: 0.56, fontFace: FG, fontSize: 8.2, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.25 });
    s.addText(g.n, { x: x + 0.16, y: 2.42, w: sw - 0.32, h: 0.48, fontFace: FG, fontSize: 7.8, color: C.muted, margin: 0, valign: "top", lineSpacingMultiple: 1.25 });
  });
  s.addText("マシン・決済端末は月額の利用料でのご提供のため、初期のご購入は不要でございます。お見積りは、区画の位置・坪数・仕様が固まりしだいご提出いたします。", {
    x: M, y: 3.06, w: CW, h: 0.24,
    fontFace: FG, fontSize: 8.8, color: C.body, margin: 0, valign: "middle",
  });

  const lw = 5.5, rx = M + lw + 0.5, rw = R - rx;
  sec(s, M, 3.48, lw, "FCプランの利点");
  [
    ["売上・利益がすべて貴社のものになります", "30坪・店頭配置で月630,880円。業務委託型の約2.6倍でございます。"],
    ["ロイヤリティが売上連動ではございません", "一律 月50,000円。売上が伸びるほど利益率が向上いたします。"],
    ["損益分岐点が低く、下振れに強い構造でございます", "30坪で1日約24杯。想定86杯に対して3倍以上の余裕がございます。"],
    ["自社業態としての知見が蓄積されます", "秋田店とあわせ、今後のご出店のご判断材料となります。"],
  ].forEach((t, i) => {
    const y = 3.86 + i * 0.56;
    s.addText("（" + (i + 1) + "）" + t[0], { x: M, y, w: lw, h: 0.2, fontFace: FG, fontSize: 8.8, bold: true, color: C.ink, margin: 0, valign: "middle" });
    s.addText(t[1], { x: M + 0.36, y: y + 0.2, w: lw - 0.36, h: 0.32, fontFace: FG, fontSize: 8.2, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.2 });
  });

  sec(s, rx, 3.48, rw, "ご留意いただきたい点");
  [
    ["初期のご投資が必要でございます", "区分ＡとＢの合計が、業務委託型と比べて増える分でございます。"],
    ["売上変動のリスクを貴社が負われます", "業務委託型の月額固定50,000円のような下支えはございません。"],
    ["食品衛生責任者の選任が必要でございます", "1名。保健所の申請手続きは弊社にて代行いたします。"],
    ["新規館のため来客数は計画値でございます", "開業は2028年度内。販売杯数は館の来客数想定に基づく試算でございます。"],
  ].forEach((t, i) => {
    const y = 3.86 + i * 0.56;
    s.addText("（" + (i + 1) + "）" + t[0], { x: rx, y, w: rw, h: 0.2, fontFace: FG, fontSize: 8.8, bold: true, color: C.ink, margin: 0, valign: "middle" });
    s.addText(t[1], { x: rx + 0.36, y: y + 0.2, w: rw - 0.36, h: 0.32, fontFace: FG, fontSize: 8.2, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.2 });
  });

  s.addShape("rect", { x: M, y: 6.28, w: CW, h: 0.72, fill: { color: C.grPale }, line: { color: C.gr, width: 1 } });
  s.addText("新規出店では基本工事のご負担が生じますが、FCプランであれば売上がそのまま貴社に帰属いたしますので、その回収を月次の利益で進めていただけます。", {
    x: M + 0.2, y: 6.28, w: CW - 0.4, h: 0.72,
    fontFace: FG, fontSize: 9, bold: true, color: C.gr, margin: 0, valign: "middle",
  });

  foot(s, [
    "※　ご投資額は、区画と仕様の確定後にあらためてお見積りをご提出いたします。金額はすべて税抜。",
  ]);
}

pres.writeFile({ fileName: path.join(__dirname, "sendai-fc.pptx") }).then((f) => console.log("written:", f));
