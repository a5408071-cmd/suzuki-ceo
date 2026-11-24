/**
 * 未来屋書店 秋田店　業務委託型とFCプランの比較
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
pres.title = "秋田店 業務委託型とFCプランの比較";

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
  s.addText("2026年11月24日　セルフカフェ株式会社", {
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
const UTL = 60000, MACH = 76000, TERM = 10000, CLOUD = 5000, SUP = 15000, ROY = 50000, ETC = 20000;
const CLN = 30000, SECF = 20000;            // 有人のため不要（参考表示）
const FX = UTL + MACH + TERM + CLOUD + SUP + ROY + ETC;   // 236,000
const IC = UTL + ETC;                                      //  80,000
const INV = 9000000, IT0 = 500000;
const YR = [
  { y: "1年目", term: "2027年3月〜2028年2月", cups: 62, note: "立ち上がり" },
  { y: "2年目", term: "2028年3月〜2029年2月", cups: 75, note: "巡航", hot: true },
  { y: "3年目", term: "2029年3月〜2030年2月", cups: 77, note: "巡航＋α" },
].map((r) => {
  const sales = r.cups * UP * DAYS;
  const mat = Math.floor(sales * 0.2);
  const iIn = Math.floor(sales * 0.25) + FIXFEE;
  const iPf = iIn - IC;
  const fCost = mat + FX;
  const fPf = sales - fCost;
  return Object.assign({}, r, { sales, mat, iIn, iCost: IC, iPf, fCost, fPf, iPfY: iPf * 12, fPfY: fPf * 12 });
});
const T_I = YR.reduce((a, r) => a + r.iPfY, 0);
const T_F = YR.reduce((a, r) => a + r.fPfY, 0);
const yen = (v) => v.toLocaleString();
const neg = (v) => "▲ " + yen(v);

/* ================================================= P1 違い */
{
  const s = pres.addSlide();
  head(s, "業務委託型とFCプランの違い", "秋田店　30坪・1日75杯（巡航）／2027年3月上旬 開業予定");

  s.addShape("rect", { x: M, y: 1.14, w: CW, h: 0.44, fill: { color: C.white }, line: { color: C.ruleS, width: 1 } });
  s.addText("【ご留意】　本資料は開業前の机上試算でございます。想定値であり、売上および利益を保証するものではございません。金額はすべて税抜でございます。", {
    x: M + 0.18, y: 1.14, w: CW - 0.36, h: 0.44,
    fontFace: FG, fontSize: 8.8, color: C.body, margin: 0, valign: "middle",
  });

  sec(s, M, 1.74, CW, "仕組みの違い", "運営はいずれも貴社スタッフ。違いは「売上の帰属」と「工事の負担」の2点でございます");
  const pw = (CW - 0.3) / 2;
  const PLAN = [
    {
      t: "業務委託型", sub: "弊社が事業主体となります", f: C.soft, bd: C.ruleT, c: C.ink,
      rows: [["運　営", "未来屋書店様（弊社からの業務委託）"], ["売上の帰属", "セルフカフェ"], ["内装工事・什器", "セルフカフェが負担"], ["貴社のご負担", "設置工事費 30〜50万円"], ["貴社の収入", "売上の25％ ＋ 月額固定 50,000円"]],
      pf: "206,250",
    },
    {
      t: "FCプラン", sub: "貴社が事業主体となられます", f: C.grPale, bd: C.gr, c: C.gr, hot: true,
      rows: [["運　営", "未来屋書店様"], ["売上の帰属", "未来屋書店様"], ["内装工事・什器", "未来屋書店様が負担（ご投資に含む）"], ["貴社のご投資", "9,000,000円"], ["貴社の収入", "売上高（経費を差し引いた残りが利益）"]],
      pf: "520,000",
    },
  ];
  PLAN.forEach((p, i) => {
    const x = M + i * (pw + 0.3);
    s.addShape("rect", { x, y: 2.10, w: pw, h: 2.52, fill: { color: C.white }, line: { color: p.bd, width: p.hot ? 1.25 : 0.75 } });
    s.addShape("rect", { x, y: 2.10, w: pw, h: 0.42, fill: { color: p.f } });
    s.addText(p.t, { x: x + 0.2, y: 2.10, w: pw * 0.42, h: 0.42, fontFace: FM, fontSize: 13, bold: true, color: p.c, margin: 0, valign: "middle" });
    s.addText(p.sub, { x: x + pw * 0.42, y: 2.10, w: pw * 0.58 - 0.2, h: 0.42, fontFace: FG, fontSize: 8.5, color: C.body, align: "right", margin: 0, valign: "middle" });
    p.rows.forEach((r, j) => {
      const y = 2.60 + j * 0.3;
      s.addText(r[0], { x: x + 0.2, y, w: 1.45, h: 0.28, fontFace: FG, fontSize: 8.8, bold: true, color: C.ink, margin: 0, valign: "middle" });
      s.addText(r[1], { x: x + 1.65, y, w: pw - 1.85, h: 0.28, fontFace: FG, fontSize: 8.8, color: C.body, margin: 0, valign: "middle" });
    });
    s.addShape("rect", { x: x + 0.2, y: 4.14, w: pw - 0.4, h: 0.006, fill: { color: C.hair } });
    s.addText("2年目（巡航）の営業利益", { x: x + 0.2, y: 4.20, w: 2.2, h: 0.34, fontFace: FG, fontSize: 8.5, color: C.body, margin: 0, valign: "middle" });
    s.addText(
      [
        { text: p.pf, options: { fontFace: FG, fontSize: 20, bold: true, color: p.c } },
        { text: " 円／月", options: { fontFace: FG, fontSize: 9.5, bold: true, color: C.body } },
      ],
      { x: x + pw - 3.0, y: 4.18, w: 2.8, h: 0.38, align: "right", margin: 0, valign: "middle" }
    );
  });
  s.addText("FCプランは 2.52倍 でございます。", {
    x: M, y: 4.70, w: CW, h: 0.22,
    fontFace: FG, fontSize: 9.5, bold: true, color: C.gr, align: "center", margin: 0, valign: "middle",
  });

  sec(s, M, 5.04, CW, "条件の比較");
  const r3 = (a, b, c) => [
    td(a, { size: 8.8 }),
    td(b, { align: "center", size: 8.8 }),
    td(c, { align: "center", size: 8.8, fill: C.grPale }),
  ];
  s.addTable(
    [
      [th("観点"), th("業務委託型", "center"), th("FCプラン", "center")],
      r3("貴社がご負担になる経費", "水道光熱費・雑費（通信費・消耗品費等）", "原料・水道光熱費・設備利用料・ロイヤリティ・雑費 ほか"),
      r3("日常業務（清掃・補充・発注）", "貴社スタッフ", "貴社スタッフ"),
      r3("清掃の外部委託・警備", "有人のため不要", "有人のため不要"),
      r3("売上変動のリスク", "小（月額固定5万円が下支え）", "貴社がご負担"),
      r3("食品衛生責任者", "不要", "1名のご選任が必要"),
      r3("保健所の営業許可", "セルフカフェ名義", "貴社名義（弊社が申請を代行）"),
    ],
    { x: M, y: 5.40, w: CW, colW: [3.0, 3.4, CW - 6.4], rowH: [0.24, 0.3, 0.24, 0.24, 0.26, 0.24, 0.24], border: TB, autoPage: false }
  );

  foot(s, [
    "※　日常業務（清掃・補充・発注）は、いずれの方式でも貴社スタッフによる運用を想定しております。人件費は両案とも未計上でございます。",
    "※　家賃は、カフェ区画が貴社の賃借面積に含まれる前提で0円としております。",
  ]);
}

/* ================================================= P2 月次損益 */
{
  const s = pres.addSlide();
  const r = YR[1];
  head(s, "月次損益の比較", "2年目・巡航時（1日75杯）　同じ条件・同じ経費区分で並べております　単位：円（税抜）");

  const hw = (CW - 0.3 * 2) / 3;
  const HERO = [
    { l: "業務委託型　営業利益", v: yen(r.iPf), f: C.soft, bd: C.ruleT, c: C.ink },
    { l: "FCプラン　営業利益", v: yen(r.fPf), f: C.white, bd: C.gr, c: C.gr },
    { l: "差　額", v: "＋" + yen(r.fPf - r.iPf), f: C.grPale, bd: C.gr, c: C.gr, hot: true },
  ];
  HERO.forEach((h, i) => {
    const x = M + i * (hw + 0.3);
    s.addShape("rect", { x, y: 1.14, w: hw, h: 0.92, fill: { color: h.f }, line: { color: h.bd, width: h.hot ? 1.25 : 0.75 } });
    s.addText(h.l, { x: x + 0.18, y: 1.19, w: hw - 0.36, h: 0.22, fontFace: FG, fontSize: 9, bold: true, color: h.c, margin: 0, valign: "middle" });
    s.addText(
      [
        { text: h.v, options: { fontFace: FG, fontSize: h.hot ? 25 : 22, bold: true, color: h.c } },
        { text: " 円／月", options: { fontFace: FG, fontSize: 9.5, bold: true, color: C.body } },
      ],
      { x: x + 0.18, y: 1.42, w: hw - 0.36, h: 0.56, margin: 0, valign: "middle" }
    );
  });
  [0, 1].forEach((i) => {
    s.addText(i === 0 ? "→" : "＝", {
      x: M + (i + 1) * hw + i * 0.3, y: 1.40, w: 0.3, h: 0.4,
      fontFace: FG, fontSize: 14, bold: true, color: C.muted, align: "center", margin: 0, valign: "middle",
    });
  });
  s.addText("年間では ＋" + yen((r.fPf - r.iPf) * 12) + "円、販売1杯あたりに換算いたしますと 139円 の差でございます。", {
    x: M, y: 2.12, w: CW, h: 0.22,
    fontFace: FG, fontSize: 9.5, bold: true, color: C.gr, align: "center", margin: 0, valign: "middle",
  });

  sec(s, M, 2.46, CW, "月次損益の内訳");
  const c1 = (v, o = {}) => td(v, Object.assign({ align: "right", size: 9 }, o));
  const c2 = (v, o = {}) => td(v, Object.assign({ align: "right", size: 9, fill: C.grPale }, o));
  const c3 = (v, o = {}) => td(v, Object.assign({ align: "right", size: 9, bold: true }, o));
  const ROWS = [
    ["セルフカフェ売上（月）", yen(r.sales), yen(r.sales), "－", {}],
    ["貴社の収入", yen(r.iIn), yen(r.sales), "＋" + yen(r.sales - r.iIn), { b: true }],
    ["　ドリンク原料（約20％）", "－", neg(r.mat), neg(r.mat), {}],
    ["　水道光熱費（書店営業時間ベース）", neg(UTL), neg(UTL), "－", {}],
    ["　マシン使用料（38,000円×2台）", "－", neg(MACH), neg(MACH), {}],
    ["　決済端末利用料（5,000円×2台）", "－", neg(TERM), neg(TERM), {}],
    ["　クラウド利用料", "－", neg(CLOUD), neg(CLOUD), {}],
    ["　営業サポート（ミニマムサポート）", "－", neg(SUP), neg(SUP), {}],
    ["　ロイヤリティ（一律・売上連動なし）", "－", neg(ROY), neg(ROY), {}],
    ["　雑費（通信費・消耗品費等）", neg(ETC), neg(ETC), "－", {}],
    ["　清掃業務委託費　※有人のため不要", "－", "（" + yen(CLN) + "）", "－", { g: true }],
    ["　セキュリティ費　※有人のため不要", "－", "（" + yen(SECF) + "）", "－", { g: true }],
    ["ランニングコスト計", neg(IC), neg(r.fCost), neg(r.fCost - IC), { b: true, cost: true }],
    ["営業利益（月）", yen(r.iPf), yen(r.fPf), "＋" + yen(r.fPf - r.iPf), { b: true, f: C.band, big: true }],
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
  s.addText("収入が ＋658,750円 増え、経費が ▲345,000円 増えるため、差引 ＋313,750円 となります。", {
    x: M, y: 6.76, w: CW, h: 0.24,
    fontFace: FG, fontSize: 9.5, bold: true, color: C.gr, align: "center", margin: 0, valign: "middle",
  });

  foot(s, [
    "※　業務委託型の「貴社の収入」は、セルフカフェ売上の25％ ＋ 月額固定50,000円でございます。FCプランは売上高がそのまま貴社に帰属いたします。",
    "※　清掃業務委託費・セキュリティ費は、貴社書店内への併設（有人）のため不要でございます。参考として括弧内に標準額を記載しております。家賃は0円。",
  ]);
}

/* ================================================= P3 3か年 */
{
  const s = pres.addSlide();
  head(s, "3か年の推移", "立ち上がりを織り込んだ3か年　単位：円（税抜）");

  sec(s, M, 1.14, CW, "年度別の収支");
  const cI = (v, o = {}) => td(v, Object.assign({ align: "right", size: 8.5 }, o));
  const cF = (v, o = {}) => td(v, Object.assign({ align: "right", size: 8.5, fill: C.grPale }, o));
  const line = (label, fi, ff, o = {}) => [
    td(label, { size: 8.5, bold: !!o.b, color: o.b ? C.ink : C.body, fill: o.lf }),
    ...YR.flatMap((r) => [cI(fi(r), { bold: !!o.b, color: o.b ? C.ink : C.body }), cF(ff(r), { bold: !!o.b, color: o.b && !o.cost ? C.gr : C.body })]),
  ];
  s.addTable(
    [
      [th("項目"), ...YR.flatMap((r) => [th(r.y + "\n業務委託型", "right"), th(r.y + "\nFCプラン", "right")])],
      line("販売杯数（1日）", (r) => r.cups + "杯", (r) => r.cups + "杯"),
      line("セルフカフェ売上（月）", (r) => yen(r.sales), (r) => yen(r.sales)),
      line("貴社の収入", (r) => yen(r.iIn), (r) => yen(r.sales), { b: true }),
      line("ランニングコスト計", () => neg(IC), (r) => neg(r.fCost), { cost: true }),
      line("営業利益（月）", (r) => yen(r.iPf), (r) => yen(r.fPf), { b: true, lf: C.band }),
      line("営業利益（年）", (r) => yen(r.iPfY), (r) => yen(r.fPfY), { b: true, lf: C.band }),
    ],
    { x: M, y: 1.50, w: CW, colW: [2.55, 1.33, 1.33, 1.33, 1.33, 1.33, 1.33], rowH: [0.36, 0.26, 0.26, 0.3, 0.28, 0.32, 0.32], border: TB, autoPage: false }
  );
  s.addShape("rect", { x: M, y: 3.82, w: CW, h: 0.44, fill: { color: C.grPale } });
  s.addText("3年間の営業利益の合計　　業務委託型 " + yen(T_I) + "円　／　FCプラン " + yen(T_F) + "円　　差額 ＋" + yen(T_F - T_I) + "円", {
    x: M + 0.2, y: 3.82, w: CW - 0.4, h: 0.44,
    fontFace: FG, fontSize: 10, bold: true, color: C.gr, margin: 0, valign: "middle",
  });

  sec(s, M, 4.52, CW, "営業利益（年）の比較");
  bars(s, {
    x: M + 1.0, y: 5.12, w: CW - 2.0, h: 1.52, max: 700, unit: "単位：万円／年", gap: 0.16,
    data: YR.flatMap((r) => [
      { v: r.iPfY / 10000, t: (r.iPfY / 10000).toFixed(1), label: r.y, sub: "業務委託型", color: C.barL, fs: 9 },
      { v: r.fPfY / 10000, t: (r.fPfY / 10000).toFixed(1), label: r.y, sub: "FCプラン", color: C.gr, fs: 9 },
    ]),
  });

  foot(s, [
    "※　販売杯数は、既にご提出済みの「開業後3か年の売上見込み」と同一でございます。1年目は立ち上がり期として巡航値の約83％で置いております。",
    "※　想定値であり、売上および利益を保証するものではございません。人件費は未計上。",
  ]);
}

/* ================================================= P4 投資と期限 */
{
  const s = pres.addSlide();
  head(s, "ご投資と回収、ご判断の期限", "ご投資は約20.3か月で回収され、3年間の累計でもFCプランが上回ります");

  const lw = 5.5, rx = M + lw + 0.5, rw = R - rx;
  sec(s, M, 1.14, lw, "ご投資額の内訳（概算）");
  const IV = [["加盟金", "1,000,000"], ["保証金", "500,000"], ["内装造作工事", "4,200,000"], ["家具什器", "1,300,000"], ["サイン工事", "500,000"], ["開業準備金", "500,000"], ["設計・申請・諸経費", "1,000,000"]];
  s.addTable(
    [
      [th("項目"), th("金額", "right")],
      ...IV.map((x) => [td(x[0], { size: 8.8 }), td(x[1], { align: "right", size: 8.8 })]),
      [td("合　計", { fill: C.bandD, bold: true, color: C.gr, size: 9.5 }), td("9,000,000", { align: "right", fill: C.bandD, bold: true, color: C.gr, size: 11 })],
    ],
    { x: M, y: 1.50, w: lw, colW: [3.5, lw - 3.5], rowH: [0.22, 0.26, 0.24, 0.26, 0.24, 0.26, 0.24, 0.28, 0.32], border: TB, autoPage: false }
  );
  s.addText("マシン・決済端末は月額の利用料でのご提供のため、初期のご購入は不要でございます。秋田店は既存造作を残置いただく前提のため、工事範囲が限定されます。仕様確定後に正式なお見積りをご提出いたします。", {
    x: M, y: 4.00, w: lw, h: 0.62,
    fontFace: FG, fontSize: 8.3, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
  });

  sec(s, rx, 1.14, rw, "投資回収と3年間のまとめ");
  s.addTable(
    [
      [th("年度"), th("FC 営業利益（年）", "right"), th("残　額", "right")],
      [td("開業時", { size: 8.8 }), td("－", { align: "right", size: 8.8 }), td("9,000,000", { align: "right", size: 8.8 })],
      [td("1年目（62杯／日）", { size: 8.8 }), td("4,667,520", { align: "right", size: 8.8 }), td("4,332,480", { align: "right", size: 8.8 })],
      [td("2年目 8.3か月目", { size: 8.8, fill: C.bandD, bold: true, color: C.gr }), td("回収完了", { align: "right", size: 8.8, fill: C.bandD, bold: true, color: C.gr }), td("0", { align: "right", size: 9.5, fill: C.bandD, bold: true, color: C.gr })],
    ],
    { x: rx, y: 1.50, w: rw, colW: [1.9, 1.7, rw - 3.6], rowH: [0.24, 0.26, 0.26, 0.32], border: TB, autoPage: false }
  );
  s.addTable(
    [
      [th("3年間の累計"), th("業務委託型", "right"), th("FCプラン", "right")],
      [td("営業利益の合計", { size: 8.8 }), td(yen(T_I), { align: "right", size: 8.8 }), td(yen(T_F), { align: "right", size: 8.8, bold: true, color: C.ink })],
      [td("初期費用・ご投資", { size: 8.8 }), td(neg(IT0), { align: "right", size: 8.8 }), td(neg(INV), { align: "right", size: 8.8 })],
      [td("差　引", { fill: C.bandD, bold: true, color: C.gr, size: 9.5 }), td(yen(T_I - IT0), { align: "right", fill: C.bandD, size: 9.5 }), td(yen(T_F - INV), { align: "right", fill: C.bandD, bold: true, color: C.gr, size: 11 })],
    ],
    { x: rx, y: 2.70, w: rw, colW: [1.9, 1.7, rw - 3.6], rowH: [0.24, 0.26, 0.26, 0.32], border: TB, autoPage: false }
  );
  s.addShape("rect", { x: rx, y: 3.92, w: rw, h: 0.70, fill: { color: C.grPale }, line: { color: C.gr, width: 1 } });
  s.addText("3年間の累計でも、FCプランが 1,880,240円 上回ります。4年目以降は年間 3,931,320円 の差となってまいります。", {
    x: rx + 0.18, y: 3.92, w: rw - 0.36, h: 0.70,
    fontFace: FG, fontSize: 9, bold: true, color: C.gr, margin: 0, valign: "middle", lineSpacingMultiple: 1.25,
  });

  sec(s, M, 4.84, CW, "FCプランで進める場合の手続きと期限", "開業 2027年3月上旬");
  s.addTable(
    [
      [th("必要となる手続き"), th("ご担当"), th("期限の目安", "right")],
      [td("ご投資 900万円の社内ご決裁", { size: 8.8, fill: C.band }), td("貴社", { size: 8.8, fill: C.band }), td("12月上旬", { align: "right", size: 9, bold: true, color: C.gr, fill: C.band })],
      [td("什器・機器の発注（リードタイム4〜6週間）", { size: 8.8, fill: C.band }), td("発注者が貴社となります", { size: 8.8, fill: C.band }), td("12月上旬", { align: "right", size: 9, bold: true, color: C.gr, fill: C.band })],
      [td("FC加盟契約の締結（業務委託契約からの変更）", { size: 8.8 }), td("両社", { size: 8.8 }), td("12月中旬", { align: "right", size: 8.8 })],
      [td("食品衛生責任者のご選任（1名）", { size: 8.8 }), td("貴社", { size: 8.8 }), td("1月中", { align: "right", size: 8.8 })],
      [td("保健所への営業許可申請（名義が貴社となります）", { size: 8.8 }), td("弊社代行", { size: 8.8 }), td("2月中", { align: "right", size: 8.8 })],
    ],
    { x: M, y: 5.20, w: CW, colW: [5.2, 2.6, CW - 7.8], rowH: [0.24, 0.3, 0.3, 0.28, 0.28, 0.28], border: TB, autoPage: false }
  );
  s.addText("什器・機器の発注が12月上旬期限でございます。FCプランでは発注の主体が貴社となりますため、ご決裁もそれまでにお願いしたく存じます。これを過ぎますと、3月上旬の開業が難しくなります。", {
    x: M, y: 7.02, w: CW, h: 0.3,
    fontFace: FG, fontSize: 9, bold: true, color: C.ink, margin: 0, valign: "middle",
  });

  foot(s, [
    "※　ご投資額は概算でございます。仕様確定後にあらためて正式なお見積りをご提出いたします。金額はすべて税抜。",
  ]);
}

pres.writeFile({ fileName: path.join(__dirname, "akita-fc.pptx") }).then((f) => console.log("written:", f));
