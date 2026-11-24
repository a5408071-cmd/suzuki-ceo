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
const TOTAL = 3;

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
const FX = 60000 + 86000 + 50000 + 20000;   // 水光熱・機械使用料・ロイヤリティ・雑費
const INV = 9000000;
const YR = [
  { y: "1年目", term: "2027年3月〜2028年2月", cups: 62, note: "立ち上がり" },
  { y: "2年目", term: "2028年3月〜2029年2月", cups: 75, note: "巡航", hot: true },
  { y: "3年目", term: "2029年3月〜2030年2月", cups: 77, note: "巡航＋α" },
].map((r) => {
  const sales = r.cups * UP * DAYS;
  const mat = Math.floor(sales * 0.2);
  const pf = sales - mat - FX;
  const fee = Math.floor(sales * 0.25) + FIXFEE;
  return Object.assign({}, r, { sales, mat, pf, fee, d: pf - fee, pfY: pf * 12, feeY: fee * 12, dY: (pf - fee) * 12 });
});
const SUM_D = YR.reduce((a, r) => a + r.dY, 0);
const SUM_PF = YR.reduce((a, r) => a + r.pfY, 0);
const yen = (v) => v.toLocaleString();

/* ================================================= P1 */
{
  const s = pres.addSlide();
  head(s, "秋田店　セルフカフェ FCプランのご提案", "30坪・1日75杯（巡航）／2027年3月上旬 開業予定　販売杯数は既にご提出の3か年見込みと同一でございます");

  s.addShape("rect", { x: M, y: 1.14, w: CW, h: 0.52, fill: { color: C.white }, line: { color: C.ruleS, width: 1 } });
  s.addText("【ご留意】　本資料は開業前の机上試算でございます。想定値であり、売上および利益を保証するものではございません。金額はすべて税抜でございます。", {
    x: M + 0.18, y: 1.14, w: CW - 0.36, h: 0.52,
    fontFace: FG, fontSize: 8.8, color: C.body, margin: 0, valign: "middle",
  });

  // ヒーロー：差額
  sec(s, M, 1.84, CW, "業務委託型との差額", "FCプランで貴社に残る利益の増加分");
  const HERO = [
    { l: "月　次（巡航時）", v: "253,750", u: "円／月" },
    { l: "年　間（巡航時）", v: "3,045,000", u: "円／年" },
    { l: "3年間の合計", v: "8,220,240", u: "円", hot: true },
  ];
  const hw = (CW - 0.22 * 2) / 3;
  HERO.forEach((h, i) => {
    const x = M + i * (hw + 0.22);
    s.addShape("rect", { x, y: 2.20, w: hw, h: 1.02, fill: { color: h.hot ? C.grPale : C.soft }, line: { color: h.hot ? C.gr : C.ruleT, width: h.hot ? 1 : 0.75 } });
    s.addText(h.l, {
      x: x + 0.18, y: 2.28, w: hw - 0.36, h: 0.22,
      fontFace: FG, fontSize: 9, bold: true, color: h.hot ? C.gr : C.ink, margin: 0, valign: "middle",
    });
    s.addText(
      [
        { text: "＋", options: { fontFace: FG, fontSize: 14, bold: true, color: h.hot ? C.gr : C.ink } },
        { text: h.v, options: { fontFace: FG, fontSize: 26, bold: true, color: h.hot ? C.gr : C.ink } },
        { text: " " + h.u, options: { fontFace: FG, fontSize: 10, bold: true, color: C.body } },
      ],
      { x: x + 0.18, y: 2.54, w: hw - 0.36, h: 0.56, margin: 0, valign: "middle" }
    );
  });

  const lw = 7.2, rx = M + lw + 0.42, rw = R - rx;
  sec(s, M, 3.42, lw, "年度別の差額");
  const rc = (t, o) => td(t, Object.assign({ align: "right", size: 9 }, o));
  s.addTable(
    [
      [th("年度"), th("販売杯数", "right"), th("業務委託型\nお受取額（月）", "right"), th("FCプラン\n償却前営業利益（月）", "right"), th("差　額", "right")],
      ...YR.map((r) => [
        td(r.y + "　" + r.note, { size: 9, fill: r.hot ? C.band : C.white }),
        rc(r.cups + "杯／日", { fill: r.hot ? C.band : C.white }),
        rc(yen(r.fee), { fill: r.hot ? C.band : C.white }),
        rc(yen(r.pf), { bold: true, color: C.ink, fill: r.hot ? C.band : C.white }),
        rc("＋" + yen(r.d), { bold: true, color: C.gr, fill: r.hot ? C.band : C.white }),
      ]),
      [td("3年間の合計（年額ベース）", { fill: C.bandD, bold: true, color: C.gr, size: 9 }), rc("－", { fill: C.bandD }), rc(yen(YR.reduce((a, r) => a + r.feeY, 0)), { fill: C.bandD }), rc(yen(SUM_PF), { fill: C.bandD, bold: true, color: C.ink }), rc("＋" + yen(SUM_D), { fill: C.bandD, bold: true, color: C.gr, size: 11 })],
    ],
    { x: M, y: 3.78, w: lw, colW: [2.0, 1.1, 1.4, 1.5, 1.2], rowH: [0.38, 0.26, 0.26, 0.26, 0.32], border: TB, autoPage: false }
  );
  s.addText("※　上記はご投資額を控除する前の差額でございます。ご投資額と回収につきましては次ページでご説明申し上げます。", {
    x: M, y: 5.36, w: lw, h: 0.2,
    fontFace: FG, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
  });

  sec(s, rx, 3.42, rw, "月次収入の比較", "2年目・巡航時");
  bars(s, {
    x: rx + 0.42, y: 3.96, w: rw - 0.74, h: 1.26, max: 60, unit: "単位：万円／月", gap: 0.45,
    data: [
      { v: 28.6, t: "28.6", label: "業務委託型", sub: "お受取額", color: C.barL, fs: 10 },
      { v: 54.0, t: "54.0", label: "FCプラン", sub: "償却前営業利益", color: C.gr, fs: 10 },
    ],
  });
  s.addText("約1.9倍でございます。", {
    x: rx, y: 5.62, w: rw, h: 0.2,
    fontFace: FG, fontSize: 9, bold: true, color: C.gr, align: "center", margin: 0, valign: "middle",
  });

  sec(s, M, 5.84, CW, "FCプランの月次損益（2年目・巡航時）");
  const FL = [
    ["売上高", "945,000"], ["ドリンク原料（20％）", "▲ 189,000"], ["清掃業務委託費", "（30,000）"],
    ["水道光熱費", "▲ 60,000"], ["機械使用料", "▲ 86,000"], ["ロイヤリティ", "▲ 50,000"],
    ["雑費", "▲ 20,000"], ["家賃", "0"], ["償却前営業利益", "540,000"],
  ];
  const fw = (CW - 0.12 * 8) / 9;
  FL.forEach((r, i) => {
    const x = M + i * (fw + 0.12), last = i === FL.length - 1;
    s.addShape("rect", { x, y: 6.20, w: fw, h: 0.62, fill: { color: last ? C.bandD : C.soft }, line: { color: last ? C.gr : C.ruleT, width: last ? 1 : 0.5 } });
    s.addText(r[0], {
      x: x + 0.06, y: 6.25, w: fw - 0.12, h: 0.22,
      fontFace: FG, fontSize: 7.2, bold: last, color: last ? C.gr : C.body, align: "center", margin: 0, valign: "middle",
    });
    s.addText(r[1], {
      x: x + 0.06, y: 6.48, w: fw - 0.12, h: 0.26,
      fontFace: FG, fontSize: last ? 10 : 8.5, bold: true, color: last ? C.gr : (r[1][0] === "（" ? C.muted : C.ink), align: "center", margin: 0, valign: "middle",
    });
  });

  foot(s, [
    "※　清掃業務委託費は貴社スタッフでご対応いただく前提のため計上しておりません。セキュリティ費は書店内への併設のため不要でございます。",
    "※　人件費は未計上。家賃はカフェ区画が貴社の賃借面積に含まれる前提で0円としております。平均荒利率は80.0％でございます。",
  ]);
}

/* ================================================= P2 */
{
  const s = pres.addSlide();
  head(s, "ご投資と回収", "ご投資は約19.6か月で回収され、以降は利益がそのまま残ります");

  const lw = 5.6, rx = M + lw + 0.5, rw = R - rx;
  sec(s, M, 1.14, lw, "ご投資額の内訳（概算）");
  const IV = [["加盟金", "1,000,000"], ["保証金", "500,000"], ["内装造作工事", "4,200,000"], ["家具什器", "1,300,000"], ["サイン工事", "500,000"], ["開業準備金", "500,000"], ["設計・申請・諸経費", "1,000,000"]];
  s.addTable(
    [
      [th("項目"), th("金額", "right")],
      ...IV.map((r) => [td(r[0], { size: 8.8 }), td(r[1], { align: "right", size: 8.8 })]),
      [td("合　計", { fill: C.bandD, bold: true, color: C.gr, size: 9.5 }), td("9,000,000", { align: "right", fill: C.bandD, bold: true, color: C.gr, size: 11 })],
    ],
    { x: M, y: 1.50, w: lw, colW: [3.6, lw - 3.6], rowH: [0.22, 0.26, 0.24, 0.26, 0.24, 0.26, 0.24, 0.28, 0.32], border: TB, autoPage: false }
  );
  s.addText("マシンおよび決済端末は月額の機械使用料でのご提供のため、初期のご購入は不要でございます。秋田店は既存造作を残置いただく前提のため、新規出店の場合より工事範囲が限定されます。仕様確定後に正式なお見積りをご提出いたします。", {
    x: M, y: 4.00, w: lw, h: 0.62,
    fontFace: FG, fontSize: 8.3, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
  });

  sec(s, rx, 1.14, rw, "投資回収の見通し");
  s.addTable(
    [
      [th("年度"), th("償却前営業利益（年）", "right"), th("残　額", "right")],
      [td("開業時", { size: 8.8 }), td("－", { align: "right", size: 8.8 }), td("9,000,000", { align: "right", size: 8.8 })],
      [td("1年目（62杯／日）", { size: 8.8 }), td("4,907,520", { align: "right", size: 8.8 }), td("4,092,480", { align: "right", size: 8.8 })],
      [td("2年目 7.6か月目", { size: 8.8, fill: C.bandD, bold: true, color: C.gr }), td("回収完了", { align: "right", size: 8.8, fill: C.bandD, bold: true, color: C.gr }), td("0", { align: "right", size: 9.5, fill: C.bandD, bold: true, color: C.gr })],
      [td("2年目（75杯／日）", { size: 8.8 }), td("6,480,000", { align: "right", size: 8.8 }), td("－", { align: "right", size: 8.8 })],
      [td("3年目（77杯／日）", { size: 8.8 }), td("6,721,920", { align: "right", size: 8.8 }), td("－", { align: "right", size: 8.8 })],
    ],
    { x: rx, y: 1.50, w: rw, colW: [1.72, 1.55, rw - 3.27], rowH: [0.24, 0.26, 0.26, 0.32, 0.26, 0.26], border: TB, autoPage: false }
  );
  s.addShape("rect", { x: rx, y: 3.26, w: rw, h: 1.14, fill: { color: C.grPale }, line: { color: C.gr, width: 1 } });
  s.addText("開業から約19.6か月でご投資を回収いたします。", {
    x: rx + 0.18, y: 3.34, w: rw - 0.36, h: 0.24,
    fontFace: FM, fontSize: 11, bold: true, color: C.gr, margin: 0, valign: "middle",
  });
  s.addText("3年契約の折り返し前に回収が完了いたします。以降は、月540,000円（3年目は560,160円）がそのまま貴社の利益として残ってまいります。3年目末までの累計では、ご投資額を差し引いて 9,109,440円 が手元に残る計算でございます。", {
    x: rx + 0.18, y: 3.58, w: rw - 0.36, h: 0.76,
    fontFace: FG, fontSize: 8.5, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
  });

  sec(s, M, 4.84, CW, "3か年の償却前営業利益");
  bars(s, {
    x: M + 1.2, y: 5.44, w: CW - 2.4, h: 1.30, max: 800, unit: "単位：万円／年", gap: 0.9,
    data: YR.map((r) => ({
      v: r.pfY / 10000, t: (r.pfY / 10000).toFixed(1), label: r.y, sub: r.cups + "杯／日",
      color: r.hot ? C.gr : C.barD, fs: 10,
    })),
  });

  foot(s, [
    "※　ご投資額は概算でございます。仕様確定後にあらためて正式なお見積りをご提出いたします。金額はすべて税抜。",
    "※　償却前営業利益（減価償却前）でございます。人件費は未計上。想定値であり、利益を保証するものではございません。",
  ]);
}

/* ================================================= P3 */
{
  const s = pres.addSlide();
  head(s, "FCプランをお薦めする理由と、ご判断の期限", "什器・機器の発注が12月上旬期限のため、ご決裁もそれまでにお願いしたく存じます");

  const lw = 5.9, rx = M + lw + 0.5, rw = R - rx;
  sec(s, M, 1.14, lw, "FCプランをお薦めする理由");
  [
    ["売上・利益がすべて貴社のものになります", "巡航時 月540,000円。業務委託型の約1.9倍。差額は3年間で8,220,240円でございます。"],
    ["ロイヤリティが売上連動ではございません", "一律 月50,000円。売上が伸びるほど利益率が向上する仕組みでございます。"],
    ["損益分岐点が低く、下振れに強い構造でございます", "1日約22杯で損益分岐。想定75杯に対して3倍以上の余裕がございます。"],
    ["マシンは月額使用料でご提供いたします", "初期のご購入は不要。更新・故障対応も弊社にて行います。"],
    ["自社業態としての知見が蓄積されます", "仙台泉をはじめ、今後のご出店のご判断材料となります。"],
  ].forEach((t, i) => {
    const y = 1.52 + i * 0.56;
    s.addText("（" + (i + 1) + "）" + t[0], { x: M, y, w: lw, h: 0.2, fontFace: FG, fontSize: 9, bold: true, color: C.ink, margin: 0, valign: "middle" });
    s.addText(t[1], { x: M + 0.38, y: y + 0.21, w: lw - 0.38, h: 0.3, fontFace: FG, fontSize: 8.2, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.2 });
  });

  sec(s, rx, 1.14, rw, "ご確認いただきたい点");
  [
    ["ご投資 900万円", "約19.6か月で回収の見通しでございます。"],
    ["売上変動のリスク", "業務委託型の月額固定50,000円のような下支えはございません。"],
    ["食品衛生責任者", "1名のご選任が必要でございます。保健所の申請手続きは弊社にて代行いたします。"],
  ].forEach((t, i) => {
    const y = 1.52 + i * 0.56;
    s.addText("（" + (i + 1) + "）" + t[0], { x: rx, y, w: rw, h: 0.2, fontFace: FG, fontSize: 9, bold: true, color: C.ink, margin: 0, valign: "middle" });
    s.addText(t[1], { x: rx + 0.38, y: y + 0.21, w: rw - 0.38, h: 0.3, fontFace: FG, fontSize: 8.2, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.2 });
  });
  s.addShape("rect", { x: rx, y: 3.30, w: rw, h: 0.86, fill: { color: C.grPale }, line: { color: C.gr, width: 1 } });
  s.addText("業務委託型でのご契約も、もちろん可能でございます。", {
    x: rx + 0.18, y: 3.36, w: rw - 0.36, h: 0.22,
    fontFace: FG, fontSize: 9, bold: true, color: C.gr, margin: 0, valign: "middle",
  });
  s.addText("ご投資を抑えて始められる点が利点でございます。一方、継続してご運営いただくご見通しがあれば、FCプランのほうが残る利益は大きくなってまいります。", {
    x: rx + 0.18, y: 3.58, w: rw - 0.36, h: 0.52,
    fontFace: FG, fontSize: 8.3, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
  });

  sec(s, M, 4.42, CW, "FCプランで進める場合の手続きと期限", "開業 2027年3月上旬");
  s.addTable(
    [
      [th("必要となる手続き"), th("ご担当"), th("期限の目安", "right")],
      [td("ご投資 900万円の社内ご決裁", { size: 8.8, fill: C.band }), td("貴社", { size: 8.8, fill: C.band }), td("12月上旬", { align: "right", size: 9, bold: true, color: C.gr, fill: C.band })],
      [td("什器・機器の発注（リードタイム4〜6週間）", { size: 8.8, fill: C.band }), td("発注者が貴社となります", { size: 8.8, fill: C.band }), td("12月上旬", { align: "right", size: 9, bold: true, color: C.gr, fill: C.band })],
      [td("FC加盟契約の締結（業務委託契約からの変更）", { size: 8.8 }), td("両社", { size: 8.8 }), td("12月中旬", { align: "right", size: 8.8 })],
      [td("食品衛生責任者のご選任（1名）", { size: 8.8 }), td("貴社", { size: 8.8 }), td("1月中", { align: "right", size: 8.8 })],
      [td("保健所への営業許可申請（名義が貴社となります）", { size: 8.8 }), td("弊社代行", { size: 8.8 }), td("2月中", { align: "right", size: 8.8 })],
    ],
    { x: M, y: 4.78, w: CW, colW: [5.2, 2.6, CW - 7.8], rowH: [0.24, 0.3, 0.3, 0.28, 0.28, 0.28], border: TB, autoPage: false }
  );
  s.addText("什器・機器の発注が12月上旬期限でございます。FCプランでは発注の主体が貴社となりますため、ご決裁もそれまでにお願いしたく存じます。これを過ぎますと、3月上旬の開業が難しくなります。", {
    x: M, y: 6.60, w: CW, h: 0.44,
    fontFace: FG, fontSize: 9, bold: true, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
  });

  foot(s, [
    "※　いずれも開業前の机上試算でございます。金額はすべて税抜。想定値であり、売上および利益を保証するものではございません。",
  ]);
}

pres.writeFile({ fileName: path.join(__dirname, "akita-fc.pptx") }).then((f) => console.log("written:", f));
