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
const INV = 9000000, IT0 = 500000;
const YR = [
  { y: "1年目", term: "2027年3月〜2028年2月", cups: 62, note: "立ち上がり" },
  { y: "2年目", term: "2028年3月〜2029年2月", cups: 75, note: "巡航", hot: true },
  { y: "3年目", term: "2029年3月〜2030年2月", cups: 77, note: "巡航＋α" },
].map((r) => {
  const sales = r.cups * UP * DAYS;
  const mat = Math.floor(sales * 0.2);
  const pf = sales - mat - FX;
  const roy = Math.floor(sales * 0.25);
  const fee = roy + FIXFEE;
  return Object.assign({}, r, { sales, mat, pf, roy, fee, pfY: pf * 12, feeY: fee * 12 });
});
const SUM_FC = YR.reduce((a, r) => a + r.pfY, 0);
const SUM_IT = YR.reduce((a, r) => a + r.feeY, 0);
const yen = (v) => v.toLocaleString();

/* ================================================= P1 */
{
  const s = pres.addSlide();
  head(s, "秋田店　業務委託型とFCプランの比較", "30坪・1日75杯（巡航）／2027年3月上旬 開業予定　立ち上がりを織り込んだ3か年で比較しております");

  s.addShape("rect", { x: M, y: 1.14, w: CW, h: 0.58, fill: { color: C.white }, line: { color: C.ruleS, width: 1 } });
  s.addText("【ご留意】　本資料は、開業前の時点における机上の試算でございます。", {
    x: M + 0.18, y: 1.19, w: CW - 0.36, h: 0.22,
    fontFace: FM, fontSize: 10.5, bold: true, color: C.gr, margin: 0, valign: "middle",
  });
  s.addText("販売杯数は、既にご提出済みの「開業後3か年の売上見込み」と同一でございます。想定値であり、売上および利益を保証するものではございません。", {
    x: M + 0.18, y: 1.42, w: CW - 0.36, h: 0.24,
    fontFace: FG, fontSize: 8.5, color: C.body, margin: 0, valign: "middle",
  });

  const lw = 6.9, rx = M + lw + 0.42, rw = R - rx;
  sec(s, M, 1.94, lw, "主要指標の比較");
  const r3 = (a, b, c, o = {}) => [
    td(a, { size: 8.8 }),
    td(b, { align: "right", size: 8.8, fill: o.f }),
    td(c, { align: "right", size: o.s || 8.8, bold: !!o.b, color: o.b ? C.gr : C.body, fill: o.f }),
  ];
  s.addTable(
    [
      [th("観点"), th("業務委託型", "right"), th("FCプラン", "right")],
      r3("運営主体", "セルフカフェ", "未来屋書店様", { b: true }),
      r3("貴社のご投資", "設置工事費 30〜50万円", "9,000,000円", { b: true }),
      r3("月次収入（2年目・巡航）", "286,250円（お受取額）", "540,000円（償却前営業利益）", { b: true }),
      r3("3年間の累計（投資控除後）", yen(SUM_IT - IT0) + "円", yen(SUM_FC - INV) + "円"),
      r3("4年目以降の年間差", "－", "＋3,211,320円", { b: true, f: C.bandD }),
      r3("投資回収の目安", "－", "約19.6か月"),
      r3("売上変動のリスク", "小（月額固定5万円を含む）", "貴社がご負担"),
      r3("食品衛生責任者", "不要", "1名のご選任が必要"),
    ],
    { x: M, y: 2.30, w: lw, colW: [2.3, 2.3, 2.3], rowH: [0.22, 0.26, 0.26, 0.3, 0.28, 0.3, 0.26, 0.26, 0.26], border: TB, autoPage: false }
  );

  sec(s, rx, 1.94, rw, "月次収入の比較", "2年目・巡航時");
  bars(s, {
    x: rx + 0.55, y: 2.52, w: rw - 0.9, h: 1.48, max: 60, unit: "単位：万円／月", gap: 0.5,
    data: [
      { v: 28.6, t: "28.6", label: "業務委託型", sub: "お受取額", color: C.barL, fs: 10 },
      { v: 54.0, t: "54.0", label: "FCプラン", sub: "償却前営業利益", color: C.gr, fs: 10 },
    ],
  });
  s.addShape("rect", { x: rx, y: 4.42, w: rw, h: 0.56, fill: { color: C.band } });
  s.addText("月次では 253,750円、FCプランが上回ります。", {
    x: rx + 0.18, y: 4.48, w: rw - 0.36, h: 0.44,
    fontFace: FG, fontSize: 9, bold: true, color: C.ink, margin: 0, valign: "middle",
  });

  sec(s, M, 5.22, CW, "3か年の累計");
  const tc = (v, o = {}) => td(v, Object.assign({ align: "right", size: 9 }, o));
  s.addTable(
    [
      [th("年度"), th("想定販売杯数", "right"), th("業務委託型　お受取額", "right"), th("FCプラン　償却前営業利益", "right")],
      ...YR.map((r) => [
        td(r.y + "　" + r.term + "（" + r.note + "）", { size: 9, fill: r.hot ? C.band : C.white }),
        tc(r.cups + "杯／日", { fill: r.hot ? C.band : C.white }),
        tc(yen(r.feeY), { fill: r.hot ? C.band : C.white }),
        tc(yen(r.pfY), { fill: r.hot ? C.band : C.white }),
      ]),
      [td("小　計", { bold: true, color: C.ink, size: 9 }), tc("－"), tc(yen(SUM_IT), { bold: true, color: C.ink }), tc(yen(SUM_FC), { bold: true, color: C.ink })],
      [td("初期費用", { size: 9 }), tc("－"), tc("▲ " + yen(IT0)), tc("▲ " + yen(INV))],
      [td("3年間の累計", { fill: C.bandD, bold: true, color: C.gr, size: 9.5 }), tc("－", { fill: C.bandD }), tc(yen(SUM_IT - IT0), { fill: C.bandD, bold: true, color: C.gr, size: 10.5 }), tc(yen(SUM_FC - INV), { fill: C.bandD, bold: true, color: C.gr, size: 10.5 })],
    ],
    { x: M, y: 5.58, w: CW, colW: [4.3, 1.7, 2.3, CW - 8.3], rowH: [0.24, 0.26, 0.26, 0.26, 0.26, 0.24, 0.3], border: TB, autoPage: false }
  );

  foot(s, [
    "※　FCプランは償却前営業利益（減価償却前）でございます。人件費は未計上。家賃はカフェ区画が貴社の賃借面積に含まれる前提で0円としております。",
    "※　金額はすべて税抜でございます。業務委託型の設置工事費は上限の50万円で計上しております。",
  ]);
}

/* ================================================= P2 */
{
  const s = pres.addSlide();
  head(s, "それぞれの計算内訳", "2年目・巡航時（1日75杯）の月次でご説明しております");

  const lw = 5.0, rx = M + lw + 0.5, rw = R - rx;
  sec(s, M, 1.14, lw, "業務委託型");
  s.addTable(
    [
      [th("項目"), th("金額", "right")],
      [td("セルフカフェ売上（1杯420円×75杯×30日）", { size: 8.5 }), td("945,000", { align: "right", size: 8.5 })],
      [td("　うち 売上の25％", { size: 8.5 }), td("236,250", { align: "right", size: 8.5 })],
      [td("　月額固定（貴社のみの特別条件）", { size: 8.5 }), td("50,000", { align: "right", size: 8.5 })],
      [td("貴社お受取額（月）", { fill: C.bandD, bold: true, color: C.gr, size: 9 }), td("286,250", { align: "right", fill: C.bandD, bold: true, color: C.gr, size: 11 })],
    ],
    { x: M, y: 1.50, w: lw, colW: [3.3, lw - 3.3], rowH: [0.22, 0.3, 0.24, 0.26, 0.32], border: TB, autoPage: false }
  );
  s.addText("運営は弊社が行い、売上金は弊社が管理いたします。貴社のご負担は設置工事費（30〜50万円）のみでございます。", {
    x: M, y: 3.00, w: lw, h: 0.44,
    fontFace: FG, fontSize: 8.3, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
  });

  sec(s, rx, 1.14, rw, "FCプラン");
  s.addTable(
    [
      [th("項目"), th("金額", "right")],
      [td("売上高（1杯420円×75杯×30日）", { size: 8.5, bold: true, color: C.ink }), td("945,000", { align: "right", size: 8.5, bold: true, color: C.ink })],
      [td("ドリンク原料（売上の20％）", { size: 8.5 }), td("▲ 189,000", { align: "right", size: 8.5 })],
      [td("清掃業務委託費（貴社スタッフ対応のため）", { size: 8.5, color: C.muted }), td("（30,000）", { align: "right", size: 8.5, color: C.muted })],
      [td("水道光熱費（書店営業時間ベース）", { size: 8.5 }), td("▲ 60,000", { align: "right", size: 8.5 })],
      [td("機械使用料（マシン2台・決済端末2台）", { size: 8.5 }), td("▲ 86,000", { align: "right", size: 8.5 })],
      [td("ロイヤリティ（一律・売上連動なし）", { size: 8.5 }), td("▲ 50,000", { align: "right", size: 8.5 })],
      [td("雑費", { size: 8.5 }), td("▲ 20,000", { align: "right", size: 8.5 })],
      [td("家賃（貴社の賃借面積に含まれる前提）", { size: 8.5 }), td("0", { align: "right", size: 8.5 })],
      [td("償却前営業利益（月）", { fill: C.bandD, bold: true, color: C.gr, size: 9 }), td("540,000", { align: "right", fill: C.bandD, bold: true, color: C.gr, size: 11 })],
      [td("平均荒利率", { fill: C.band, bold: true, color: C.ink, size: 9 }), td("80.0％", { align: "right", fill: C.band, bold: true, color: C.ink, size: 9 })],
    ],
    { x: rx, y: 1.50, w: rw, colW: [3.5, rw - 3.5], rowH: [0.22, 0.26, 0.24, 0.26, 0.24, 0.26, 0.24, 0.22, 0.26, 0.3, 0.26], border: TB, autoPage: false }
  );
  s.addText("セキュリティ費は、貴社書店内への併設のため計上しておりません。清掃を外部委託とされる場合は月30,000円が加わります。損益分岐点は1日約22杯でございます。", {
    x: rx, y: 4.46, w: rw, h: 0.44,
    fontFace: FG, fontSize: 8.3, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
  });

  sec(s, M, 5.10, CW, "ご投資額の内訳（FCプラン・概算）");
  const IV = [["加盟金", "1,000,000"], ["保証金", "500,000"], ["内装造作工事", "4,200,000"], ["家具什器", "1,300,000"], ["サイン工事", "500,000"], ["開業準備金", "500,000"], ["設計・申請・諸経費", "1,000,000"]];
  const iw = (CW - 0.16 * 3) / 4;
  IV.forEach((r, i) => {
    const x = M + (i % 4) * (iw + 0.16), y = 5.46 + Math.floor(i / 4) * 0.34;
    s.addShape("rect", { x, y, w: iw, h: 0.3, fill: { color: i % 2 ? C.soft : C.white }, line: { color: C.ruleT, width: 0.5 } });
    s.addText(r[0], { x: x + 0.12, y, w: iw * 0.56, h: 0.3, fontFace: FG, fontSize: 8.3, color: C.ink, margin: 0, valign: "middle" });
    s.addText(r[1], { x: x + iw * 0.56, y, w: iw * 0.44 - 0.12, h: 0.3, fontFace: FG, fontSize: 8.3, color: C.body, align: "right", margin: 0, valign: "middle" });
  });
  s.addShape("rect", { x: M + 3 * (iw + 0.16), y: 5.80, w: iw, h: 0.3, fill: { color: C.bandD } });
  s.addText("合　計", { x: M + 3 * (iw + 0.16) + 0.12, y: 5.80, w: iw * 0.5, h: 0.3, fontFace: FG, fontSize: 8.5, bold: true, color: C.gr, margin: 0, valign: "middle" });
  s.addText("9,000,000", { x: M + 3 * (iw + 0.16) + iw * 0.5, y: 5.80, w: iw * 0.5 - 0.12, h: 0.3, fontFace: FG, fontSize: 9.5, bold: true, color: C.gr, align: "right", margin: 0, valign: "middle" });
  s.addText("マシンおよび決済端末は月額の機械使用料でのご提供でございます。初期のご購入は不要です。秋田店は既存造作を残置いただく前提のため、新規出店の場合より工事範囲が限定されます。金額は概算で、仕様確定後に正式なお見積りをご提出いたします。", {
    x: M, y: 6.30, w: CW, h: 0.44,
    fontFace: FG, fontSize: 8.3, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
  });

  foot(s, [
    "※　金額はすべて税抜でございます。想定値であり、売上および利益を保証するものではございません。",
  ]);
}

/* ================================================= P3 */
{
  const s = pres.addSlide();
  head(s, "ご判断にあたって", "3年で区切られるか、その先も続けられるかで結論が変わります");

  s.addShape("rect", { x: M, y: 1.14, w: CW, h: 0.92, fill: { color: C.grPale }, line: { color: C.gr, width: 1 } });
  s.addText("3年間ではほぼ互角、3年1か月で逆転いたします。", {
    x: M + 0.2, y: 1.20, w: CW - 0.4, h: 0.24,
    fontFace: FM, fontSize: 12, bold: true, color: C.gr, margin: 0, valign: "middle",
  });
  s.addText("3年間の累計は、業務委託型 9,389,200円に対しFCプラン 9,109,440円。その差は279,760円でございます。4年目以降は年間 3,211,320円 の差がつきますので、5年間では 6,142,880円 FCプランが上回ります。3年でいったん区切られるご予定であれば業務委託型、その先も続けられるご見通しであればFCプランが有利でございます。", {
    x: M + 0.2, y: 1.44, w: CW - 0.4, h: 0.56,
    fontFace: FG, fontSize: 8.8, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
  });

  const lw = 5.1, rx = M + lw + 0.5, rw = R - rx;
  sec(s, M, 2.26, lw, "FCプランの利点");
  [
    ["売上・利益がすべて貴社のものになります", "巡航時 月540,000円。業務委託型の約1.9倍でございます。"],
    ["ロイヤリティが売上連動ではございません", "一律 月50,000円。売上が伸びるほど利益率が向上いたします。"],
    ["マシンは月額使用料でご提供いたします", "初期のご購入は不要。更新・故障対応も弊社にて行います。"],
    ["自社業態としての知見が蓄積されます", "仙台泉をはじめ、今後のご出店のご判断材料となります。"],
  ].forEach((t, i) => {
    const y = 2.64 + i * 0.52;
    s.addText("（" + (i + 1) + "）" + t[0], { x: M, y, w: lw, h: 0.2, fontFace: FG, fontSize: 8.8, bold: true, color: C.ink, margin: 0, valign: "middle" });
    s.addText(t[1], { x: M + 0.36, y: y + 0.2, w: lw - 0.36, h: 0.28, fontFace: FG, fontSize: 8, color: C.body, margin: 0, valign: "middle" });
  });

  sec(s, rx, 2.26, rw, "FCプランの留意点");
  [
    ["初期投資が必要でございます", "900万円。回収の目安は約19.6か月（立ち上がりを含む）でございます。"],
    ["3年では投資を回収しきれても差は出ません", "累計で業務委託型を下回ります。4年目以降で優位に転じます。"],
    ["売上変動のリスクを貴社が負われます", "業務委託型の月額固定50,000円のような下支えがございません。"],
    ["食品衛生責任者の選任が必要でございます", "1名。保健所の申請手続きは弊社にて代行いたします。"],
  ].forEach((t, i) => {
    const y = 2.64 + i * 0.52;
    s.addText("（" + (i + 1) + "）" + t[0], { x: rx, y, w: rw, h: 0.2, fontFace: FG, fontSize: 8.8, bold: true, color: C.ink, margin: 0, valign: "middle" });
    s.addText(t[1], { x: rx + 0.36, y: y + 0.2, w: rw - 0.36, h: 0.28, fontFace: FG, fontSize: 8, color: C.body, margin: 0, valign: "middle" });
  });

  sec(s, M, 4.80, CW, "契約形態を変更される場合の手続きと期限", "開業 2027年3月上旬");
  s.addTable(
    [
      [th("必要となる手続き"), th("ご担当"), th("期限の目安", "right")],
      [td("FC加盟契約の締結（業務委託契約からの変更）", { size: 8.5 }), td("両社", { size: 8.5 }), td("12月中旬", { align: "right", size: 8.5 })],
      [td("ご投資 900万円の社内ご決裁", { size: 8.5 }), td("貴社", { size: 8.5 }), td("12月上旬", { align: "right", size: 8.5, bold: true, color: C.gr })],
      [td("食品衛生責任者のご選任（1名）", { size: 8.5 }), td("貴社", { size: 8.5 }), td("1月中", { align: "right", size: 8.5 })],
      [td("保健所への営業許可申請（名義が貴社となります）", { size: 8.5 }), td("弊社代行", { size: 8.5 }), td("2月中", { align: "right", size: 8.5 })],
      [td("什器・機器の発注（リードタイム4〜6週間）", { size: 8.5, fill: C.band }), td("発注者が貴社となります", { size: 8.5, fill: C.band }), td("12月上旬", { align: "right", size: 8.5, bold: true, color: C.gr, fill: C.band })],
    ],
    { x: M, y: 5.16, w: CW, colW: [5.2, 2.6, CW - 7.8], rowH: [0.22, 0.26, 0.26, 0.26, 0.26, 0.28], border: TB, autoPage: false }
  );
  s.addText("什器・機器の発注は12月上旬が期限でございます。発注の主体が変わりますため、ご決裁は12月上旬までにお願いしたく存じます。これを過ぎますと、3月上旬の開業が難しくなります。", {
    x: M, y: 6.78, w: CW, h: 0.44,
    fontFace: FG, fontSize: 8.8, bold: true, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
  });

  foot(s, [
    "※　いずれも開業前の机上試算でございます。金額はすべて税抜。想定値であり、売上および利益を保証するものではございません。",
    "※　ご投資額は概算でございます。仕様確定後にあらためて正式なお見積りをご提出いたします。",
  ]);
}

pres.writeFile({ fileName: path.join(__dirname, "akita-compare.pptx") }).then((f) => console.log("written:", f));
