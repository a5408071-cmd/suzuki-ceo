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
const UP = 420, DAYS = 30;
const MACH = 76000, TERM = 10000, CLOUD = 5000, SUP = 15000, ROY = 50000;
const UTL = { 30: 60000, 60: 70000 }, ETC = { 30: 20000, 60: 30000 };
const CLN = 30000, SECF = 20000;
const INV_LO = 7000000, INV_HI = 10000000;
const CASE = [
  { t: "30坪", pos: "店頭配置", cups: 86, tb: 30, hot: true },
  { t: "30坪", pos: "店奥配置", cups: 60, tb: 30 },
  { t: "60坪", pos: "店頭配置", cups: 98, tb: 60, ref: true },
  { t: "60坪", pos: "店奥配置", cups: 69, tb: 60, ref: true },
].map((c) => {
  const sales = c.cups * UP * DAYS;
  const mat = Math.floor(sales * 0.2);
  const fx = UTL[c.tb] + MACH + TERM + CLOUD + SUP + ROY + ETC[c.tb];
  const pf = sales - mat - fx;
  return Object.assign({}, c, { sales, mat, fx, cost: mat + fx, pf, pfY: pf * 12, lo: INV_LO / pf, hi: INV_HI / pf });
});
const B = CASE[0];
const yen = (v) => v.toLocaleString();
const neg = (v) => "▲ " + yen(v);

/* ================================================= P1 概要 */
{
  const s = pres.addSlide();
  head(s, "仙台泉（新規出店）　セルフカフェ FCプランのご提案", "30坪・1日86杯（店頭配置）／2028年度内 開業予定　貴社が運営主体となられる場合の試算でございます");

  s.addShape("rect", { x: M, y: 1.14, w: CW, h: 0.62, fill: { color: C.white }, line: { color: C.ruleS, width: 1 } });
  s.addText("基本工事（床・壁・天井・照明）は、貴社のご負担でございます。", {
    x: M + 0.18, y: 1.18, w: CW - 0.36, h: 0.24,
    fontFace: FG, fontSize: 9.5, bold: true, color: C.gr, margin: 0, valign: "middle",
  });
  s.addText("他の書店区画と同等の仕上がりが必要となるためでございます。本資料で申し上げる「ご投資」は、その基本工事が完了した状態からの内装造作・什器・サイン等（業務委託型であれば弊社が負担しておりました部分）に、加盟金・保証金等を加えた総額でございます。", {
    x: M + 0.18, y: 1.42, w: CW - 0.36, h: 0.3,
    fontFace: FG, fontSize: 8.5, color: C.body, margin: 0, valign: "middle",
  });

  sec(s, M, 1.92, CW, "FCプランの枠組み");
  const pw = (CW - 0.26 * 2) / 3;
  const FR = [
    { h: "運営と売上", rows: [["運　営", "未来屋書店様"], ["売上の帰属", "すべて未来屋書店様"], ["日常業務", "貴社スタッフ（清掃・補充・発注）"], ["貴社の収入", "売上高から経費を差し引いた残り"]] },
    { h: "弊社にお支払いいただくもの", rows: [["ロイヤリティ", "月 50,000円（一律・売上連動なし）"], ["マシン使用料", "月 76,000円（38,000円×2台）"], ["決済端末・クラウド", "月 15,000円"], ["営業サポート", "月 15,000円（ミニマムサポート）"]] },
    { h: "ご負担いただく経費", rows: [["ドリンク原料", "売上の約20％"], ["水道光熱費", "書店営業時間ベース"], ["雑費", "通信費・消耗品費等"], ["清掃・警備", "有人のため不要でございます"]] },
  ];
  FR.forEach((g, i) => {
    const x = M + i * (pw + 0.26);
    s.addShape("rect", { x, y: 2.28, w: pw, h: 1.78, fill: { color: C.white }, line: { color: C.ruleT, width: 0.75 } });
    s.addShape("rect", { x, y: 2.28, w: pw, h: 0.36, fill: { color: C.grPale } });
    s.addText(g.h, { x: x + 0.16, y: 2.28, w: pw - 0.32, h: 0.36, fontFace: FG, fontSize: 9, bold: true, color: C.gr, margin: 0, valign: "middle" });
    g.rows.forEach((r, j) => {
      const y = 2.74 + j * 0.32;
      s.addText(r[0], { x: x + 0.16, y, w: 1.35, h: 0.3, fontFace: FG, fontSize: 8.3, bold: true, color: C.ink, margin: 0, valign: "middle" });
      s.addText(r[1], { x: x + 1.54, y, w: pw - 1.7, h: 0.3, fontFace: FG, fontSize: 8.3, color: C.body, margin: 0, valign: "middle" });
    });
  });

  sec(s, M, 4.26, CW, "主要指標（30坪・店頭配置）");
  const kw = (CW - 0.24 * 3) / 4;
  [
    { l: "セルフカフェ売上（月）", v: yen(B.sales), u: "円" },
    { l: "償却前営業利益（月）", v: yen(B.pf), u: "円", hot: true },
    { l: "平均荒利率", v: "80.0", u: "％" },
    { l: "損益分岐点", v: "約24", u: "杯／日" },
  ].forEach((k, i) => {
    const x = M + i * (kw + 0.24);
    s.addShape("rect", { x, y: 4.62, w: kw, h: 0.94, fill: { color: k.hot ? C.grPale : C.soft }, line: { color: k.hot ? C.gr : C.ruleT, width: k.hot ? 1 : 0.75 } });
    s.addText(k.l, { x: x + 0.16, y: 4.68, w: kw - 0.32, h: 0.22, fontFace: FG, fontSize: 8.5, bold: true, color: k.hot ? C.gr : C.ink, margin: 0, valign: "middle" });
    s.addText(
      [{ text: k.v, options: { fontFace: FG, fontSize: k.hot ? 22 : 19, bold: true, color: k.hot ? C.gr : C.ink } },
       { text: " " + k.u, options: { fontFace: FG, fontSize: 9.5, bold: true, color: C.body } }],
      { x: x + 0.16, y: 4.92, w: kw - 0.32, h: 0.56, margin: 0, valign: "middle" }
    );
  });
  s.addText("損益分岐点は1日約24杯でございます。想定の86杯に対して3倍以上の余裕があり、下振れに強い構造でございます。", {
    x: M, y: 5.72, w: CW, h: 0.24,
    fontFace: FG, fontSize: 9, bold: true, color: C.gr, margin: 0, valign: "middle",
  });

  sec(s, M, 6.06, CW, "前提条件");
  [
    ["販売単価・営業日数", "1杯 420円（税抜）／月30日で算出しております。"],
    ["販売杯数", "既にご提出済みの売上試算と同一でございます（館の来客数は秋田店の1.3倍、販売杯数は1.15倍）。"],
    ["家賃", "カフェ区画が貴社の賃借面積に含まれる前提で0円としております。人件費は未計上でございます。"],
  ].forEach((t, i) => {
    const y = 6.42 + i * 0.24;
    s.addText("（" + (i + 1) + "）", { x: M, y, w: 0.5, h: 0.22, fontFace: FG, fontSize: 8.5, color: C.ink, margin: 0, valign: "middle" });
    s.addText(t[0], { x: M + 0.5, y, w: 2.3, h: 0.22, fontFace: FG, fontSize: 8.5, bold: true, color: C.ink, margin: 0, valign: "middle" });
    s.addText(t[1], { x: M + 2.85, y, w: CW - 2.85, h: 0.22, fontFace: FG, fontSize: 8.5, color: C.body, margin: 0, valign: "middle" });
  });

  foot(s, [
    "※　本資料は開業前の机上試算でございます。現地での営業実績はなく、館の来客数は計画値でございます。想定値であり、売上および利益を保証するものではございません。金額はすべて税抜。",
  ]);
}

/* ================================================= P2 月次損益 */
{
  const s = pres.addSlide();
  head(s, "月次損益と、坪数・配置別の試算", "単位：円（税抜）　償却前営業利益（減価償却前）でございます");

  const lw = 4.6, rx = M + lw + 0.5, rw = R - rx;
  const ng = (v) => "▲" + yen(v);
  sec(s, M, 1.14, lw, "月次損益の内訳", "30坪・店頭配置");
  const rr = (l, v, o = {}) => [
    td(l, { size: 8.8, bold: !!o.b, color: o.g ? C.muted : (o.b ? C.ink : C.body), fill: o.f }),
    td(v, { align: "right", size: o.big ? 11 : 8.8, bold: !!o.b, color: o.g ? C.muted : (o.b ? (o.gr ? C.gr : C.ink) : C.body), fill: o.f }),
  ];
  s.addTable(
    [
      [th("項目"), th("金額", "right")],
      rr("売上高（1杯420円×86杯×30日）", yen(B.sales), { b: true }),
      rr("ドリンク原料（約20％）", neg(B.mat)),
      rr("水道光熱費（書店営業時間ベース）", neg(UTL[30])),
      rr("マシン使用料（38,000円×2台）", neg(MACH)),
      rr("決済端末利用料（5,000円×2台）", neg(TERM)),
      rr("クラウド利用料", neg(CLOUD)),
      rr("営業サポート（ミニマムサポート）", neg(SUP)),
      rr("ロイヤリティ（一律・売上連動なし）", neg(ROY)),
      rr("雑費（通信費・消耗品費等）", neg(ETC[30])),
      rr("清掃業務委託費　※有人のため不要", "（" + yen(CLN) + "）", { g: true }),
      rr("セキュリティ費　※有人のため不要", "（" + yen(SECF) + "）", { g: true }),
      rr("家賃（貴社の賃借面積に含まれる前提）", "0"),
      rr("ランニングコスト計", neg(B.cost), { b: true }),
      rr("償却前営業利益（月）", yen(B.pf), { b: true, gr: true, f: C.bandD, big: true }),
      rr("平均荒利率", "80.0％", { b: true, f: C.band }),
    ],
    { x: M, y: 1.50, w: lw, colW: [3.5, lw - 3.5], rowH: [0.24, 0.3, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.28, 0.34, 0.28], border: TB, autoPage: false }
  );

  sec(s, rx, 1.14, rw, "坪数・配置別の試算");
  const cc = (v, o = {}) => td(v, Object.assign({ align: "right", size: 8.3 }, o));
  s.addTable(
    [
      [th("項目"), ...CASE.map((c) => th(c.t + "\n" + c.pos, "right"))],
      [td("販売杯数（1日）", { size: 8.3 }), ...CASE.map((c) => cc(c.cups + "杯", { fill: c.hot ? C.grPale : C.white }))],
      [td("売上高（月）", { size: 8.3 }), ...CASE.map((c) => cc(yen(c.sales), { fill: c.hot ? C.grPale : C.white }))],
      [td("ランニングコスト計", { size: 8.3 }), ...CASE.map((c) => cc(ng(c.cost), { fill: c.hot ? C.grPale : C.white }))],
      [td("償却前営業利益（月）", { size: 8.3, bold: true, color: C.ink, fill: C.band }), ...CASE.map((c) => cc(yen(c.pf), { bold: true, color: C.gr, size: 10, fill: c.hot ? C.bandD : C.band }))],
      [td("償却前営業利益（年）", { size: 8.3, bold: true, color: C.ink, fill: C.band }), ...CASE.map((c) => cc(yen(c.pfY), { bold: true, color: C.ink, fill: c.hot ? C.bandD : C.band }))],
      [td("損益分岐点（1日）", { size: 8.3 }), ...CASE.map((c) => cc(c.tb === 30 ? "約24杯" : "約26杯", { fill: c.hot ? C.grPale : C.white }))],
    ],
    { x: rx, y: 1.50, w: rw, colW: [1.35, (rw - 1.35) / 4, (rw - 1.35) / 4, (rw - 1.35) / 4, (rw - 1.35) / 4], rowH: [0.34, 0.26, 0.26, 0.26, 0.3, 0.28, 0.26], border: TB, autoPage: false }
  );

  sec(s, rx, 3.70, rw, "償却前営業利益（月）");
  bars(s, {
    x: rx + 0.5, y: 4.30, w: rw - 0.8, h: 1.46, max: 80, unit: "単位：万円／月", gap: 0.22,
    data: CASE.map((c) => ({
      v: c.pf / 10000, t: (c.pf / 10000).toFixed(1), label: c.t, sub: c.pos.replace("配置", ""),
      color: c.hot ? C.gr : (c.ref ? "D9D9D9" : C.barL), fs: 9,
    })),
  });
  s.addShape("rect", { x: rx, y: 6.26, w: rw, h: 0.68, fill: { color: C.band } });
  s.addText("配置により大きく変わります。店頭配置であれば月630,880円、店奥配置では368,800円でございます（30坪の場合）。", {
    x: rx + 0.18, y: 6.26, w: rw - 0.36, h: 0.68,
    fontFace: FG, fontSize: 8.5, color: C.ink, margin: 0, valign: "middle", lineSpacingMultiple: 1.25,
  });

  foot(s, [
    "※　清掃業務委託費・セキュリティ費は、貴社書店内への併設（有人）のため不要でございます。参考として括弧内に標準額を記載しております。",
    "※　60坪はご参考でございます。人件費は未計上。想定値であり、売上および利益を保証するものではございません。",
  ]);
}

/* ================================================= P3 投資と回収 */
{
  const s = pres.addSlide();
  head(s, "ご投資と回収、ご判断の材料", "基本工事が完了した状態からのご投資でございます");

  sec(s, M, 1.14, CW, "ご投資の目安");
  s.addShape("rect", { x: M, y: 1.50, w: CW, h: 1.22, fill: { color: C.grPale }, line: { color: C.gr, width: 1.25 } });
  s.addText("基本工事を除く総額", {
    x: M + 0.24, y: 1.58, w: 4.2, h: 0.26,
    fontFace: FG, fontSize: 9.5, bold: true, color: C.gr, margin: 0, valign: "middle",
  });
  s.addText(
    [{ text: "700", options: { fontFace: FG, fontSize: 26, bold: true, color: C.gr } },
     { text: " 万円 　〜 　", options: { fontFace: FG, fontSize: 12, bold: true, color: C.body } },
     { text: "1,000", options: { fontFace: FG, fontSize: 26, bold: true, color: C.gr } },
     { text: " 万円 程度", options: { fontFace: FG, fontSize: 12, bold: true, color: C.body } }],
    { x: M + 0.24, y: 1.86, w: 4.6, h: 0.6, margin: 0, valign: "middle" }
  );
  s.addText(
    [{ text: "含まれるもの　", options: { fontFace: FG, fontSize: 8.3, bold: true, color: C.gr } },
     { text: "内装造作工事／家具什器／サイン工事／加盟金／保証金／開業準備金／設計・申請・諸経費", options: { fontFace: FG, fontSize: 8.3, color: C.ink, breakLine: true } },
     { text: "含まれないもの　", options: { fontFace: FG, fontSize: 8.3, bold: true, color: C.gr } },
     { text: "基本工事（床・壁・天井・照明）", options: { fontFace: FG, fontSize: 8.3, color: C.ink, breakLine: true } },
     { text: "現地を確認しておらず、また新規物件のため、現時点では概算でございます。仕様により変動いたします。マシン・決済端末は月額利用料でのご提供のため、初期購入は不要でございます。", options: { fontFace: FG, fontSize: 8.3, color: C.body } }],
    { x: M + 5.0, y: 1.58, w: CW - 5.24, h: 1.08, margin: 0, valign: "top", lineSpacingMultiple: 1.3 }
  );

  sec(s, M, 2.92, CW, "投資回収の目安");
  const cc = (v, o = {}) => td(v, Object.assign({ align: "right", size: 8.8 }, o));
  s.addTable(
    [
      [th("区分"), th("償却前営業利益（月）", "right"), th("700万円の場合", "right"), th("1,000万円の場合", "right")],
      ...CASE.map((c) => [
        td(c.t + "　" + c.pos + (c.ref ? "（参考）" : ""), { size: 8.8, bold: !!c.hot, color: c.hot ? C.gr : C.body, fill: c.hot ? C.bandD : C.white }),
        cc(yen(c.pf), { bold: !!c.hot, color: c.hot ? C.gr : C.body, fill: c.hot ? C.bandD : C.white }),
        cc("約 " + c.lo.toFixed(1) + " か月", { bold: true, color: c.hot ? C.gr : C.ink, fill: c.hot ? C.bandD : C.white }),
        cc("約 " + c.hi.toFixed(1) + " か月", { bold: true, color: c.hot ? C.gr : C.ink, fill: c.hot ? C.bandD : C.white }),
      ]),
    ],
    { x: M, y: 3.28, w: CW, colW: [3.4, 2.5, 2.4, CW - 8.3], rowH: [0.24, 0.3, 0.28, 0.28, 0.28], border: TB, autoPage: false }
  );
  s.addText("30坪・店頭配置であれば、ご投資700万円で約11か月、1,000万円でも約16か月で回収の見通しでございます。いずれも開業から2年以内でございます。", {
    x: M, y: 4.76, w: CW, h: 0.24,
    fontFace: FG, fontSize: 9, bold: true, color: C.gr, margin: 0, valign: "middle",
  });

  const lw = 5.3, rx = M + lw + 0.5, rw = R - rx;
  sec(s, M, 5.12, lw, "FCプランの利点");
  [
    ["売上・利益がすべて貴社のものになります", "30坪・店頭配置で月630,880円、年間7,570,560円でございます。"],
    ["ロイヤリティが売上連動ではございません", "一律 月50,000円。売上が伸びるほど利益率が向上いたします。"],
    ["損益分岐点が低く、下振れに強い構造です", "1日約24杯。想定86杯に対して3倍以上の余裕がございます。"],
    ["基本工事のご負担も利益で回収いただけます", "売上がそのまま帰属いたしますので、回収を月次の利益で進められます。"],
  ].forEach((t, i) => {
    const y = 5.50 + i * 0.5;
    s.addText("（" + (i + 1) + "）" + t[0], { x: M, y, w: lw, h: 0.2, fontFace: FG, fontSize: 8.8, bold: true, color: C.ink, margin: 0, valign: "middle" });
    s.addText(t[1], { x: M + 0.36, y: y + 0.2, w: lw - 0.36, h: 0.28, fontFace: FG, fontSize: 8.2, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.2 });
  });

  sec(s, rx, 5.12, rw, "ご留意いただきたい点");
  [
    ["ご投資は概算でございます", "現地未確認かつ新規物件のため、仕様確定後にお見積りをご提出いたします。"],
    ["売上変動のリスクを貴社が負われます", "来客数は館の計画値であり、実績の裏付けはございません。"],
    ["配置により結果が大きく変わります", "店頭配置と店奥配置で、月の営業利益に262,080円の差が生じます。"],
    ["食品衛生責任者の選任が必要でございます", "1名。保健所の申請手続きは弊社にて代行いたします。"],
  ].forEach((t, i) => {
    const y = 5.50 + i * 0.5;
    s.addText("（" + (i + 1) + "）" + t[0], { x: rx, y, w: rw, h: 0.2, fontFace: FG, fontSize: 8.8, bold: true, color: C.ink, margin: 0, valign: "middle" });
    s.addText(t[1], { x: rx + 0.36, y: y + 0.2, w: rw - 0.36, h: 0.28, fontFace: FG, fontSize: 8.2, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.2 });
  });

  foot(s, [
    "※　ご投資額は概算でございます。区画の位置・坪数・仕様が確定しだい、あらためて正式なお見積りをご提出いたします。金額はすべて税抜。",
  ]);
}

pres.writeFile({ fileName: path.join(__dirname, "sendai-fc.pptx") }).then((f) => console.log("written:", f));
