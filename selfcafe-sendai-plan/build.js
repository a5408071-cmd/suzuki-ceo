/**
 * 未来屋書店 仙台泉（新店）セルフカフェ レイアウトイメージ（30坪）
 * A4横 1ページ／現地未確認のイメージ図
 */
const pptxgen = require("pptxgenjs");

const C = {
  bg: "FFFFFF", ink: "1A1A1A", body: "333333", muted: "767676",
  ruleS: "1A1A1A", ruleT: "808080", hair: "C8C8C8",
  band: "F2F2F2", soft: "F7F8F9", bandD: "E9F1EC",
  gr: "106838", grPale: "E9F1EC", barL: "BFBFBF", white: "FFFFFF",
  wall: "1A1A1A", furn: "FFFFFF", chair: "D9D9D9", cnt: "ECECEC",
};
const FM = "Yu Mincho", FG = "Yu Gothic";
const W = 11.69, H = 8.27, M = 0.58, CW = W - M * 2, R = M + CW;

const pres = new pptxgen();
pres.defineLayout({ name: "A4L", width: W, height: H });
pres.layout = "A4L";
pres.author = "セルフカフェ株式会社";
pres.title = "仙台泉 レイアウトイメージ（30坪）";

const s = pres.addSlide();
s.background = { color: C.bg };

/* ---------------- 体裁 ---------------- */
s.addText("仙台泉店（仮称）　セルフカフェ　レイアウトイメージ", {
  x: M, y: 0.36, w: CW - 3.3, h: 0.36,
  fontFace: FM, fontSize: 17, bold: true, color: C.ink, margin: 0, valign: "middle",
});
s.addText("30坪・店頭配置を想定した参考図でございます（現地未確認）", {
  x: M, y: 0.72, w: CW - 3.3, h: 0.2,
  fontFace: FG, fontSize: 8.5, color: C.body, margin: 0, valign: "middle",
});
s.addText("未来屋書店株式会社　御中", {
  x: R - 3.3, y: 0.34, w: 3.3, h: 0.2,
  fontFace: FG, fontSize: 9, color: C.ink, align: "right", margin: 0, valign: "middle",
});
s.addText("2026年11月27日　セルフカフェ株式会社", {
  x: R - 3.3, y: 0.54, w: 3.3, h: 0.18,
  fontFace: FG, fontSize: 8, color: C.muted, align: "right", margin: 0, valign: "middle",
});
s.addShape("rect", { x: M, y: 0.98, w: CW, h: 0.019, fill: { color: C.gr } });
s.addShape("rect", { x: M, y: 1.025, w: CW, h: 0.006, fill: { color: C.ruleS } });

let SEC = 0;
const sec = (x, y, w, t, note) => {
  SEC++;
  s.addText(SEC + "．" + t, { x, y, w, h: 0.24, fontFace: FM, fontSize: 11.5, bold: true, color: C.ink, margin: 0, valign: "middle" });
  if (note) s.addText(note, { x: x + w - 3.2, y, w: 3.2, h: 0.24, fontFace: FG, fontSize: 8, color: C.muted, align: "right", margin: 0, valign: "middle" });
  s.addShape("rect", { x, y: y + 0.25, w, h: 0.008, fill: { color: C.gr } });
};
const td = (t, o = {}) => ({
  text: t,
  options: { fill: { color: o.fill || C.white }, color: o.color || C.body, bold: !!o.bold, fontSize: o.size || 8.5, fontFace: FG, align: o.align || "left", valign: "middle", margin: [0.02, 0.08, 0.02, 0.08] },
});
const TB = { type: "solid", color: C.ruleT, pt: 0.5 };

/* ================= 平面図 ================= */
const PW = 11000, PD = 9000;          // 区画 間口×奥行（mm）
const PX0 = 0.98, PY0 = 1.80, PIW = 6.05;
const SC = PIW / PW;
const X = (mm) => PX0 + mm * SC;
const Y = (mm) => PY0 + mm * SC;
const Z = (mm) => mm * SC;
const PIH = Z(PD);

sec(M, 1.14, PIW + 0.40, "平面イメージ（30坪／59席）", "上＝モール通路・書店入口側");

// 床
s.addShape("rect", { x: X(0), y: Y(0), w: Z(PW), h: Z(PD), fill: { color: "FCFCFC" }, line: { color: C.hair, width: 0.5 } });

// 壁（開口 x2,000〜9,000 を除く）
const WT = 150;
const wall = (x0, y0, x1, y1) => s.addShape("rect", { x: X(x0), y: Y(y0), w: Z(x1 - x0), h: Z(y1 - y0), fill: { color: C.wall } });
wall(0, 0, 2000, WT); wall(9000, 0, PW, WT);           // 上
wall(0, PD - WT, PW, PD);                               // 下（奥）
wall(0, 0, WT, PD); wall(PW - WT, 0, PW, PD);           // 左右

// 開口（店頭側）
s.addShape("rect", { x: X(2000), y: Y(0), w: Z(7000), h: Z(70), fill: { color: C.gr } });
s.addText("開口（書店売場・モール通路側／間口 約7.0m）", {
  x: X(1700), y: PY0 - 0.22, w: Z(5100), h: 0.19,
  fontFace: FG, fontSize: 7.5, bold: true, color: C.gr, align: "left", margin: 0, valign: "middle",
});
s.addShape("rect", { x: X(7000), y: PY0 - 0.235, w: Z(2000), h: 0.175, fill: { color: C.gr } });
s.addText("サイン", { x: X(7000), y: PY0 - 0.235, w: Z(2000), h: 0.175, fontFace: FG, fontSize: 7, bold: true, color: "FFFFFF", align: "center", margin: 0, valign: "middle" });

// 家具描画ヘルパー
const box = (x0, y0, w, h, fill, line) =>
  s.addShape("rect", { x: X(x0), y: Y(y0), w: Z(w), h: Z(h), fill: { color: fill }, line: { color: line || C.ruleT, width: 0.75 } });
const chair = (cx, cy) =>
  s.addShape("ellipse", { x: X(cx - 210), y: Y(cy - 210), w: Z(420), h: Z(420), fill: { color: C.chair }, line: { color: C.hair, width: 0.5 } });
const lbl = (x0, y0, w, h, t, o = {}) =>
  s.addText(t, { x: X(x0), y: Y(y0), w: Z(w), h: Z(h), fontFace: FG, fontSize: o.fs || 7.2, bold: o.b !== false, color: o.c || C.ink, align: "center", margin: 0, valign: "middle" });

/* --- 奥側：ドリンクカウンター（給排水） --- */
box(900, 7950, 4400, 900, C.cnt);
box(1100, 8030, 900, 740, C.gr, C.gr);
box(2150, 8030, 900, 740, C.gr, C.gr);
lbl(1100, 8030, 900, 740, "マシン\n①", { c: "FFFFFF", fs: 7 });
lbl(2150, 8030, 900, 740, "マシン\n②", { c: "FFFFFF", fs: 7 });
lbl(3150, 8030, 2050, 740, "カップ・備品／決済端末", { fs: 7 });

/* --- 返却・ゴミ --- */
box(5700, 8150, 1400, 700, C.cnt);
lbl(5700, 8150, 1400, 700, "返却・ゴミ", { fs: 7 });

/* --- ストック --- */
box(9500, 7950, 1350, 900, C.band);
lbl(9500, 7950, 1350, 900, "ストック\n備品", { fs: 7 });

/* --- 2人掛けテーブル 8卓（16席） --- */
const T2W = 700, T2D = 700;
[1150, 2450].forEach((ty) => {
  [1200, 3250, 5300, 7350].forEach((tx) => {
    box(tx, ty, T2W, T2D, C.furn);
    chair(tx - 330, ty + T2D / 2);
    chair(tx + T2W + 330, ty + T2D / 2);
  });
});

/* --- 4人掛けテーブル 8卓（32席） --- */
const T4W = 1200, T4D = 800;
[4100, 6250].forEach((ty) => {
  [1200, 3300, 5400, 7500].forEach((tx) => {
    box(tx, ty, T4W, T4D, C.furn);
    chair(tx + 300, ty - 330);
    chair(tx + 900, ty - 330);
    chair(tx + 300, ty + T4D + 330);
    chair(tx + 900, ty + T4D + 330);
  });
});

/* --- 壁付カウンター席 11席 --- */
box(10250, 900, 600, 6600, C.cnt);
for (let i = 0; i < 11; i++) chair(9900, 1200 + i * 600);
s.addText("壁付カウンター席 11席", {
  x: X(7900), y: Y(460), w: Z(2950), h: 0.18,
  fontFace: FG, fontSize: 7.2, bold: true, color: C.ink, align: "right", margin: 0, valign: "middle",
});

/* --- 寸法・注記（図の下） --- */
const dy = PY0 + PIH + 0.11;
s.addShape("rect", { x: X(0), y: dy, w: Z(PW), h: 0.008, fill: { color: C.ruleT } });
s.addShape("rect", { x: X(0), y: dy - 0.04, w: 0.008, h: 0.088, fill: { color: C.ruleT } });
s.addShape("rect", { x: X(PW), y: dy - 0.04, w: 0.008, h: 0.088, fill: { color: C.ruleT } });
s.addText("間口 11,000", { x: X(0), y: dy + 0.05, w: Z(PW), h: 0.18, fontFace: FG, fontSize: 8, color: C.muted, align: "center", margin: 0, valign: "middle" });
s.addText("奥行 9,000", { x: X(0) - 0.72, y: PY0 + PIH / 2 - 0.11, w: 1.2, h: 0.22, fontFace: FG, fontSize: 8, color: C.muted, align: "center", margin: 0, valign: "middle", rotate: 270 });
s.addText("奥の壁面1面にドリンクカウンター（給排水・電源）を集約し、工事範囲を最小限とする構成でございます。", {
  x: PX0, y: dy + 0.26, w: PIW, h: 0.18,
  fontFace: FG, fontSize: 7.8, bold: true, color: C.gr, align: "center", margin: 0, valign: "middle",
});

/* ================= 右欄 ================= */
const RX = 7.28, RW = R - RX;

sec(RX, 1.14, RW, "想定諸元");
s.addTable(
  [
    [td("想定面積", { bold: true, fill: C.soft }), td("99.0㎡（約30坪）")],
    [td("間口×奥行", { bold: true, fill: C.soft }), td("11,000mm × 9,000mm")],
    [td("想定席数", { bold: true, fill: C.soft }), td("59席", { bold: true, color: C.gr })],
    [td("　4人掛けテーブル", { fill: C.white }), td("8卓／32席")],
    [td("　2人掛けテーブル", { fill: C.white }), td("8卓／16席")],
    [td("　壁付カウンター席", { fill: C.white }), td("11席")],
    [td("ドリンクマシン", { bold: true, fill: C.soft }), td("2台")],
    [td("1席あたり面積", { bold: true, fill: C.soft }), td("約1.7㎡／席")],
  ],
  { x: RX, y: 1.50, w: RW, colW: [1.52, RW - 1.52], rowH: [0.26, 0.26, 0.28, 0.24, 0.24, 0.24, 0.26, 0.26], border: TB, autoPage: false }
);

sec(RX, 3.70, RW, "凡例");
const leg = [
  [C.gr, "ドリンクマシン（2台）"],
  [C.cnt, "カウンター・備品・返却"],
  [C.furn, "テーブル"],
  [C.chair, "客席（椅子・スツール）"],
  [C.wall, "壁"],
];
leg.forEach((g, i) => {
  const y = 4.06 + i * 0.235;
  s.addShape("rect", { x: RX, y: y + 0.045, w: 0.22, h: 0.145, fill: { color: g[0] }, line: { color: C.ruleT, width: 0.75 } });
  s.addText(g[1], { x: RX + 0.32, y, w: RW - 0.32, h: 0.235, fontFace: FG, fontSize: 8.3, color: C.body, margin: 0, valign: "middle" });
});

sec(RX, 5.32, RW, "レイアウトの考え方");
[
  ["給排水", "ドリンクカウンターを奥の壁面1面に集約し、給排水・電源の引込みを最短にしております。工事費の抑制に直結いたします。"],
  ["動線", "開口から入って正面奥がカウンターとなり、初めてのお客様でも迷わない配置でございます。返却・ゴミはカウンター横にまとめております。"],
  ["席構成", "滞在時間の短い2人掛けを通路側に、長めの4人掛けを奥側に配置しております。壁付カウンター席はお1人様の利用を想定しております。"],
  ["視認性", "通路側を客席とすることで、賑わいが外から見える構成としております。サインは開口上部を想定しております。"],
].forEach((t, i) => {
  const y = 5.68 + i * 0.49;
  s.addText(t[0], { x: RX, y, w: 0.72, h: 0.2, fontFace: FG, fontSize: 8.3, bold: true, color: C.gr, margin: 0, valign: "middle" });
  s.addText(t[1], { x: RX + 0.76, y: y - 0.02, w: RW - 0.76, h: 0.46, fontFace: FG, fontSize: 7.8, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.12 });
});

/* ---------------- 脚注 ---------------- */
s.addShape("rect", { x: M, y: H - 0.78, w: CW, h: 0.006, fill: { color: C.hair } });
[
  "※　本図は現地未確認のイメージ図でございます。区画の形状・寸法・開口位置・柱位置・設備位置はいずれも確定しておらず、実際のレイアウトとは異なります。",
  "※　席数は弊社既存店（30坪／約59席）の実績をもとにした想定値でございます。区画形状の確定後に、あらためて図面をご提出いたします。",
].forEach((t, i) =>
  s.addText(t, { x: M, y: H - 0.70 + i * 0.15, w: CW - 2.6, h: 0.15, fontFace: FG, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle" })
);
s.addText("セルフカフェ株式会社", { x: R - 2.4, y: H - 0.3, w: 2.4, h: 0.16, fontFace: FG, fontSize: 7.5, color: C.muted, align: "right", margin: 0, valign: "middle" });

const out = __dirname + "/sendai-plan.pptx";
pres.writeFile({ fileName: out }).then(() => console.log("written:", out));
