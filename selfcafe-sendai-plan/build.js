/**
 * 未来屋書店 仙台泉（新店）セルフカフェ レイアウトイメージ（30坪・カウンター中心）
 * A4横 1ページ／現地未確認のイメージ図
 */
const pptxgen = require("pptxgenjs");

const C = {
  bg: "FFFFFF", ink: "1A1A1A", body: "333333", muted: "767676",
  ruleS: "1A1A1A", ruleT: "808080", hair: "C8C8C8",
  band: "F2F2F2", soft: "F7F8F9", bandD: "E9F1EC",
  gr: "106838", grPale: "E9F1EC", white: "FFFFFF",
  floor: "FBFBFA", wall: "1A1A1A",
  cnt: "E8E8E6", cntE: "9A9A96",
  tbl: "FFFFFF", tblE: "8C8C8C",
  seat: "D4D4D1", comm: "F4F2ED",
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
s.addText("30坪・店頭配置／カウンター席を中心とした構成の参考図でございます（現地未確認）", {
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
const PW = 11000, PD = 9000;
const PX0 = 0.92, PY0 = 1.78, PIW = 6.40;
const SC = PIW / PW;
const X = (mm) => PX0 + mm * SC;
const Y = (mm) => PY0 + mm * SC;
const Z = (mm) => mm * SC;
const PIH = Z(PD);

sec(M, 1.14, PIW + 0.30, "平面イメージ（30坪／59席・うちカウンター39席）", "上＝モール通路・書店入口側");

// 床
s.addShape("rect", { x: X(0), y: Y(0), w: Z(PW), h: Z(PD), fill: { color: C.floor }, line: { color: C.hair, width: 0.5 } });

// 壁
const WT = 150;
const wall = (x0, y0, x1, y1) => s.addShape("rect", { x: X(x0), y: Y(y0), w: Z(x1 - x0), h: Z(y1 - y0), fill: { color: C.wall } });
wall(0, 0, 2000, WT); wall(9000, 0, PW, WT);
wall(0, PD - WT, PW, PD);
wall(0, 0, WT, PD); wall(PW - WT, 0, PW, PD);

// 開口
s.addShape("rect", { x: X(2000), y: Y(0), w: Z(7000), h: Z(70), fill: { color: C.gr } });
s.addText("開口（書店売場・モール通路側／間口 約7.0m）", {
  x: X(1700), y: PY0 - 0.22, w: Z(5100), h: 0.19,
  fontFace: FG, fontSize: 7.5, bold: true, color: C.gr, margin: 0, valign: "middle",
});
s.addShape("rect", { x: X(7000), y: PY0 - 0.235, w: Z(2000), h: 0.175, fill: { color: C.gr } });
s.addText("サイン", { x: X(7000), y: PY0 - 0.235, w: Z(2000), h: 0.175, fontFace: FG, fontSize: 7, bold: true, color: "FFFFFF", align: "center", margin: 0, valign: "middle" });

/* --- 描画ヘルパー --- */
const box = (x0, y0, w, h, fill, line, lw) =>
  s.addShape("rect", { x: X(x0), y: Y(y0), w: Z(w), h: Z(h), fill: { color: fill }, line: { color: line || C.tblE, width: lw === undefined ? 0.6 : lw } });
const ell = (cx, cy, d, fill, line) =>
  s.addShape("ellipse", { x: X(cx - d / 2), y: Y(cy - d / 2), w: Z(d), h: Z(d), fill: { color: fill }, line: line ? { color: line, width: 0.6 } : { type: "none" } });
const stool = (cx, cy) => ell(cx, cy, 380, C.seat);
const chair = (cx, cy) =>
  s.addShape("roundRect", { x: X(cx - 210), y: Y(cy - 200), w: Z(420), h: Z(400), rectRadius: 0.2, fill: { color: C.seat }, line: { type: "none" } });
const zl = (x0, y0, w, t, o = {}) =>
  s.addText(t, { x: X(x0), y: Y(y0), w: Z(w), h: 0.17, fontFace: FG, fontSize: o.fs || 7, bold: true, color: o.c || C.gr, align: o.al || "center", margin: 0, valign: "middle" });
const vl = (cx, cy, t) =>
  s.addText(t, { x: X(cx) - 0.85, y: Y(cy) - 0.085, w: 1.7, h: 0.17, fontFace: FG, fontSize: 7, bold: true, color: C.gr, align: "center", margin: 0, valign: "middle", rotate: 270 });

/* --- A フロントカウンター席 11席（通路に面する） --- */
box(2200, 800, 6600, 550, C.cnt, C.cntE);
for (let i = 0; i < 11; i++) stool(2500 + i * 600, 1720);
zl(2200, 210, 6600, "フロントカウンター席 11席");

/* --- B 左壁カウンター席 6席 --- */
box(150, 2700, 600, 3600, C.cnt, C.cntE);
for (let i = 0; i < 6; i++) stool(1090, 3000 + i * 600);
vl(450, 4500, "壁付カウンター 6席");

/* --- C 右壁カウンター席 10席 --- */
box(10250, 1900, 600, 6000, C.cnt, C.cntE);
for (let i = 0; i < 10; i++) stool(9910, 2200 + i * 600);
vl(10550, 4900, "壁付カウンター 10席");

/* --- D 中央コミュナルテーブル 12席 --- */
box(4900, 3000, 1200, 3600, C.comm, C.gr, 1);
for (let i = 0; i < 6; i++) { stool(4540, 3300 + i * 600); stool(6460, 3300 + i * 600); }
zl(4300, 2560, 2400, "コミュナルテーブル 12席");

/* --- E/F 4人掛けテーブル 4卓 16席 --- */
const T4W = 1200, T4D = 800;
[[1950, 3150], [1950, 5350], [7500, 3150], [7500, 5350]].forEach(([tx, ty]) => {
  box(tx, ty, T4W, T4D, C.tbl);
  chair(tx + 300, ty - 320); chair(tx + 900, ty - 320);
  chair(tx + 300, ty + T4D + 320); chair(tx + 900, ty + T4D + 320);
});

/* --- G 2人掛け丸テーブル 2卓 4席 --- */
[2700, 8000].forEach((cx) => {
  ell(cx, 7350, 750, C.tbl, C.tblE);
  stool(cx - 655, 7350); stool(cx + 655, 7350);
});

/* --- 奥側：ドリンクカウンター --- */
box(900, 7950, 4400, 900, C.cnt, C.cntE);
box(1100, 8030, 900, 740, C.gr, C.gr);
box(2150, 8030, 900, 740, C.gr, C.gr);
zl(1100, 8180, 900, "マシン①", { c: "FFFFFF" });
zl(2150, 8180, 900, "マシン②", { c: "FFFFFF" });
zl(3150, 8180, 2050, "カップ・備品／決済端末", { c: C.ink });

/* --- 返却・ゴミ／ストック --- */
box(5700, 8150, 1400, 700, C.cnt, C.cntE);
zl(5700, 8330, 1400, "返却・ゴミ", { c: C.ink });
box(9400, 7950, 1450, 900, C.band, C.cntE);
zl(9400, 8230, 1450, "ストック備品", { c: C.ink });

/* --- 寸法 --- */
const dy = PY0 + PIH + 0.10;
s.addShape("rect", { x: X(0), y: dy, w: Z(PW), h: 0.008, fill: { color: C.ruleT } });
s.addShape("rect", { x: X(0), y: dy - 0.04, w: 0.008, h: 0.088, fill: { color: C.ruleT } });
s.addShape("rect", { x: X(PW), y: dy - 0.04, w: 0.008, h: 0.088, fill: { color: C.ruleT } });
s.addText("間口 11,000", { x: X(0), y: dy + 0.04, w: Z(PW), h: 0.18, fontFace: FG, fontSize: 8, color: C.muted, align: "center", margin: 0, valign: "middle" });
s.addText("奥行 9,000", { x: X(0) - 0.70, y: PY0 + PIH / 2 - 0.11, w: 1.2, h: 0.22, fontFace: FG, fontSize: 8, color: C.muted, align: "center", margin: 0, valign: "middle", rotate: 270 });

/* ================= 右欄 ================= */
const RX = 7.52, RW = R - RX;

sec(RX, 1.14, RW, "想定諸元");
s.addTable(
  [
    [td("想定面積", { bold: true, fill: C.soft }), td("99.0㎡（約30坪）")],
    [td("間口×奥行", { bold: true, fill: C.soft }), td("11,000mm × 9,000mm")],
    [td("想定席数", { bold: true, fill: C.soft }), td("59席", { bold: true, color: C.gr })],
    [td("うちカウンター席", { bold: true, fill: C.soft }), td("39席（全体の66％）", { bold: true, color: C.gr })],
    [td("ドリンクマシン", { bold: true, fill: C.soft }), td("2台")],
    [td("1席あたり面積", { bold: true, fill: C.soft }), td("約1.7㎡／席")],
  ],
  { x: RX, y: 1.50, w: RW, colW: [1.42, RW - 1.42], rowH: [0.26, 0.26, 0.27, 0.27, 0.26, 0.26], border: TB, autoPage: false }
);

sec(RX, 3.22, RW, "席構成");
s.addTable(
  [
    [td("フロントカウンター席", { fill: C.white }), td("11席", { align: "right" })],
    [td("壁付カウンター席（左右）", { fill: C.white }), td("16席", { align: "right" })],
    [td("コミュナルテーブル", { fill: C.white }), td("12席", { align: "right" })],
    [td("4人掛けテーブル（4卓）", { fill: C.white }), td("16席", { align: "right" })],
    [td("2人掛け丸テーブル（2卓）", { fill: C.white }), td("4席", { align: "right" })],
    [td("合計", { bold: true, color: C.ink, fill: C.bandD }), td("59席", { bold: true, color: C.gr, align: "right", fill: C.bandD })],
  ],
  { x: RX, y: 3.58, w: RW, colW: [RW - 0.95, 0.95], rowH: [0.25, 0.25, 0.25, 0.25, 0.25, 0.29], border: TB, autoPage: false }
);

sec(RX, 5.32, RW, "レイアウトの考え方");
[
  ["カウンター\n中心", "全席の約3分の2をカウンター席としております。お1人様・短時間のご利用に対応し、席の回転を高める構成でございます。"],
  ["視認性", "通路に面してフロントカウンターを設けております。ご利用の風景が外から見えることが、最も確実な集客につながります。"],
  ["給排水", "奥の壁面1面にドリンクカウンターを集約し、給排水・電源の引込みを最短としております。工事範囲の抑制に直結いたします。"],
].forEach((t, i) => {
  const y = 5.70 + i * 0.52;
  s.addText(t[0], { x: RX, y, w: 0.78, h: 0.44, fontFace: FG, fontSize: 8.3, bold: true, color: C.gr, margin: 0, valign: "top", lineSpacingMultiple: 1.1 });
  s.addText(t[1], { x: RX + 0.82, y: y - 0.02, w: RW - 0.82, h: 0.5, fontFace: FG, fontSize: 7.8, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.12 });
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
