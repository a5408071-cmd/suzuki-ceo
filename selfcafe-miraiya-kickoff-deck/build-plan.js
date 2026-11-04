/**
 * 秋田店 セルフカフェ区画（30坪）図面落とし込み — A4横 3ページ
 * ベース図：未来屋書店 秋田店 平面図（DWG.NO 02-1 / 縮尺1/160 / 2023.06.14）
 * 縮尺検証：区画実測 208.05m2 = 62.94坪（図面表記 62.97坪、誤差0.05%）
 */
const pptxgen = require("pptxgenjs");
const path = require("path");

const C = {
  bg: "FFFFFF", ink: "1A1A1A", green: "106838", greenDeep: "0C5230",
  gold: "8A6316", rule: "333333", line: "BFBFBF", lightLine: "DCDCDC",
  head: "EDEDED", band: "F7F7F7", tint: "EFF5F1", goldTint: "FBF5E8",
  body: "333333", muted: "6B6B6B", white: "FFFFFF", red: "A52A2A",
};
const F = { jp: "Yu Gothic", num: "Arial" };
const HAS_JP = /[ぁ-んァ-ヶ一-龥々〜ー]/;
const numFace = (s) => (HAS_JP.test(String(s)) ? F.jp : F.num);
const A = path.join(__dirname, "assets");
const W = 11.69, H = 8.27, M = 0.42, CW = W - M * 2, R = M + CW;
const DOC = "未来屋書店様 秋田店 セルフカフェ区画 図面落とし込み";

const pres = new pptxgen();
pres.defineLayout({ name: "A4L", width: W, height: H });
pres.layout = "A4L";
pres.author = "セルフカフェ株式会社";
pres.title = "秋田店 セルフカフェ区画（30坪）図面落とし込み";

let PAGE = 0;
function head(s, title, sub, doc) {
  PAGE++;
  s.background = { color: C.bg };
  s.addText(doc || DOC, { x: M, y: 0.34, w: 7.4, h: 0.18, fontFace: F.jp, fontSize: 8, color: C.muted, margin: 0, valign: "middle" });
  s.addText(String(PAGE), { x: R - 1.0, y: 0.34, w: 1.0, h: 0.18, fontFace: F.num, fontSize: 8, color: C.muted, align: "right", margin: 0, valign: "middle" });
  s.addShape("rect", { x: M, y: 0.56, w: CW, h: 0.014, fill: { color: C.rule } });
  s.addText(title, { x: M, y: 0.62, w: CW - 3.6, h: 0.34, fontFace: F.jp, fontSize: 14, bold: true, color: C.ink, margin: 0, valign: "middle" });
  if (sub) s.addText(sub, { x: R - 3.6, y: 0.62, w: 3.6, h: 0.34, fontFace: F.jp, fontSize: 8, color: C.muted, align: "right", margin: 0, valign: "middle" });
  s.addShape("rect", { x: M, y: 0.99, w: CW, h: 0.008, fill: { color: C.line } });
}
const secAt = (s, x, y, w, t) => {
  s.addShape("rect", { x, y: y + 0.03, w: 0.055, h: 0.16, fill: { color: C.greenDeep } });
  s.addText(t, { x: x + 0.14, y, w: w - 0.14, h: 0.22, fontFace: F.jp, fontSize: 10, bold: true, color: C.ink, margin: 0, valign: "middle" });
};
const sec = (s, y, t, w) => secAt(s, M, y, w || CW, t);
const th = (t, al) => ({ text: t, options: { fill: { color: C.head }, color: C.ink, bold: true, fontSize: 8.5, fontFace: F.jp, align: al || "left", valign: "middle", margin: [0.02, 0.08, 0.02, 0.08] } });
const td = (t, o = {}) => ({ text: t, options: { fill: { color: o.fill || C.white }, color: o.color || C.body, bold: !!o.bold, fontSize: o.size || 8.5, fontFace: o.face || numFace(t), align: o.align || "left", valign: "middle", margin: [0.02, 0.08, 0.02, 0.08] } });
function bullets(s, x, y, w, items, opt = {}) {
  const pitch = opt.pitch || 0.175, fs = opt.size || 8.5;
  items.forEach((t, i) => {
    const yy = y + i * pitch;
    const mark = typeof t === "object" ? (t.m || "・") : "・";
    const txt = typeof t === "object" ? t.t : t;
    const hl = typeof t === "object" && t.hl;
    s.addText(mark, { x, y: yy, w: 0.22, h: pitch, fontFace: F.jp, fontSize: fs - 0.5, bold: !!hl, color: hl ? C.gold : C.muted, margin: 0, valign: "middle" });
    s.addText(txt, { x: x + 0.24, y: yy, w: w - 0.24, h: pitch, fontFace: F.jp, fontSize: fs, bold: !!hl, color: hl ? C.gold : C.body, margin: 0, valign: "middle" });
  });
}

/* ===================== 図面の座標系 =====================
   クリップ範囲（PDFポイント）と配置先（インチ）の対応。
   1pt = 56.4232mm（実寸）。アスペクト比は厳密に保持。            */
const CX0 = 20, CY0 = 185, CX1 = 300, CY1 = 578;
const IX = M, IY = 1.10, IW = 4.30, IH = IW * (CY1 - CY0) / (CX1 - CX0);
const S = IW / (CX1 - CX0);                 // inch / pt
const X = (p) => IX + (p - CX0) * S;
const Y = (p) => IY + (p - CY0) * S;
const D = (p) => p * S;
const MMPT = 420.0 / 1191 * 160;            // 56.4232 mm / pt

// カフェ区画の実測座標（PDFポイント・回転補正済み）
const CAFE = { x0: 37.24, y0: 210.36, x1: 271.36, y1: 557.07 };
// 案A：厨房前（既存厨房の東面から東側へ）— 推奨
const KE = 102.3;       // 既存厨房の東面
const PA_TOP = 341.98;  // 区画北端から 7,426mm
const PA = [
  { x0: KE, y0: PA_TOP, x1: 233.09, y1: 417.91 },
  { x0: KE, y0: 417.91, x1: 246.47, y1: 510.55 },
  { x0: KE, y0: 510.55, x1: 271.36, y1: 557.07 },
];
// 案B：厨房込み・西側縦型
const PB = { x0: 37.24, y0: 347.13, x1: 186.19, y1: 557.07 };
// 既存厨房（既存造作・現状のまま）の外形
const KIT = { x0: 48.0, y0: 338.3, x1: 102.3, y1: 533.7 };

function planBase(s) {
  s.addImage({ path: path.join(A, "akita-plan-cafe.png"), x: IX, y: IY, w: IW, h: IH });
  s.addShape("rect", { x: IX, y: IY, w: IW, h: IH, fill: { type: "none" }, line: { color: C.line, width: 0.75 } });
}
const fillRect = (s, r, color, tr) =>
  s.addShape("rect", { x: X(r.x0), y: Y(r.y0), w: D(r.x1 - r.x0), h: D(r.y1 - r.y0), fill: { color, transparency: tr }, line: { type: "none" } });
const seg = (s, x1, y1, x2, y2, o = {}) =>
  s.addShape("line", { x: X(x1), y: Y(y1), w: D(x2 - x1), h: D(y2 - y1), line: { color: o.color || C.green, width: o.w || 1.75, dashType: o.dash || "solid" } });
function tag(s, cx, cy, w, h, text, o = {}) {
  s.addShape("rect", { x: cx - w / 2, y: cy - h / 2, w, h, fill: { color: o.fill || C.white, transparency: o.tr === undefined ? 8 : o.tr }, line: { color: o.border || C.green, width: 0.75 } });
  s.addText(text, { x: cx - w / 2, y: cy - h / 2, w, h, fontFace: F.jp, fontSize: o.size || 8, bold: o.bold !== false, color: o.color || C.greenDeep, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 0.92 });
}
function kitchen(s) {
  s.addShape("rect", {
    x: X(KIT.x0), y: Y(KIT.y0), w: D(KIT.x1 - KIT.x0), h: D(KIT.y1 - KIT.y0),
    fill: { type: "none" }, line: { color: C.gold, width: 1.5, dashType: "dash" },
  });
}
function mallNote(s) {
  s.addText("モール共用通路（南面）", {
    x: X(60), y: Y(561), w: D(150), h: 0.17,
    fontFace: F.jp, fontSize: 7.5, color: C.muted, align: "left", margin: 0, valign: "middle",
  });
}
function legend(s, y) {
  const items = [
    [C.green, "セルフカフェ区画（30坪）"],
    [C.gold, "既存厨房（既存造作・現状のまま）"],
  ];
  items.forEach((it, i) => {
    const x = IX + i * 1.92;
    s.addShape("rect", { x, y: y + 0.045, w: 0.14, h: 0.09, fill: { color: it[0], transparency: i === 0 ? 55 : 100 }, line: { color: it[0], width: 0.75, dashType: i === 1 ? "dash" : "solid" } });
    s.addText(it[1], { x: x + 0.2, y, w: 1.9, h: 0.18, fontFace: F.jp, fontSize: 7, color: C.muted, margin: 0, valign: "middle" });
  });
}
const foot = (s, lines) => {
  s.addShape("rect", { x: M, y: H - 0.84, w: CW, h: 0.008, fill: { color: C.lightLine } });
  lines.forEach((t, i) =>
    s.addText(t, { x: M, y: H - 0.76 + i * 0.145, w: CW - 4.1, h: 0.14, fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle" })
  );
  s.addText("セルフカフェ株式会社　052-879-3557 ／ info@selfcafe.jp", {
    x: R - 4.0, y: H - 0.3, w: 4.0, h: 0.15, fontFace: F.jp, fontSize: 7.5, color: C.muted, align: "right", margin: 0, valign: "middle",
  });
};

/* ================================================= P1 案A（推奨） */
{
  const s = pres.addSlide();
  head(s, "１．セルフカフェ区画 30坪 ─ 案Ａ：厨房前（推奨）",
    "ベース図：貴社ご提供 平面図 DWG.NO 02-1／縮尺 1/160");

  planBase(s);
  // 既存厨房は区画外
  s.addShape("rect", {
    x: X(KIT.x0), y: Y(KIT.y0), w: D(KIT.x1 - KIT.x0), h: D(KIT.y1 - KIT.y0),
    fill: { color: C.gold, transparency: 86 }, line: { color: C.gold, width: 1.5, dashType: "dash" },
  });
  PA.forEach((r) => fillRect(s, r, C.green, 82));
  // 新設する区画境界（太線）
  seg(s, KE, PA_TOP, KE, 557.07, { color: C.greenDeep, w: 2.25 });
  seg(s, KE, PA_TOP, 233.09, PA_TOP, { color: C.greenDeep, w: 2.25 });
  // 既存区画の外周に沿う部分（細線）
  seg(s, 233.09, PA_TOP, 233.09, 417.91, { color: C.greenDeep, w: 1 });
  seg(s, 233.09, 417.91, 246.47, 417.91, { color: C.greenDeep, w: 1 });
  seg(s, 246.47, 417.91, 246.47, 510.55, { color: C.greenDeep, w: 1 });
  seg(s, 246.47, 510.55, 271.36, 510.55, { color: C.greenDeep, w: 1 });

  tag(s, (X(KE) + X(233.09)) / 2, Y(PA_TOP) + 0.13, 0.62, 0.17, "7,380", { size: 7.5, color: C.greenDeep });
  tag(s, X(271.36) - 0.36, Y(551), 0.64, 0.17, "9,539", { size: 7.5, color: C.greenDeep });
  tag(s, X(240), Y(430), 0.66, 0.17, "12,136", { size: 7.5, color: C.greenDeep });
  tag(s, (X(KE) + X(246.47)) / 2, Y(470), 1.42, 0.38, "セルフカフェ\n30.00坪（99.17㎡）", { size: 8 });
  tag(s, X(75), Y(420), 0.92, 0.38, "既存厨房\n区画外", { size: 7.5, color: C.gold, border: C.gold });
  tag(s, X(172), Y(278), 1.30, 0.34, "未来屋書店様\n32.94坪（108.88㎡）", { size: 8, color: C.muted, border: C.line });
  mallNote(s);
  legend(s, IY + IH + 0.06);

  const RX = IX + IW + 0.3, RW = R - RX;
  secAt(s, RX, 1.10, RW, "案Ａの諸元");
  s.addTable(
    [
      [th("項目"), th("数値・内容")],
      [td("セルフカフェ区画", { size: 9 }), td("99.17㎡（30.00坪）／ 奥行 12,136mm", { bold: true, color: C.green, size: 9 })],
      [td("間口（南面）", { size: 9 }), td("9,539mm　※北側は 7,380mm", { size: 9 })],
      [td("西端の位置", { size: 9 }), td("既存厨房の東面（区画西端から 3,671mm）", { size: 9 })],
      [td("既存厨房", { size: 9 }), td("区画外。約10坪は貴社側に残ります", { bold: true, color: C.gold, size: 9 })],
      [td("カフェ区画 全体", { size: 9 }), td("208.16㎡（62.97坪）※図面表記", { size: 9 })],
      [td("未来屋書店様側", { size: 9 }), td("108.88㎡（32.94坪）※既存厨房を含む", { size: 9 })],
      [td("想定席数", { size: 9 }), td("約59席（土浦店 30坪／59席／75杯と同等）", { size: 9 })],
    ],
    { x: RX, y: 1.40, w: RW, colW: [1.42, RW - 1.42], rowH: [0.2, 0.26, 0.24, 0.24, 0.24, 0.24, 0.24, 0.24], border: { type: "solid", color: C.line, pt: 0.75 }, autoPage: false }
  );

  secAt(s, RX, 3.36, RW, "既存厨房を区画外とする理由");
  bullets(s, RX, 3.66, RW, [
    "業務委託型のため、設置工事費は弊社の負担でございます。既存厨房を区画に含めますと、厨房設備の撤去と床・壁の復旧が生じ、工事費と工期の双方が膨らみます。",
    "弊社のカフェは全自動抽出機を中心とした構成で、厨房設備を必要といたしません。厨房を区画外とすることで、工事範囲を床・内装・機器設置に限定できます。",
    "給排水は、隣接する既存厨房から分岐して引き込む想定でございます。引込距離が最短となり、配管工事を最小限に抑えられます。",
    "モール共用通路に面する南面を 9,539mm 確保でき、通路からの視認性は十分でございます。",
  ], { pitch: 0.46, size: 8.5 });

  secAt(s, RX, 5.60, RW, "この案の前提となるご確認事項");
  s.addShape("rect", { x: RX, y: 5.90, w: RW, h: 0.62, fill: { color: C.goldTint }, line: { color: C.gold, width: 1 } });
  s.addText("既存厨房の給排水を、区画内へ分岐してご使用いただけますでしょうか。", {
    x: RX + 0.16, y: 5.96, w: RW - 0.32, h: 0.2, fontFace: F.jp, fontSize: 9.5, bold: true, color: C.gold, margin: 0, valign: "middle",
  });
  s.addText("分岐位置・口径・排水勾配の確保可否により、本案の成立が決まります。", {
    x: RX + 0.16, y: 6.17, w: RW - 0.32, h: 0.26, fontFace: F.jp, fontSize: 8.5, color: C.body, margin: 0, valign: "top",
  });
  bullets(s, RX, 6.60, RW, [
    "区画外となる既存厨房 約10坪 の扱い（貴社でご利用か、クリエ様の原状回復対象か）。",
    "30坪の位置は本日のご相談用のたたき台でございます。現地内見のうえ確定させていただきます。",
  ], { pitch: 0.2, size: 8.5 });

  foot(s, [
    "※　本図は区画位置のご相談を目的とした資料であり、工事図面ではございません。",
    "※　縮尺 1/160。図示寸法は図面からの実測値であり、現地実測により差異が生じる場合がございます。",
  ]);
}

/* ================================================= P2 案B と比較 */
{
  const s = pres.addSlide();
  head(s, "２．比較案 ─ 案Ｂ：厨房込み（西側縦型）", "同一縮尺・同一ベース図");

  planBase(s);
  fillRect(s, PB, C.green, 82);
  s.addShape("rect", {
    x: X(PB.x0), y: Y(PB.y0), w: D(PB.x1 - PB.x0), h: D(PB.y1 - PB.y0),
    fill: { type: "none" }, line: { color: C.greenDeep, width: 2 },
  });
  kitchen(s);
  s.addShape("rect", {
    x: X(KIT.x0), y: Y(347.13), w: D(KIT.x1 - KIT.x0), h: D(KIT.y1 - 347.13),
    fill: { color: C.red, transparency: 76 }, line: { type: "none" },
  });
  tag(s, (X(PB.x0) + X(PB.x1)) / 2, Y(PB.y0) + 0.13, 0.62, 0.17, "8,400", { size: 7.5, color: C.greenDeep });
  tag(s, X(PB.x1) - 0.40, Y(385), 0.68, 0.17, "11,850", { size: 7.5, color: C.greenDeep });
  tag(s, X(PB.x1) - 0.74, Y(452), 1.38, 0.38, "セルフカフェ\n30.00坪（99.17㎡）", { size: 8 });
  tag(s, X(75), Y(318), 1.02, 0.30, "既存厨房が\n区画内に入る", { size: 7.5, color: C.red, border: C.red });
  tag(s, X(172), Y(278), 1.30, 0.34, "未来屋書店様\n32.94坪（108.88㎡）", { size: 8, color: C.muted, border: C.line });
  mallNote(s);
  legend(s, IY + IH + 0.06);

  const RX = IX + IW + 0.3, RW = R - RX;
  secAt(s, RX, 1.10, RW, "２案の比較");
  const cw2 = (RW - 1.52) / 2;
  s.addTable(
    [
      [th("観点"), th("案Ａ　厨房前", "center"), th("案Ｂ　厨房込み", "center")],
      [td("面積", { size: 8.5 }), td("30.00坪", { align: "center", size: 8.5 }), td("30.00坪", { align: "center", size: 8.5 })],
      [td("形状", { size: 8.5 }), td("間口9,539 × 奥行12,136", { align: "center", size: 8.5 }), td("8,400 × 11,850", { align: "center", size: 8.5 })],
      [td("既存厨房", { size: 8.5 }), td("区画外", { align: "center", bold: true, color: C.green, size: 8.5 }), td("区画内", { align: "center", color: C.red, size: 8.5 })],
      [td("厨房設備の撤去・復旧", { size: 8.5 }), td("不要", { align: "center", bold: true, color: C.green, size: 8.5 }), td("必要", { align: "center", bold: true, color: C.red, size: 8.5 })],
      [td("弊社負担の工事費", { size: 8.5 }), td("最小", { align: "center", bold: true, color: C.green, size: 8.5 }), td("増加", { align: "center", bold: true, color: C.red, size: 8.5 })],
      [td("給排水", { size: 8.5 }), td("既存厨房から分岐（最短）", { align: "center", size: 8.5 }), td("既存位置をそのまま使用", { align: "center", size: 8.5 })],
      [td("モール通路への接道", { size: 8.5 }), td("南面 9,539mm", { align: "center", bold: true, color: C.green, size: 8.5 }), td("南面 7,743mm ＋ 南西の角", { align: "center", size: 8.5 })],
      [td("貴社にお戻しする形", { size: 8.5 }), td("北側 ＋ 西側（厨房）", { align: "center", size: 8.5 }), td("北側 ＋ 東側", { align: "center", size: 8.5 })],
      [td("弊社の評価", { size: 8.5, bold: true, fill: C.band }), td("推奨", { align: "center", bold: true, color: C.white, size: 9, fill: C.green }), td("次善", { align: "center", bold: true, color: C.muted, size: 9, fill: C.band })],
    ],
    { x: RX, y: 1.40, w: RW, colW: [1.52, cw2, cw2], rowH: [0.2, 0.22, 0.26, 0.22, 0.24, 0.22, 0.26, 0.26, 0.24, 0.26], border: { type: "solid", color: C.line, pt: 0.75 }, autoPage: false }
  );

  secAt(s, RX, 4.08, RW, "判断の分かれ目");
  s.addShape("rect", { x: RX, y: 4.38, w: RW, h: 0.94, fill: { color: C.goldTint }, line: { color: C.gold, width: 1 } });
  s.addText("既存厨房の給排水を区画内へ分岐してご使用いただけるか。この一点でございます。", {
    x: RX + 0.16, y: 4.46, w: RW - 0.32, h: 0.22, fontFace: F.jp, fontSize: 9.5, bold: true, color: C.gold, margin: 0, valign: "middle",
  });
  s.addText("分岐が可能であれば案Ａが、費用・工期の双方で有利でございます。分岐が難しく、区画内に給排水を新設する必要がある場合は、既存の設備をそのまま使える案Ｂが現実的となります。本日、この前提のご確認をお願いしたく存じます。", {
    x: RX + 0.16, y: 4.68, w: RW - 0.32, h: 0.58, fontFace: F.jp, fontSize: 8.5, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.1,
  });

  secAt(s, RX, 5.52, RW, "案Ｂを採る場合の留意点");
  bullets(s, RX, 5.82, RW, [
    "既存厨房の設備（シンク・冷蔵機器・排気フード等）の撤去と、床・壁・天井の復旧が工事範囲に加わります。弊社負担の設置工事費が増加し、工期も延びます。",
    "弊社のカフェは全自動抽出機を中心とした構成のため、厨房設備そのものは使用いたしません。区画に含めても、設備を活かす場面は限られます。",
    "一方で、給排水・電源・排気が既存位置のまま使えるため、設備工事のリスクは最も小さくなります。",
  ], { pitch: 0.46, size: 8.5 });

  foot(s, [
    "※　両案とも面積は 30.00坪ちょうどで作図しております。",
    "※　区画の確定にあたっては、館（イオンモール様）の区画変更申請の要否についても併せてご確認をお願いいたします。",
  ]);
}

/* ================================================= P3 読み取り事実と確認事項 */
{
  const s = pres.addSlide();
  head(s, "３．ご提供図面から読み取れた事項と、ご確認のお願い");

  sec(s, 1.10, "ご提供いただいた図面");
  s.addTable(
    [
      [th("図面", "center"), th("内容"), th("縮尺", "center"), th("日付", "center"), th("枚数", "center")],
      [
        td("①", { align: "center", bold: true, color: C.green, size: 9 }),
        td("未来屋書店 秋田店 平面図（DWG.NO 02-1）　イオンモール秋田 2F／店舗番号 271-1", { size: 8.5 }),
        td("1/160", { align: "center", size: 8.5 }), td("2023.06.14", { align: "center", size: 8.5 }), td("1", { align: "center", size: 8.5 }),
      ],
      [
        td("②", { align: "center", bold: true, color: C.green, size: 9 }),
        td("CAFÉ de CRIÉ イオンモール秋田店　内装実施図一式（pokka create co.ltd 作成）", { size: 8.5 }),
        td("1/50 ほか", { align: "center", size: 8.5 }), td("2016.05.25", { align: "center", size: 8.5 }), td("26", { align: "center", size: 8.5 }),
      ],
    ],
    { x: M, y: 1.40, w: CW, colW: [0.44, 7.25, 1.0, 1.1, 1.06], rowH: [0.2, 0.3, 0.3], border: { type: "solid", color: C.line, pt: 0.75 }, autoPage: false }
  );

  const LW = 5.28, RX = M + LW + 0.29, RW = R - RX;

  secAt(s, M, 2.42, LW, "②の構成（全26枚）");
  bullets(s, M, 2.72, LW, [
    "平面図・床伏図（2枚）　什器配置と床仕上げ",
    "天井伏図（2枚）　照明・空調・排煙の位置",
    "躯体・間仕切図（1枚）　厨房まわりの区画",
    "断面詳細図（2枚）　壁・床の納まり",
    "展開図（4枚）　ファサードおよび店内立面",
    "什器詳細図（12枚）　カウンター・返却台・ショーケース・棚",
    "建具・サイン詳細図（3枚）",
  ], { pitch: 0.185, size: 8.5 });

  s.addShape("rect", { x: M, y: 4.08, w: LW, h: 0.58, fill: { color: C.band }, line: { color: C.line, width: 0.75 } });
  s.addText("設備図（給排水・電気容量・動力・換気）は②に含まれておりません。", {
    x: M + 0.14, y: 4.14, w: LW - 0.28, h: 0.2, fontFace: F.jp, fontSize: 8.5, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
  s.addText("残置範囲の判断と、弊社機器の電源・給排水計画に必要でございます。クリエ様へのご確認の際に、併せてご手配をお願いしたく存じます。", {
    x: M + 0.14, y: 4.34, w: LW - 0.28, h: 0.26, fontFace: F.jp, fontSize: 8, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.05,
  });

  secAt(s, M, 4.80, LW, "②から判明した既存店の諸元");
  s.addTable(
    [
      [th("項目"), th("記載内容")],
      [td("専有面積", { size: 8.5 }), td("132.5㎡（43坪）", { bold: true, color: C.ink, size: 8.5 })],
      [td("客席数", { size: 8.5 }), td("禁煙席 73席／51卓（1P席 39%）", { size: 8.5 })],
      [td("設計・施工", { size: 8.5 }), td("pokka create co.ltd（2016年5月）", { size: 8.5 })],
    ],
    { x: M, y: 5.10, w: 3.22, colW: [1.16, 2.06], rowH: [0.2, 0.24, 0.24, 0.24], border: { type: "solid", color: C.line, pt: 0.75 }, autoPage: false }
  );
  s.addImage({ path: path.join(A, "akita-crie-plan.png"), x: M + 3.38, y: 5.10, w: 1.90, h: 1.525 });
  s.addShape("rect", { x: M + 3.38, y: 5.10, w: 1.90, h: 1.525, fill: { type: "none" }, line: { color: C.line, width: 0.75 } });
  s.addText("②-1枚目　平面図（クリエ様 既存店・縮尺1/50）", {
    x: M + 3.38, y: 6.66, w: 1.90, h: 0.30, fontFace: F.jp, fontSize: 7, color: C.muted, margin: 0, valign: "top", lineSpacingMultiple: 1.05,
  });

  secAt(s, RX, 2.42, RW, "本日ご確認いただきたい事項");
  bullets(s, RX, 2.72, RW, [
    { t: "カフェ区画 62.97坪 と、クリエ様の専有面積 43坪 の差（約20坪）の扱い。書店側のカフェ席を含む範囲でしょうか、あるいは改装により区画が広がったものでしょうか。30坪をどこから切り出すかに直結いたします。", hl: true, m: "★" },
    { t: "平面図の凡例では、カフェ区画内の什器はすべて「既存転用什器」、厨房は「既存造作（現状のまま）」として図示されております。この前提が現時点でも有効か（クリエ様の原状回復義務がどこまで免除されるか）をご教示ください。", hl: true, m: "★" },
    { t: "既存厨房の給排水を、弊社区画内へ分岐してご使用いただけるか（分岐位置・口径・排水勾配）。あわせて分電盤の位置・容量。②に設備図が含まれていないため、現況のご提供をお願いいたします。", hl: true, m: "★" },
    { t: "区画外となる既存厨房 約10坪 の扱い。貴社でご利用になるか、クリエ様の原状回復の対象とされるかをご教示ください。", hl: true, m: "★" },
    "既存の排煙オペレーター（区画内に2か所図示）および空調の吹出位置について、区画を分割した場合の扱い。",
    "カフェ区画の西側外周および南側外周の仕様（外壁か、共用部との境界か）。ファサードの計画に必要でございます。",
    "区画を分割する場合の、館（イオンモール様）への区画変更申請の要否と、所要期間。",
  ], { pitch: 0.555, size: 8.5 });

  s.addShape("rect", { x: RX, y: 6.52, w: RW, h: 0.62, fill: { color: C.tint }, line: { color: C.green, width: 1 } });
  s.addText("11月の現地内見では、上記★の4点を実地で確認してまいります。内見の日程をご調整いただけますと幸いでございます。", {
    x: RX + 0.16, y: 6.62, w: RW - 0.32, h: 0.42, fontFace: F.jp, fontSize: 9, bold: true, color: C.greenDeep, margin: 0, valign: "middle", lineSpacingMultiple: 1.1,
  });

  foot(s, [
    "※　本資料は、貴社よりご提供いただいた図面2件に基づき、弊社にて作成したものでございます。",
  ]);
}

/* ================================================= 別紙1 秋田店 ご確認事項 */
const DOC2 = "未来屋書店様 事前打合せ 別紙（ご確認事項）";
{
  const s = pres.addSlide();
  head(s, "別紙１　秋田店　ご確認事項", "2026年11月4日 事前打合せ資料より", DOC2);

  const lw = 6.4, rx = M + lw + 0.35, rw = R - rx;
  secAt(s, M, 1.10, lw, "カフェ・ド・クリエ様へご確認をお願いしたい事項");
  s.addText("弊社が工事に取り掛かるにあたり、解約協議の中であわせてご確認いただけますと、工期と費用の精度が上がります。", {
    x: M, y: 1.34, w: lw, h: 0.18, fontFace: F.jp, fontSize: 8, color: C.body, margin: 0, valign: "middle",
  });
  const chk = [
    ["★", "原状回復の範囲", "どこまで残置いただけるか。工期と費用に最も影響します。"],
    ["★", "給排水の残置と分岐の可否", "経路・口径・止水栓の位置。弊社区画へ分岐してご使用いただけるかを含みます。"],
    ["★", "電気容量・分電盤の残置", "容量と回路数。カフェ機器の消費電力に足りるかを確認いたします。"],
    ["★", "既存厨房の扱い", "区画外となる約10坪を貴社でご利用か、クリエ様の原状回復対象とされるか。"],
    ["", "造作の残置可否", "カウンター・間仕切り・天井・照明・床仕上げ。"],
    ["", "厨房設備の撤去範囲", "撤去後の床・壁の状態（下地が露出するか、仕上げまで復旧されるか）。"],
    ["", "排気ダクト・グリストラップ", "残置か撤去か。残置の場合は閉塞処理の要否。"],
    ["", "看板・サインの撤去範囲", "下地・電源の残置可否。弊社サインの取り付けに影響します。"],
    ["", "閉店日と退去完了日", "弊社の工事着手可能日となります。"],
    ["", "館への原状回復届", "提出の有無と内容。弊社の工事申請との整合が必要です。"],
    ["", "引き渡し時の立会い", "弊社施工担当を同席させていただけますと、認識の齟齬を防げます。"],
  ];
  chk.forEach((c, i) => {
    const y = 1.58 + i * 0.47;
    const hl = c[0] === "★";
    s.addShape("rect", { x: M, y, w: lw, h: 0.43, fill: { color: hl ? C.goldTint : (i % 2 ? C.band : C.white) }, line: { color: C.lightLine, width: 0.5 } });
    s.addText(hl ? "★" : String(i + 1), {
      x: M + 0.06, y: y + 0.03, w: 0.3, h: 0.19,
      fontFace: F.jp, fontSize: 8, bold: true, color: hl ? C.gold : C.muted, align: "center", margin: 0, valign: "middle",
    });
    s.addText(c[1], {
      x: M + 0.42, y: y + 0.03, w: lw - 0.56, h: 0.19,
      fontFace: F.jp, fontSize: 9, bold: true, color: hl ? C.gold : C.ink, margin: 0, valign: "middle",
    });
    s.addText(c[2], {
      x: M + 0.42, y: y + 0.22, w: lw - 0.56, h: 0.19,
      fontFace: F.jp, fontSize: 8.5, color: C.body, margin: 0, valign: "middle",
    });
  });
  s.addText("★印の4点は、工期と工事費に直結いたします。優先してご確認いただけますと幸いでございます。", {
    x: M, y: 6.82, w: lw, h: 0.18, fontFace: F.jp, fontSize: 8.5, bold: true, color: C.gold, margin: 0, valign: "middle",
  });

  secAt(s, rx, 1.10, rw, "レイアウト作成に先立ち必要となる情報");
  bullets(s, rx, 1.36, rw, [
    "区画図面（平面図・展開図／CADデータ）",
    "既存カフェ区画の実測寸法",
    "電気容量および分電盤の位置",
    "給排水の位置・系統",
    "排気ダクト・空調の系統",
    "現況写真（客席・カウンター・厨房まわり）",
    "天井高および床仕上げの納まり",
    "館の内装制限・工事区分",
  ], { pitch: 0.19, size: 8.5 });

  secAt(s, rx, 3.00, rw, "残置の範囲が、工期と費用を決めます");
  s.addTable(
    [
      [th("閉店後の工程"), th("期間", "center")],
      [td("残置物の撤去・下地調整", { size: 8.5 }), td("3〜5日", { align: "center", size: 8.5 })],
      [td("内装工事（造作・床・サイン）", { size: 8.5 }), td("2〜3週間", { align: "center", size: 8.5 })],
      [td("機器設置・試運転・什器搬入", { size: 8.5 }), td("3〜5日", { align: "center", size: 8.5 })],
      [td("合計", { bold: true, color: C.ink, size: 9, fill: C.tint }), td("約1か月", { align: "center", bold: true, color: C.green, size: 9, fill: C.tint })],
    ],
    { x: rx, y: 3.26, w: rw, colW: [rw - 1.0, 1.0], rowH: [0.2, 0.24, 0.24, 0.24, 0.28], border: { type: "solid", color: C.line, pt: 0.75 }, autoPage: false }
  );

  s.addShape("roundRect", { x: rx, y: 4.62, w: rw, h: 1.00, rectRadius: 0.04, fill: { color: C.goldTint }, line: { color: "E2CFA6", width: 0.75 } });
  s.addText("スケルトン戻しとなる場合", {
    x: rx + 0.16, y: 4.70, w: rw - 0.32, h: 0.2, fontFace: F.jp, fontSize: 8.5, bold: true, color: C.gold, margin: 0, valign: "middle",
  });
  s.addText("下地からの復旧が必要となり、工期が2週間程度延びる見込みでございます。2027年3月上旬の開業が厳しくなります。", {
    x: rx + 0.16, y: 4.92, w: rw - 0.32, h: 0.64, fontFace: F.jp, fontSize: 8, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
  });

  s.addShape("roundRect", { x: rx, y: 5.82, w: rw, h: 1.26, rectRadius: 0.04, fill: { color: C.band }, line: { color: C.lightLine, width: 0.75 } });
  s.addText("ご相談のお願い", {
    x: rx + 0.16, y: 5.90, w: rw - 0.32, h: 0.2, fontFace: F.jp, fontSize: 8.5, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
  s.addText("カフェ・ド・クリエ様との解約協議の中で、造作・床・給排水・電気を可能な限り現状のまま残していただく方向でご調整いただけますと幸いでございます。工期の短縮と工事費の抑制の両方につながります。", {
    x: rx + 0.16, y: 6.12, w: rw - 0.32, h: 0.90, fontFace: F.jp, fontSize: 8, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
  });

  foot(s, [
    "※　工期は目安でございます。館（イオンモール様）の工事申請に要する期間は含んでおりません。",
  ]);
}

/* ================================================= 別紙2 ご相談ポイント・費用・工程 */
{
  const s = pres.addSlide();
  head(s, "別紙２　秋田店　ご相談ポイント・費用・今後の進め方", "", DOC2);

  const lw = 6.4, rx = M + lw + 0.35, rw = R - rx;
  secAt(s, M, 1.10, lw, "レイアウト作成に伴うご相談ポイント");
  const pts = [
    ["席数と席種の構成", "カウンター席／4人掛け／ソファ席の比率。30坪で約55〜60席を想定しております。"],
    ["マシン・決済端末の位置", "電源および給排水の位置に制約されます。壁面か、島配置かをご相談させてください。"],
    ["書籍の持ち込み運用", "ご購入前の書籍・雑誌の持ち込みを可とするか。可の場合、動線と汚損対策の整理が必要です。"],
    ["レジとの関係", "書店レジとカフェ区画の距離。お客様の動線と、スタッフの目が届く範囲。"],
    ["サイン表記", "「MIRAIYA Book & Cafe」の最終表記と、館の看板規定への適合。"],
    ["既存什器の再利用範囲", "テーブル・椅子・造作のうち、再利用可能なものの選別。"],
    ["バックヤード", "原料・備品のストック場所、ゴミの一時保管場所の確保。"],
    ["清掃・補充の動線", "1日15分の作業を、書店業務の流れの中でどこに組み込むか。"],
  ];
  pts.forEach((pt, i) => {
    const y = 1.36 + i * 0.52;
    s.addShape("rect", { x: M, y, w: lw, h: 0.48, fill: { color: i % 2 ? C.band : C.white }, line: { color: C.lightLine, width: 0.5 } });
    s.addText((i + 1) + "．" + pt[0], {
      x: M + 0.14, y: y + 0.03, w: lw - 0.28, h: 0.2, fontFace: F.jp, fontSize: 8.5, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(pt[1], {
      x: M + 0.14, y: y + 0.23, w: lw - 0.28, h: 0.22, fontFace: F.jp, fontSize: 8, color: C.body, margin: 0, valign: "middle",
    });
  });

  secAt(s, M, 5.66, lw, "区画の見え方（本日のご相談）");
  bullets(s, M, 5.92, lw, [
    { m: "①", t: "書店売場との境界の作り方", hl: true },
    "　什器で仕切るか、床仕上げのみで区切るか",
    { m: "②", t: "店頭（モール通路側）の見え方", hl: true },
    "　サインの位置と、カフェを前面に出すか書店を前面にするか",
  ], { pitch: 0.205, size: 8.5 });
  s.addText("※ 30坪をどこから切り出すかにつきましては、本編 P1・P2 にて2案をご提案しております。", {
    x: M, y: 6.80, w: lw, h: 0.18, fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
  });

  secAt(s, rx, 1.10, rw, "費用・契約に関するご確認事項");
  bullets(s, rx, 1.36, rw, [
    "業務委託型のため、工事費のお見積りのご提出はございません。貴社にご負担いただく費用は、設置工事費（30〜50万円・税抜）のみでございます。",
    "防犯カメラの映像（個人情報）の取扱い。保存期間・閲覧権限等の取り決め案を弊社よりご提示いたします。覚書の要否をご教示ください。",
    "インターネット回線は貴社にて新規お引き込みをお願いしております。お申込みから開通まで1〜2か月を要します。",
    "水道光熱費・通信費・消耗品費は貴社ご負担での運用を想定しております。ご認識に相違がないかご確認ください。",
  ], { pitch: 0.50, size: 8.5 });

  secAt(s, rx, 3.44, rw, "本日お決めいただきたいこと");
  const dec = [
    ["①", "11月上旬の現地内見の日程", "2〜3候補をいただけますと幸いです。"],
    ["②", "30坪の候補領域の方向性", "案Ａ（厨房前）か案Ｂ（厨房込み）か。"],
    ["③", "クリエ様へご確認いただく事項", "別紙１の11項目のうち、協議にのせていただく範囲。"],
    ["④", "仙台の確認事項の回答期限の目安", "いつ頃までにご回答いただけるか。"],
  ];
  dec.forEach((d, i) => {
    const y = 3.66 + i * 0.54;
    s.addShape("rect", { x: rx, y, w: rw, h: 0.50, fill: { color: i % 2 ? C.band : C.white }, line: { color: C.lightLine, width: 0.5 } });
    s.addText(d[0], {
      x: rx + 0.08, y: y + 0.04, w: 0.26, h: 0.2, fontFace: F.jp, fontSize: 9, bold: true, color: C.green, margin: 0, valign: "middle",
    });
    s.addText(d[1], {
      x: rx + 0.36, y: y + 0.04, w: rw - 0.46, h: 0.2, fontFace: F.jp, fontSize: 8.5, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(d[2], {
      x: rx + 0.36, y: y + 0.24, w: rw - 0.46, h: 0.24, fontFace: F.jp, fontSize: 8, color: C.body, margin: 0, valign: "middle",
    });
  });

  secAt(s, rx, 5.88, rw, "主要工程");
  const sch = [
    ["11月上旬", "現地内見（弊社施工担当同行）"],
    ["11月中旬", "レイアウト図のご提出／解約条件の確定"],
    ["12月上旬", "回線お申込み・什器機器の発注"],
    ["1月末", "クリエ様 閉店・退去"],
    ["2月", "内装工事・機器設置（約1か月）"],
    ["3月上旬", "開業"],
  ];
  s.addTable(
    [
      [th("時期"), th("内容")],
      ...sch.map((r, i) => {
        const last = i === sch.length - 1;
        const f = last ? C.tint : (i % 2 ? C.band : C.white);
        return [
          td(r[0], { bold: true, color: last ? C.green : C.ink, size: 8, fill: f }),
          td(r[1], { size: 8, fill: f, bold: last, color: last ? C.green : C.body }),
        ];
      }),
    ],
    { x: rx, y: 6.12, w: rw, colW: [0.95, rw - 0.95], rowH: [0.18, ...sch.map(() => 0.175)], border: { type: "solid", color: C.line, pt: 0.75 }, autoPage: false }
  );

  foot(s, [
    "※　インターネット回線はお申込みから開通まで1〜2か月を要します。12月中のお手続きをお願いいたします。",
    "※　工期約1か月は、造作・床・給排水・電気を残置いただくことを前提としております。",
  ]);
}

/* ================================================= 別紙3 仙台 */
{
  const s = pres.addSlide();
  head(s, "別紙３　仙台新規出店　ご確認事項", "売上試算は確認事項が揃いしだい作成いたします", DOC2);

  const lw = 4.6, rx = M + lw + 0.35, rw = R - rx;
  secAt(s, M, 1.10, lw, "ご提供いただいた物件情報");
  const info = [
    ["商業施設", "旧アリオ仙台泉（SENDAI IZUMI Reborn PROJECT）"],
    ["所在地", "仙台市泉区泉中央1-5-1／泉中央駅 徒歩2分"],
    ["開業予定", "2028年度内予定"],
    ["館の規模", "総賃貸面積 約20,000㎡／専門店50〜70店舗"],
    ["商業集積", "SELVA等と合わせ約40,000㎡・約180店舗"],
    ["交通", "地下鉄・バス 1日約3.1万人／年間約1,130万人"],
    ["書店フロア", "4F"],
    ["書店の坪数", "446坪（新店全体）"],
    ["書籍売場来店客数", "13,000人／月（想定）"],
    ["セルフカフェ坪数", "30坪（60坪の試算もご希望）"],
    ["水回り", "なし"],
    ["電源・回線", "新規物件のため工事対応"],
    ["営業時間", "書店・カフェとも未定"],
  ];
  s.addTable(
    [
      [th("項目"), th("内容")],
      ...info.map((r) => [td(r[0], { bold: true, color: C.ink, size: 8 }), td(r[1], { size: 8 })]),
    ],
    { x: M, y: 1.36, w: lw, colW: [1.45, 3.15], rowH: [0.2, ...info.map(() => 0.235)], border: { type: "solid", color: C.line, pt: 0.75 }, autoPage: false }
  );

  secAt(s, M, 4.80, lw, "仙台の進め方");
  bullets(s, M, 5.06, lw, [
    "本日　　物件情報・確認事項のすり合わせ",
    "後日　　確認事項が揃いしだい、30坪・60坪の2案で売上試算をご提出",
    "その後　出店条件のご提示・ご協議",
  ], { pitch: 0.21, size: 8.5 });

  secAt(s, M, 5.92, lw, "ご参考：弊社既存店の販売実績");
  s.addShape("roundRect", { x: M, y: 6.18, w: lw, h: 0.80, rectRadius: 0.04, fill: { color: C.band }, line: { color: C.lightLine, width: 0.75 } });
  s.addText("書店併設型の既存店（2026年5〜8月平均）　1日 40〜45杯程度\n商業施設内の大型店　1日 約123杯", {
    x: M + 0.16, y: 6.28, w: lw - 0.32, h: 0.60, fontFace: F.jp, fontSize: 8.5, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.35,
  });

  secAt(s, rx, 1.10, rw, "本日ご確認させていただきたい事項");
  s.addShape("roundRect", { x: rx, y: 1.36, w: rw, h: 1.72, rectRadius: 0.04, fill: { color: C.goldTint }, line: { color: "E2CFA6", width: 0.75 } });
  s.addText("【最重要】来客数の定義をそろえさせてください", {
    x: rx + 0.18, y: 1.44, w: rw - 0.36, h: 0.2, fontFace: F.jp, fontSize: 9.5, bold: true, color: C.gold, margin: 0, valign: "middle",
  });
  s.addText("秋田店・土浦店の試算は「カフェ区画の来客数」（秋田3,760人／月）を分母としております。一方、今回の13,000人／月は「書籍売場への来店客数」であり、分母が異なります。同じ購入率を当てはめると過大な試算となってしまいます。\n\nつきましては、13,000人／月の算出方法（書店への入店客数か、レジ通過客数か等）をご教示いただくとともに、同じ定義での秋田店・土浦店の数値をあわせてご提供いただけませんでしょうか。既存店の販売実績と突き合わせることで、仙台の販売杯数に根拠を持たせることができます。", {
    x: rx + 0.18, y: 1.66, w: rw - 0.36, h: 1.34, fontFace: F.jp, fontSize: 8, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
  });

  const grp = (y, t) => {
    s.addText(t, { x: rx, y, w: rw, h: 0.2, fontFace: F.jp, fontSize: 9, bold: true, color: C.ink, margin: 0, valign: "middle" });
    s.addShape("rect", { x: rx, y: y + 0.21, w: rw, h: 0.008, fill: { color: C.lightLine } });
  };

  grp(3.20, "Ａ．売上試算の前提に関する事項");
  bullets(s, rx, 3.44, rw, [
    { m: "★", t: "13,000人／月の算出方法と、同じ定義での秋田店・土浦店の数値", hl: true },
    "館全体の想定来客数（年間・月間）",
    "4Fのフロア構成と、核となるテナント",
    "カフェ区画の想定位置（通路・エスカレーターとの関係）",
    "書店の想定営業時間／カフェの希望営業時間",
    "想定席数（30坪・60坪それぞれ）",
    "書籍の持ち込み可否など、運用のお考え",
  ], { pitch: 0.185, size: 8.5 });

  grp(4.86, "Ｂ．工事・設備に関する事項");
  bullets(s, rx, 5.10, rw, [
    { m: "★", t: "給排水の新設可否と引込ルート（水回りなしのため）", hl: true },
    { m: "★", t: "電源容量および回線の引込条件・申請先", hl: true },
    "館側の内装制限・工事区分（Ａ工事／Ｂ工事／Ｃ工事）",
    "什器・サインに関する館の規定",
  ], { pitch: 0.185, size: 8.5 });

  grp(5.94, "Ｃ．出店条件・スケジュールに関する事項");
  bullets(s, rx, 6.18, rw, [
    "開業時期の確度（2028年度内のいつ頃か）",
    "出店条件（賃料・共益費等）の概算",
    "ご想定の契約形態（業務委託型／FC）",
    "貴社内のご検討スケジュール",
  ], { pitch: 0.185, size: 8.5 });

  foot(s, [
    "※　★印の3点は、売上試算の根拠および工事費・出店可否そのものに影響いたします。",
    "※　売上試算は確認事項が揃いしだい、30坪・60坪の2案で作成のうえご提出いたします。本日時点では未着手でございます。",
  ]);
}

/* ================================================= 別紙4 仙台 試算Q&A */
{
  const s = pres.addSlide();
  head(s, "別紙４　仙台新規出店　売上試算のためのご質問（ご回答記入シート）", "", DOC2);

  // 試算の組み立て
  s.addShape("roundRect", { x: M, y: 1.08, w: CW, h: 0.80, rectRadius: 0.04, fill: { color: C.goldTint }, line: { color: "E2CFA6", width: 0.75 } });
  s.addText("試算の組み立て　①カフェ区画の来客数（人／月）　→　②購入率　→　③販売杯数　→　④客単価・原価　→　⑤損益", {
    x: M + 0.16, y: 1.14, w: CW - 0.32, h: 0.2,
    fontFace: F.jp, fontSize: 9, bold: true, color: C.gold, margin: 0, valign: "middle",
  });
  s.addText("既存店では、秋田店 3,760人／月に対し 85杯／日（購入率 約68%）、土浦店 4,027人／月に対し 75杯／日（約56%）と幅がございます。仙台の 13,000人／月は「書籍売場への来店客数」で分母の定義が異なるため、そのまま当てはめますと過大な試算となってしまいます。同じ定義での既存店の数値をいただければ、仙台の販売杯数に根拠を持たせることができます。", {
    x: M + 0.16, y: 1.36, w: CW - 0.32, h: 0.46,
    fontFace: F.jp, fontSize: 8, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
  });

  const grpRow = (t) => [{
    text: t,
    options: { colspan: 4, fill: { color: C.tint }, color: C.greenDeep, bold: true, fontSize: 8.5, fontFace: F.jp, valign: "middle", margin: [0.02, 0.1, 0.02, 0.1] },
  }];
  const qRow = (no, q, eff, hl) => [
    td(no, { align: "center", bold: true, size: 8, color: hl ? C.gold : C.muted, fill: hl ? C.goldTint : C.white }),
    td((hl ? "★　" : "") + q, { size: 8, color: hl ? C.gold : C.body, bold: !!hl, fill: hl ? C.goldTint : C.white }),
    td(eff, { size: 7.5, color: C.muted, align: "center", fill: hl ? C.goldTint : C.white }),
    td("", { fill: C.white }),
  ];

  const QS = [
    ["g", "Ⅰ．来客数 ─ 試算の分母となる数値でございます（最優先）"],
    ["1", "13,000人／月は、どのように算出された数値でしょうか。", "分母そのもの", 1],
    ["2", "同じ定義での秋田店・土浦店の月間来客数。", "購入率の検証", 1],
    ["3", "館全体の想定来客数（年間・月間）。", "立地ポテンシャル"],
    ["4", "13,000人／月の根拠とされた類似館・既存店。", "妥当性の検証"],
    ["g", "Ⅱ．立地・区画 ─ カフェ区画への来店率に影響いたします"],
    ["5", "4Fのフロア構成（核テナント、ＥＳ・ＥＶの位置）。", "回遊・視認性"],
    ["6", "セルフカフェ区画の想定位置（通路に面するか、奥か）。", "来店率"],
    ["7", "4Fへの動線（直通ＥＶ・ＥＳの有無）。", "フロア到達率"],
    ["8", "書店446坪の売場構成（主要ジャンルとレジの位置）。", "滞在時間"],
    ["g", "Ⅲ．営業条件 ─ 販売杯数に直結いたします"],
    ["9", "館の営業時間と、書店の想定営業時間。", "営業時間"],
    ["10", "カフェの希望営業時間（書店営業時間に合わせる想定か）。", "稼働時間"],
    ["11", "想定席数（弊社想定：30坪＝約59席／60坪＝約73席）。", "席数の上限"],
    ["12", "ご購入前の書籍・雑誌の持ち込み可否。", "滞在時間"],
    ["g", "Ⅳ．費用・条件 ─ 損益および出店可否に影響いたします"],
    ["13", "給排水の新設可否と引込ルート（水回りなしのため）。", "工事費・可否", 1],
    ["14", "電源容量および回線の引込条件・申請先。", "工事費"],
    ["15", "館の工事区分（Ａ／Ｂ／Ｃ工事）と内装制限。", "工事費"],
    ["16", "賃料・共益費の扱い（貴社の賃借面積に含まれるか）。", "固定費"],
  ];
  s.addTable(
    [[th("", "center"), th("ご質問"), th("試算への影響", "center"), th("ご回答欄")],
     ...QS.map((r) => (r[0] === "g" ? grpRow(r[1]) : qRow(r[0], r[1], r[2], r[3])))],
    {
      x: M, y: 1.98, w: CW, colW: [0.36, 4.64, 1.25, 4.60],
      rowH: [0.2, ...QS.map((r) => (r[0] === "g" ? 0.2 : 0.235))],
      border: { type: "solid", color: C.line, pt: 0.75 }, autoPage: false,
    }
  );

  s.addText("★印の3点（Ｑ1・Ｑ2・Ｑ13）が揃いしだい、30坪・60坪の2案で売上試算を作成し、あらためてご提出いたします。他の項目は未定のままでも試算の着手は可能でございます。", {
    x: M, y: 6.84, w: CW, h: 0.2,
    fontFace: F.jp, fontSize: 8.5, bold: true, color: C.gold, margin: 0, valign: "middle",
  });
  s.addText("※　Ｑ1につきましては、書店への入店客数か、レジ通過客数か、館の通行量からの按分か、といった算出の考え方をご教示いただけますと幸いでございます。", {
    x: M, y: 7.08, w: CW, h: 0.2,
    fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
  });

  foot(s, [
    "※　ご回答欄にご記入のうえご返送いただく形でも、打合せの中で口頭にてご教示いただく形でも差し支えございません。",
  ]);
}

pres.writeFile({ fileName: path.join(__dirname, "akita-plan-30tsubo.pptx") }).then((f) => console.log("written:", f));
