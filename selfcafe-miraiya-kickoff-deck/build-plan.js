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
function head(s, title, sub) {
  PAGE++;
  s.background = { color: C.bg };
  s.addText(DOC, { x: M, y: 0.34, w: 7.4, h: 0.18, fontFace: F.jp, fontSize: 8, color: C.muted, margin: 0, valign: "middle" });
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
// 案A：西側縦型
const PA = { x0: 37.24, y0: 347.13, x1: 186.19, y1: 557.07 };
// 案B：南側横型（右端は既存区画の段差に沿う）
const PB_TOP = 412.95;
const PB = [
  { x0: 37.24, y0: PB_TOP, x1: 233.09, y1: 417.91 },
  { x0: 37.24, y0: 417.91, x1: 246.47, y1: 510.55 },
  { x0: 37.24, y0: 510.55, x1: 271.36, y1: 557.07 },
];
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
  head(s, "１．セルフカフェ区画 30坪 ─ 案Ａ：西側縦型（推奨）",
    "ベース図：貴社ご提供 平面図 DWG.NO 02-1／縮尺 1/160");

  planBase(s);
  fillRect(s, PA, C.green, 82);
  s.addShape("rect", {
    x: X(PA.x0), y: Y(PA.y0), w: D(PA.x1 - PA.x0), h: D(PA.y1 - PA.y0),
    fill: { type: "none" }, line: { color: C.greenDeep, width: 2 },
  });
  kitchen(s);
  // 寸法
  tag(s, (X(PA.x0) + X(PA.x1)) / 2, Y(PA.y0) + 0.13, 0.62, 0.17, "8,400", { size: 7.5, color: C.greenDeep });
  tag(s, X(PA.x1) - 0.40, Y(385), 0.68, 0.17, "11,850", { size: 7.5, color: C.greenDeep });
  // ラベル
  tag(s, X(PA.x1) - 0.74, Y(452), 1.38, 0.38, "セルフカフェ\n30.00坪（99.17㎡）", { size: 8 });
  tag(s, X(172), Y(278), 1.30, 0.34, "未来屋書店様\n32.94坪（108.88㎡）", { size: 8, color: C.muted, border: C.line });
  tag(s, X(75), Y(300), 0.78, 0.17, "既存厨房", { size: 7.5, color: C.gold, border: C.gold });
  mallNote(s);
  legend(s, IY + IH + 0.06);

  const RX = IX + IW + 0.3, RW = R - RX;
  secAt(s, RX, 1.10, RW, "案Ａの諸元");
  s.addTable(
    [
      [th("項目"), th("数値・内容")],
      [td("セルフカフェ区画", { size: 9 }), td("8,400mm × 11,850mm ／ 99.17㎡（30.00坪）", { bold: true, color: C.green, size: 9 })],
      [td("カフェ区画 全体", { size: 9 }), td("208.16㎡（62.97坪）※図面表記", { size: 9 })],
      [td("未来屋書店様側", { size: 9 }), td("108.88㎡（32.94坪）", { size: 9 })],
      [td("既存厨房", { size: 9 }), td("区画内に収まる（幅約3.0m × 長さ約11.0m）", { size: 9 })],
      [td("天井高・柱スパン", { size: 9 }), td("CH=3,000 ／ 8,000（図面記載）", { size: 9 })],
      [td("想定席数", { size: 9 }), td("約59席（土浦店 30坪／59席／75杯と同等）", { size: 9 })],
    ],
    { x: RX, y: 1.40, w: RW, colW: [1.52, RW - 1.52], rowH: [0.2, 0.26, 0.26, 0.26, 0.26, 0.26, 0.26], border: { type: "solid", color: C.line, pt: 0.75 }, autoPage: false }
  );

  secAt(s, RX, 3.36, RW, "この切り出し方をお薦めする理由");
  bullets(s, RX, 3.66, RW, [
    "既存厨房を区画内にほぼ全量収められるため、給排水・電源の既存位置をそのまま使える見込みでございます。",
    "工事範囲が内装と機器設置に限定され、閉店後 約1か月の工期と、3月上旬の開業予定を確保しやすくなります。",
    "モール共用通路に面する南面と、売場南西の角を確保でき、通路からの視認性を損ないません。",
    "未来屋書店様にお戻しする約33坪が、既存書架（031区画）に隣接してまとまります。",
  ], { pitch: 0.34, size: 8.5 });

  secAt(s, RX, 5.12, RW, "作図の前提");
  bullets(s, RX, 5.42, RW, [
    "貴社ご提供の平面図に、縮尺を実測検証のうえ重ねております。区画面積の実測値は 208.05㎡（62.94坪）で、図面表記 208.16㎡（62.97坪）との差は 0.05% でございます。",
    "30坪の位置は本日のご相談用のたたき台でございます。既存厨房の残置範囲が確定次第、正式な区画図として作成いたします。",
    "席数および什器の詳細レイアウトは、現地実測のうえ別途作図いたします。",
  ], { pitch: 0.44, size: 8.5 });

  foot(s, [
    "※　本図は区画位置のご相談を目的とした資料であり、工事図面ではございません。",
    "※　縮尺 1/160。図示寸法は図面からの実測値であり、現地実測により差異が生じる場合がございます。",
  ]);
}

/* ================================================= P2 案B と比較 */
{
  const s = pres.addSlide();
  head(s, "２．比較案 ─ 案Ｂ：南側横型", "同一縮尺・同一ベース図");

  planBase(s);
  PB.forEach((r) => fillRect(s, r, C.green, 82));
  // 切り出しライン（既存区画境界に沿わない新設部分）
  seg(s, PB[0].x0, PB_TOP, 233.09, PB_TOP, { color: C.greenDeep, w: 2 });
  seg(s, 233.09, PB_TOP, 233.09, 417.91, { color: C.greenDeep, w: 2 });
  seg(s, 233.09, 417.91, 246.47, 417.91, { color: C.greenDeep, w: 2 });
  seg(s, 246.47, 417.91, 246.47, 510.55, { color: C.greenDeep, w: 2 });
  seg(s, 246.47, 510.55, 271.36, 510.55, { color: C.greenDeep, w: 2 });
  kitchen(s);
  tag(s, (X(PB[0].x0) + X(233.09)) / 2, Y(PB_TOP) + 0.13, 0.62, 0.17, "8,131", { size: 7.5, color: C.greenDeep });
  tag(s, (X(120) + X(220)) / 2, Y(500), 1.26, 0.40, "セルフカフェ\n30.00坪（99.17㎡）", { size: 8 });
  tag(s, X(160), Y(300), 1.26, 0.34, "未来屋書店様\n32.94坪（108.88㎡）", { size: 8, color: C.muted, border: C.line });
  // 区画外となる厨房部分
  s.addShape("rect", {
    x: X(KIT.x0), y: Y(KIT.y0), w: D(KIT.x1 - KIT.x0), h: D(PB_TOP - KIT.y0),
    fill: { color: C.red, transparency: 74 }, line: { type: "none" },
  });
  tag(s, X(75), Y(372), 1.06, 0.30, "厨房の約4割が\n区画外", { size: 7.5, color: C.red, border: C.red });
  mallNote(s);
  legend(s, IY + IH + 0.06);

  const RX = IX + IW + 0.3, RW = R - RX;
  secAt(s, RX, 1.10, RW, "２案の比較");
  const cw2 = (RW - 1.46) / 2;
  s.addTable(
    [
      [th("観点"), th("案Ａ　西側縦型", "center"), th("案Ｂ　南側横型", "center")],
      [td("面積", { size: 8.5 }), td("30.00坪（99.17㎡）", { align: "center", size: 8.5 }), td("30.00坪（99.17㎡）", { align: "center", size: 8.5 })],
      [td("形状", { size: 8.5 }), td("8,400 × 11,850", { align: "center", size: 8.5 }), td("奥行 8,131 × 全幅", { align: "center", size: 8.5 })],
      [td("既存厨房", { size: 8.5 }), td("ほぼ全量を区画内に確保", { align: "center", bold: true, color: C.green, size: 8.5 }), td("約4割が区画外となる", { align: "center", bold: true, color: C.red, size: 8.5 })],
      [td("モール通路への接道", { size: 8.5 }), td("南面 約7.7m ＋ 南西の角", { align: "center", size: 8.5 }), td("南面 全幅 約13.2m", { align: "center", bold: true, color: C.green, size: 8.5 })],
      [td("貴社にお戻しする形", { size: 8.5 }), td("北側 ＋ 東側のＬ字", { align: "center", size: 8.5 }), td("北側の整形（長方形）", { align: "center", bold: true, color: C.green, size: 8.5 })],
      [td("工事範囲・工期", { size: 8.5 }), td("最小。残置を最大限活かせる", { align: "center", bold: true, color: C.green, size: 8.5 }), td("厨房の再構築が生じ得る", { align: "center", color: C.red, size: 8.5 })],
      [td("弊社の評価", { size: 8.5, bold: true, fill: C.band }), td("推奨", { align: "center", bold: true, color: C.white, size: 9, fill: C.green }), td("要調整", { align: "center", bold: true, color: C.muted, size: 9, fill: C.band })],
    ],
    { x: RX, y: 1.40, w: RW, colW: [1.46, cw2, cw2], rowH: [0.2, 0.24, 0.24, 0.3, 0.26, 0.26, 0.3, 0.26], border: { type: "solid", color: C.line, pt: 0.75 }, autoPage: false }
  );

  secAt(s, RX, 3.64, RW, "判断の分かれ目");
  s.addShape("rect", { x: RX, y: 3.94, w: RW, h: 0.86, fill: { color: C.goldTint }, line: { color: C.gold, width: 1 } });
  s.addText("２案の優劣は、既存厨房をどこまで残置いただけるかで決まります。", {
    x: RX + 0.16, y: 4.02, w: RW - 0.32, h: 0.22, fontFace: F.jp, fontSize: 9.5, bold: true, color: C.gold, margin: 0, valign: "middle",
  });
  s.addText("厨房が現状のまま使える前提であれば案Ａが最も有利でございます。厨房がスケルトン戻しとなる場合は、通路への接道が長く、貴社の売場も整形となる案Ｂが有力となります。本日は、この前提のご確認をお願いしたく存じます。", {
    x: RX + 0.16, y: 4.24, w: RW - 0.32, h: 0.50, fontFace: F.jp, fontSize: 8.5, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.1,
  });

  secAt(s, RX, 5.00, RW, "調整案（ご参考）");
  bullets(s, RX, 5.30, RW, [
    "南西Ｌ字型　─　厨房のある西側ストリップ（幅約4.2m）と、南側の帯（奥行約6.0m）を組み合わせる切り方でございます。厨房を全量確保しながら、南面の接道も長く取れます。",
    "一方で、貴社にお戻しする区画・弊社区画の双方がＬ字となるため、什器計画が複雑になります。案Ａ・案Ｂでまとまらない場合の次善案としてご提示いたします。",
  ], { pitch: 0.56, size: 8.5 });

  foot(s, [
    "※　両案とも面積は 30.00坪ちょうどで作図しております（案Ｂを既存区画の段差ラインで切る場合は 29.07坪）。",
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
    { t: "既存厨房の給排水の引込位置と能力、および分電盤の位置・容量。②に設備図が含まれていないため、現況のご提供をお願いいたします。", hl: true, m: "★" },
    "既存の排煙オペレーター（区画内に2か所図示）および空調の吹出位置について、区画を分割した場合の扱い。",
    "カフェ区画の西側外周および南側外周の仕様（外壁か、共用部との境界か）。ファサードの計画に必要でございます。",
    "区画を分割する場合の、館（イオンモール様）への区画変更申請の要否と、所要期間。",
  ], { pitch: 0.62, size: 8.5 });

  s.addShape("rect", { x: RX, y: 6.62, w: RW, h: 0.62, fill: { color: C.tint }, line: { color: C.green, width: 1 } });
  s.addText("11月の現地内見では、上記★の3点を実地で確認してまいります。内見の日程をご調整いただけますと幸いでございます。", {
    x: RX + 0.16, y: 6.72, w: RW - 0.32, h: 0.42, fontFace: F.jp, fontSize: 9, bold: true, color: C.greenDeep, margin: 0, valign: "middle", lineSpacingMultiple: 1.1,
  });

  foot(s, [
    "※　本資料は、貴社よりご提供いただいた図面2件に基づき、弊社にて作成したものでございます。",
  ]);
}

pres.writeFile({ fileName: path.join(__dirname, "akita-plan-30tsubo.pptx") }).then((f) => console.log("written:", f));
