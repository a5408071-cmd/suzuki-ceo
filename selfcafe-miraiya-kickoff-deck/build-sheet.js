/**
 * 未来屋書店 秋田店 スケジュールと役割分担 — A4横1枚
 * 前提：30坪／残置（居抜き）／2027年2月末オープン
 */
const pptxgen = require("pptxgenjs");
const path = require("path");

const C = {
  bg: "FFFDF7", ink: "0A3D23", green: "106838", greenDeep: "0C5230",
  greenLabel: "1B8149", gold: "A8761F", goldLine: "C79A44", goldTint: "FBF4E4",
  body: "4B5550", muted: "7E8882", footer: "A8B0AA",
  tint: "F1F8F3", tintLine: "D2E7DA", warmLine: "EFEAE0", grayBand: "F4F2EE", white: "FFFFFF",
};
const F = { jp: "Yu Gothic", num: "Arial" };
const HAS_JP = /[ぁ-んァ-ヶ一-龥々〜ー]/;
const numFace = (s) => (HAS_JP.test(String(s)) ? F.jp : F.num);
const A = path.join(__dirname, "assets");

const W = 11.69, H = 8.27, M = 0.42, CW = W - M * 2, R = M + CW;

const pres = new pptxgen();
pres.defineLayout({ name: "A4L", width: W, height: H });
pres.layout = "A4L";
pres.author = "セルフカフェ株式会社";
pres.title = "未来屋書店 秋田店 スケジュールと役割分担";

const s = pres.addSlide();
s.background = { color: C.bg };

/* ---------------- ヘッダー ---------------- */
s.addShape("rect", { x: M, y: 0.34, w: 0.2, h: 0.026, fill: { color: C.green } });
s.addText("未来屋書店 秋田店　カフェ区画 業態転換", {
  x: M + 0.3, y: 0.25, w: 7.0, h: 0.2,
  fontFace: F.jp, fontSize: 9.5, bold: true, color: C.greenLabel, margin: 0, valign: "middle",
});
s.addText("スケジュールと役割分担", {
  x: M, y: 0.46, w: 8.0, h: 0.44,
  fontFace: F.jp, fontSize: 22, bold: true, color: C.ink, margin: 0, valign: "middle",
});
s.addImage({ path: path.join(A, "logo-green.png"), x: R - 1.55, y: 0.3, w: 1.55, h: 0.221 });
s.addText("2026年10月17日　セルフカフェ株式会社", {
  x: R - 3.6, y: 0.58, w: 3.6, h: 0.2,
  fontFace: F.jp, fontSize: 8.5, color: C.muted, align: "right", margin: 0, valign: "middle",
});

// 前提チップ
const chips = [
  ["区画", "30坪"],
  ["契約形態", "業務委託型・3年"],
  ["御社の初期費用（設置工事費）", "30〜50万円"],
  ["オープン目標", "2027年2月末"],
  ["閉店後の工期", "約1か月"],
];
const chw = (CW - 0.14 * 4) / 5;
chips.forEach((c, i) => {
  const x = M + i * (chw + 0.14);
  s.addShape("roundRect", { x, y: 1.0, w: chw, h: 0.42, rectRadius: 0.05, fill: { color: C.tint }, line: { color: C.tintLine, width: 0.75 } });
  s.addText(c[0], {
    x: x + 0.14, y: 1.03, w: chw - 0.28, h: 0.17,
    fontFace: F.jp, fontSize: 7, color: C.muted, margin: 0, valign: "middle",
  });
  s.addText(c[1], {
    x: x + 0.14, y: 1.19, w: chw - 0.28, h: 0.2,
    fontFace: numFace(c[1]), fontSize: 10.5, bold: true, color: C.green, margin: 0, valign: "middle",
  });
});

/* ---------------- スケジュール表 ---------------- */
const th = (t, al) => ({
  text: t,
  options: { fill: { color: C.greenDeep }, color: C.white, bold: true, fontSize: 9, fontFace: F.jp, align: al || "left", valign: "middle", margin: [0.02, 0.1, 0.02, 0.1] },
});
const td = (t, o = {}) => ({
  text: t,
  options: {
    fill: { color: o.fill || C.white }, color: o.color || C.body, bold: !!o.bold,
    fontSize: o.size || 8.5, fontFace: o.face || F.jp, align: o.align || "left",
    valign: "middle", margin: [0.02, 0.1, 0.02, 0.1],
  },
});
const SC = { fill: C.tint, color: C.green, bold: true, align: "center", size: 8 };
const MY = { fill: C.goldTint, color: C.gold, bold: true, align: "center", size: 8 };
const BO = { fill: C.grayBand, color: C.ink, bold: true, align: "center", size: 8 };

s.addTable(
  [
    [th("時期"), th("内容"), th("担当", "center")],
    [td("11月上旬", { bold: true, color: C.ink }), td("施工担当による現地内見（実測・電源・給排水・既存ネット回線・残置設備）"), td("両社", BO)],
    [td("11月中旬", { bold: true, color: C.ink }), td("レイアウト図・什器リストのご提出／設置工事費のご案内（30〜50万円）"), td("セルフカフェ", SC)],
    [td("11月中旬", { bold: true, color: C.ink }), td("クリエ様の解約条件・原状回復範囲の確定（残置の範囲を確定）"), td("未来屋書店様", MY)],
    [td("11月下旬", { bold: true, color: C.ink }), td("仕様の確定／業務委託契約の締結（カメラ映像の取扱いを含む）"), td("両社", BO)],
    [td("11月下旬〜12月", { bold: true, color: C.ink }), td("館（イオンモール様）への工事申請"), td("両社", BO)],
    [td("12月上旬", { bold: true, color: C.gold }), td("ネット回線の新規お申込み（開通まで1〜2か月）", { bold: true, color: C.ink, fill: C.goldTint }), td("未来屋書店様", MY)],
    [td("12月上旬", { bold: true, color: C.gold }), td("什器・機器の発注（リードタイム4〜6週間）", { bold: true, color: C.ink, fill: C.goldTint }), td("セルフカフェ", SC)],
    [td("1月末", { bold: true, color: C.ink }), td("カフェ・ド・クリエ 閉店・退去（造作・床・設備は残置）"), td("未来屋書店様", MY)],
    [td("2月（約1か月）", { bold: true, color: C.gold }), td("内装工事 → 機器設置・試運転 → 什器搬入／回線開通の確認", { bold: true, color: C.ink, fill: C.goldTint }), td("セルフカフェ", SC)],
    [td("2月下旬", { bold: true, color: C.ink }), td("オープン準備・研修"), td("両社", BO)],
    [td("2月末", { bold: true, color: C.gold, fill: C.goldTint }), td("オープン", { bold: true, color: C.gold, fill: C.goldTint, size: 10 }), td("—", { fill: C.goldTint, align: "center" })],
  ],
  {
    x: M, y: 1.56, w: CW, colW: [1.5, 7.74, 1.61],
    rowH: [0.25, 0.245, 0.245, 0.245, 0.245, 0.245, 0.245, 0.245, 0.245, 0.245, 0.245, 0.29],
    border: { type: "solid", color: C.warmLine, pt: 0.75 },
    autoPage: false,
  }
);

/* ---------------- 前提のアラート ---------------- */
const ay = 4.66;
s.addShape("roundRect", { x: M, y: ay, w: CW, h: 0.7, rectRadius: 0.05, fill: { color: C.greenDeep }, line: { type: "none" } });
s.addText(
  [
    { text: "工期約1か月は造作・床・給排水・電気が残置される前提です。ネット回線は12月中のお申込みが必要です。", options: { fontFace: F.jp, fontSize: 9.5, bold: true, color: C.white, breakLine: true } },
    { text: "スケルトン戻しとなる場合は工期が延び、2月末オープンが厳しくなります。回線は申込から開通まで1〜2か月かかります。また1/31閉店→2/28オープンは正味28日のためバッファがありません。閉店を1月中旬〜下旬に前倒しできると安全です。", options: { fontFace: F.jp, fontSize: 8.5, color: "AFC2B8" } },
  ],
  { x: M + 0.26, y: ay, w: CW - 0.52, h: 0.7, margin: 0, valign: "middle", lineSpacingMultiple: 1.2 }
);

/* ---------------- 下段：役割分担 ---------------- */
const by = 5.46, bh = 2.42;
const lw = 4.05, rw = CW - lw - 0.24, rx = M + lw + 0.24;

// 左：セルフカフェ
s.addShape("roundRect", { x: M, y: by, w: lw, h: bh, rectRadius: 0.06, fill: { color: C.tint }, line: { color: C.tintLine, width: 0.75 } });
s.addText("セルフカフェ（弊社）で進めること", {
  x: M + 0.22, y: by + 0.12, w: lw - 0.44, h: 0.24,
  fontFace: F.jp, fontSize: 10, bold: true, color: C.green, margin: 0, valign: "middle",
});
s.addShape("rect", { x: M + 0.22, y: by + 0.38, w: lw - 0.44, h: 0.01, fill: { color: C.tintLine } });
[
  "施工担当の内見日程の調整（11月上旬）",
  "レイアウト図・什器リストの作成（30坪）",
  "設置工事費のご案内（30〜50万円）",
  "業務委託契約書のドラフト送付",
  "カメラ映像の取扱いに関する取り決めのご提示",
  "館への工事申請用 図面・仕様書の作成",
  "什器・機器の発注（12月上旬）",
  "内装工事・機器設置・試運転",
  "開業スケジュール確定版の提示（内見後）",
  "オープン後の集客・お客様対応",
].forEach((t, i) => {
  const y = by + 0.50 + i * 0.185;
  s.addText("・", { x: M + 0.24, y, w: 0.16, h: 0.185, fontFace: F.jp, fontSize: 8.5, color: C.green, margin: 0, valign: "middle" });
  s.addText(t, { x: M + 0.42, y, w: lw - 0.64, h: 0.185, fontFace: F.jp, fontSize: 8.5, color: C.ink, margin: 0, valign: "middle" });
});

// 右：未来屋書店様
s.addShape("roundRect", { x: rx, y: by, w: rw, h: bh, rectRadius: 0.06, fill: { color: C.white }, line: { color: C.goldLine, width: 0.75 } });
s.addText("未来屋書店様にご確認・ご準備いただきたいこと", {
  x: rx + 0.22, y: by + 0.12, w: rw - 0.44, h: 0.24,
  fontFace: F.jp, fontSize: 10, bold: true, color: C.gold, margin: 0, valign: "middle",
});
s.addShape("rect", { x: rx + 0.22, y: by + 0.38, w: rw - 0.44, h: 0.01, fill: { color: C.warmLine } });

const sub = (rw - 0.44 - 0.3) / 2;
s.addText("ご確認事項（★＝工期・コストに直結）", {
  x: rx + 0.22, y: by + 0.44, w: sub, h: 0.18,
  fontFace: F.jp, fontSize: 8, bold: true, color: C.muted, margin: 0, valign: "middle",
});
[
  ["★", "クリエ様の解約条件・閉店日の確定"],
  ["★", "原状回復の範囲（造作・床・設備の残置）"],
  ["★", "残置する既存設備（給排水・電気容量）"],
  ["★", "ネット回線の新規敷設（12月中に申込）"],
  ["", "開業日の確定（契約期間の起算日）"],
  ["", "カメラ映像（個人情報）の取扱いの取決め"],
  ["", "館への工事申請（様式・期限・搬入ルート）"],
  ["", "サイン表記の確定（MIRAIYA Book & Cafe）"],
  ["", "運営体制（営業時間・清掃補充の担当分け）"],
  ["", "社内稟議の進捗・決裁完了の目処"],
].forEach((t, i) => {
  const y = by + 0.62 + i * 0.17;
  s.addText(t[0] || "・", {
    x: rx + 0.24, y, w: 0.18, h: 0.17,
    fontFace: F.jp, fontSize: t[0] ? 7 : 8, bold: !!t[0], color: t[0] ? C.gold : C.muted, margin: 0, valign: "middle",
  });
  s.addText(t[1], {
    x: rx + 0.44, y, w: sub - 0.2, h: 0.17,
    fontFace: F.jp, fontSize: 8, bold: !!t[0], color: C.ink, margin: 0, valign: "middle",
  });
});

const sx = rx + 0.22 + sub + 0.3;
s.addText("ご提供いただきたい資料", {
  x: sx, y: by + 0.44, w: sub, h: 0.18,
  fontFace: F.jp, fontSize: 8, bold: true, color: C.muted, margin: 0, valign: "middle",
});
[
  "区画図面（平面図・展開図／CADデータ）",
  "電気容量・分電盤の位置",
  "給排水の位置・系統",
  "既存ネット回線の有無・引込経路",
  "現況写真（既存設備・厨房まわり）",
  "館の工事規定・申請様式",
  "クリエ様の原状回復条件がわかる資料",
  "ご契約者情報（会社名・所在地・代表者名）",
  "業務委託料のお振込先口座",
].forEach((t, i) => {
  const y = by + 0.62 + i * 0.17;
  s.addText("・", { x: sx + 0.02, y, w: 0.16, h: 0.17, fontFace: F.jp, fontSize: 8, color: C.muted, margin: 0, valign: "middle" });
  s.addText(t, { x: sx + 0.2, y, w: sub - 0.2, h: 0.17, fontFace: F.jp, fontSize: 8, color: C.ink, margin: 0, valign: "middle" });
});

/* ---------------- フッター ---------------- */
s.addText("※ 業務委託型のため、工事費のお見積り提出はございません。御社にご負担いただくのは設置工事費（30〜50万円・税抜）のみです。", {
  x: M, y: 7.94, w: CW, h: 0.15,
  fontFace: F.jp, fontSize: 7.5, color: C.gold, margin: 0, valign: "middle",
});
s.addText("※ 業務委託契約書は別途ご送付済みです。別表は仕様確定後に作成いたします。工期・納期は目安で、館の工事申請期間は未織り込みです。", {
  x: M, y: 8.08, w: 7.0, h: 0.15,
  fontFace: F.jp, fontSize: 7.5, color: C.footer, margin: 0, valign: "middle",
});
s.addText("セルフカフェ株式会社　052-879-3557 ／ info@selfcafe.jp", {
  x: R - 3.6, y: 8.08, w: 3.6, h: 0.15,
  fontFace: F.jp, fontSize: 7.5, color: C.footer, align: "right", margin: 0, valign: "middle",
});

pres.writeFile({ fileName: path.join(__dirname, "akita-schedule-sheet.pptx") }).then((f) => console.log("written:", f));
