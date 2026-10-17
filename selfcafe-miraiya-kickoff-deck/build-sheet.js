/**
 * 未来屋書店 秋田店 進行スケジュールおよびご確認事項 — A4横1枚（フォーマル書式）
 * 前提：30坪／業務委託型3年／残置／2027年3月上旬 開業
 */
const pptxgen = require("pptxgenjs");
const path = require("path");

const C = {
  bg: "FFFFFF", ink: "1A1A1A", green: "106838", greenDeep: "0C5230",
  gold: "8A6316", rule: "333333", line: "BFBFBF", lightLine: "DCDCDC",
  head: "EDEDED", band: "F7F7F7", body: "333333", muted: "6B6B6B", white: "FFFFFF",
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
pres.title = "未来屋書店 秋田店 進行スケジュールおよびご確認事項";

const s = pres.addSlide();
s.background = { color: C.bg };

/* ================= 発信ヘッダー ================= */
s.addText("2026年10月17日", {
  x: R - 3.0, y: 0.30, w: 3.0, h: 0.18,
  fontFace: F.jp, fontSize: 9, color: C.body, align: "right", margin: 0, valign: "middle",
});
s.addText("株式会社未来屋書店　御中", {
  x: M, y: 0.50, w: 5.0, h: 0.24,
  fontFace: F.jp, fontSize: 11.5, color: C.ink, margin: 0, valign: "middle",
});
s.addImage({ path: path.join(A, "logo-green.png"), x: R - 1.45, y: 0.48, w: 1.45, h: 0.207 });
s.addText("セルフカフェ株式会社", {
  x: R - 4.6, y: 0.70, w: 4.6, h: 0.18,
  fontFace: F.jp, fontSize: 9.5, color: C.ink, align: "right", margin: 0, valign: "middle",
});
s.addText("愛知県名古屋市緑区鴻仏目1-102ベルインビルB棟3階　TEL 052-879-3557", {
  x: R - 5.4, y: 0.87, w: 5.4, h: 0.16,
  fontFace: F.jp, fontSize: 7.5, color: C.muted, align: "right", margin: 0, valign: "middle",
});

/* ================= 表題 ================= */
s.addShape("rect", { x: M, y: 1.10, w: CW, h: 0.016, fill: { color: C.rule } });
s.addText("秋田店 カフェ区画 業態転換　進行スケジュールおよびご確認事項", {
  x: M, y: 1.15, w: CW, h: 0.32,
  fontFace: F.jp, fontSize: 15, bold: true, color: C.ink, align: "center", margin: 0, valign: "middle",
});
s.addShape("rect", { x: M, y: 1.50, w: CW, h: 0.008, fill: { color: C.line } });

s.addText("平素より格別のお引き立てを賜り、厚く御礼申し上げます。標記の件につきまして、現時点における進行スケジュールおよびご確認事項を下記のとおりご案内申し上げます。", {
  x: M, y: 1.56, w: CW, h: 0.18,
  fontFace: F.jp, fontSize: 9, color: C.body, margin: 0, valign: "middle",
});

/* ================= 共通スタイル ================= */
const sec = (y, t) => {
  s.addShape("rect", { x: M, y: y + 0.03, w: 0.055, h: 0.16, fill: { color: C.greenDeep } });
  s.addText(t, {
    x: M + 0.14, y, w: CW - 0.14, h: 0.22,
    fontFace: F.jp, fontSize: 10, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
};
const th = (t, al) => ({
  text: t,
  options: { fill: { color: C.head }, color: C.ink, bold: true, fontSize: 8.5, fontFace: F.jp, align: al || "left", valign: "middle", margin: [0.02, 0.1, 0.02, 0.1] },
});
const td = (t, o = {}) => ({
  text: t,
  options: {
    fill: { color: o.fill || C.white }, color: o.color || C.body, bold: !!o.bold,
    fontSize: o.size || 8.5, fontFace: o.face || numFace(t), align: o.align || "left",
    valign: "middle", margin: [0.02, 0.1, 0.02, 0.1],
  },
});

/* ================= 1．前提条件 ================= */
sec(1.82, "１．前提条件");
const pw = [1.7, 2.3, 2.75, 2.1, 2.0];
s.addTable(
  [
    [th("区画", "center"), th("契約形態", "center"), th("貴社ご負担の初期費用", "center"), th("開業予定", "center"), th("閉店後の工期", "center")],
    [
      td("30坪", { align: "center", bold: true, color: C.ink, size: 9.5 }),
      td("業務委託型・3年", { align: "center", bold: true, color: C.ink, size: 9.5 }),
      td("30〜50万円（税抜・設置工事費）", { align: "center", bold: true, color: C.ink, size: 9.5 }),
      td("2027年3月上旬", { align: "center", bold: true, color: C.green, size: 9.5 }),
      td("約1か月（残置前提）", { align: "center", bold: true, color: C.ink, size: 9.5 }),
    ],
  ],
  { x: M, y: 2.04, w: CW, colW: pw, rowH: [0.22, 0.28], border: { type: "solid", color: C.line, pt: 0.75 }, autoPage: false }
);

/* ================= 2．進行スケジュール ================= */
sec(2.64, "２．進行スケジュール");
const who = (t) => {
  const col = t === "弊社" ? C.green : t === "貴社" ? C.gold : C.muted;
  return td(t, { align: "center", bold: true, color: col, size: 8.5 });
};
const rows = [
  ["11月上旬", "現地内見（実測、電源・給排水、既存回線、残置設備のご確認）", "両社"],
  ["11月中旬", "レイアウト図・什器リストのご提出、設置工事費のご案内", "弊社"],
  ["11月中旬", "カフェ・ド・クリエ様の解約条件および原状回復範囲の確定", "貴社"],
  ["11月下旬", "仕様の確定および業務委託契約の締結（カメラ映像の取扱いを含む）", "両社"],
  ["11月下旬〜12月", "館（イオンモール様）への工事申請", "両社"],
  ["12月上旬", "インターネット回線の新規お申込み（開通まで1〜2か月）", "貴社"],
  ["12月上旬", "什器・機器の発注（リードタイム4〜6週間）", "弊社"],
  ["1月末", "カフェ・ド・クリエ様 閉店・退去（造作・床・設備は残置）", "貴社"],
  ["2月（約1か月）", "内装工事、機器設置・試運転、什器搬入、回線開通の確認", "弊社"],
  ["2月下旬", "開業準備・研修", "両社"],
];
const body = rows.map((r, i) => {
  const f = i % 2 ? C.band : C.white;
  return [td(r[0], { bold: true, color: C.ink, fill: f }), td(r[1], { fill: f }), who(r[2])];
});
s.addTable(
  [
    [th("時期"), th("内容"), th("担当", "center")],
    ...body,
    [
      td("3月上旬", { bold: true, color: C.green, fill: "EFF5F1", size: 9.5 }),
      td("開業", { bold: true, color: C.green, fill: "EFF5F1", size: 9.5 }),
      td("—", { align: "center", color: C.muted, fill: "EFF5F1" }),
    ],
  ],
  {
    x: M, y: 2.86, w: CW, colW: [1.45, 7.8, 1.6],
    rowH: [0.21, 0.215, 0.215, 0.215, 0.215, 0.215, 0.215, 0.215, 0.215, 0.215, 0.215, 0.25],
    border: { type: "solid", color: C.line, pt: 0.75 },
    autoPage: false,
  }
);

/* ================= 3〜5．役割分担 ================= */
const colW3 = (CW - 0.6) / 3, cy = 5.58;
const cols = [
  {
    t: "３．弊社にて対応いたします事項",
    items: [
      "現地内見の日程調整（11月上旬）",
      "レイアウト図および什器リストの作成（30坪）",
      "設置工事費のご案内（30〜50万円・税抜）",
      "業務委託契約書のドラフト送付",
      "カメラ映像の取扱いに関する取り決めのご提示",
      "館への工事申請用 図面・仕様書の作成",
      "什器・機器の発注（12月上旬）",
      "内装工事、機器設置および試運転",
      "開業スケジュール確定版のご提示（内見後）",
      "開業後の集客およびお客様対応",
    ],
  },
  {
    t: "４．貴社にてご確認・ご決定をお願いする事項",
    note: "※印は工期に直結いたします",
    items: [
      "※ カフェ・ド・クリエ様の解約条件・閉店日",
      "※ 原状回復の範囲（造作・床・設備の残置）",
      "※ 残置いただく既存設備（給排水・電気容量）",
      "※ インターネット回線の新規敷設（12月中）",
      "開業日の確定（契約期間の起算日となります）",
      "カメラ映像（個人情報）の取扱いの取り決め",
      "館への工事申請（様式・期限・搬入ルート）",
      "サイン表記の確定（MIRAIYA Book & Cafe）",
      "運営体制（営業時間、清掃・補充のご担当）",
      "社内稟議の進捗および決裁完了の目処",
    ],
  },
  {
    t: "５．ご提供をお願いする資料・情報",
    items: [
      "区画図面（平面図・展開図／CADデータ）",
      "電気容量および分電盤の位置",
      "給排水の位置・系統",
      "既存インターネット回線の有無・引込経路",
      "現況写真（既存設備、厨房まわり）",
      "館の工事規定および申請様式",
      "カフェ・ド・クリエ様の原状回復条件がわかる資料",
      "ご契約者情報（会社名・所在地・代表者名）",
      "業務委託料のお振込先口座",
    ],
  },
];
cols.forEach((c, ci) => {
  const x = M + ci * (colW3 + 0.3);
  s.addShape("rect", { x, y: cy + 0.03, w: 0.055, h: 0.16, fill: { color: C.greenDeep } });
  s.addText(c.t, {
    x: x + 0.14, y: cy, w: colW3 - 0.14, h: 0.22,
    fontFace: F.jp, fontSize: 9.5, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
  s.addShape("rect", { x, y: cy + 0.24, w: colW3, h: 0.008, fill: { color: C.lightLine } });
  if (c.note) {
    s.addText(c.note, {
      x, y: cy + 0.28, w: colW3, h: 0.14,
      fontFace: F.jp, fontSize: 7, color: C.gold, margin: 0, valign: "middle",
    });
  }
  const y0 = cy + 0.44;
  c.items.forEach((t, i) => {
    const y = y0 + i * 0.155;
    const star = t.startsWith("※");
    s.addText(star ? "" : "・", {
      x: x + 0.02, y, w: 0.14, h: 0.155,
      fontFace: F.jp, fontSize: 8, color: C.muted, margin: 0, valign: "middle",
    });
    s.addText(t, {
      x: x + (star ? 0.02 : 0.16), y, w: colW3 - (star ? 0.02 : 0.16), h: 0.155,
      fontFace: F.jp, fontSize: 8, bold: star, color: star ? C.gold : C.body, margin: 0, valign: "middle",
    });
  });
});

/* ================= 備考 ================= */
s.addShape("rect", { x: M, y: 7.62, w: CW, h: 0.008, fill: { color: C.lightLine } });
[
  "※1　工期約1か月は、造作・床・給排水・電気を残置いただくことを前提としております。スケルトン戻しとなる場合は工期が延長いたします。",
  "※2　業務委託型のため、工事費のお見積りのご提出はございません。貴社にご負担いただく費用は、設置工事費（30〜50万円・税抜）のみでございます。",
  "※3　業務委託契約書は別途ご送付済みです。別表は仕様確定後に作成のうえ添付いたします。館の工事申請に要する期間は本表に織り込んでおりません。",
].forEach((t, i) => {
  s.addText(t, {
    x: M, y: 7.70 + i * 0.145, w: CW, h: 0.14,
    fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
  });
});

pres.writeFile({ fileName: path.join(__dirname, "akita-schedule-sheet.pptx") }).then((f) => console.log("written:", f));
