/**
 * セルフカフェ 業務委託型FC資料 — 全25ページ 再設計版
 *
 * 元資料（Googleスライド）は表・グラフ・図解が画像貼付だったため、
 * すべてPowerPointネイティブの図形・表・グラフ・アイコンで組み直したもの。
 * 写真は「写真枠」（点線プレースホルダ）として用意してある。
 *
 * 実行前に `node build-icons.js` でアイコンPNGを生成しておくこと。
 */
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

// ---------------------------------------------------------------- palette
const C = {
  bg: "FFFCF5", // 既存資料の地色
  green: "1B5E20", // ブランド緑（既存タイトル色）
  greenDeep: "12401A", // 濃色パネル
  greenMid: "2E7D32",
  greenTint: "F0F5EF",
  greenPale: "A9C6AC", // 濃色パネル上のサブテキスト
  brown: "5D4037",
  brownLt: "8D6E63", // 既存ヘッダー罫線色
  ink: "3B3833",
  inkSub: "6E6A61",
  muted: "9A948A",
  cream: "FAF3E3",
  creamDk: "E9DEC4",
  line: "E6DFCD",
  white: "FFFFFF",
  gold: "B8893B",
};

const F = { jp: "Noto Sans JP", num: "Arial" };

// ---------------------------------------------------------------- geometry
const M = 0.729; // 左右マージン（元資料 52.5pt）
const CW = 11.875; // コンテンツ幅（元資料 52.5〜907.5pt）
const RIGHT = M + CW; // 12.604
const SW = 13.333;
const SH = 7.5;

const ASSETS = path.join(__dirname, "assets");
const ICONS = path.join(ASSETS, "icons");
const shadow = () => ({ type: "outer", color: "6B6250", blur: 9, offset: 1.4, angle: 90, opacity: 0.14 });

// ---------------------------------------------------------------- helpers
function shell(slide, title, pageNo) {
  slide.background = { color: C.bg };
  slide.addText(title, {
    x: M, y: 0.554, w: 9.3, h: 0.42,
    fontFace: F.jp, fontSize: 30, color: C.green,
    align: "left", valign: "middle", margin: 0,
  });
  // 既存全ページ共通のヘッダー罫線（52.5→907.5pt, y=87.8pt, 1.5pt）
  slide.addShape("rect", { x: M, y: 1.219, w: CW, h: 0.021, fill: { color: C.brownLt } });
  slide.addImage({ path: path.join(ASSETS, "logo_gray.png"), x: 10.297, y: 0.878, w: 2.214, h: 0.315 });
  slide.addText(String(pageNo), {
    x: RIGHT - 1.0, y: 6.94, w: 1.0, h: 0.24,
    fontFace: "Cambria", fontSize: 13, color: "000000", align: "right", valign: "middle", margin: 0,
  });
}

function kicker(slide, text, y) {
  slide.addText(text, {
    x: M, y: y, w: CW, h: 0.28,
    fontFace: F.jp, fontSize: 11.5, bold: true, color: C.brown, margin: 0, valign: "middle",
  });
}

function card(slide, x, y, w, h, opts = {}) {
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.07,
    fill: { color: opts.fill || C.white },
    line: { color: opts.line || C.line, width: 0.75 },
    shadow: opts.noShadow ? undefined : shadow(),
  });
}

function panel(slide, x, y, w, h) {
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.09,
    fill: { color: C.greenDeep }, line: { type: "none" }, shadow: shadow(),
  });
}

/** アイコン画像を置く。name は Lucide 名、tone は build-icons.js の色キー */
function icon(slide, name, tone, x, y, size) {
  slide.addImage({ path: path.join(ICONS, `${name}-${tone}.png`), x, y, w: size, h: size });
}

/** 薄緑の丸い座布団つきアイコン */
function iconBadge(slide, name, x, y, opts = {}) {
  const d = opts.d || 0.62;
  slide.addShape("ellipse", {
    x, y, w: d, h: d,
    fill: { color: opts.fill || C.greenTint }, line: { type: "none" },
  });
  const s = d * 0.56;
  icon(slide, name, opts.tone || "green", x + (d - s) / 2, y + (d - s) / 2, s);
}

/** 連番チップ */
function chip(slide, x, y, label, opts = {}) {
  const size = opts.size || 0.42;
  slide.addShape("roundRect", {
    x, y, w: size, h: size, rectRadius: 0.09,
    fill: { color: opts.fill || C.green }, line: { type: "none" },
  });
  slide.addText(label, {
    x, y, w: size, h: size,
    fontFace: opts.face || F.num, fontSize: opts.fontSize || 13, bold: true,
    color: opts.color || C.white, align: "center", valign: "middle", margin: 0,
  });
}

/** 写真枠。opts.img に assets/photos/ 内のファイル名を渡すと、
 *  ファイルが存在する場合は写真を挿入し、無ければ点線プレースホルダを描く。 */
const PHOTOS = path.join(ASSETS, "photos");
function photoSlot(slide, x, y, w, h, caption, opts = {}) {
  const img = opts.img ? path.join(PHOTOS, opts.img) : null;
  if (img && fs.existsSync(img)) {
    slide.addImage({ path: img, x, y, w, h, sizing: { type: "cover", w, h } });
    slide.addShape("rect", { x, y, w, h, fill: { type: "none" }, line: { color: "D8CFBB", width: 1 } });
    if (caption) {
      const ch = 0.36;
      const cw2 = Math.min(w - 0.24, 0.4 + caption.length * ((opts.captionSize || 10) / 72) * 1.06);
      slide.addShape("roundRect", {
        x: x + 0.12, y: y + h - ch - 0.12, w: cw2, h: ch, rectRadius: 0.05,
        fill: { color: C.white, transparency: 10 }, line: { type: "none" },
      });
      slide.addText(caption, {
        x: x + 0.28, y: y + h - ch - 0.12, w: cw2 - 0.28, h: ch,
        fontFace: F.jp, fontSize: opts.captionSize || 10, bold: true, color: C.ink,
        margin: 0, valign: "middle",
      });
    }
    return;
  }
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.06,
    fill: { color: opts.fill || C.cream },
    line: { color: opts.line || "CFC4A8", width: 1, dashType: "dash" },
  });
  const iconSize = Math.min(0.62, h * 0.24);
  const cy = caption ? y + h / 2 - 0.34 : y + h / 2 - iconSize / 2;
  icon(slide, "LuImage", "muted", x + w / 2 - iconSize / 2, cy, iconSize);
  slide.addText("写真を挿入", {
    x: x + 0.10, y: cy + iconSize + 0.04, w: w - 0.20, h: 0.24,
    fontFace: F.jp, fontSize: 9.5, color: C.muted, align: "center", valign: "middle", margin: 0,
  });
  if (caption) {
    slide.addText(caption, {
      x: x + 0.10, y: cy + iconSize + 0.28, w: w - 0.20, h: 0.46,
      fontFace: F.jp, fontSize: opts.captionSize || 10.5, bold: true,
      color: opts.captionColor || C.brown,
      align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.15,
    });
  }
}

/** 数値＋単位のスタットタイル */
function statTile(slide, x, y, w, h, value, unit, label, opts = {}) {
  const dark = !!opts.dark;
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.07,
    fill: { color: dark ? C.greenDeep : C.white },
    line: dark ? { type: "none" } : { color: C.line, width: 0.75 },
    shadow: shadow(),
  });
  slide.addText(
    [
      { text: value, options: { fontFace: F.num, fontSize: opts.valueSize || 30, bold: true, color: dark ? C.white : C.green } },
      { text: unit ? " " + unit : "", options: { fontFace: F.jp, fontSize: 12, bold: true, color: dark ? C.greenPale : C.brown } },
    ],
    { x: x + 0.24, y: y + 0.14, w: w - 0.48, h: h * 0.5, align: "left", valign: "middle", margin: 0 }
  );
  slide.addText(label, {
    x: x + 0.24, y: y + h * 0.55, w: w - 0.48, h: h * 0.36,
    fontFace: F.jp, fontSize: 10.5, color: dark ? C.greenPale : C.inkSub,
    align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.15,
  });
}

// ================================================================ deck
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
pres.author = "セルフカフェ株式会社";
pres.title = "セルフカフェ 業務委託型制度";

/* =============================================================== p.1 表紙 */
{
  const s = pres.addSlide();
  s.background = { color: C.greenDeep };

  s.addImage({ path: path.join(ASSETS, "logo_white.png"), x: 0.90, y: 0.72, w: 2.66, h: 0.378 });

  s.addText("業務委託型制度", {
    x: 0.90, y: 1.62, w: 9.0, h: 1.02,
    fontFace: F.jp, fontSize: 52, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  s.addText("PARTNER PROGRAM", {
    x: 0.90, y: 2.68, w: 9.0, h: 0.30,
    fontFace: F.num, fontSize: 12, bold: true, color: C.greenPale, charSpacing: 4, margin: 0, valign: "middle",
  });
  s.addText("場所をご提供いただくだけ。運営は本部が担う、新しいパートナー制度。", {
    x: 0.90, y: 3.08, w: 9.4, h: 0.38,
    fontFace: F.jp, fontSize: 16, color: C.white, margin: 0, valign: "middle",
  });

  // 実績チップ
  const chips = [
    ["LuStore", "全国71店舗"],
    ["LuCoins", "初期費用 30万〜50万円"],
    ["LuTimer", "1日15分程度の運営"],
  ];
  let cx = 0.90;
  chips.forEach(([ic, label]) => {
    const w = 0.66 + label.length * 0.166;
    s.addShape("roundRect", {
      x: cx, y: 3.68, w, h: 0.50, rectRadius: 0.25,
      fill: { color: "1D5624" }, line: { color: "3E7A45", width: 0.75 },
    });
    icon(s, ic, "white", cx + 0.20, 3.68 + 0.135, 0.23);
    s.addText(label, {
      x: cx + 0.52, y: 3.68, w: w - 0.62, h: 0.50,
      fontFace: F.jp, fontSize: 11.5, bold: true, color: C.white, valign: "middle", margin: 0,
    });
    cx += w + 0.20;
  });

  // 下部：写真ストリップ（4枠）
  const bandY = 4.56, bandH = SH - bandY;
  const sw = (SW - 3 * 0.06) / 4;
  const caps = ["店舗外観", "店内（席・照明）", "ドリンクマシン", "利用中の店内"];
  caps.forEach((cap, i) => {
    const x = i * (sw + 0.06);
    s.addShape("rect", { x, y: bandY, w: sw, h: bandH, fill: { color: C.cream }, line: { type: "none" } });
    icon(s, "LuImage", "muted", x + sw / 2 - 0.30, bandY + 0.72, 0.60);
    s.addText("写真を挿入", {
      x: x + 0.10, y: bandY + 1.40, w: sw - 0.20, h: 0.24,
      fontFace: F.jp, fontSize: 10, color: C.muted, align: "center", valign: "middle", margin: 0,
    });
    s.addText(cap, {
      x: x + 0.10, y: bandY + 1.66, w: sw - 0.20, h: 0.30,
      fontFace: F.jp, fontSize: 11.5, bold: true, color: C.brown, align: "center", valign: "middle", margin: 0,
    });
  });
}

/* =============================================================== p.2 業務委託型とは */
{
  const s = pres.addSlide();
  shell(s, "業務委託型とは", 2);
  kicker(s, "通常のFC加盟とは異なり、貴社は場所をご提供いただくだけ／運営主体は本部", 1.36);

  const colW = [2.30, 4.55, 5.02];
  const tx = M, ty = 2.00;

  s.addShape("roundRect", {
    x: tx + colW[0] + colW[1] - 0.02, y: ty - 0.32, w: colW[2] + 0.04, h: 5.15,
    rectRadius: 0.07, fill: { color: C.greenTint }, line: { type: "none" },
  });
  icon(s, "LuBadgeCheck", "green", tx + colW[0] + colW[1] + 0.16, ty - 0.30, 0.24);
  s.addText("貴社にご提案するのはこちら", {
    x: tx + colW[0] + colW[1] + 0.46, y: ty - 0.32, w: colW[2] - 0.60, h: 0.28,
    fontFace: F.jp, fontSize: 10.5, bold: true, color: C.greenMid, margin: 0, valign: "middle",
  });

  const hdr = (t, fill, color) => ({
    text: t,
    options: { fill: { color: fill }, color, bold: true, fontSize: 12, fontFace: F.jp, align: "left", valign: "middle", margin: [0.06, 0.14, 0.06, 0.14] },
  });
  const lab = (t) => ({
    text: t,
    options: { fill: { color: "F6F2E6" }, color: C.brown, bold: true, fontSize: 11.5, fontFace: F.jp, valign: "middle", margin: [0.06, 0.14, 0.06, 0.14] },
  });
  const old = (t) => ({
    text: t,
    options: { fill: { color: C.white }, color: C.inkSub, fontSize: 10.5, fontFace: F.jp, valign: "middle", margin: [0.08, 0.16, 0.08, 0.16] },
  });
  const nw = (t) => ({
    text: t,
    options: { fill: { color: "F2F7F2" }, color: C.green, bold: true, fontSize: 11, fontFace: F.jp, valign: "middle", margin: [0.08, 0.16, 0.08, 0.16] },
  });

  s.addTable(
    [
      [hdr("項目", C.creamDk, C.brown), hdr("通常のFC加盟", "8A8272", C.white), hdr("業務委託型", C.green, C.white)],
      [lab("運営主体"), old("加盟者様が独立オーナーとして運営"), nw("本部が運営を主導、貴社は場所提供のみ")],
      [lab("初期費用"), old("加盟金100万円＋工事費用等\n（開業資金目安：居抜き700万円〜／スケルトン工事1,000万円〜）"), nw("30万〜50万円程度（機器設置費のみ）")],
      [lab("出店スペース"), old("独立した店舗区画（25坪〜）が必要"), nw("既存スペースの一角に間借り設置可能")],
      [lab("日常業務"), old("加盟者様が運営全般を担当"), nw("清掃・補充のみ、1日15分程度\n（既存スタッフで対応可）")],
      [lab("売上金の流れ"), old("本部が売上金を管理し、ロイヤリティなどの費用を控除した残額を加盟者様へ振込"), nw("本部が売上金を管理し、25%を業務委託料として貴社へお支払い")],
      [lab("契約・撤退"), old("独立出店が前提の長期契約"), nw("契約期間の縛りあり（最低3年〜）／本部審査あり。原状回復の負担も小さい")],
    ],
    {
      x: tx, y: ty, w: CW, colW,
      rowH: [0.50, 0.60, 0.86, 0.58, 0.72, 0.82, 0.77],
      border: { type: "solid", color: C.line, pt: 0.75 },
      autoPage: false,
    }
  );
}

/* =============================================================== p.3 業務委託型のメリット */
{
  const s = pres.addSlide();
  shell(s, "業務委託型のメリット", 3);
  kicker(s, "貴社にとってのメリット／既存事業の一角を、負担をかけずに収益資産へ", 1.36);

  const data = [
    ["LuCoins", "遊休スペースが収益化", "レジ横・待合スペースなどの", "デッドスペースが毎月の収益源に。",
      ["加盟金なし、機器設置費（30万〜50万円程度）のみで開始", "業務委託料として売上の25%を毎月お支払い（収益イメージは次ページ参照）"]],
    ["LuTimer", "運営負担はほぼゼロ", "日常業務は清掃・補充のみ、", "1日15分程度。",
      ["既存スタッフの稼働内で対応可能、新規採用は不要", "トラブル対応も24時間警備会社・遠隔監視で完結、現場対応はほぼ不要"]],
    ["LuMapPin", "新しい来店動機に", "Wi-Fi・電源完備のカフェが", "新しい接点に。",
      ["Googleマップ・カフェ検索アプリで「近くのカフェ」を探す人と出会える", "書店・美容室・整体院・病院の待合など、待ち時間が生まれる業態と特に好相性"]],
    ["LuShieldCheck", "低リスクで導入できる", "通常のFC加盟（700万円〜）と比べ、", "意思決定のハードルが低い。",
      ["独立出店ではないため原状回復の負担も小さい", "内装・システム導入まで本部がフルサポート、店舗運営のノウハウをそのまま活用"]],
  ];

  const cw = 5.77, ch = 2.42, gx = 0.325, gy = 0.26;
  data.forEach((d, i) => {
    const x = M + (i % 2) * (cw + gx);
    const y = 1.80 + Math.floor(i / 2) * (ch + gy);
    card(s, x, y, cw, ch);
    iconBadge(s, d[0], x + 0.30, y + 0.24, { d: 0.60 });
    s.addText(d[1], {
      x: x + 1.04, y: y + 0.24, w: cw - 1.34, h: 0.60,
      fontFace: F.jp, fontSize: 15.5, bold: true, color: C.green, valign: "middle", margin: 0,
    });
    s.addText(
      [
        { text: d[2], options: { color: C.inkSub } },
        { text: d[3], options: { color: C.green, bold: true } },
      ],
      { x: x + 0.30, y: y + 0.92, w: cw - 0.60, h: 0.50, fontFace: F.jp, fontSize: 11, margin: 0, valign: "top", lineSpacingMultiple: 1.2 }
    );
    s.addShape("rect", { x: x + 0.30, y: y + 1.50, w: cw - 0.60, h: 0.011, fill: { color: C.line } });
    s.addText(
      d[4].map((t, j) => ({ text: t, options: { bullet: { characterCode: "25AA", indent: 11 }, breakLine: j < d[4].length - 1 } })),
      { x: x + 0.30, y: y + 1.62, w: cw - 0.60, h: 0.68, fontFace: F.jp, fontSize: 10, color: C.ink, margin: 0, valign: "top", paraSpaceAfter: 5 }
    );
  });
}

/* =============================================================== p.4 会社概要 */
{
  const s = pres.addSlide();
  shell(s, "会社概要", 4);

  const px = M, py = 1.55, pw = 4.55, ph = 5.30;
  panel(s, px, py, pw, ph);
  s.addImage({ path: path.join(ASSETS, "logo_white.png"), x: px + 0.52, y: py + 0.56, w: 3.51, h: 0.50 });
  s.addText("セルフカフェ株式会社", {
    x: px + 0.52, y: py + 1.32, w: pw - 1.04, h: 0.42,
    fontFace: F.jp, fontSize: 19, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  s.addText("SELF CAFE Inc.", {
    x: px + 0.52, y: py + 1.76, w: pw - 1.04, h: 0.30,
    fontFace: F.num, fontSize: 11.5, color: C.greenPale, charSpacing: 2, margin: 0, valign: "middle",
  });
  s.addShape("rect", { x: px + 0.52, y: py + 2.20, w: 1.10, h: 0.021, fill: { color: C.greenPale } });
  s.addText("Wi-Fi・電源完備の無人カフェを企画・開発・運営しています。", {
    x: px + 0.52, y: py + 2.40, w: pw - 1.04, h: 0.72,
    fontFace: F.jp, fontSize: 12, color: C.greenPale, margin: 0, valign: "top", lineSpacingMultiple: 1.35,
  });
  photoSlot(s, px + 0.52, py + 3.26, pw - 1.04, 1.62, "セルフカフェ 盛岡駅前店", {
    fill: "1D5624", line: "3E7A45", captionSize: 9.5, captionColor: "C6DCC8", img: "p02-storefront.jpg",
  });

  const lx = 5.75, lw = RIGHT - lx;
  const items = [
    ["会社名", "セルフカフェ株式会社"],
    ["代表取締役", "鈴木 大基"],
    ["設立日", "2024年5月1日"],
    ["資本金", "2,000万円"],
    ["事業内容", "無人カフェの経営"],
    ["公式HP", "https://selfcafe.jp/"],
  ];
  const rh = 0.74;
  items.forEach((it, i) => {
    const y = 1.60 + i * rh;
    s.addText(it[0], {
      x: lx, y: y, w: 1.55, h: rh - 0.08,
      fontFace: F.jp, fontSize: 11.5, bold: true, color: C.brown, valign: "middle", margin: 0,
    });
    s.addText(it[1], {
      x: lx + 1.62, y: y, w: lw - 1.62, h: rh - 0.08,
      fontFace: F.jp, fontSize: 14.5, color: C.ink, valign: "middle", margin: 0,
    });
    s.addShape("rect", { x: lx, y: y + rh - 0.06, w: lw, h: 0.011, fill: { color: C.line } });
  });

  s.addShape("roundRect", { x: lx, y: 6.20, w: lw, h: 0.52, rectRadius: 0.06, fill: { color: C.cream }, line: { type: "none" } });
  s.addText("※2024年5月1日よりウッドデザインパーク(株)から分社化", {
    x: lx + 0.22, y: 6.20, w: lw - 0.44, h: 0.52,
    fontFace: F.jp, fontSize: 11, color: C.brown, valign: "middle", margin: 0,
  });
}

/* =============================================================== p.5 社会背景と市場性 */
{
  const s = pres.addSlide();
  shell(s, "社会背景と市場性", 5);
  kicker(s, "現代社会が抱える3つの課題", 1.36);

  const items = [
    ["LuUsers", "深刻な労働力不足", "サービス業の人手確保が困難に", "少子高齢化と働き手の価値観の変化でサービス業の人手確保が難しくなり、人件費が上昇して経営負担が増している。"],
    ["LuLaptop", "サードプレイスの枯渇", "自宅以外の作業場所が足りない", "リモートワークの普及で自宅以外の手頃な作業場所が不足し、カフェやコワーキングは価格や混雑の課題が残り、学習スペースも制限が多く気軽に利用できない。"],
    ["LuBuilding2", "遊休資産の有効活用", "空室・デッドスペースが埋まらない", "ビル空室やデッドスペースが増え、従来のテナント誘致だけでは埋まらず、新たな活用モデルの開発が求められている。"],
  ];
  const cw = (CW - 2 * 0.28) / 3, ch = 3.10;
  items.forEach((it, i) => {
    const x = M + i * (cw + 0.28), y = 1.86;
    card(s, x, y, cw, ch);
    iconBadge(s, it[0], x + 0.30, y + 0.28, { d: 0.66 });
    s.addText(String(i + 1).padStart(2, "0"), {
      x: x + cw - 0.90, y: y + 0.28, w: 0.60, h: 0.40,
      fontFace: F.num, fontSize: 20, bold: true, color: C.creamDk, align: "right", valign: "middle", margin: 0,
    });
    s.addText(it[1], {
      x: x + 0.30, y: y + 1.06, w: cw - 0.60, h: 0.38,
      fontFace: F.jp, fontSize: 15.5, bold: true, color: C.green, margin: 0, valign: "middle",
    });
    s.addText(it[2], {
      x: x + 0.30, y: y + 1.46, w: cw - 0.60, h: 0.26,
      fontFace: F.jp, fontSize: 10.5, bold: true, color: C.gold, margin: 0, valign: "middle",
    });
    s.addShape("rect", { x: x + 0.30, y: y + 1.80, w: cw - 0.60, h: 0.011, fill: { color: C.line } });
    s.addText(it[3], {
      x: x + 0.30, y: y + 1.92, w: cw - 0.60, h: 1.04,
      fontFace: F.jp, fontSize: 11, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.4,
    });
  });

  panel(s, M, 5.28, CW, 1.56);
  icon(s, "LuZap", "white", M + 0.42, 5.56, 0.34);
  s.addText("この3つの課題に同時に応えるのが、「無人カフェ × 遊休スペース」というモデルです。", {
    x: M + 0.92, y: 5.50, w: CW - 1.34, h: 0.46,
    fontFace: F.jp, fontSize: 19, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  s.addText("人手をかけずに運営でき、働く・学ぶ場所を求める需要を受け止め、空いたスペースを収益資産に変えられます。", {
    x: M + 0.92, y: 6.04, w: CW - 1.34, h: 0.44,
    fontFace: F.jp, fontSize: 11.5, color: C.greenPale, margin: 0, valign: "middle",
  });
}

/* =============================================================== p.6 人手不足対策 */
{
  const s = pres.addSlide();
  shell(s, "人手不足対策：無人管理の優位性", 6);

  const px = M, py = 1.55, pw = 4.35, ph = 5.30;
  panel(s, px, py, pw, ph);
  s.addText("THE ADVANTAGE", {
    x: px + 0.46, y: py + 0.52, w: pw - 0.92, h: 0.28,
    fontFace: F.num, fontSize: 10.5, bold: true, color: C.greenPale, charSpacing: 3, margin: 0, valign: "middle",
  });
  s.addText("「人」に\n頼らない\n経営へ", {
    x: px + 0.46, y: py + 1.00, w: pw - 0.92, h: 2.20,
    fontFace: F.jp, fontSize: 31, bold: true, color: C.white, margin: 0, valign: "top", lineSpacingMultiple: 1.24,
  });
  s.addShape("rect", { x: px + 0.46, y: py + 3.36, w: 1.10, h: 0.021, fill: { color: C.greenPale } });
  s.addText("無人運営がもたらす、3つの構造的な強み。人に依存しないから、規模を広げても運営品質とコストが崩れません。", {
    x: px + 0.46, y: py + 3.60, w: pw - 0.92, h: 1.18,
    fontFace: F.jp, fontSize: 11.5, color: C.greenPale, margin: 0, valign: "top", lineSpacingMultiple: 1.4,
  });
  icon(s, "LuUsers", "pale", px + pw - 1.24, py + ph - 1.10, 0.72);

  const cx = 5.35, cw = RIGHT - cx, chh = 1.62;
  const items = [
    ["LuClipboardList", "管理・教育コストの大幅削減", "スタッフ管理に伴う工数や人的リスクを抑え、運営にかかる負担を軽減します。"],
    ["LuPiggyBank", "省人力運営で、無駄な人件費をカット", "必要最小限の作業で回る仕組みにより、人件費を抑えながら安定した収益を確保できます。"],
    ["LuBadgeCheck", "安定した収益と品質", "属人性を排除することで、常に均一したサービス品質を維持し、運営の安定性を高めます。"],
  ];
  items.forEach((it, i) => {
    const y = 1.55 + i * (chh + 0.22);
    card(s, cx, y, cw, chh);
    iconBadge(s, it[0], cx + 0.34, y + 0.34, { d: 0.66 });
    s.addText(it[1], {
      x: cx + 1.16, y: y + 0.30, w: cw - 1.50, h: 0.42,
      fontFace: F.jp, fontSize: 16, bold: true, color: C.green, margin: 0, valign: "middle",
    });
    s.addText(it[2], {
      x: cx + 1.16, y: y + 0.78, w: cw - 1.50, h: 0.62,
      fontFace: F.jp, fontSize: 11.5, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
    });
  });
}

/* =============================================================== p.7 既存事業とのシナジー */
{
  const s = pres.addSlide();
  shell(s, "既存事業とのシナジー", 7);

  panel(s, M, 1.40, CW, 1.28);
  s.addText("カフェ併設がもたらす、既存事業へのプラス効果", {
    x: M + 0.42, y: 1.52, w: CW - 0.84, h: 0.28,
    fontFace: F.jp, fontSize: 10.5, bold: true, color: C.greenPale, margin: 0, valign: "middle",
  });
  s.addText("「待ち時間」の解消や新規集客のフックとして、カフェは最適の組み合わせです。", {
    x: M + 0.42, y: 1.86, w: CW - 0.84, h: 0.62,
    fontFace: F.jp, fontSize: 21, bold: true, color: C.white, margin: 0, valign: "middle",
  });

  const items = [
    ["LuHourglass", "待ち時間の解消", "「待つ時間」を「くつろぐ時間」へ", "書店・美容室・整体院・病院の待合・コインランドリーなど、待ち時間が生まれる場所にカフェがあるだけで顧客満足度が上がります。"],
    ["LuSearch", "新規集客の新しい入り口", "「カフェ」で検索されて見つかる", "Googleマップやカフェ検索アプリ・SNSで「近くのカフェ」を探す人にも見つけてもらえる。カフェ目的で立ち寄った人が、既存事業を初めて知る新しいお客様になります。"],
    ["LuShoppingBag", "滞在時間・回遊率アップ", "「ついで買い」のきっかけに", "コーヒー片手に店内をゆっくり見て回れるようになり、滞在時間が伸びて追加購入・再来店のきっかけが増えます。"],
    ["LuRuler", "省スペース・低コスト", "セルフ型だから始めやすい", "専任スタッフ不要。レジ横や待合の小さなスペースに設置でき、大掛かりな工事や人員追加なしで導入できます。"],
  ];
  const cw = 5.77, chh = 1.82, gx = 0.325, gy = 0.22;
  items.forEach((it, i) => {
    const x = M + (i % 2) * (cw + gx);
    const y = 2.90 + Math.floor(i / 2) * (chh + gy);
    card(s, x, y, cw, chh);
    iconBadge(s, it[0], x + 0.28, y + 0.28, { d: 0.58 });
    const tagW = 0.19 + it[1].length * 0.135;
    s.addShape("roundRect", { x: x + 0.98, y: y + 0.28, w: tagW, h: 0.30, rectRadius: 0.15, fill: { color: C.green }, line: { type: "none" } });
    s.addText(it[1], {
      x: x + 0.98, y: y + 0.28, w: tagW, h: 0.30,
      fontFace: F.jp, fontSize: 9.5, bold: true, color: C.white, align: "center", valign: "middle", margin: 0,
    });
    s.addText(it[2], {
      x: x + 0.98, y: y + 0.64, w: cw - 1.26, h: 0.34,
      fontFace: F.jp, fontSize: 14.5, bold: true, color: C.green, margin: 0, valign: "middle",
    });
    s.addText(it[3], {
      x: x + 0.28, y: y + 1.04, w: cw - 0.56, h: 0.64,
      fontFace: F.jp, fontSize: 10.5, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.26,
    });
  });

  s.addText("※待合スペース・レジ横カウンターなど、限られたスペースでも設置可能です。", {
    x: M, y: 6.80, w: CW, h: 0.24,
    fontFace: F.jp, fontSize: 9.5, color: C.muted, margin: 0, valign: "middle",
  });
}

/* =============================================================== p.8 セルフカフェとは */
{
  const s = pres.addSlide();
  shell(s, "セルフカフェとは", 8);

  const lw2 = 6.55;
  const rows = [
    ["LuHouse", "お家のような", "寛ぎすぎないが、設備が整った集中環境。"],
    ["LuCoffee", "カフェのような", "食事はないが、時間を気にせず美味しいコーヒーが飲める。"],
    ["LuBookOpen", "図書館のような", "静かすぎない、適度な賑やかさがある。"],
  ];
  rows.forEach((r, i) => {
    const y = 1.55 + i * 1.26;
    card(s, M, y, lw2, 1.10);
    iconBadge(s, r[0], M + 0.30, y + 0.24, { d: 0.62 });
    s.addText(r[1], {
      x: M + 1.10, y: y + 0.20, w: lw2 - 1.40, h: 0.36,
      fontFace: F.jp, fontSize: 16, bold: true, color: C.green, margin: 0, valign: "middle",
    });
    s.addText(r[2], {
      x: M + 1.10, y: y + 0.58, w: lw2 - 1.40, h: 0.34,
      fontFace: F.jp, fontSize: 11.5, color: C.ink, margin: 0, valign: "middle",
    });
  });

  panel(s, M, 5.44, lw2, 1.40);
  s.addText("「誰もが気軽に使える」", {
    x: M + 0.40, y: 5.58, w: lw2 - 0.80, h: 0.38,
    fontFace: F.jp, fontSize: 14, color: C.greenPale, margin: 0, valign: "middle",
  });
  s.addText("第三の居場所", {
    x: M + 0.40, y: 5.98, w: lw2 - 1.40, h: 0.62,
    fontFace: F.jp, fontSize: 27, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  icon(s, "LuCoffee", "pale", M + lw2 - 1.14, 5.84, 0.66);

  // 右：展開エリアパネル
  const rx = 7.60, rw = RIGHT - rx;
  card(s, rx, 1.55, rw, 5.29);
  icon(s, "LuStore", "green", rx + 0.34, 1.80, 0.30);
  s.addText("展開エリア（2026年7月時点）", {
    x: rx + 0.74, y: 1.78, w: rw - 1.08, h: 0.30,
    fontFace: F.jp, fontSize: 10.5, bold: true, color: C.brown, margin: 0, valign: "middle",
  });
  s.addText(
    [
      { text: "71", options: { fontFace: F.num, fontSize: 46, bold: true, color: C.green } },
      { text: " 店舗", options: { fontFace: F.jp, fontSize: 15, bold: true, color: C.green } },
    ],
    { x: rx + 0.34, y: 2.14, w: rw - 0.68, h: 0.66, margin: 0, valign: "middle", align: "left" }
  );
  s.addShape("rect", { x: rx + 0.34, y: 2.92, w: rw - 0.68, h: 0.011, fill: { color: C.line } });

  const prefs = [["愛知県", 36], ["大阪府", 18], ["東京都", 7], ["岩手県", 2], ["岐阜県", 2],
    ["埼玉県", 1], ["千葉県", 1], ["静岡県", 1], ["滋賀県", 1], ["三重県", 1], ["鳥取県", 1]];
  prefs.forEach((p, i) => {
    const y = 3.04 + i * 0.252;
    s.addText(p[0], { x: rx + 0.34, y, w: 2.0, h: 0.24, fontFace: F.jp, fontSize: 10.5, color: C.ink, margin: 0, valign: "middle" });
    s.addText(String(p[1]) + " 店舗", {
      x: rx + rw - 1.90, y, w: 1.56, h: 0.24,
      fontFace: F.jp, fontSize: 10.5, bold: true, color: C.green, align: "right", margin: 0, valign: "middle",
    });
  });

  s.addShape("roundRect", { x: rx + 0.34, y: 5.90, w: rw - 0.68, h: 0.72, rectRadius: 0.06, fill: { color: C.cream }, line: { type: "none" } });
  s.addText("※平均滞在時間は2〜3時間、1回のご利用で2杯程度購入される傾向があります。", {
    x: rx + 0.52, y: 5.90, w: rw - 1.04, h: 0.72,
    fontFace: F.jp, fontSize: 10, color: C.brown, margin: 0, valign: "middle", lineSpacingMultiple: 1.25,
  });
}

/* =============================================================== p.9 実績データ：店舗展開数 */
{
  const s = pres.addSlide();
  shell(s, "実績データ：店舗展開数の推移", 9);
  kicker(s, "営業中店舗の開店月ベース累計（2022.09〜2026.07／単位：店舗）", 1.36);

  const series = [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 4, 8, 8, 9, 9, 12, 17, 19, 21, 22, 23,
    23, 25, 28, 31, 32, 34, 39, 42, 44, 44, 46, 47, 48, 48, 50, 50, 51, 55, 58, 61, 64, 67, 71];
  const labels = series.map((_, i) =>
    i === 0 ? "2022.09" : i === 12 ? "2023.09" : i === 24 ? "2024.09" : i === 36 ? "2025.09" : i === 46 ? "2026.07" : ""
  );

  s.addChart(
    pres.ChartType.area,
    [{ name: "累計店舗数", labels, values: series }],
    {
      x: M, y: 1.80, w: 8.52, h: 4.52,
      chartColors: [C.greenMid], chartColorsOpacity: 34,
      showLegend: false, showTitle: false, showValue: false,
      lineSize: 3.5, lineSmooth: false,
      catAxisLabelColor: C.inkSub, catAxisLabelFontFace: F.num, catAxisLabelFontSize: 10,
      catAxisLabelFrequency: 1, catAxisLabelRotate: 0, catAxisMultiLevelLabels: false,
      catAxisLineShow: true, catAxisLineColor: C.creamDk,
      catGridLine: { style: "none" },
      valAxisLabelColor: C.muted, valAxisLabelFontFace: F.num, valAxisLabelFontSize: 10,
      valAxisMinVal: 0, valAxisMaxVal: 75, valAxisMajorUnit: 25,
      valAxisLineShow: false,
      valGridLine: { color: C.line, size: 0.75 },
      dataBorder: { pt: 0, color: C.greenMid },
    }
  );

  // 成長を示す右上方向の矢印＋現在地の強調
  s.addShape("line", {
    x: M + 8.52 * 0.28, y: 2.98, w: 8.52 * 0.58, h: 2.52, flipV: true,
    line: { color: C.gold, width: 4.5, endArrowType: "triangle" },
  });
  const cxp = M + 8.52 - 0.62;
  const cyp = 1.80 + 0.10 + (1 - 71 / 75) * (4.52 - 0.52);
  s.addShape("ellipse", { x: cxp - 0.09, y: cyp - 0.09, w: 0.18, h: 0.18, fill: { color: C.gold }, line: { color: C.white, width: 1.75 } });
  s.addShape("roundRect", {
    x: cxp - 2.02, y: cyp + 0.20, w: 1.84, h: 0.44, rectRadius: 0.06,
    fill: { color: C.greenDeep }, line: { type: "none" },
  });
  s.addText(
    [
      { text: "現在 ", options: { fontFace: F.jp, fontSize: 10, color: C.greenPale } },
      { text: "71", options: { fontFace: F.num, fontSize: 15, bold: true, color: C.white } },
      { text: " 店舗", options: { fontFace: F.jp, fontSize: 10.5, bold: true, color: C.white } },
    ],
    { x: cxp - 2.02, y: cyp + 0.20, w: 1.84, h: 0.44, align: "center", margin: 0, valign: "middle" }
  );

  const tx = 9.55, tw = RIGHT - tx;
  statTile(s, tx, 1.80, tw, 1.42, "71", "店舗", "2026年7月時点の営業中店舗数", { dark: true, valueSize: 34 });
  statTile(s, tx, 3.36, tw, 1.42, "11", "都府県", "東北〜中国地方に広がる展開エリア");
  statTile(s, tx, 4.92, tw, 1.42, "47", "ヶ月", "サービス開始からの営業月数");

  s.addText("2022年9月の1号店OPENから、驚異的なスピードで出店エリアを拡大。店舗ネットワークは着実に成長を続けています。", {
    x: M, y: 6.42, w: 8.52, h: 0.46,
    fontFace: F.jp, fontSize: 11.5, color: C.ink, margin: 0, valign: "middle", lineSpacingMultiple: 1.25,
  });
}

/* =============================================================== p.10 実績データ：月間利用者数 */
{
  const s = pres.addSlide();
  shell(s, "実績データ：月間利用者数の推移", 10);
  kicker(s, "全社月間利用者数（各月の合計）／店舗数の拡大に比例して拡大（2023.01〜2026.07）", 1.36);

  const labels = ["2023.01", "2023.07", "2024.01", "2024.07", "2025.01", "2025.07", "2026.01", "2026.07"];
  const users = [895, 1419, 9722, 26383, 46017, 74260, 80293, 101343];
  const stores = [1, 1, 9, 22, 32, 46, 51, 71];

  s.addChart(
    [
      {
        type: pres.ChartType.bar,
        data: [{ name: "月間利用者数（人・左軸）", labels, values: users }],
        options: {
          chartColors: [C.greenMid], barGapWidthPct: 36,
          showValue: true, dataLabelPosition: "outEnd",
          dataLabelColor: C.green, dataLabelFontFace: F.num, dataLabelFontSize: 9,
          dataLabelFormatCode: "#,##0",
        },
      },
      {
        type: pres.ChartType.line,
        data: [{ name: "累計店舗数（店・右軸）", labels, values: stores }],
        options: {
          chartColors: [C.gold], lineSize: 2.5, lineSmooth: false,
          lineDataSymbol: "circle", lineDataSymbolSize: 7,
          lineDataSymbolLineColor: C.gold, lineDataSymbolLineSize: 1.25,
          secondaryValAxis: true, secondaryCatAxis: true,
        },
      },
    ],
    {
      x: M, y: 1.80, w: 8.52, h: 4.52,
      showTitle: false,
      showLegend: true, legendPos: "t", legendColor: C.inkSub, legendFontFace: F.jp, legendFontSize: 10,
      catAxes: [
        { catAxisLabelColor: C.inkSub, catAxisLabelFontFace: F.num, catAxisLabelFontSize: 10, catAxisLineColor: C.creamDk, catGridLine: { style: "none" } },
        { catAxisHidden: true },
      ],
      valAxes: [
        {
          valAxisLabelColor: C.muted, valAxisLabelFontFace: F.num, valAxisLabelFontSize: 10,
          valAxisLabelFormatCode: "#,##0",
          valAxisMinVal: 0, valAxisMaxVal: 110000, valAxisMajorUnit: 55000,
          valAxisLineShow: false, valGridLine: { color: C.line, size: 0.75 },
        },
        {
          valAxisLabelColor: C.muted, valAxisLabelFontFace: F.num, valAxisLabelFontSize: 10,
          valAxisMinVal: 0, valAxisMaxVal: 75, valAxisMajorUnit: 25,
          valAxisLineShow: false, valGridLine: { style: "none" },
        },
      ],
    }
  );

  // 成長を示す右上方向の矢印＋現在バッジ
  s.addShape("line", {
    x: M + 8.52 * 0.24, y: 3.30, w: 8.52 * 0.60, h: 2.30, flipV: true,
    line: { color: C.gold, width: 4.5, endArrowType: "triangle" },
  });
  const lbx = M + 0.5 + (8.52 - 0.9) * (7.5 / 8) - 0.52;
  s.addShape("roundRect", {
    x: lbx, y: 1.92, w: 1.04, h: 0.32, rectRadius: 0.16,
    fill: { color: C.gold }, line: { type: "none" },
  });
  s.addText("現在", {
    x: lbx, y: 1.92, w: 1.04, h: 0.32,
    fontFace: F.jp, fontSize: 10.5, bold: true, color: C.white, align: "center", margin: 0, valign: "middle",
  });

  const tx = 9.55, tw = RIGHT - tx;
  statTile(s, tx, 1.80, tw, 1.42, "101,343", "人", "2026年7月の月間利用者数", { dark: true, valueSize: 26 });
  statTile(s, tx, 3.36, tw, 1.42, "104", "倍", "立ち上げ当初比の月間利用者数", { valueSize: 30 });
  statTile(s, tx, 4.92, tw, 1.42, "100,000", "杯", "月間販売数を突破", { valueSize: 26 });

  s.addText("立ち上げ当初から、月間利用者数の規模は着実に拡大を続けています。新規出店直後の一時的な希薄化を挟みながらも、ほぼ店舗数に比例するペースで成長しています。", {
    x: M, y: 6.42, w: 8.52, h: 0.46,
    fontFace: F.jp, fontSize: 11.5, color: C.ink, margin: 0, valign: "middle", lineSpacingMultiple: 1.25,
  });
}

/* =============================================================== p.11 利用システム */
{
  const s = pres.addSlide();
  shell(s, "利用システム：究極の利便性と価格", 11);

  const lw2 = 6.10;
  const rows = [
    ["LuCoffee", "ドリンク一杯（390円〜）", "ドリンク購入だけで、Wi-Fi・電源が完備された空間を自由にご利用いただけます。"],
    ["LuClock", "滞在時間無制限", "追加料金や延長料金を気にせず、仕事や勉強に没頭できる環境を提供します。"],
    ["LuSmartphone", "サブスクプラン", "ヘビーユーザー向けに、月額定額で通い放題のプランも用意。最低1杯147円〜で継続利用を促せます。"],
  ];
  rows.forEach((r, i) => {
    const y = 1.55 + i * 1.80;
    card(s, M, y, lw2, 1.62);
    iconBadge(s, r[0], M + 0.32, y + 0.32, { d: 0.66 });
    s.addText(r[1], {
      x: M + 1.14, y: y + 0.28, w: lw2 - 1.46, h: 0.40,
      fontFace: F.jp, fontSize: 16, bold: true, color: C.green, margin: 0, valign: "middle",
    });
    s.addText(r[2], {
      x: M + 1.14, y: y + 0.74, w: lw2 - 1.46, h: 0.64,
      fontFace: F.jp, fontSize: 11.5, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
    });
  });

  const rx = 7.30, rw = RIGHT - rx;
  card(s, rx, 1.55, rw, 5.42);
  s.addText("すべての店舗に共通する設備・ルール", {
    x: rx + 0.32, y: 1.78, w: rw - 0.64, h: 0.30,
    fontFace: F.jp, fontSize: 10.5, bold: true, color: C.brown, margin: 0, valign: "middle",
  });
  const feats = [
    ["LuCalendarCheck", "無人・年中無休"], ["LuWifi", "高速Wi-Fi"], ["LuPlug", "電源完備"],
    ["LuClock", "席時間無制限"], ["LuBookOpen", "仕事・勉強"], ["LuMessagesSquare", "会話OK"],
  ];
  const tw2 = (rw - 0.64 - 2 * 0.22) / 3, th = 1.42;
  feats.forEach(([ic, label], i) => {
    const x = rx + 0.32 + (i % 3) * (tw2 + 0.22);
    const y = 2.22 + Math.floor(i / 3) * (th + 0.22);
    s.addShape("roundRect", { x, y, w: tw2, h: th, rectRadius: 0.07, fill: { color: C.greenTint }, line: { color: "D7E4D7", width: 0.75 } });
    icon(s, ic, "green", x + tw2 / 2 - 0.24, y + 0.26, 0.48);
    s.addText(label, {
      x: x + 0.06, y: y + 0.86, w: tw2 - 0.12, h: 0.40,
      fontFace: F.jp, fontSize: 11.5, bold: true, color: C.green, align: "center", valign: "middle", margin: 0,
    });
  });
  s.addShape("roundRect", { x: rx + 0.32, y: 5.62, w: rw - 0.64, h: 1.06, rectRadius: 0.06, fill: { color: C.cream }, line: { type: "none" } });
  s.addText("予約不要・登録不要。ドリンク1杯から、いつでも気軽に使える価格設計にしています。", {
    x: rx + 0.52, y: 5.62, w: rw - 1.04, h: 1.06,
    fontFace: F.jp, fontSize: 11, color: C.brown, margin: 0, valign: "middle", lineSpacingMultiple: 1.3,
  });
}

/* =============================================================== p.12 ユーザー層 */
{
  const s = pres.addSlide();
  shell(s, "セルフカフェ_ユーザー層", 12);

  const cw = (CW - 0.32) / 2;

  // 左：年齢構成
  card(s, M, 1.42, cw, 5.42);
  iconBadge(s, "LuUsers", M + 0.32, 1.66, { d: 0.62 });
  s.addText("若年層から絶大な支持", {
    x: M + 1.10, y: 1.66, w: cw - 1.42, h: 0.62,
    fontFace: F.jp, fontSize: 17, bold: true, color: C.green, margin: 0, valign: "middle",
  });
  s.addText(
    [
      { text: "29歳までのユーザーが", options: { color: C.ink } },
      { text: "約50%", options: { color: C.green, bold: true } },
      { text: "。デジタルネイティブ世代にとって「無人」は心理的ハードルが低く、快適な空間です。", options: { color: C.ink } },
    ],
    { x: M + 0.32, y: 2.42, w: cw - 0.64, h: 0.66, fontFace: F.jp, fontSize: 11.5, margin: 0, valign: "top", lineSpacingMultiple: 1.32 }
  );
  s.addText("（歳）", {
    x: M + cw - 1.00, y: 6.44, w: 0.70, h: 0.22,
    fontFace: F.jp, fontSize: 9, color: C.muted, align: "right", valign: "middle", margin: 0,
  });
  s.addChart(
    pres.ChartType.bar,
    [{ name: "構成比", labels: ["〜14", "15〜19", "20〜29", "30〜39", "40〜49", "50〜59", "60〜"], values: [3, 24, 29, 16, 13, 4, 1] }],
    {
      x: M + 0.26, y: 3.22, w: cw - 0.52, h: 3.32,
      barDir: "col", chartColors: [C.greenPale, C.green, C.green, C.greenMid, C.greenMid, C.creamDk, C.creamDk],
      varyColors: true, barGapWidthPct: 40,
      showLegend: false, showTitle: false,
      showValue: true, dataLabelPosition: "outEnd", dataLabelColor: C.green,
      dataLabelFontFace: F.num, dataLabelFontSize: 10, dataLabelFormatCode: '0"%"',
      catAxisLabelColor: C.inkSub, catAxisLabelFontFace: F.num, catAxisLabelFontSize: 10,
      catAxisLabelRotate: 0, catAxisLineColor: C.creamDk, catGridLine: { style: "none" },
      valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 36, valGridLine: { style: "none" },
    }
  );

  // 右：利用目的
  const rx = M + cw + 0.32;
  card(s, rx, 1.42, cw, 5.42);
  iconBadge(s, "LuGraduationCap", rx + 0.32, 1.66, { d: 0.62 });
  s.addText("「勉強」が主目的", {
    x: rx + 1.10, y: 1.66, w: cw - 1.42, h: 0.62,
    fontFace: F.jp, fontSize: 17, bold: true, color: C.green, margin: 0, valign: "middle",
  });
  s.addText(
    [
      { text: "利用目的の", options: { color: C.ink } },
      { text: "62.6%が「勉強」", options: { color: C.green, bold: true } },
      { text: "。一般的なカフェが「休憩」を主目的とするのに対し、高い生産性を求める層が集まります。", options: { color: C.ink } },
    ],
    { x: rx + 0.32, y: 2.42, w: cw - 0.64, h: 0.66, fontFace: F.jp, fontSize: 11.5, margin: 0, valign: "top", lineSpacingMultiple: 1.32 }
  );
  s.addChart(
    pres.ChartType.doughnut,
    [{ name: "利用目的", labels: ["勉強 62.6%", "休憩・雑談 22.8%", "仕事・打ち合わせ 13.0%", "その他 1.6%"], values: [62.6, 22.8, 13.0, 1.6] }],
    {
      x: rx + 0.30, y: 3.20, w: cw - 0.60, h: 3.34,
      holeSize: 56,
      chartColors: [C.green, C.gold, C.greenPale, C.creamDk],
      showLegend: true, legendPos: "b", legendColor: C.inkSub, legendFontFace: F.jp, legendFontSize: 10,
      showTitle: false, showValue: false,
    }
  );
}

/* =============================================================== p.13 / p.14 店舗毎の利用状況 */
[
  { page: 13, suffix: "①", stores: [["ささしまライブ店（愛知県名古屋市）", "store-sasashima.jpg"], ["新瑞橋店（愛知県名古屋市）", "store-aratama.jpg"], ["名駅西口店（愛知県名古屋市）", "store-meieki.jpg"], ["印西牧の原店（千葉県印西市）", "store-inzai.jpg"], ["栄店（愛知県名古屋市）", "store-sakae.jpg"], ["御器所店（愛知県名古屋市）", "store-gokiso.jpg"]] },
  { page: 14, suffix: "②", stores: [["盛岡駅前店（岩手県盛岡市）", "store-morioka.jpg"], ["天満店（大阪府大阪市）", "store-tenma.jpg"], ["浜松新橋店（静岡県浜松市）", "store-hamamatsu.jpg"], ["谷町九丁目店（大阪府大阪市）", "store-tanimachi.jpg"], ["相川駅前店（大阪府大阪市）", "store-aikawa.jpg"], ["あべの南店（大阪府大阪市）", "store-abeno.jpg"]] },
].forEach(({ page, suffix, stores }) => {
  const s = pres.addSlide();
  shell(s, `店舗毎の利用状況 ${suffix}`, page);
  kicker(s, "店内カメラで見た実際の様子。平日日中でも席が埋まる店舗が多数あります。", 1.36);

  const gx = 0.26, gy = 0.24;
  const cw = (CW - 2 * gx) / 3, chh = (6.86 - 1.82 - gy) / 2;
  stores.forEach(([name, img], i) => {
    const x = M + (i % 3) * (cw + gx);
    const y = 1.82 + Math.floor(i / 3) * (chh + gy);
    photoSlot(s, x, y, cw, chh, name, { captionSize: 10, img });
  });
});

/* =============================================================== p.15 比較表 */
{
  const s = pres.addSlide();
  shell(s, "比較表：いいところ取りのビジネスモデル", 15);
  kicker(s, "カフェの「気軽さ」とオフィスの「機能」を融合させた唯一無二のポジション", 1.36);

  const colW = [2.35, 3.55, 2.98, 2.97];
  const ty = 2.02;

  s.addShape("roundRect", {
    x: M + colW[0] - 0.02, y: ty - 0.32, w: colW[1] + 0.04, h: 4.54,
    rectRadius: 0.07, fill: { color: C.greenTint }, line: { type: "none" },
  });
  icon(s, "LuBadgeCheck", "green", M + colW[0] + 0.16, ty - 0.30, 0.24);
  s.addText("いいところ取りのポジション", {
    x: M + colW[0] + 0.46, y: ty - 0.32, w: colW[1] - 0.60, h: 0.28,
    fontFace: F.jp, fontSize: 10.5, bold: true, color: C.greenMid, margin: 0, valign: "middle",
  });

  const hdr = (t, fill, color) => ({
    text: t,
    options: { fill: { color: fill }, color, bold: true, fontSize: 12, fontFace: F.jp, align: "left", valign: "middle", margin: [0.06, 0.14, 0.06, 0.14] },
  });
  const lab = (t) => ({
    text: t,
    options: { fill: { color: "F6F2E6" }, color: C.brown, bold: true, fontSize: 11.5, fontFace: F.jp, valign: "middle", margin: [0.06, 0.14, 0.06, 0.14] },
  });
  const good = (t) => ({
    text: [
      { text: "◎ ", options: { color: C.green, bold: true, fontSize: 13 } },
      { text: t, options: { color: C.green, bold: true, fontSize: 11 } },
    ],
    options: { fill: { color: "F2F7F2" }, fontFace: F.jp, valign: "middle", margin: [0.08, 0.16, 0.08, 0.16] },
  });
  const so = (mark, t) => ({
    text: [
      { text: mark + " ", options: { color: mark === "×" ? "B0553F" : C.gold, bold: true, fontSize: 13 } },
      { text: t, options: { color: C.inkSub, fontSize: 10.5 } },
    ],
    options: { fill: { color: C.white }, fontFace: F.jp, valign: "middle", margin: [0.08, 0.16, 0.08, 0.16] },
  });

  s.addTable(
    [
      [hdr("比較項目", C.creamDk, C.brown), hdr("セルフカフェ", C.green, C.white), hdr("コワーキング", "8A8272", C.white), hdr("シェアオフィス", "8A8272", C.white)],
      [lab("利用の手軽さ"), good("予約不要・1杯〜"), so("△", "登録や予約が必要"), so("×", "法人契約がメイン")],
      [lab("空間の雰囲気"), good("適度な賑わい"), so("△", "静かすぎて緊張する"), so("×", "閉鎖的な個室中心")],
      [lab("ドリンクの質"), good("本格コーヒー"), so("△", "サーバーが中心"), so("×", "持ち込みが基本")],
      [lab("運営人員"), good("完全無人"), so("△", "スタッフ常駐"), so("△", "管理人が必要")],
      [lab("導入コスト"), good("低投資パッケージ"), so("×", "高額な内装・設備"), so("×", "莫大な建築コスト")],
    ],
    {
      x: M, y: ty, w: CW, colW,
      rowH: [0.52, 0.74, 0.74, 0.74, 0.74, 0.74],
      border: { type: "solid", color: C.line, pt: 0.75 },
      autoPage: false,
    }
  );

  s.addText("※コワーキング／シェアオフィスの評価は、一般的なサービス形態を想定した当社比較です。", {
    x: M, y: 6.58, w: CW, h: 0.24,
    fontFace: F.jp, fontSize: 9.5, color: C.muted, margin: 0, valign: "middle",
  });
}

/* =============================================================== p.16 無人開閉店システム */
{
  const s = pres.addSlide();
  shell(s, "独自の無人開閉店システム", 16);

  const lw2 = 5.55;
  panel(s, M, 1.50, lw2, 5.30);
  icon(s, "LuKeyRound", "white", M + 0.46, 1.84, 0.42);
  s.addText("警備会社と連携した\n安心の自動化", {
    x: M + 0.46, y: 2.40, w: lw2 - 0.92, h: 1.00,
    fontFace: F.jp, fontSize: 24, bold: true, color: C.white, margin: 0, valign: "top", lineSpacingMultiple: 1.24,
  });
  s.addText("独自のスマートロックシステムにより、店舗開錠／施錠を遠隔で完全自動化。オーナー様は現地に足を運ぶこと無く、営業時間を管理できます。", {
    x: M + 0.46, y: 3.56, w: lw2 - 0.92, h: 1.02,
    fontFace: F.jp, fontSize: 11.5, color: C.greenPale, margin: 0, valign: "top", lineSpacingMultiple: 1.4,
  });
  ["任意の営業時間を分単位でコントロール", "不審者の侵入を防ぐセキュリティと連携", "無人ならではの24時間営業も容易"].forEach((b, i) => {
    const y = 4.76 + i * 0.52;
    icon(s, "LuCheck", "white", M + 0.46, y + 0.06, 0.22);
    s.addText(b, {
      x: M + 0.80, y: y, w: lw2 - 1.26, h: 0.34,
      fontFace: F.jp, fontSize: 12, bold: true, color: C.white, margin: 0, valign: "middle",
    });
  });

  const rx = M + lw2 + 0.32, rw = RIGHT - rx;
  const items = [
    ["LuBellRing", "非常ボタン", "非常・緊急時に押すと、警備員が駆け付けます。"],
    ["LuCctv", "画像センサー", "非常時の映像を監視センターへ送信。スピーカーで威嚇可能。閉店後は侵入者を感知して通報します。"],
    ["LuFlame", "火災センサー", "火災の発生を感知して、警備員が駆け付けます。"],
    ["LuKeyRound", "出入口の管理", "電気錠・電磁錠や自動ドアと連携し、閉店／開店時間に自動施錠・開錠。閉店時は自動で侵入警戒をセットします。"],
  ];
  const cw2 = (rw - 0.26) / 2, ch2 = (5.30 - 0.24) / 2;
  items.forEach((it, i) => {
    const x = rx + (i % 2) * (cw2 + 0.26);
    const y = 1.50 + Math.floor(i / 2) * (ch2 + 0.24);
    card(s, x, y, cw2, ch2);
    iconBadge(s, it[0], x + 0.26, y + 0.26, { d: 0.60 });
    s.addText(it[1], {
      x: x + 0.26, y: y + 0.94, w: cw2 - 0.52, h: 0.34,
      fontFace: F.jp, fontSize: 14, bold: true, color: C.green, margin: 0, valign: "middle",
    });
    s.addText(it[2], {
      x: x + 0.26, y: y + 1.32, w: cw2 - 0.52, h: 1.02,
      fontFace: F.jp, fontSize: 10.5, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
    });
  });

  s.addText("※警備会社は環境により異なる場合がございます。", {
    x: M, y: 6.92, w: 8.0, h: 0.24,
    fontFace: F.jp, fontSize: 9.5, color: C.muted, margin: 0, valign: "middle",
  });
}

/* =============================================================== p.17 モニタリング＆サポート */
{
  const s = pres.addSlide();
  shell(s, "店内モニタリング＆サポート基盤", 17);
  kicker(s, "緊急時に強い、安心のサポート", 1.36);

  const items = [
    ["LuBellRing", "非常ボタンを店内各所に", "災害・体調不良・不審者・お客様同士のトラブルなど、あらゆる緊急事態に備えています。"],
    ["LuVideo", "6台のカメラで常時監視", "店内の状況を常時把握。異常があれば即座に確認・対応できる体制を整えています。"],
    ["LuMic", "双方向音声に対応", "カメラはスタッフとお客様が直接やり取りできる双方向音声に対応。その場で声がけやサポートを行えます。"],
    ["LuMessageCircle", "LINEチャットで即時対応", "お客様からのお問い合わせにも即時対応。質問対応・意見収集の窓口としても機能します。"],
  ];
  const cw = (CW - 3 * 0.26) / 4, chh = 2.72;
  items.forEach((it, i) => {
    const x = M + i * (cw + 0.26), y = 1.78;
    card(s, x, y, cw, chh);
    iconBadge(s, it[0], x + 0.26, y + 0.26, { d: 0.62 });
    s.addText(it[1], {
      x: x + 0.26, y: y + 0.98, w: cw - 0.52, h: 0.42,
      fontFace: F.jp, fontSize: 13.5, bold: true, color: C.green, margin: 0, valign: "middle",
    });
    s.addText(it[2], {
      x: x + 0.26, y: y + 1.50, w: cw - 0.52, h: 1.06,
      fontFace: F.jp, fontSize: 10.5, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
    });
  });

  photoSlot(s, M, 4.78, 5.55, 2.06, "モニタリング画面／店内カメラの様子", { captionSize: 10.5, img: "p16-support.jpg" });

  const rx = M + 5.55 + 0.32, rw = RIGHT - rx;
  panel(s, rx, 4.78, rw, 2.06);
  icon(s, "LuShieldCheck", "white", rx + 0.42, 5.04, 0.36);
  s.addText("現場に駆けつける必要は、ほぼありません。", {
    x: rx + 0.92, y: 5.00, w: rw - 1.34, h: 0.44,
    fontFace: F.jp, fontSize: 17, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  s.addText("24時間体制の警備会社と遠隔監視システム、そして本部のチャット対応が重なることで、無人でも「見守られている店舗」を実現しています。", {
    x: rx + 0.92, y: 5.52, w: rw - 1.34, h: 0.90,
    fontFace: F.jp, fontSize: 11.5, color: C.greenPale, margin: 0, valign: "top", lineSpacingMultiple: 1.35,
  });
}

/* =============================================================== p.18 パートナー制度の3つの特徴 */
{
  const s = pres.addSlide();
  shell(s, "パートナー制度の3つの特徴", 18);
  kicker(s, "低コスト・遊休スペース活用・省人力運営の3点が、導入判断のハードルを下げます。", 1.36);

  const items = [
    ["LuWallet", "低初期コスト", "1/10以下", "他社カフェと比較し、1/10以下のコストで出店が可能です。"],
    ["LuBuilding2", "遊休スペースの活用", "デッドスペース → 収益資産", "既存店の空室やデッドスペースを、収益資産へ再構築します。"],
    ["LuTimer", "1日15分程度の運営", "清掃・補充のみ", "無人化により人件費を極限までカット。日常業務は清掃・補充のみです。"],
  ];
  const cw = (CW - 2 * 0.32) / 3, chh = 3.06;
  items.forEach((it, i) => {
    const x = M + i * (cw + 0.32), y = 1.86;
    card(s, x, y, cw, chh);
    s.addShape("ellipse", { x: x + cw / 2 - 0.52, y: y + 0.38, w: 1.04, h: 1.04, fill: { color: C.greenTint }, line: { type: "none" } });
    icon(s, it[0], "green", x + cw / 2 - 0.29, y + 0.61, 0.58);
    s.addText(it[1], {
      x: x + 0.24, y: y + 1.62, w: cw - 0.48, h: 0.42,
      fontFace: F.jp, fontSize: 17, bold: true, color: C.green, align: "center", valign: "middle", margin: 0,
    });
    s.addText(it[2], {
      x: x + 0.24, y: y + 2.06, w: cw - 0.48, h: 0.30,
      fontFace: F.jp, fontSize: 11, bold: true, color: C.gold, align: "center", valign: "middle", margin: 0,
    });
    s.addText(it[3], {
      x: x + 0.34, y: y + 2.38, w: cw - 0.68, h: 0.62,
      fontFace: F.jp, fontSize: 11, color: C.ink, align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.3,
    });
  });

  panel(s, M, 5.28, CW, 1.56);
  icon(s, "LuSparkles", "white", M + 0.42, 5.56, 0.34);
  s.addText("新しく人を採用・配置する必要はありません。", {
    x: M + 0.92, y: 5.50, w: CW - 1.34, h: 0.46,
    fontFace: F.jp, fontSize: 19, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  s.addText("既存スタッフの稼働内で、もう一つの収益源をつくれます。", {
    x: M + 0.92, y: 6.04, w: CW - 1.34, h: 0.44,
    fontFace: F.jp, fontSize: 11.5, color: C.greenPale, margin: 0, valign: "middle",
  });
}

/* =============================================================== p.19 収益シミュレーション */
{
  const s = pres.addSlide();
  shell(s, "収益シミュレーション", 19);
  kicker(s, "既存の収益シミュレーション（25坪・70坪）の営業総利益をもとにした2パターン比較", 1.36);

  const tw = 8.15;
  const colW = [3.05, 2.55, 2.55];
  const hdr = (t, fill, color, align) => ({
    text: t,
    options: { fill: { color: fill }, color, bold: true, fontSize: 12, fontFace: F.jp, align: align || "left", valign: "middle", margin: [0.06, 0.16, 0.06, 0.16] },
  });
  const lab = (t, big) => ({
    text: t,
    options: { fill: { color: "F6F2E6" }, color: big ? C.green : C.brown, bold: true, fontSize: big ? 13.5 : 11.5, fontFace: F.jp, valign: "middle", margin: [0.06, 0.16, 0.06, 0.16] },
  });
  const val = (t, opts = {}) => ({
    text: t,
    options: {
      fill: { color: opts.fill || C.white }, color: opts.color || C.ink,
      bold: !!opts.bold, fontSize: opts.size || 13, fontFace: F.num,
      align: "right", valign: "middle", margin: [0.06, 0.18, 0.06, 0.18],
    },
  });

  s.addTable(
    [
      [hdr("項目", C.creamDk, C.brown), hdr("25坪パターン", C.green, C.white, "right"), hdr("70坪パターン", C.green, C.white, "right")],
      [lab("営業総利益（売上－原価）"), val("888,100円"), val("1,387,600円")],
      [lab("業務委託料（25%）", true), val("222,025円", { bold: true, color: C.green, size: 17, fill: "F2F7F2" }), val("346,900円", { bold: true, color: C.green, size: 17, fill: "F2F7F2" })],
      [lab("年間換算"), val("2,664,300円", { color: C.inkSub, size: 12 }), val("4,162,800円", { color: C.inkSub, size: 12 })],
    ],
    {
      x: M, y: 1.94, w: tw, colW,
      rowH: [0.52, 0.86, 1.06, 0.80],
      border: { type: "solid", color: C.line, pt: 0.75 },
      autoPage: false,
    }
  );

  s.addShape("roundRect", { x: M, y: 5.44, w: tw, h: 1.06, rectRadius: 0.06, fill: { color: C.cream }, line: { type: "none" } });
  s.addText("パートナー様にご負担いただく経費：水道光熱費／通信費／人件費／その他雑費\n（いずれも実費精算。具体的な金額は店舗の稼働状況により変動します）", {
    x: M + 0.24, y: 5.44, w: tw - 0.48, h: 1.06,
    fontFace: F.jp, fontSize: 10.5, color: C.brown, margin: 0, valign: "middle", lineSpacingMultiple: 1.3,
  });

  const rx = M + tw + 0.32, rw = RIGHT - rx;
  panel(s, rx, 1.94, rw, 2.30);
  icon(s, "LuCalculator", "white", rx + 0.34, 2.20, 0.34);
  s.addText("貴社の受取イメージ", {
    x: rx + 0.34, y: 2.64, w: rw - 0.68, h: 0.32,
    fontFace: F.jp, fontSize: 12, bold: true, color: C.greenPale, margin: 0, valign: "middle",
  });
  s.addText("月 22万〜35万円", {
    x: rx + 0.34, y: 3.00, w: rw - 0.68, h: 0.50,
    fontFace: F.jp, fontSize: 21, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  s.addText("年 266万〜416万円", {
    x: rx + 0.34, y: 3.52, w: rw - 0.68, h: 0.36,
    fontFace: F.jp, fontSize: 13, color: C.greenPale, margin: 0, valign: "middle",
  });

  card(s, rx, 4.46, rw, 2.04);
  s.addText("計算の前提", {
    x: rx + 0.30, y: 4.66, w: rw - 0.60, h: 0.28,
    fontFace: F.jp, fontSize: 10.5, bold: true, color: C.brown, margin: 0, valign: "middle",
  });
  s.addText(
    ["営業総利益（売上－原価）の25%を業務委託料としてお支払い", "売上金の管理は本部が実施", "上記は既存店の実績をもとにした試算値"]
      .map((t, i, a) => ({ text: t, options: { bullet: { characterCode: "25AA", indent: 11 }, breakLine: i < a.length - 1 } })),
    { x: rx + 0.30, y: 4.98, w: rw - 0.60, h: 1.34, fontFace: F.jp, fontSize: 10.5, color: C.ink, margin: 0, valign: "top", paraSpaceAfter: 5 }
  );
}

/* =============================================================== p.20 有料オプション一覧 */
{
  const s = pres.addSlide();
  shell(s, "有料オプション一覧", 20);
  kicker(s, "すべて任意でお選びいただけます（金額は税抜き表示）", 1.36);

  const colW = [2.35, 1.45, 1.45, 6.625];
  const hdr = (t, align) => ({
    text: t,
    options: { fill: { color: C.green }, color: C.white, bold: true, fontSize: 12, fontFace: F.jp, align: align || "left", valign: "middle", margin: [0.06, 0.14, 0.06, 0.14] },
  });
  const lab = (t) => ({
    text: t,
    options: { fill: { color: "F6F2E6" }, color: C.brown, bold: true, fontSize: 11.5, fontFace: F.jp, valign: "middle", margin: [0.06, 0.14, 0.06, 0.14] },
  });
  const yen = (t) => ({
    text: t,
    options: { fill: { color: C.white }, color: C.green, bold: true, fontSize: 12.5, fontFace: F.num, align: "right", valign: "middle", margin: [0.06, 0.16, 0.06, 0.16] },
  });
  const note = (runs) => ({
    text: runs,
    options: { fill: { color: C.white }, fontFace: F.jp, valign: "middle", margin: [0.08, 0.16, 0.08, 0.16] },
  });
  const b = (t) => ({ text: t, options: { color: C.green, bold: true, fontSize: 10, breakLine: true } });
  const p = (t, last) => ({ text: t, options: { color: C.ink, fontSize: 9.5, breakLine: !last } });

  s.addTable(
    [
      [hdr("項目"), hdr("初期費用", "right"), hdr("月額費用", "right"), hdr("備考")],
      [lab("公式LINE作成"), yen("15,000円"), yen("0円"), note([
        b("FC店の公式LINEを新規作成できます。"),
        p("メリット：独自の発信や告知、広告が打てます。"),
        p("デメリット：セルフカフェ公式LINEから切り離されるため、本来手放しでよいLINE問い合わせが直接届きます。"),
        p("本部の発信と連携できないため、本部情報はパートナー様側で都度発信いただく必要があります。", true),
      ])],
      [lab("顧客用コピー機"), yen("50,000円"), yen("20,000円"), note([
        b("PayPay決済のみ"),
        p("※本体代は別途"),
        p("※グローバルIPの取得が必要です", true),
      ])],
      [lab("営業サポートプラン\n（フルサポート）"), yen("0円"), yen("55,000円"), note([
        b("顧客チャット対応"),
        p("セルフカフェ公式LINEやHPからの問い合わせに対応"),
        b("ドリンクマシン管理・補充"),
        p("メーカー対応（エリア限定）", true),
      ])],
      [lab("営業サポートプラン\n（ミニマムサポート）"), yen("0円"), yen("15,000円"), note([
        b("顧客チャット対応"),
        p("セルフカフェ公式LINEやHPからの問い合わせに対応", true),
      ])],
    ],
    {
      x: M, y: 1.94, w: CW, colW,
      rowH: [0.50, 1.34, 0.92, 1.12, 0.86],
      border: { type: "solid", color: C.line, pt: 0.75 },
      autoPage: false,
    }
  );
}

/* =============================================================== p.21 ドリンクマシン */
{
  const s = pres.addSlide();
  shell(s, "ドリンクマシンに関して", 21);

  const lw2 = 4.35;
  card(s, M, 1.50, lw2, 5.35);
  s.addShape("roundRect", { x: M + 0.30, y: 1.76, w: lw2 - 0.60, h: 0.52, rectRadius: 0.26, fill: { color: C.green }, line: { type: "none" } });
  icon(s, "LuCoffee", "white", M + 0.52, 1.86, 0.32);
  s.addText("トータルサポートパック", {
    x: M + 0.92, y: 1.76, w: lw2 - 1.22, h: 0.52,
    fontFace: F.jp, fontSize: 13.5, bold: true, color: C.white, valign: "middle", margin: 0,
  });
  [["付属", "コーヒーサーバー、備品ラック"], ["メンテナンス", "月1回"], ["機種", "100RS"]].forEach((sp, i) => {
    const y = 2.48 + i * 0.62;
    s.addText(sp[0], {
      x: M + 0.30, y, w: 1.26, h: 0.34,
      fontFace: F.jp, fontSize: 10.5, bold: true, color: C.brown, margin: 0, valign: "middle",
    });
    s.addText(sp[1], {
      x: M + 1.56, y, w: lw2 - 1.86, h: 0.34,
      fontFace: F.jp, fontSize: 11.5, color: C.ink, margin: 0, valign: "middle",
    });
    s.addShape("rect", { x: M + 0.30, y: y + 0.40, w: lw2 - 0.60, h: 0.011, fill: { color: C.line } });
  });
  photoSlot(s, M + 0.30, 4.42, lw2 - 0.60, 2.16, "ドリンクマシン本体（100RS）", { captionSize: 10.5, img: "p24-machine.jpg" });

  const rx = M + lw2 + 0.32, rw = RIGHT - rx;
  s.addText("＜設置イメージ＞", {
    x: rx, y: 1.50, w: rw, h: 0.34,
    fontFace: F.jp, fontSize: 13, bold: true, color: C.brown, margin: 0, valign: "middle",
  });
  photoSlot(s, rx, 1.94, rw, 4.91, "店内に設置したドリンクマシン", { captionSize: 11, img: "p24-install.jpg" });
}

/* =============================================================== p.22 FAQ */
{
  const s = pres.addSlide();
  shell(s, "よくある質問 (FAQ)", 22);

  const qa = [
    ["準備は難しいですか？", "内装、システム導入まで本部がフルサポートするため、初めての方でも安心です。"],
    ["トラブル時の対応は？", "24時間体制の警備会社と遠隔監視システムにより、オーナー様が現場に駆けつける必要はほぼありません。また顧客とのLINEチャットを設けており、そこで質問対応・意見収集を行っています。"],
    ["人材採用については？", "新しく人材を採用する必要はありません。清掃管理・原料補充はご自身でもご対応可能ですが、業務委託として委託することも可能です。"],
    ["集客はどうすれば？", "数々のメディア取材や、本部によるSEO/MEO対策をはじめとしたWebマーケティングによって集客を支援します。"],
  ];
  const cw = 5.77, chh = 2.42, gx = 0.325, gy = 0.26;
  qa.forEach(([q, a], i) => {
    const x = M + (i % 2) * (cw + gx);
    const y = 1.50 + Math.floor(i / 2) * (chh + gy);
    card(s, x, y, cw, chh);
    chip(s, x + 0.30, y + 0.28, "Q", { size: 0.44, fontSize: 15 });
    s.addText(q, {
      x: x + 0.90, y: y + 0.28, w: cw - 1.20, h: 0.44,
      fontFace: F.jp, fontSize: 15, bold: true, color: C.green, valign: "middle", margin: 0,
    });
    s.addShape("rect", { x: x + 0.30, y: y + 0.92, w: cw - 0.60, h: 0.011, fill: { color: C.line } });
    chip(s, x + 0.30, y + 1.08, "A", { size: 0.44, fontSize: 15, fill: C.creamDk, color: C.brown });
    s.addText(a, {
      x: x + 0.90, y: y + 1.06, w: cw - 1.20, h: 1.10,
      fontFace: F.jp, fontSize: 11, color: C.ink, valign: "top", margin: 0, lineSpacingMultiple: 1.32,
    });
  });
}

/* =============================================================== p.23 万全のサポート体制 */
{
  const s = pres.addSlide();
  shell(s, "万全のサポート体制", 23);
  kicker(s, "開業前から開業後まで、本部が一貫して伴走します。", 1.36);

  const cols = [
    ["LuFileCheck", "開業前サポート", ["パートナー制度のご説明", "収益試算表の作成", "現地調査／店舗デザイン／設計", "工事業者手配", "運営研修"]],
    ["LuLifeBuoy", "開業後サポート", ["店舗運営相談", "マシンの定期メンテナンス", "WEBサイト作成／更新", "SEO対策（HP作成・定期更新／プレスリリース等）", "MEO対策（Google MAP設置・定期投稿／更新等）", "改修工事サポート"]],
  ];
  const cw = (CW - 0.36) / 2, chh = 3.86;
  cols.forEach(([ic, title, list], i) => {
    const x = M + i * (cw + 0.36), y = 1.80;
    card(s, x, y, cw, chh);
    iconBadge(s, ic, x + 0.34, y + 0.30, { d: 0.66 });
    s.addText(title, {
      x: x + 1.16, y: y + 0.30, w: cw - 1.50, h: 0.66,
      fontFace: F.jp, fontSize: 19, bold: true, color: C.green, valign: "middle", margin: 0,
    });
    s.addShape("rect", { x: x + 0.34, y: y + 1.10, w: cw - 0.68, h: 0.011, fill: { color: C.line } });
    list.forEach((t, j) => {
      const ry = y + 1.28 + j * 0.42;
      icon(s, "LuCheck", "green", x + 0.36, ry + 0.06, 0.20);
      s.addText(t, {
        x: x + 0.70, y: ry, w: cw - 1.04, h: 0.34,
        fontFace: F.jp, fontSize: 11.5, color: C.ink, valign: "middle", margin: 0,
      });
    });
  });

  panel(s, M, 5.90, CW, 0.94);
  icon(s, "LuUsers", "white", M + 0.42, 6.21, 0.32);
  s.addText("新たに人を採用・配置する必要はありません。", {
    x: M + 0.90, y: 5.90, w: CW - 1.32, h: 0.94,
    fontFace: F.jp, fontSize: 18, bold: true, color: C.white, valign: "middle", margin: 0,
  });
}

/* =============================================================== p.24 開業までの流れ */
{
  const s = pres.addSlide();
  shell(s, "開業までの流れ", 24);

  const steps = [
    ["LuUsers", "面談", "制度のご説明と\nご要望のヒアリング"],
    ["LuFileSearch", "加盟審査", "立地・条件を\n本部で審査"],
    ["LuPenLine", "加盟申込み", "契約締結と\n設置場所の確定"],
    ["LuHammer", "工事", "機器設置と\nシステム導入"],
    ["LuStore", "運営開始", "研修を経て\nオープン"],
  ];
  const gap = 0.42;
  const cw = (CW - 4 * gap) / 5, chh = 3.16;
  const cy = 3.14;

  // 「2〜3か月程度」のブラケット（加盟申込み〜運営開始）
  const bx = M + 2 * (cw + gap);
  const bw = CW - (bx - M);
  s.addShape("roundRect", {
    x: bx + bw / 2 - 1.10, y: 2.16, w: 2.20, h: 0.52, rectRadius: 0.26,
    fill: { color: C.greenTint }, line: { color: "CBDCCB", width: 0.75 },
  });
  icon(s, "LuClock", "green", bx + bw / 2 - 0.86, 2.29, 0.26);
  s.addText("2〜3か月程度", {
    x: bx + bw / 2 - 0.54, y: 2.16, w: 1.60, h: 0.52,
    fontFace: F.jp, fontSize: 12.5, bold: true, color: C.green, valign: "middle", margin: 0,
  });
  s.addShape("rect", { x: bx + 0.02, y: 2.68, w: 0.014, h: 0.34, fill: { color: "9DBBA0" } });
  s.addShape("rect", { x: bx + bw - 0.03, y: 2.68, w: 0.014, h: 0.34, fill: { color: "9DBBA0" } });
  s.addShape("rect", { x: bx + 0.02, y: 3.00, w: bw - 0.04, h: 0.014, fill: { color: "9DBBA0" } });

  steps.forEach(([ic, title, body], i) => {
    const x = M + i * (cw + gap);
    const last = i === steps.length - 1;
    if (last) panel(s, x, cy, cw, chh); else card(s, x, cy, cw, chh);
    chip(s, x + cw / 2 - 0.21, cy + 0.26, String(i + 1).padStart(2, "0"), {
      size: 0.42, fontSize: 12, fill: last ? "2C6B34" : C.green,
    });
    s.addShape("ellipse", {
      x: x + cw / 2 - 0.42, y: cy + 0.84, w: 0.84, h: 0.84,
      fill: { color: last ? "1D5624" : C.greenTint }, line: { type: "none" },
    });
    icon(s, ic, last ? "white" : "green", x + cw / 2 - 0.23, cy + 1.03, 0.46);
    s.addText(title, {
      x: x + 0.12, y: cy + 2.02, w: cw - 0.24, h: 0.40,
      fontFace: F.jp, fontSize: 15.5, bold: true, color: last ? C.white : C.green,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(body, {
      x: x + 0.10, y: cy + 2.44, w: cw - 0.20, h: 0.58,
      fontFace: F.jp, fontSize: 10, color: last ? C.greenPale : C.inkSub,
      align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.25,
    });
    if (!last) {
      s.addShape("triangle", {
        x: x + cw + 0.10, y: cy + chh / 2 - 0.11, w: 0.22, h: 0.22,
        fill: { color: "C9BFA6" }, line: { type: "none" }, rotate: 90,
      });
    }
  });

  s.addText("※工事内容・立地条件により期間は変動します。", {
    x: M, y: 6.52, w: CW, h: 0.24,
    fontFace: F.jp, fontSize: 9.5, color: C.muted, margin: 0, valign: "middle",
  });
}

/* =============================================================== p.25 お問い合わせ */
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  s.addText("25", {
    x: RIGHT - 1.0, y: 6.94, w: 1.0, h: 0.24,
    fontFace: "Cambria", fontSize: 13, color: "000000", align: "right", valign: "middle", margin: 0,
  });

  panel(s, M, 0.60, CW, 6.10);
  s.addImage({ path: path.join(ASSETS, "logo_white.png"), x: M + 0.80, y: 1.14, w: 2.66, h: 0.378 });
  s.addText("お問い合わせ", {
    x: M + 0.80, y: 1.82, w: 7.0, h: 0.72,
    fontFace: F.jp, fontSize: 40, bold: true, color: C.white, valign: "middle", margin: 0,
  });
  s.addText("会社名・ご氏名・ご連絡先をご記入の上、下記までご連絡ください。", {
    x: M + 0.80, y: 2.60, w: 7.1, h: 0.36,
    fontFace: F.jp, fontSize: 14, color: C.greenPale, valign: "middle", margin: 0,
  });
  s.addShape("rect", { x: M + 0.80, y: 3.20, w: 1.40, h: 0.021, fill: { color: C.greenPale } });

  [
    ["LuPhone", "お電話", "XXX-XXXX-XXXX（担当：XXX）"],
    ["LuMail", "メール", "aaaaa@aaaa.aaa"],
    ["LuGlobe", "公式HP", "https://selfcafe.jp/"],
  ].forEach(([ic, label, value], i) => {
    const y = 3.60 + i * 0.86;
    s.addShape("ellipse", { x: M + 0.80, y, w: 0.56, h: 0.56, fill: { color: "1D5624" }, line: { type: "none" } });
    icon(s, ic, "white", M + 0.96, y + 0.16, 0.24);
    s.addText(label, {
      x: M + 1.54, y: y, w: 1.16, h: 0.56,
      fontFace: F.jp, fontSize: 11, color: C.greenPale, valign: "middle", margin: 0,
    });
    s.addText(value, {
      x: M + 2.70, y: y, w: 5.3, h: 0.56,
      fontFace: F.num, fontSize: 17, bold: true, color: C.white, valign: "middle", margin: 0,
    });
  });

  photoSlot(s, M + 8.30, 1.14, 2.80, 5.02, "セルフカフェ 店内", {
    fill: "1D5624", line: "3E7A45", captionSize: 10, captionColor: "C6DCC8", img: "store-morioka.jpg",
  });
}

const out = path.join(__dirname, "selfcafe-fc-deck.pptx");
pres.writeFile({ fileName: out }).then(() => console.log("written:", out, fs.statSync(out).size, "bytes"));
