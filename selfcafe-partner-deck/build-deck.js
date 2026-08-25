/**
 * セルフカフェ パートナー制度（FC）資料 — 全27ページ
 *
 * デザインは「業務委託型制度」資料の実測値を踏襲（配色・座標・書体）。
 * 共通の実績数値は業務委託型資料に合わせている。
 * FC固有の条件（加盟金・開業資金・収益シミュレーション）はFC資料の値。
 *
 * 実行前に `node build-assets.js` を走らせること。
 */
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

// ---------------------------------------------------------------- palette（業務委託型資料の実測値）
const C = {
  bg: "FFFDF7",
  ink: "0A3D23", // 見出し
  green: "106838", // ブランド緑・数値
  greenDeep: "0C5230", // 濃緑パネル
  greenLabel: "1B8149", // セクション名
  gold: "A8761F",
  goldLine: "C79A44",
  goldTint: "FBF4E4",
  body: "4B5550",
  muted: "7E8882",
  grayText: "8E8B84", // 比較列
  footer: "A8B0AA",
  tint: "F1F8F3", // 薄緑面
  tintLine: "D2E7DA",
  warmLine: "EFEAE0",
  grayBand: "F4F2EE",
  white: "FFFFFF",
  midGreen: "A9C6B5",
  near: "1B211D",
  // 表紙
  cvPale: "8CC5A5",
  cvSub: "D6E3DB",
  cvBody: "AFC2B8",
  cvSmall: "8FA398",
};

const F = { jp: "Yu Gothic", num: "Arial" };

// ---------------------------------------------------------------- geometry
const M = 0.722; // 左マージン
const CW = 11.889; // コンテンツ幅（0.722 → 12.611）
const R = M + CW;
const TOP = 1.806; // コンテンツ開始
const BOT = 6.861; // コンテンツ下端

const A = path.join(__dirname, "assets");
const ICON = path.join(A, "icons");
const LOGO_G = path.join(A, "logo-green.png");
const LOGO_W = path.join(A, "logo-white.png");
const FOOTER = "セルフカフェ パートナー制度";

// ---------------------------------------------------------------- helpers
const pad2 = (n) => String(n).padStart(2, "0");

// ページ／セクションの自動採番（ページを挿入しても番号がズレない）
let PAGE = 0;
let SEC = 0;
/** 表紙・お問い合わせなど、セクション番号を持たないページ */
function bare() { PAGE++; return PAGE; }

function shell(s, secName, title, lead) {
  PAGE++; SEC++;
  const secNo = SEC, pageNo = PAGE;
  s.background = { color: C.bg };
  s.addShape("rect", { x: M, y: 0.597, w: 0.222, h: 0.028, fill: { color: C.green } });
  s.addText(
    [
      { text: pad2(secNo), options: { fontFace: F.num, fontSize: 10.5, bold: true, color: C.gold } },
      { text: "   " + secName, options: { fontFace: F.jp, fontSize: 10.5, bold: true, color: C.greenLabel } },
    ],
    { x: 1.042, y: 0.486, w: 7.0, h: 0.194, margin: 0, valign: "middle" }
  );
  s.addText(title, {
    x: M, y: 0.722, w: 9.722, h: 0.556,
    fontFace: F.jp, fontSize: 29, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
  if (lead) {
    s.addText(lead, {
      x: M, y: 1.431, w: 10.972, h: 0.25,
      fontFace: F.jp, fontSize: 12.5, color: C.body, margin: 0, valign: "middle",
    });
  }
  s.addImage({ path: LOGO_G, x: 10.778, y: 0.528, w: 1.833, h: 0.261 });
  s.addText(FOOTER, {
    x: M, y: 7.083, w: 5.556, h: 0.167,
    fontFace: F.jp, fontSize: 9, color: C.footer, margin: 0, valign: "middle",
  });
  s.addText(pad2(pageNo), {
    x: 11.222, y: 7.069, w: 1.389, h: 0.167,
    fontFace: F.num, fontSize: 10, bold: true, color: C.muted, align: "right", margin: 0, valign: "middle",
  });
}

/** 白カード */
function card(s, x, y, w, h, opts = {}) {
  s.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.06,
    fill: { color: opts.fill || C.white },
    line: { color: opts.line || C.warmLine, width: 0.75 },
  });
}
/** 薄緑カード */
function tintCard(s, x, y, w, h) {
  card(s, x, y, w, h, { fill: C.tint, line: C.tintLine });
}
/** 濃緑パネル */
function panel(s, x, y, w, h) {
  s.addShape("roundRect", { x, y, w, h, rectRadius: 0.06, fill: { color: C.greenDeep }, line: { type: "none" } });
}

function icon(s, name, tone, x, y, size) {
  s.addImage({ path: path.join(ICON, `${name}-${tone}.png`), x, y, w: size, h: size });
}

/** KPIチップ列。items=[{v,u,l}]、幅はCWを等分 */
function chipRow(s, y, items, opts = {}) {
  const gap = 0.167;
  const h = opts.h || 0.667;
  const w = (CW - gap * (items.length - 1)) / items.length;
  items.forEach((it, i) => {
    const x = M + i * (w + gap);
    const gold = !!it.gold;
    s.addShape("roundRect", {
      x, y, w, h, rectRadius: 0.05,
      fill: { color: gold ? C.goldTint : C.tint },
      line: { color: gold ? "E7D3A8" : C.tintLine, width: 0.75 },
    });
    s.addText(
      [
        { text: it.v, options: { fontFace: F.num, fontSize: opts.vSize || 20, bold: true, color: gold ? C.gold : C.green } },
        { text: it.u ? " " + it.u : "", options: { fontFace: F.jp, fontSize: 11, bold: true, color: gold ? C.gold : C.green } },
      ],
      { x: x + 0.181, y: y + 0.083, w: w - 0.36, h: 0.306, margin: 0, valign: "middle" }
    );
    s.addText(it.l, {
      x: x + 0.181, y: y + 0.402, w: w - 0.32, h: 0.181,
      fontFace: F.jp, fontSize: 9, color: C.muted, margin: 0, valign: "middle",
    });
  });
}

/** 写真枠。opts.img に assets/photos/ 内のファイル名を渡すと、
 *  ファイルが存在する場合は写真を挿入し、無ければ従来の点線プレースホルダを描く。
 *  写真の差し替えは assets/photos/ の同名ファイルを置き換えて再ビルドするだけでよい。 */
const PHOTOS = path.join(A, "photos");
const CROPDIR = path.join(PHOTOS, ".crops");
fs.mkdirSync(CROPDIR, { recursive: true });
/** 枠の縦横比に合わせて中央トリミングするジョブ一覧（書き出し前に sharp で処理） */
const CROPS = new Map();
function croppedPath(src, w, h, focus) {
  const base = path.basename(src, path.extname(src));
  const out = path.join(CROPDIR, `${base}-${Math.round(w * 100)}x${Math.round(h * 100)}.jpg`);
  CROPS.set(out, { src, aspect: w / h, focus: focus == null ? 0.5 : focus });
  return out;
}
function photoSlot(s, x, y, w, h, caption, opts = {}) {
  const img = opts.img ? path.join(PHOTOS, opts.img) : null;
  if (img && fs.existsSync(img)) {
    s.addImage({ path: croppedPath(img, w, h, opts.focus), x, y, w, h });
    s.addShape("rect", { x, y, w, h, fill: { type: "none" }, line: { color: "D8D2C6", width: 1 } });
    if (caption) {
      const ch = 0.36;
      const cw2 = Math.min(w - 0.28, 0.4 + caption.length * ((opts.capSize || 9.5) / 72) * 1.06);
      s.addShape("roundRect", {
        x: x + 0.14, y: y + h - ch - 0.14, w: cw2, h: ch, rectRadius: 0.05,
        fill: { color: C.white, transparency: 10 }, line: { type: "none" },
      });
      s.addText(caption, {
        x: x + 0.32, y: y + h - ch - 0.14, w: cw2 - 0.3, h: ch,
        fontFace: F.jp, fontSize: opts.capSize || 9.5, bold: true, color: C.ink,
        margin: 0, valign: "middle",
      });
    }
    return;
  }
  s.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.05,
    fill: { color: opts.fill || C.grayBand },
    line: { color: opts.line || "DED8CC", width: 1, dashType: "dash" },
  });
  const is = Math.min(0.52, h * 0.22);
  const cy = y + h / 2 - (caption ? 0.30 : is / 2);
  icon(s, "LuImage", "muted", x + w / 2 - is / 2, cy, is);
  s.addText("写真を挿入", {
    x: x + 0.08, y: cy + is + 0.04, w: w - 0.16, h: 0.2,
    fontFace: F.jp, fontSize: 9, color: opts.subColor || C.muted, align: "center", valign: "middle", margin: 0,
  });
  if (caption) {
    s.addText(caption, {
      x: x + 0.08, y: cy + is + 0.25, w: w - 0.16, h: 0.42,
      fontFace: F.jp, fontSize: opts.capSize || 9.5, bold: true, color: opts.capColor || C.ink,
      align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.15,
    });
  }
}

/** 見出し付き小ラベル（カード内の英字ラベル） */
function eyebrowIn(s, x, y, w, text, color) {
  s.addText(text, {
    x, y, w, h: 0.18,
    fontFace: F.num, fontSize: 9, bold: true, color: color || C.gold, charSpacing: 1.5, margin: 0, valign: "middle",
  });
}

/** 出典・注記 */
function note(s, y, text) {
  s.addText(text, {
    x: M, y, w: CW, h: 0.2,
    fontFace: F.jp, fontSize: 8.5, color: C.footer, margin: 0, valign: "middle",
  });
}

/** 表の共通スタイル */
const th = (t, opts = {}) => ({
  text: t,
  options: {
    fill: { color: opts.fill || C.greenDeep }, color: opts.color || C.white, bold: true,
    fontSize: 10.5, fontFace: F.jp, align: opts.align || "left", valign: "middle",
    margin: [0.05, 0.14, 0.05, 0.14],
  },
});
const tl = (t) => ({
  text: t,
  options: {
    fill: { color: C.grayBand }, color: C.ink, bold: true, fontSize: 10.5, fontFace: F.jp,
    valign: "middle", margin: [0.05, 0.14, 0.05, 0.14],
  },
});
const td = (t, opts = {}) => ({
  text: t,
  options: {
    fill: { color: opts.fill || C.white }, color: opts.color || C.body,
    bold: !!opts.bold, fontSize: opts.size || 10.5, fontFace: opts.face || F.jp,
    align: opts.align || "left", valign: "middle", margin: [0.05, 0.14, 0.05, 0.14],
  },
});

// ================================================================ deck
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "セルフカフェ株式会社";
pres.title = "セルフカフェ パートナー制度";

/* ===================================================== p1 表紙 */
{
  const s = pres.addSlide();
  bare();
  s.addImage({ path: path.join(A, "cover-bg.png"), x: 0, y: 0, w: 13.333, h: 7.5 });
  s.addImage({ path: LOGO_W, x: 0.889, y: 2.083, w: 2.861, h: 0.407 });

  s.addShape("rect", { x: 0.889, y: 3.083, w: 0.278, h: 0.028, fill: { color: C.goldLine } });
  s.addText("PARTNER PROGRAM", {
    x: 1.306, y: 2.986, w: 4.167, h: 0.194,
    fontFace: F.num, fontSize: 10.5, bold: true, color: C.cvPale, charSpacing: 2, margin: 0, valign: "middle",
  });
  s.addText("自社物件・遊休スペースを活かす、無人カフェFC", {
    x: 0.889, y: 3.278, w: 7.778, h: 0.306,
    fontFace: F.jp, fontSize: 15, color: C.cvSub, margin: 0, valign: "middle",
  });
  s.addText("パートナー制度", {
    x: 0.889, y: 3.556, w: 8.889, h: 1.028,
    fontFace: F.jp, fontSize: 54, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  s.addText("内装・システム導入から集客まで本部が伴走。\n低投資で始められる、無人カフェのフランチャイズ制度です。", {
    x: 0.889, y: 4.861, w: 7.5, h: 0.639,
    fontFace: F.jp, fontSize: 12.5, color: C.cvBody, margin: 0, valign: "top", lineSpacingMultiple: 1.35,
  });

  s.addShape("rect", { x: 0.889, y: 6.056, w: 0.778, h: 0.028, fill: { color: C.goldLine } });
  const stats = [
    { x: 0.889, w: 3.472, v: "71", u: "店舗", l: "全国展開（年内10店舗以上 出店予定）" },
    { x: 4.833, w: 2.083, v: "約10", u: "万人", l: "月間利用者数" },
    { x: 7.389, w: 2.5, v: "650", u: "万円〜", l: "初期費用（居抜き・20坪想定）" },
  ];
  stats.forEach((st) => {
    s.addText(
      [
        { text: st.v, options: { fontFace: F.num, fontSize: 26, bold: true, color: C.white } },
        { text: " " + st.u, options: { fontFace: F.jp, fontSize: 12, bold: true, color: C.white } },
      ],
      { x: st.x, y: 6.361, w: st.w, h: 0.417, margin: 0, valign: "middle" }
    );
    s.addText(st.l, {
      x: st.x, y: 6.806, w: st.w + 0.6, h: 0.181,
      fontFace: F.jp, fontSize: 9.5, color: C.cvSmall, margin: 0, valign: "middle",
    });
  });
  [4.389, 6.944].forEach((x) =>
    s.addShape("rect", { x, y: 6.389, w: 0.013, h: 0.583, fill: { color: "3C5F4B" } })
  );
  s.addText("セルフカフェ株式会社", {
    x: 9.722, y: 6.778, w: 2.889, h: 0.194,
    fontFace: F.jp, fontSize: 9.5, color: C.cvSmall, align: "right", margin: 0, valign: "middle",
  });
}

/* ===================================================== p2 会社概要 */
{
  const s = pres.addSlide();
  shell(s,  "会社概要", "会社概要", "無人カフェ「セルフカフェ」の企画・開発・運営を行っています。");

  // 実写バナー（元画像 1712x412 の横長比率に合わせて全幅で使用）
  photoSlot(s, M, TOP, CW, 2.78, "セルフカフェ 盛岡駅前店（2025年4月オープン）",
    { img: "p02-storefront.jpg", capSize: 10 });

  const iy = TOP + 3.0;
  const items = [
    ["会社名", "セルフカフェ株式会社"], ["代表取締役", "鈴木 大基"], ["資本金", "2,000万円"],
    ["事業内容", "無人カフェの経営"], ["設立日", "2024年5月1日"], ["公式HP", "https://selfcafe.jp/"],
  ];
  const cw = (CW - 2 * 0.5) / 3, rh = 0.82;
  items.forEach((it, i) => {
    const x = M + (i % 3) * (cw + 0.5);
    const y = iy + Math.floor(i / 3) * rh;
    s.addText(it[0], {
      x, y, w: cw, h: 0.2,
      fontFace: F.jp, fontSize: 9, bold: true, color: C.muted, margin: 0, valign: "middle",
    });
    s.addText(it[1], {
      x, y: y + 0.23, w: cw, h: 0.32,
      fontFace: F.jp, fontSize: 13, color: C.ink, margin: 0, valign: "middle",
    });
    s.addShape("rect", { x, y: y + 0.62, w: cw, h: 0.011, fill: { color: C.warmLine } });
  });
  s.addShape("roundRect", {
    x: M, y: iy + 2 * rh + 0.14, w: CW, h: 0.5, rectRadius: 0.05,
    fill: { color: C.tint }, line: { type: "none" },
  });
  s.addText("※ 2024年5月1日よりウッドデザインパーク株式会社から分社化", {
    x: M + 0.24, y: iy + 2 * rh + 0.14, w: CW - 0.48, h: 0.5,
    fontFace: F.jp, fontSize: 10, color: C.green, margin: 0, valign: "middle",
  });
}

/* ===================================================== p3 社会背景と市場性 */
{
  const s = pres.addSlide();
  shell(s,  "市場背景", "社会背景と市場性", "現代社会が抱える3つの課題が、無人カフェの需要をつくっています。");

  const items = [
    ["LuUsers", "深刻な労働力不足",
      "少子高齢化と働き手の価値観の変化でサービス業の人手確保が難しくなり、人件費が上昇して経営負担が増している。",
      "人に依存しない店舗フォーマットが必要"],
    ["LuLaptop", "サードプレイスの枯渇",
      "リモートワークの普及で自宅以外の手頃な作業場所が不足し、カフェやコワーキングは価格や混雑の課題が残り、学習スペースも制限が多く気軽に利用できない。",
      "手頃で気軽に使える作業場所への需要"],
    ["LuBuilding2", "遊休資産の有効活用",
      "ビル空室やデッドスペースが増え、従来のテナント誘致だけでは埋まらず、新たな活用モデルの開発が求められている。",
      "空きスペースを収益化する新しいモデル"],
  ];
  const w = (CW - 2 * 0.28) / 3, h = 4.0;
  items.forEach((it, i) => {
    const x = M + i * (w + 0.28), y = TOP;
    card(s, x, y, w, h);
    eyebrowIn(s, x + 0.28, y + 0.3, 2.0, `ISSUE ${pad2(i + 1)}`);
    icon(s, it[0], "green", x + w - 0.86, y + 0.26, 0.44);
    s.addText(it[1], {
      x: x + 0.28, y: y + 0.62, w: w - 0.56, h: 0.42,
      fontFace: F.jp, fontSize: 16, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addShape("rect", { x: x + 0.28, y: y + 1.14, w: w - 0.56, h: 0.011, fill: { color: C.warmLine } });
    s.addText(it[2], {
      x: x + 0.28, y: y + 1.28, w: w - 0.56, h: 1.7,
      fontFace: F.jp, fontSize: 10.5, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.45,
    });
    s.addShape("roundRect", {
      x: x + 0.28, y: y + h - 0.86, w: w - 0.56, h: 0.56, rectRadius: 0.05,
      fill: { color: C.tint }, line: { type: "none" },
    });
    icon(s, "LuTrendingUp", "green", x + 0.44, y + h - 0.71, 0.26);
    s.addText(it[3], {
      x: x + 0.78, y: y + h - 0.86, w: w - 1.06, h: 0.56,
      fontFace: F.jp, fontSize: 10, bold: true, color: C.green, margin: 0, valign: "middle",
    });
  });

  note(s, 6.62, "※ 課題認識は当社の市場理解に基づく整理です。");
}

/* ===================================================== p4 人手不足対策 */
{
  const s = pres.addSlide();
  shell(s,  "優位性", "人手不足対策：無人管理の優位性", "運営から「人」の変数を外すと、コスト・品質・安定性が同時に改善します。");

  const pw = 4.6;
  panel(s, M, TOP, pw, 4.72);
  eyebrowIn(s, M + 0.44, TOP + 0.44, 2.4, "CORE CONCEPT", C.cvPale);
  s.addText("「人」に\n頼らない経営へ", {
    x: M + 0.44, y: TOP + 0.86, w: pw - 0.88, h: 1.5,
    fontFace: F.jp, fontSize: 28, bold: true, color: C.white, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
  });
  s.addShape("rect", { x: M + 0.44, y: TOP + 2.5, w: 0.78, h: 0.028, fill: { color: C.goldLine } });
  s.addText("スタッフの採用・教育・シフト管理という、店舗運営でもっとも重い変動要素を仕組みで置き換えます。", {
    x: M + 0.44, y: TOP + 2.72, w: pw - 0.88, h: 1.3,
    fontFace: F.jp, fontSize: 11.5, color: C.cvBody, margin: 0, valign: "top", lineSpacingMultiple: 1.5,
  });
  icon(s, "LuUsers", "pale", M + pw - 1.14, TOP + 3.72, 0.62);

  const cx = M + pw + 0.42, cw = R - cx, ch = 1.44;
  const items = [
    ["LuClipboardList", "管理・教育コストの大幅削減", "スタッフ管理に伴う工数や人的リスクを抑え、運営にかかる負担を軽減します。"],
    ["LuPiggyBank", "省人力運営で、無駄な人件費をカット", "必要最小限の作業で回る仕組みにより、人件費を抑えながら安定した収益を確保できます。"],
    ["LuBadgeCheck", "安定した収益と品質", "属人性を排除することで、常に均一したサービス品質を維持し、運営の安定性を高めます。"],
  ];
  items.forEach((it, i) => {
    const y = TOP + i * (ch + 0.20);
    card(s, cx, y, cw, ch);
    s.addText(String(i + 1), {
      x: cx + 0.3, y: y + 0.3, w: 0.4, h: 0.4,
      fontFace: F.num, fontSize: 22, bold: true, color: C.midGreen, margin: 0, valign: "middle",
    });
    s.addText(it[1], {
      x: cx + 0.82, y: y + 0.28, w: cw - 1.6, h: 0.4,
      fontFace: F.jp, fontSize: 15, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(it[2], {
      x: cx + 0.82, y: y + 0.72, w: cw - 1.2, h: 0.52,
      fontFace: F.jp, fontSize: 10.5, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.35,
    });
    icon(s, it[0], "green", cx + cw - 0.72, y + 0.3, 0.36);
  });
}

/* ===================================================== p5 不動産活用（FC固有） */
{
  const s = pres.addSlide();
  shell(s,  "不動産活用", "持ち物件オーナー様へ", "自社物件なら家賃コストがゼロ。売上をダイレクトに利益へ繋げられます。");

  chipRow(s, TOP, [
    { v: "0", u: "円", l: "自社物件なら新規の家賃負担なし" },
    { v: "20", u: "坪〜", l: "必要面積の目安（40席想定）" },
    { v: "650", u: "万円〜", l: "初期費用（居抜き・20坪想定）" },
    { v: "15", u: "分／日", l: "日常のオーナー業務（清掃・補充）" },
  ]);

  const y2 = TOP + 0.92;
  const cw = (CW - 0.42) / 2;
  const items = [
    ["LuCoins", "高い営業利益率", "自社物件なら「家賃コスト」がゼロ。固定費を抑え、売上をダイレクトに利益へ繋げます。",
      ["家賃という最大の固定費が不要", "空室期間の機会損失を収益に転換", "テナント誘致に頼らない自社運用"]],
    ["LuShoppingBag", "相乗効果", "既存事業の「待ち時間」解消や、新規集客のフックとして、カフェは最適の組み合わせです。",
      ["書店・美容室・整体院・病院の待合と好相性", "「カフェ」で検索されて新規客に見つかる", "滞在時間が伸び、ついで買いが生まれる"]],
  ];
  items.forEach((it, i) => {
    const x = M + i * (cw + 0.42);
    card(s, x, y2, cw, 2.44);
    icon(s, it[0], "green", x + 0.3, y2 + 0.3, 0.42);
    s.addText(it[1], {
      x: x + 0.88, y: y2 + 0.28, w: cw - 1.18, h: 0.44,
      fontFace: F.jp, fontSize: 17, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(it[2], {
      x: x + 0.3, y: y2 + 0.82, w: cw - 0.6, h: 0.6,
      fontFace: F.jp, fontSize: 11, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.4,
    });
    s.addShape("rect", { x: x + 0.3, y: y2 + 1.48, w: cw - 0.6, h: 0.011, fill: { color: C.warmLine } });
    it[3].forEach((b, j) => {
      const by = y2 + 1.6 + j * 0.26;
      icon(s, "LuCheck", "green", x + 0.32, by + 0.03, 0.18);
      s.addText(b, {
        x: x + 0.6, y: by, w: cw - 0.9, h: 0.24,
        fontFace: F.jp, fontSize: 10, color: C.ink, margin: 0, valign: "middle",
      });
    });
  });

  const phw = (CW - 0.3) / 2;
  photoSlot(s, M, y2 + 2.66, phw, 1.62, "自社物件への設置イメージ", { img: "p05-install.jpg", capSize: 10 });
  photoSlot(s, M + phw + 0.3, y2 + 2.66, phw, 1.62, "店舗外観（盛岡駅前店）", { img: "p02-storefront.jpg", capSize: 10 });
}

/* ===================================================== 活用可能性（新規／佐藤提案） */
{
  const s = pres.addSlide();
  shell(s, "活用可能性", "カフェの先にある、場所の可能性",
    "「集客できる」で終わらせず、その場所で何ができるかを一緒に設計します。");

  const items = [
    ["LuUsers", "USE 01", "まちづくり・地域コミュニティ",
      "「人が集まる場所」として、地域の拠点になります。",
      ["世代を越えて人が集まる常設スペース", "自治体・地域団体との連携の受け皿に"]],
    ["LuGraduationCap", "USE 02", "セミナー・イベント会場",
      "営業時間外や一部区画を、貸しスペースとして二次活用できます。",
      ["勉強会・セミナー・ワークショップ", "席・Wi-Fi・電源をそのまま使える"]],
    ["LuShoppingBag", "USE 03", "物販スペース",
      "棚や壁面を使い、物販の売り場としても活用できます。",
      ["地域産品・書籍・自社商品の販売", "滞在が長いぶん、商品が目に触れる"]],
  ];
  const w = (CW - 2 * 0.28) / 3, h = 2.62;
  items.forEach((it, i) => {
    const x = M + i * (w + 0.28), y = TOP;
    card(s, x, y, w, h);
    eyebrowIn(s, x + 0.3, y + 0.3, 2.0, it[1]);
    icon(s, it[0], "green", x + w - 0.84, y + 0.26, 0.42);
    s.addText(it[2], {
      x: x + 0.3, y: y + 0.58, w: w - 0.6, h: 0.42,
      fontFace: F.jp, fontSize: 15.5, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(it[3], {
      x: x + 0.3, y: y + 1.06, w: w - 0.6, h: 0.56,
      fontFace: F.jp, fontSize: 10.5, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.4,
    });
    s.addShape("rect", { x: x + 0.3, y: y + 1.7, w: w - 0.6, h: 0.011, fill: { color: C.warmLine } });
    it[4].forEach((b, j) => {
      const by = y + 1.84 + j * 0.32;
      icon(s, "LuCheck", "green", x + 0.32, by + 0.04, 0.19);
      s.addText(b, {
        x: x + 0.6, y: by, w: w - 0.9, h: 0.28,
        fontFace: F.jp, fontSize: 10, color: C.ink, margin: 0, valign: "middle",
      });
    });
  });

  // 下段：地方都市の優位性（全幅パネル）
  const by = TOP + h + 0.24, bh = 1.94;
  panel(s, M, by, CW, bh);
  eyebrowIn(s, M + 0.36, by + 0.26, 2.4, "LOCAL ADVANTAGE", C.cvPale);
  s.addText("地方都市ほど、競合がいない。", {
    x: M + 0.36, y: by + 0.56, w: 6.4, h: 0.5,
    fontFace: F.jp, fontSize: 22, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  icon(s, "LuMapPin", "pale", M + 6.6, by + 0.58, 0.46);
  s.addText("都市部に比べて同種の施設が少なく、開店直後から認知が広がりやすい傾向があります。\n岩手県 盛岡駅前店は1日80杯ペースで稼働しています（出店事例のページ参照）。", {
    x: M + 7.4, y: by + 0.34, w: CW - 7.8, h: 1.36,
    fontFace: F.jp, fontSize: 11, color: C.cvBody, margin: 0, valign: "middle", lineSpacingMultiple: 1.5,
  });
  s.addShape("rect", { x: M + 7.1, y: by + 0.36, w: 0.013, h: bh - 0.72, fill: { color: "2A7A4E" } });

  note(s, 6.72, "※ 二次利用（貸しスペース・物販）の可否は立地・契約条件により異なります。個別にご相談ください。");
}

/* ===================================================== p6 セルフカフェとは */
{
  const s = pres.addSlide();
  shell(s,  "ブランド", "セルフカフェとは",
    "“誰もが気軽に使える”第三の居場所を、全国71店舗で展開中。年内にさらに10店舗以上の出店を予定しています。");

  const lw = 7.3;
  const rows = [
    ["LuHouse", "お家のような", "寛ぎすぎないが、設備が整った集中環境。"],
    ["LuCoffee", "カフェのような", "食事はないが、時間を気にせず美味しいコーヒーが飲める。"],
    ["LuBookOpen", "図書館のような", "静かすぎない、適度な賑やかさがある。"],
  ];
  rows.forEach((r, i) => {
    const y = TOP + i * 0.98;
    card(s, M, y, lw, 0.86);
    icon(s, r[0], "green", M + 0.28, y + 0.22, 0.42);
    s.addText(r[1], {
      x: M + 0.88, y: y + 0.14, w: 2.0, h: 0.3,
      fontFace: F.jp, fontSize: 14, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(r[2], {
      x: M + 0.88, y: y + 0.44, w: lw - 1.2, h: 0.28,
      fontFace: F.jp, fontSize: 10.5, color: C.body, margin: 0, valign: "middle",
    });
  });

  panel(s, M, TOP + 3.0, lw, 1.06);
  s.addText("“誰もが気軽に使える”", {
    x: M + 0.36, y: TOP + 3.12, w: lw - 0.72, h: 0.3,
    fontFace: F.jp, fontSize: 11.5, color: C.cvPale, margin: 0, valign: "middle",
  });
  s.addText("第三の居場所", {
    x: M + 0.36, y: TOP + 3.44, w: lw - 1.2, h: 0.5,
    fontFace: F.jp, fontSize: 24, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  icon(s, "LuCoffee", "pale", M + lw - 1.0, TOP + 3.4, 0.52);

  const st = [["2〜3時間", "平均滞在時間"], ["2杯程度", "1人あたり購入数"], ["24時間", "営業も設定可能"]];
  const sw = (lw - 0.32) / 3;
  st.forEach((t, i) => {
    const x = M + i * (sw + 0.16), y = TOP + 4.24;
    tintCard(s, x, y, sw, 0.72);
    s.addText(t[0], {
      x: x + 0.18, y: y + 0.08, w: sw - 0.36, h: 0.32,
      fontFace: F.jp, fontSize: 15, bold: true, color: C.green, margin: 0, valign: "middle",
    });
    s.addText(t[1], {
      x: x + 0.18, y: y + 0.4, w: sw - 0.36, h: 0.22,
      fontFace: F.jp, fontSize: 9, color: C.muted, margin: 0, valign: "middle",
    });
  });

  // 右：展開エリア
  const rx = M + lw + 0.42, rw = R - rx;
  card(s, rx, TOP, rw, 4.96);
  eyebrowIn(s, rx + 0.3, TOP + 0.28, 2.0, "NETWORK");
  s.addText(
    [
      { text: "71", options: { fontFace: F.num, fontSize: 40, bold: true, color: C.green } },
      { text: " 店舗", options: { fontFace: F.jp, fontSize: 14, bold: true, color: C.green } },
    ],
    { x: rx + 0.3, y: TOP + 0.54, w: rw - 0.6, h: 0.6, margin: 0, valign: "middle" }
  );
  s.addText("2026年7月時点／11都府県で展開中", {
    x: rx + 0.3, y: TOP + 1.16, w: rw - 0.6, h: 0.22,
    fontFace: F.jp, fontSize: 9, color: C.muted, margin: 0, valign: "middle",
  });
  s.addShape("rect", { x: rx + 0.3, y: TOP + 1.48, w: rw - 0.6, h: 0.011, fill: { color: C.warmLine } });
  const prefs = [["愛知県", 36], ["大阪府", 18], ["東京都", 7], ["岩手県", 2], ["岐阜県", 2], ["埼玉県", 1], ["千葉県", 1], ["静岡県", 1], ["滋賀県", 1], ["三重県", 1], ["鳥取県", 1]];
  prefs.forEach((p, i) => {
    const y = TOP + 1.58 + i * 0.232;
    s.addText(p[0], {
      x: rx + 0.3, y, w: 1.8, h: 0.23,
      fontFace: F.jp, fontSize: 9.5, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(String(p[1]), {
      x: rx + rw - 1.1, y, w: 0.8, h: 0.23,
      fontFace: F.num, fontSize: 10.5, bold: true, color: C.green, align: "right", margin: 0, valign: "middle",
    });
  });
  s.addShape("roundRect", {
    x: rx + 0.3, y: TOP + 4.2, w: rw - 0.6, h: 0.56, rectRadius: 0.05,
    fill: { color: C.goldTint }, line: { type: "none" },
  });
  s.addText("出店予定：年内にさらに10店舗以上", {
    x: rx + 0.46, y: TOP + 4.2, w: rw - 0.9, h: 0.56,
    fontFace: F.jp, fontSize: 10, bold: true, color: C.gold, margin: 0, valign: "middle",
  });
}

/* ===================================================== p7 店舗展開数の推移（FCに無かったページ） */
{
  const s = pres.addSlide();
  shell(s,  "実績データ", "店舗展開数の推移",
    "2022年9月の1号店OPENから、驚異的なスピードで出店エリアを拡大。出店は今後も継続していきます。");

  // 業務委託型資料と同一の系列（2022.09〜2026.07、営業中店舗の開店月ベース累計）
  const series = [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 4, 8, 8, 9, 9, 12, 17, 19, 21, 22, 23,
    23, 25, 28, 31, 32, 34, 39, 42, 44, 44, 46, 47, 48, 48, 50, 50, 51, 55, 58, 61, 64, 67, 71];
  const labels = series.map((_, i) =>
    i === 0 ? "2022.09" : i === 12 ? "2023.09" : i === 24 ? "2024.09" : i === 36 ? "2025.09" : i === 46 ? "2026.07" : ""
  );

  const bx = M, by = TOP, bw = CW, bh = 3.06;
  card(s, bx, by, bw, bh);
  s.addText("営業中店舗の開店月ベース累計", {
    x: bx + 0.3, y: by + 0.2, w: 5.0, h: 0.26,
    fontFace: F.jp, fontSize: 11, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
  s.addText("2022.09 〜 2026.07（単位：店舗）", {
    x: bx + 0.3, y: by + 0.46, w: 5.0, h: 0.22,
    fontFace: F.jp, fontSize: 9, color: C.muted, margin: 0, valign: "middle",
  });
  s.addShape("roundRect", {
    x: bx + bw - 4.06, y: by + 0.22, w: 3.76, h: 0.46, rectRadius: 0.23,
    fill: { color: C.gold }, line: { type: "none" },
  });
  s.addText(
    [
      { text: "サービス開始以来、", options: { fontFace: F.jp, fontSize: 11, bold: true, color: C.white } },
      { text: "47", options: { fontFace: F.num, fontSize: 15, bold: true, color: C.white } },
      { text: " ヶ月で ", options: { fontFace: F.jp, fontSize: 11, bold: true, color: C.white } },
      { text: "71", options: { fontFace: F.num, fontSize: 15, bold: true, color: C.white } },
      { text: " 店舗", options: { fontFace: F.jp, fontSize: 11, bold: true, color: C.white } },
    ],
    { x: bx + bw - 4.06, y: by + 0.22, w: 3.76, h: 0.46, align: "center", margin: 0, valign: "middle" }
  );
  s.addChart(pres.ChartType.area, [{ name: "店舗数", labels, values: series }], {
    x: bx + 0.16, y: by + 0.78, w: bw - 0.4, h: bh - 1.0,
    chartColors: [C.green], chartColorsOpacity: 32, lineSize: 3.5, lineSmooth: false,
    showLegend: false, showTitle: false, showValue: false,
    catAxisLabelColor: C.muted, catAxisLabelFontFace: F.num, catAxisLabelFontSize: 9,
    catAxisLabelFrequency: 1, catAxisLabelRotate: 0, catAxisMultiLevelLabels: false,
    catAxisLineShow: false, catGridLine: { style: "none" },
    valAxisLabelColor: C.muted, valAxisLabelFontFace: F.num, valAxisLabelFontSize: 9,
    valAxisMinVal: 0, valAxisMaxVal: 75, valAxisMajorUnit: 25,
    valAxisLineShow: false, valGridLine: { color: C.warmLine, size: 0.75 },
  });

  // 成長を示す右上方向の矢印（グラフ上のオーバーレイ）
  s.addShape("line", {
    x: bx + bw * 0.30, y: by + 1.06, w: bw * 0.56, h: 1.36, flipV: true,
    line: { color: C.gold, width: 4.5, endArrowType: "triangle" },
  });
  // 現在地の強調：最終点マーカー＋コールアウト
  const cxp = bx + 0.16 + (bw - 0.4) - 0.47;
  const cyp = by + 0.78 + 0.10 + (1 - 71 / 75) * (bh - 1.0 - 0.42);
  s.addShape("ellipse", { x: cxp - 0.09, y: cyp - 0.09, w: 0.18, h: 0.18, fill: { color: C.gold }, line: { color: C.white, width: 1.75 } });
  s.addShape("roundRect", {
    x: cxp - 1.82, y: cyp + 0.18, w: 1.68, h: 0.4, rectRadius: 0.05,
    fill: { color: C.greenDeep }, line: { type: "none" },
  });
  s.addText(
    [
      { text: "現在 ", options: { fontFace: F.jp, fontSize: 9, color: C.cvPale } },
      { text: "71", options: { fontFace: F.num, fontSize: 14, bold: true, color: C.white } },
      { text: " 店舗", options: { fontFace: F.jp, fontSize: 9.5, bold: true, color: C.white } },
    ],
    { x: cxp - 1.82, y: cyp + 0.18, w: 1.68, h: 0.4, align: "center", margin: 0, valign: "middle" }
  );

  const ms = [["2022.09", "1", "店舗"], ["2023.09", "2", "店舗"], ["2024.09", "23", "店舗"],
    ["2025.09", "48", "店舗"], ["2026.07", "71", "店舗"], ["年内", "＋10", "店舗以上"]];
  const mw = (CW - 5 * 0.16) / 6, my = by + bh + 0.22;
  ms.forEach((m, i) => {
    const x = M + i * (mw + 0.16);
    const last = i === 5;
    s.addShape("roundRect", {
      x, y: my, w: mw, h: 0.92, rectRadius: 0.05,
      fill: { color: last ? C.goldTint : C.tint }, line: { color: last ? "E7D3A8" : C.tintLine, width: 0.75 },
    });
    s.addText(m[0], {
      x, y: my + 0.12, w: mw, h: 0.2,
      fontFace: F.num, fontSize: 9, bold: true, color: last ? C.gold : C.muted, align: "center", margin: 0, valign: "middle",
    });
    s.addText(
      [
        { text: m[1], options: { fontFace: F.num, fontSize: 19, bold: true, color: last ? C.gold : C.ink } },
        { text: " " + m[2], options: { fontFace: F.jp, fontSize: 9.5, bold: true, color: last ? C.gold : C.ink } },
      ],
      { x, y: my + 0.34, w: mw, h: 0.42, align: "center", margin: 0, valign: "middle" }
    );
  });

  s.addText(
    [
      { text: "現在は ", options: { color: C.ink } },
      { text: "11都府県・全国71店舗", options: { color: C.green, bold: true } },
      { text: "。年内にさらに ", options: { color: C.ink } },
      { text: "10店舗以上", options: { color: C.green, bold: true } },
      { text: " の出店を予定。", options: { color: C.ink } },
    ],
    { x: M, y: my + 1.06, w: CW, h: 0.26, fontFace: F.jp, fontSize: 12, margin: 0, valign: "middle" }
  );
  note(s, my + 1.36, "出典：社内管理台帳（営業中店舗の開店月ベース累計）／2026年7月時点。「年内」は今後の出店予定です。");
}

/* ===================================================== p8 月間利用者数の推移 */
{
  const s = pres.addSlide();
  shell(s,  "実績データ", "月間利用者数の推移",
    "ブランドの認知拡大とともに、月間の利用者数は約10万人規模まで拡大しました。");

  const labels = ["23.01", "23.07", "24.01", "24.07", "25.01", "25.07", "26.01", "26.07"];
  const values = [895, 1419, 9722, 26383, 46017, 74260, 80293, 101343];
  const approx = ["約900人", "約1,400人", "約9,700人", "約2.6万人", "約4.6万人", "約7.4万人", "約8.0万人", "約10万人"];
  // 薄→濃のグラデーションで直近を強調（最終棒はブランド緑）
  const ramp = ["D8E8DE", "CCE0D4", "B8D4C2", "9CC3AA", "7BAE8F", "579970", "2E824F", "106838"];

  const bw = 8.4, bh = 4.06;
  card(s, M, TOP, bw, bh);
  s.addText("全社月間利用者数（各月の合計）／店舗数の拡大に比例して拡大", {
    x: M + 0.3, y: TOP + 0.2, w: bw - 0.6, h: 0.26,
    fontFace: F.jp, fontSize: 11, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
  s.addText("2023.01 〜 2026.07（半期ごとの推移）", {
    x: M + 0.3, y: TOP + 0.46, w: bw - 0.6, h: 0.22,
    fontFace: F.jp, fontSize: 9, color: C.muted, margin: 0, valign: "middle",
  });
  // 自前描画の棒グラフ：直近の伸びを強調するため高さは非線形（値^1.3）でスケール
  const px0 = M + 0.42, pw2 = bw - 0.84;
  const baseY = TOP + 3.5, Hmax = 2.42;
  const slotW = pw2 / values.length, barW = 0.64;
  s.addShape("rect", { x: px0 - 0.06, y: baseY, w: pw2 + 0.12, h: 0.014, fill: { color: C.warmLine } });
  values.forEach((v, i) => {
    const hgt = Math.max(0.035, Hmax * Math.pow(v / values[7], 1.3));
    const x = px0 + i * slotW + (slotW - barW) / 2;
    const last = i === values.length - 1;
    s.addShape("rect", { x, y: baseY - hgt, w: barW, h: hgt, fill: { color: ramp[i] }, line: { type: "none" } });
    if (last) {
      s.addShape("roundRect", {
        x: x + barW / 2 - 1.0, y: baseY - hgt - 0.58, w: 2.0, h: 0.5, rectRadius: 0.25,
        fill: { color: C.gold }, line: { type: "none" },
      });
      s.addText(
        [
          { text: "現在 ", options: { fontFace: F.jp, fontSize: 11, color: C.white } },
          { text: "約10万人", options: { fontFace: F.jp, fontSize: 16, bold: true, color: C.white } },
        ],
        { x: x + barW / 2 - 1.0, y: baseY - hgt - 0.58, w: 2.0, h: 0.5, align: "center", margin: 0, valign: "middle" }
      );
    } else {
      s.addText(approx[i], {
        x: x + barW / 2 - 0.75, y: baseY - hgt - 0.3, w: 1.5, h: 0.24,
        fontFace: F.jp, fontSize: 10, bold: true, color: i >= 4 ? C.green : C.muted, align: "center", margin: 0, valign: "middle",
      });
    }
    s.addText(labels[i], {
      x: px0 + i * slotW, y: baseY + 0.08, w: slotW, h: 0.22,
      fontFace: F.num, fontSize: 10.5, bold: true, color: C.muted, align: "center", margin: 0, valign: "middle",
    });
  });

  // 出店ストーリーの注釈（1号店 22.09／2号店 23.09）
  // コネクタは数値ラベルを避けて左にずらし、基線上の金色ドットに落とす
  const anns = [
    { ax: px0 + 0.14, y: baseY - 1.06, lines: [
      ["22.09", { fontFace: F.num, fontSize: 9.5, bold: true, color: C.gold }],
      ["1号店・ささしまライブ店 OPEN", { fontFace: F.jp, fontSize: 10.5, bold: true, color: C.ink }],
    ]},
    { ax: px0 + 1.89, y: baseY - 1.98, lines: [
      ["23.09", { fontFace: F.num, fontSize: 9.5, bold: true, color: C.gold }],
      ["2店舗目・千種店 OPEN", { fontFace: F.jp, fontSize: 10.5, bold: true, color: C.ink }],
      ["ここから店舗展開開始", { fontFace: F.jp, fontSize: 10, bold: true, color: C.green }],
    ]},
  ];
  anns.forEach((an) => {
    an.lines.forEach((ln, j) => {
      s.addText(ln[0], {
        x: an.ax + 0.14, y: an.y + j * 0.22, w: 2.6, h: 0.2,
        ...ln[1], margin: 0, valign: "middle",
      });
    });
    const yb = an.y + an.lines.length * 0.22 + 0.05;
    s.addShape("line", {
      x: an.ax, y: yb, w: 0, h: baseY - 0.07 - yb,
      line: { color: C.gold, width: 1, dashType: "dash" },
    });
    s.addShape("ellipse", { x: an.ax - 0.04, y: baseY - 0.075, w: 0.08, h: 0.08, fill: { color: C.gold }, line: { type: "none" } });
  });

  // 成長を示す右上方向の矢印（ラベルの間を抜ける軌道）
  s.addShape("line", {
    x: px0 + 2.62, y: baseY - 1.66, w: 2.14, h: 0.86, flipV: true,
    line: { color: C.gold, width: 4.5, endArrowType: "triangle" },
  });

  const rx = M + bw + 0.42, rw = R - rx;
  const tiles = [
    ["約10", "万人", "2026年7月の月間利用者数"],
    ["約100", "倍", "3年半での月間利用者数の伸び"],
    ["1,400", "人前後", "1店舗あたりの月間利用者数"],
  ];
  tiles.forEach((t, i) => {
    const y = TOP + i * 1.42;
    const dark = i === 0;
    if (dark) panel(s, rx, y, rw, 1.24); else tintCard(s, rx, y, rw, 1.24);
    s.addText(
      [
        { text: t[0], options: { fontFace: F.num, fontSize: 26, bold: true, color: dark ? C.white : C.green } },
        { text: " " + t[1], options: { fontFace: F.jp, fontSize: 12, bold: true, color: dark ? C.cvPale : C.green } },
      ],
      { x: rx + 0.26, y: y + 0.2, w: rw - 0.52, h: 0.44, margin: 0, valign: "middle" }
    );
    s.addText(t[2], {
      x: rx + 0.26, y: y + 0.72, w: rw - 0.52, h: 0.4,
      fontFace: F.jp, fontSize: 10, color: dark ? C.cvBody : C.muted, margin: 0, valign: "top", lineSpacingMultiple: 1.2,
    });
  });

  note(s, 6.6, "出典：社内管理台帳（各月の全社利用者数）／2026年7月時点。数値は概数、棒の高さは成長イメージを強調した表現です。");
}

/* ===================================================== p9 利用システム */
{
  const s = pres.addSlide();
  shell(s,  "利用システム", "利用システム：究極の利便性と価格",
    "ドリンク1杯で、時間制限なし。だから毎日使われ、リピートが積み上がります。");

  const items = [
    ["LuCoffee", "420円〜", "1杯あたり", "ドリンク一杯で入店",
      "ドリンク購入だけで、Wi-Fi・電源が完備された空間を自由にご利用いただけます。"],
    ["LuClock", "0円", "延長・追加料金", "滞在時間無制限",
      "追加料金や延長料金を気にせず、仕事や勉強に没頭できる環境を提供します。"],
    ["LuSmartphone", "147円〜", "サブスク1杯換算", "サブスクプラン",
      "ヘビーユーザー向けに、月額定額で通い放題のプランも用意。継続利用を促しやすい仕組みです。"],
  ];
  const w = (CW - 2 * 0.28) / 3, h = 2.7;
  items.forEach((it, i) => {
    const x = M + i * (w + 0.28), y = TOP;
    card(s, x, y, w, h);
    icon(s, it[0], "green", x + 0.3, y + 0.3, 0.44);
    s.addText(
      [
        { text: it[1], options: { fontFace: F.num, fontSize: 24, bold: true, color: C.green } },
      ],
      { x: x + w - 2.3, y: y + 0.26, w: 2.0, h: 0.42, align: "right", margin: 0, valign: "middle" }
    );
    s.addText(it[2], {
      x: x + w - 2.3, y: y + 0.68, w: 2.0, h: 0.2,
      fontFace: F.jp, fontSize: 8.5, color: C.muted, align: "right", margin: 0, valign: "middle",
    });
    s.addText(it[3], {
      x: x + 0.3, y: y + 1.06, w: w - 0.6, h: 0.4,
      fontFace: F.jp, fontSize: 16, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addShape("rect", { x: x + 0.3, y: y + 1.54, w: w - 0.6, h: 0.011, fill: { color: C.warmLine } });
    s.addText(it[4], {
      x: x + 0.3, y: y + 1.68, w: w - 0.6, h: 0.86,
      fontFace: F.jp, fontSize: 10.5, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.42,
    });
  });

  const feats = [
    ["LuCalendarCheck", "無人・年中無休"], ["LuWifi", "高速Wi-Fi"], ["LuPlug", "電源完備"],
    ["LuClock", "席時間無制限"], ["LuBookOpen", "仕事・勉強"], ["LuMessagesSquare", "会話OK"],
  ];
  const fy = TOP + h + 0.3;
  const fw = (CW - 5 * 0.16) / 6;
  feats.forEach(([ic, label], i) => {
    const x = M + i * (fw + 0.16);
    tintCard(s, x, fy, fw, 1.0);
    icon(s, ic, "green", x + fw / 2 - 0.19, fy + 0.18, 0.38);
    s.addText(label, {
      x: x + 0.04, y: fy + 0.62, w: fw - 0.08, h: 0.28,
      fontFace: F.jp, fontSize: 10, bold: true, color: C.green, align: "center", margin: 0, valign: "middle",
    });
  });
  note(s, fy + 1.14, "※ 予約不要・登録不要。ドリンク1杯からご利用いただけます。");
}

/* ===================================================== p10 ユーザー層 */
{
  const s = pres.addSlide();
  shell(s,  "ユーザー層", "セルフカフェのユーザー層",
    "「勉強」を目的に、若年層が繰り返し来店する——長時間・高頻度の利用構造です。");

  const cw = (CW - 0.42) / 2;

  card(s, M, TOP, cw, 4.72);
  s.addText("1", {
    x: M + 0.3, y: TOP + 0.26, w: 0.34, h: 0.34,
    fontFace: F.num, fontSize: 20, bold: true, color: C.midGreen, margin: 0, valign: "middle",
  });
  s.addText("若年層から絶大な支持", {
    x: M + 0.72, y: TOP + 0.24, w: cw - 1.0, h: 0.38,
    fontFace: F.jp, fontSize: 17, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
  s.addText(
    [
      { text: "29歳までのユーザーが", options: { color: C.body } },
      { text: "約50%", options: { color: C.green, bold: true } },
      { text: "。デジタルネイティブ世代にとって「無人」は心理的ハードルが低く、快適な空間です。", options: { color: C.body } },
    ],
    { x: M + 0.3, y: TOP + 0.72, w: cw - 0.6, h: 0.6, fontFace: F.jp, fontSize: 10.5, margin: 0, valign: "top", lineSpacingMultiple: 1.35 }
  );
  s.addChart(pres.ChartType.bar,
    [{ name: "構成比", labels: ["〜14", "15〜19", "20〜29", "30〜39", "40〜49", "50〜59", "60〜"], values: [3, 24, 29, 16, 13, 4, 1] }],
    {
      x: M + 0.2, y: TOP + 1.4, w: cw - 0.4, h: 3.06,
      barDir: "col", chartColors: [C.midGreen, C.green, C.green, C.green, C.green, C.tintLine, C.tintLine],
      varyColors: true, barGapWidthPct: 45,
      showLegend: false, showTitle: false,
      showValue: true, dataLabelPosition: "outEnd", dataLabelColor: C.green,
      dataLabelFontFace: F.num, dataLabelFontSize: 9.5, dataLabelFormatCode: '0"%"',
      catAxisLabelColor: C.muted, catAxisLabelFontFace: F.num, catAxisLabelFontSize: 9,
      catAxisLabelRotate: 0, catAxisLineShow: false, catGridLine: { style: "none" },
      valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 36, valGridLine: { style: "none" },
    }
  );
  s.addText("（歳）", {
    x: M + cw - 1.0, y: TOP + 4.4, w: 0.66, h: 0.2,
    fontFace: F.jp, fontSize: 8.5, color: C.footer, align: "right", margin: 0, valign: "middle",
  });

  const rx = M + cw + 0.42;
  card(s, rx, TOP, cw, 4.72);
  s.addText("2", {
    x: rx + 0.3, y: TOP + 0.26, w: 0.34, h: 0.34,
    fontFace: F.num, fontSize: 20, bold: true, color: C.midGreen, margin: 0, valign: "middle",
  });
  s.addText("「勉強」が主目的", {
    x: rx + 0.72, y: TOP + 0.24, w: cw - 1.0, h: 0.38,
    fontFace: F.jp, fontSize: 17, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
  s.addText(
    [
      { text: "利用目的の", options: { color: C.body } },
      { text: "62.6%が「勉強」", options: { color: C.green, bold: true } },
      { text: "。一般的なカフェが「休憩」を主目的とするのに対し、高い生産性を求める層が集まります。", options: { color: C.body } },
    ],
    { x: rx + 0.3, y: TOP + 0.72, w: cw - 0.6, h: 0.6, fontFace: F.jp, fontSize: 10.5, margin: 0, valign: "top", lineSpacingMultiple: 1.35 }
  );
  s.addChart(pres.ChartType.doughnut,
    [{
      name: "利用目的",
      labels: ["勉強 62.6%", "休憩・雑談 22.8%", "仕事・打ち合わせ 13.0%", "作業 0.2%", "その他 1.4%"],
      values: [62.6, 22.8, 13.0, 0.2, 1.4],
    }],
    {
      x: rx + 0.2, y: TOP + 1.36, w: cw - 0.4, h: 3.16,
      holeSize: 58, chartColors: [C.green, C.gold, C.midGreen, C.warmLine, C.tintLine],
      showLegend: true, legendPos: "b", legendColor: C.body, legendFontFace: F.jp, legendFontSize: 9.5,
      showTitle: false, showValue: false,
    }
  );
}

/* ===================================================== p11 / p12 店舗の様子 */
[
  {
    area: "愛知・千葉エリア",
    stores: [["ささしまライブ店", "愛知県名古屋市", "store-sasashima.jpg"], ["名駅西口店", "愛知県名古屋市", "store-meieki.jpg"], ["栄店", "愛知県名古屋市", "store-sakae.jpg"],
      ["新瑞橋店", "愛知県名古屋市", "store-aratama.jpg"], ["御器所店", "愛知県名古屋市", "store-gokiso.jpg"], ["印西牧の原店", "千葉県印西市", "store-inzai.jpg"]],
  },
  {
    area: "岩手・静岡・大阪エリア",
    stores: [["盛岡駅前店", "岩手県盛岡市", "store-morioka.jpg"], ["浜松新橋店", "静岡県浜松市", "store-hamamatsu.jpg"], ["相川駅前店", "大阪府大阪市", "store-aikawa.jpg"],
      ["天満店", "大阪府大阪市", "store-tenma.jpg"], ["あべの南店", "大阪府大阪市", "store-abeno.jpg"], ["谷町九丁目店", "大阪府大阪市", "store-tanimachi.jpg"]],
  },
].forEach(({ area, stores }) => {
  const s = pres.addSlide();
  shell(s,  "店舗の様子", `店舗毎の利用状況 ／ ${area}`,
    "店内カメラで見た実際の様子。平日日中でも席が埋まる店舗が多数あります。");

  const gx = 0.24, gy = 0.26;
  const w = (CW - 2 * gx) / 3, h = (BOT - TOP - gy) / 2;
  stores.forEach((st, i) => {
    const x = M + (i % 3) * (w + gx);
    const y = TOP + Math.floor(i / 3) * (h + gy);
    photoSlot(s, x, y, w, h - 0.34, null, { img: st[2] });
    s.addText(st[0], {
      x, y: y + h - 0.32, w: w - 1.3, h: 0.26,
      fontFace: F.jp, fontSize: 11, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(st[1], {
      x: x + w - 1.3, y: y + h - 0.32, w: 1.3, h: 0.26,
      fontFace: F.jp, fontSize: 9, color: C.muted, align: "right", margin: 0, valign: "middle",
    });
  });
});

/* ===================================================== p13 比較表 */
{
  const s = pres.addSlide();
  shell(s,  "ポジショニング", "比較表：いいところ取りのビジネスモデル",
    "カフェの「気軽さ」とオフィスの「機能」、その両方を満たす選択肢がありませんでした。");

  const colW = [1.9, 3.6, 3.2, 3.189];
  const ty = TOP + 0.1;

  // 本命列の面を先に敷く
  s.addShape("roundRect", {
    x: M + colW[0], y: ty, w: colW[1], h: 4.62, rectRadius: 0.05,
    fill: { color: C.tint }, line: { type: "none" },
  });

  const good = (t) => ({
    text: [
      { text: "◎ ", options: { color: C.green, bold: true, fontSize: 12.5 } },
      { text: t, options: { color: C.green, bold: true, fontSize: 10.5 } },
    ],
    options: { fill: { color: C.tint }, fontFace: F.jp, valign: "middle", margin: [0.05, 0.16, 0.05, 0.16] },
  });
  const other = (mark, t) => ({
    text: [
      { text: mark + " ", options: { color: mark === "×" ? "B0553F" : C.gold, bold: true, fontSize: 12.5 } },
      { text: t, options: { color: C.grayText, fontSize: 10 } },
    ],
    options: { fill: { color: C.white }, fontFace: F.jp, valign: "middle", margin: [0.05, 0.16, 0.05, 0.16] },
  });

  s.addTable(
    [
      [th("比較項目", { fill: C.grayBand, color: C.ink }), th("セルフカフェ"),
        th("コワーキング", { fill: "8E8B84" }), th("シェアオフィス", { fill: "8E8B84" })],
      [tl("利用の手軽さ"), good("予約不要・1杯〜"), other("△", "登録や予約が必要"), other("×", "法人契約がメイン")],
      [tl("空間の雰囲気"), good("適度な賑わい"), other("△", "静かすぎて緊張する"), other("×", "閉鎖的な個室中心")],
      [tl("ドリンクの質"), good("本格コーヒー"), other("△", "サーバーが中心"), other("×", "持ち込みが基本")],
      [tl("運営人員"), good("完全無人"), other("△", "スタッフ常駐"), other("△", "管理人が必要")],
      [tl("導入コスト"), good("低投資パッケージ"), other("×", "高額な内装・設備"), other("×", "莫大な建築コスト")],
    ],
    {
      x: M, y: ty, w: CW, colW,
      rowH: [0.42, 0.78, 0.78, 0.78, 0.78, 0.78],
      border: { type: "solid", color: C.warmLine, pt: 0.75 },
      autoPage: false,
    }
  );

  s.addShape("roundRect", {
    x: M, y: 6.28, w: CW, h: 0.52, rectRadius: 0.05,
    fill: { color: C.greenDeep }, line: { type: "none" },
  });
  s.addText("カフェの「気軽さ」とオフィスの「機能」を融合させた、唯一無二のポジション。", {
    x: M + 0.3, y: 6.28, w: CW - 0.6, h: 0.52,
    fontFace: F.jp, fontSize: 12.5, bold: true, color: C.white, margin: 0, valign: "middle",
  });
}

/* ===================================================== p14 無人開閉店システム */
{
  const s = pres.addSlide();
  shell(s,  "運営システム", "独自の無人開閉店システム",
    "警備会社と連携した安心の自動化。オーナー様が現地に足を運ぶ必要はありません。");

  const pw = 5.6;
  panel(s, M, TOP, pw, 4.72);
  icon(s, "LuKeyRound", "white", M + 0.44, TOP + 0.4, 0.44);
  s.addText("スマートロックで、\n開錠／施錠を完全自動化", {
    x: M + 0.44, y: TOP + 1.0, w: pw - 0.88, h: 1.1,
    fontFace: F.jp, fontSize: 21, bold: true, color: C.white, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
  });
  s.addText("独自のスマートロックシステムにより、店舗の開錠／施錠を遠隔で完全自動化。現地に足を運ばずに営業時間を管理できます。", {
    x: M + 0.44, y: TOP + 2.24, w: pw - 0.88, h: 1.0,
    fontFace: F.jp, fontSize: 11, color: C.cvBody, margin: 0, valign: "top", lineSpacingMultiple: 1.45,
  });
  ["任意の営業時間を分単位でコントロール", "不審者の侵入を防ぐセキュリティと連携", "無人ならではの24時間営業も容易"].forEach((b, i) => {
    const y = TOP + 3.34 + i * 0.44;
    s.addText(String(i + 1), {
      x: M + 0.44, y, w: 0.24, h: 0.3,
      fontFace: F.num, fontSize: 11, bold: true, color: C.cvPale, margin: 0, valign: "middle",
    });
    s.addText(b, {
      x: M + 0.76, y, w: pw - 1.2, h: 0.3,
      fontFace: F.jp, fontSize: 11, bold: true, color: C.white, margin: 0, valign: "middle",
    });
  });

  const rx = M + pw + 0.42, rw = R - rx;
  const items = [
    ["LuBellRing", "非常ボタン", "非常・緊急時に押すと、警備員が駆け付けます。"],
    ["LuCctv", "画像センサー", "非常時の映像を監視センターへ送信。スピーカーで威嚇可能。閉店後は侵入者を感知して通報します。"],
    ["LuFlame", "火災センサー", "火災の発生を感知して、警備員が駆け付けます。"],
    ["LuKeyRound", "出入口の管理", "電気錠・電磁錠や自動ドアと連携し、閉店／開店時間に自動施錠・開錠。閉店時は自動で侵入警戒をセットします。"],
  ];
  const cw = (rw - 0.24) / 2, ch = (4.72 - 0.24) / 2;
  items.forEach((it, i) => {
    const x = rx + (i % 2) * (cw + 0.24);
    const y = TOP + Math.floor(i / 2) * (ch + 0.24);
    card(s, x, y, cw, ch);
    icon(s, it[0], "green", x + 0.26, y + 0.26, 0.38);
    s.addText(it[1], {
      x: x + 0.26, y: y + 0.76, w: cw - 0.52, h: 0.3,
      fontFace: F.jp, fontSize: 13, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(it[2], {
      x: x + 0.26, y: y + 1.1, w: cw - 0.52, h: 1.0,
      fontFace: F.jp, fontSize: 10, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.35,
    });
  });
  note(s, 6.62, "※ 警備会社は環境により異なる場合がございます。");
}

/* ===================================================== p15 モニタリング */
{
  const s = pres.addSlide();
  shell(s,  "運営システム", "店内モニタリング＆サポート基盤",
    "緊急時に強い、安心のサポート体制で無人運営を支えます。");

  const pw = 5.6;
  panel(s, M, TOP, pw, 3.06);
  eyebrowIn(s, M + 0.44, TOP + 0.32, 2.0, "SUPPORT", C.cvPale);
  s.addText("災害・体調不良・不審者。\nあらゆる緊急事態に備える。", {
    x: M + 0.44, y: TOP + 0.68, w: pw - 0.88, h: 1.0,
    fontFace: F.jp, fontSize: 19, bold: true, color: C.white, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
  });
  s.addText("店内の各所に非常ボタンを設置し、お客様同士のトラブルまで幅広く対応。6台以上のカメラによるモニタリングで、お問い合わせにも即時対応します。", {
    x: M + 0.44, y: TOP + 1.82, w: pw - 0.88, h: 0.72,
    fontFace: F.jp, fontSize: 10.5, color: C.cvBody, margin: 0, valign: "top", lineSpacingMultiple: 1.45,
  });
  // パネル内の設備アイコン列
  [["LuBellRing", "非常ボタン"], ["LuCctv", "常時モニタリング"], ["LuFlame", "火災センサー"]].forEach((ic, i) => {
    const ix = M + 0.44 + i * 1.68 + 0.5;
    icon(s, ic[0], "pale", ix - 0.15, TOP + 2.52, 0.3);
    s.addText(ic[1], {
      x: ix - 0.75, y: TOP + 2.84, w: 1.5, h: 0.16,
      fontFace: F.jp, fontSize: 8, color: C.cvPale, align: "center", margin: 0, valign: "middle",
    });
  });

  // 警備会社との連携（写真枠の代わり）
  tintCard(s, M, TOP + 3.22, pw, 1.5);
  icon(s, "LuShieldCheck", "green", M + 0.3, TOP + 3.5, 0.5);
  s.addText("警備会社と連携（ALSOK・SECOM 等）", {
    x: M + 1.0, y: TOP + 3.42, w: pw - 1.3, h: 0.3,
    fontFace: F.jp, fontSize: 13, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
  s.addText("非常時には警備員が駆けつけ。スマートロック・侵入検知と連動した警備体制です。", {
    x: M + 1.0, y: TOP + 3.74, w: pw - 1.3, h: 0.56,
    fontFace: F.jp, fontSize: 10, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.35,
  });
  s.addText("※ 連携する警備会社は物件・環境により異なります。", {
    x: M + 1.0, y: TOP + 4.36, w: pw - 1.3, h: 0.2,
    fontFace: F.jp, fontSize: 8, color: C.muted, margin: 0, valign: "middle",
  });

  const rx = M + pw + 0.42, rw = R - rx;
  const items = [
    ["LuVideo", "6台以上", "カメラモニタリング", "店内を多角的に常時監視。死角を作らない配置で異常を早期に把握します。"],
    ["LuMic", "双方向", "音声対応", "カメラはスタッフとお客様が直接やり取りできる双方向音声に対応。その場で声がけやサポートを行えます。"],
    ["LuMessageCircle", "即時", "LINEチャット対応", "お客様からのお問い合わせに即時対応。質問対応・意見収集の窓口としても機能します。"],
  ];
  const ch = (4.72 - 0.4) / 3;
  items.forEach((it, i) => {
    const y = TOP + i * (ch + 0.2);
    card(s, rx, y, rw, ch);
    icon(s, it[0], "green", rx + 0.28, y + 0.26, 0.36);
    s.addText(
      [
        { text: it[1], options: { fontFace: F.num, fontSize: 18, bold: true, color: C.green } },
        { text: "  " + it[2], options: { fontFace: F.jp, fontSize: 13, bold: true, color: C.ink } },
      ],
      { x: rx + 0.8, y: y + 0.24, w: rw - 1.1, h: 0.4, margin: 0, valign: "middle" }
    );
    s.addText(it[3], {
      x: rx + 0.8, y: y + 0.68, w: rw - 1.1, h: 0.6,
      fontFace: F.jp, fontSize: 10, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.35,
    });
  });
}

/* ===================================================== p16 パートナー制度の3つの特徴 */
{
  const s = pres.addSlide();
  shell(s,  "パートナー制度", "パートナー制度の3つの特徴",
    "投資を抑え、既存の資産と人員のまま始められます。");

  const items = [
    ["LuWallet", "POINT 01", "1/10 以下", "低初期コスト",
      "他社カフェ業態と比較し、1/10以下のコストで出店が可能。初期費用650万円〜から始められます。",
      ["加盟金 100万円", "初期費用 650万円〜"]],
    ["LuBuilding2", "POINT 02", "0 円の新規賃料", "遊休スペースの活用",
      "既存店の空室やデッドスペースを、収益資産へ再構築。限られたスペースの一角に設置できます。",
      ["レジ横", "待合スペース", "空き区画"]],
    ["LuTimer", "POINT 03", "15 分／日", "1日15分程度の運営",
      "清掃・補充のみ。無人化により人件費を極限までカット。既存スタッフの稼働内で対応でき、新規採用は不要です。",
      ["清掃", "原料補充"]],
  ];
  const w = (CW - 2 * 0.3) / 3, h = 4.24;
  items.forEach((it, i) => {
    const x = M + i * (w + 0.3), y = TOP;
    card(s, x, y, w, h);
    eyebrowIn(s, x + 0.3, y + 0.3, 2.0, it[1]);
    icon(s, it[0], "green", x + w - 0.82, y + 0.26, 0.4);
    s.addText(it[2], {
      x: x + 0.3, y: y + 0.62, w: w - 0.6, h: 0.5,
      fontFace: F.num, fontSize: 24, bold: true, color: C.green, margin: 0, valign: "middle",
    });
    s.addText(it[3], {
      x: x + 0.3, y: y + 1.18, w: w - 0.6, h: 0.36,
      fontFace: F.jp, fontSize: 15, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addShape("rect", { x: x + 0.3, y: y + 1.64, w: w - 0.6, h: 0.011, fill: { color: C.warmLine } });
    s.addText(it[4], {
      x: x + 0.3, y: y + 1.78, w: w - 0.6, h: 1.5,
      fontFace: F.jp, fontSize: 11, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.5,
    });
    let tx = x + 0.3;
    it[5].forEach((tag) => {
      const tw = 0.3 + tag.length * 0.135;
      const ty2 = y + h - 0.62;
      s.addShape("roundRect", {
        x: tx, y: ty2, w: tw, h: 0.32, rectRadius: 0.16,
        fill: { color: C.tint }, line: { color: C.tintLine, width: 0.75 },
      });
      s.addText(tag, {
        x: tx, y: ty2, w: tw, h: 0.32,
        fontFace: F.jp, fontSize: 9, bold: true, color: C.green, align: "center", margin: 0, valign: "middle",
      });
      tx += tw + 0.14;
    });
  });

  // 下部の帯：3つの特徴のまとめ
  const bandY = TOP + h + 0.2;
  panel(s, M, bandY, CW, 0.62);
  [["LuUsers", "新規採用は不要"], ["LuStore", "既存スペースをそのまま活用"], ["LuPiggyBank", "低投資・低リスクで開始"]].forEach((b2, i) => {
    const bx2 = M + 0.5 + i * (CW / 3);
    icon(s, b2[0], "pale", bx2, bandY + 0.16, 0.3);
    s.addText(b2[1], {
      x: bx2 + 0.42, y: bandY, w: CW / 3 - 1.0, h: 0.62,
      fontFace: F.jp, fontSize: 12, bold: true, color: C.white, margin: 0, valign: "middle",
    });
  });
}

/* ===================================================== p17 初期費用（他社比較） */
{
  const s = pres.addSlide();
  shell(s,  "初期費用", "初期費用：他社カフェ業態との比較",
    "同じ「カフェ出店」でも、必要な初期投資はまったく違います。");

  const colW = [2.4, 2.6, 3.4, 3.489];
  s.addShape("roundRect", {
    x: M + colW[0], y: TOP, w: colW[1], h: 4.56, rectRadius: 0.05,
    fill: { color: C.tint }, line: { type: "none" },
  });
  const mine = (t, big) => ({
    text: t,
    options: {
      fill: { color: C.tint }, color: C.green, bold: true, fontSize: big ? 13 : 11,
      fontFace: F.num, align: "right", valign: "middle", margin: [0.05, 0.16, 0.05, 0.16],
    },
  });
  const oth = (t) => ({
    text: t,
    options: {
      fill: { color: C.white }, color: C.grayText, fontSize: 10.5, fontFace: F.jp,
      align: "right", valign: "middle", margin: [0.05, 0.16, 0.05, 0.16],
    },
  });

  s.addTable(
    [
      [th("項目", { fill: C.grayBand, color: C.ink }), th("セルフカフェ", { align: "right" }),
        th("C社", { fill: "8E8B84", align: "right" }), th("D社", { fill: "8E8B84", align: "right" })],
      [tl("加盟金"), mine("100万円"), oth("300万円〜"), oth("150万円〜")],
      [tl("店舗工事費用"), mine("400万円〜"), oth("8,000万円〜"), oth("2,000万円〜")],
      [tl("保証金"), mine("50万円〜"), oth("100万円〜"), oth("150万円〜")],
      [tl("研修費"), mine("なし"), oth("15万円"), oth("20万円／1名")],
      [tl("開業資金"), mine("550万円〜", true), oth("8,400万円〜"), oth("3,000万円〜")],
      [tl("ロイヤリティ"), mine("月5万円（一律）"), oth("1席1,500円／月"), oth("売上の2%／月")],
    ],
    {
      x: M, y: TOP, w: CW, colW,
      rowH: [0.42, 0.62, 0.62, 0.62, 0.62, 0.78, 0.72],
      border: { type: "solid", color: C.warmLine, pt: 0.75 },
      autoPage: false,
    }
  );

  s.addShape("roundRect", {
    x: M, y: 6.24, w: CW, h: 0.56, rectRadius: 0.05,
    fill: { color: C.greenDeep }, line: { type: "none" },
  });
  s.addText(
    [
      { text: "開業資金は他社カフェ業態の ", options: { color: C.white } },
      { text: "1/5 〜 1/15", options: { color: C.cvPale, bold: true } },
      { text: "。ロイヤリティも売上連動ではなく月5万円の一律です。", options: { color: C.white } },
    ],
    { x: M + 0.3, y: 6.24, w: CW - 0.6, h: 0.56, fontFace: F.jp, fontSize: 12, bold: true, margin: 0, valign: "middle" }
  );
  note(s, 6.9, "");
}

/* ===================================================== p18 初期費用・売上イメージ */
{
  const s = pres.addSlide();
  shell(s,  "初期費用", "初期費用・売上イメージ",
    "出店方法は2パターン。居抜き・既存店併設なら初期費用を大きく抑えられます。");

  const cw = (CW - 0.42) / 2;
  const pats = [
    ["① 既存店併設 / 居抜き", "650万円〜", "内装が残る区画や自社店舗の一角を活用。工事範囲を最小化できます。", true],
    ["② スケルトン工事", "950万円〜", "内装のない状態からの造作。レイアウトを自由に設計できます。", false],
  ];
  pats.forEach((p, i) => {
    const x = M + i * (cw + 0.42), y = TOP;
    if (p[3]) panel(s, x, y, cw, 1.86); else card(s, x, y, cw, 1.86);
    const fg = p[3] ? C.white : C.ink;
    s.addText(p[0], {
      x: x + 0.3, y: y + 0.24, w: cw - 0.6, h: 0.32,
      fontFace: F.jp, fontSize: 13, bold: true, color: p[3] ? C.cvPale : C.green, margin: 0, valign: "middle",
    });
    s.addText("初期費用目安", {
      x: x + 0.3, y: y + 0.6, w: cw - 0.6, h: 0.2,
      fontFace: F.jp, fontSize: 9, color: p[3] ? C.cvBody : C.muted, margin: 0, valign: "middle",
    });
    s.addText(p[1], {
      x: x + 0.3, y: y + 0.8, w: cw - 0.6, h: 0.5,
      fontFace: F.num, fontSize: 30, bold: true, color: fg, margin: 0, valign: "middle",
    });
    s.addText(p[2], {
      x: x + 0.3, y: y + 1.34, w: cw - 0.6, h: 0.42,
      fontFace: F.jp, fontSize: 10, color: p[3] ? C.cvBody : C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
    });
  });

  const by = TOP + 2.06;
  card(s, M, by, CW, 2.7);
  s.addText("パートナー制度 初期費用内訳（20坪想定・居抜きパターン）", {
    x: M + 0.3, y: by + 0.22, w: 7.0, h: 0.28,
    fontFace: F.jp, fontSize: 11.5, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
  s.addText(
    [
      { text: "合計 ", options: { fontFace: F.jp, fontSize: 10, color: C.muted } },
      { text: "650万円〜", options: { fontFace: F.num, fontSize: 15, bold: true, color: C.green } },
    ],
    { x: M + CW - 2.6, y: by + 0.18, w: 2.3, h: 0.36, align: "right", margin: 0, valign: "middle" }
  );
  const bd = [
    ["LuHandCoins", "加盟金", "100万円", ""],
    ["LuHammer", "出店工事費用", "400万円〜", "設備・電気・大工・内装工事等"],
    ["LuStore", "出店準備金", "100万円〜", "カメラ・モニター・机・椅子等の備品類"],
    ["LuReceipt", "別途保証金", "50万円〜", "預り金。原料やロイヤリティの支払い確認が取れない場合に差し引き"],
  ];
  const bw = (CW - 0.6 - 3 * 0.2) / 4;
  bd.forEach((b, i) => {
    const x = M + 0.3 + i * (bw + 0.2), y = by + 0.66;
    tintCard(s, x, y, bw, 1.82);
    icon(s, b[0], "green", x + 0.22, y + 0.2, 0.32);
    s.addText(b[1], {
      x: x + 0.22, y: y + 0.6, w: bw - 0.44, h: 0.26,
      fontFace: F.jp, fontSize: 10.5, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(b[2], {
      x: x + 0.22, y: y + 0.86, w: bw - 0.44, h: 0.36,
      fontFace: F.num, fontSize: 17, bold: true, color: C.green, margin: 0, valign: "middle",
    });
    if (b[3]) {
      s.addText(b[3], {
        x: x + 0.22, y: y + 1.24, w: bw - 0.44, h: 0.5,
        fontFace: F.jp, fontSize: 8.5, color: C.muted, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
      });
    }
  });
  note(s, 6.68, "※ 金額は目安です。物件条件・工事範囲により変動します。");
}

/* ===================================================== p19 / p20 収益シミュレーション */
[
  {
    label: "20坪（40席）", tag: "STANDARD",
    rev: ["売上高", "1杯420円 × 75杯 × 30日", "945,000円"],
    rows: [
      ["ドリンク原料", "原料・カップ等（20%）", "-189,000円"],
      ["家賃", "仮定坪単価：10,000円", "-200,000円"],
      ["清掃費", "清掃パートナー様への報酬", "-30,000円"],
      ["水道光熱費", "水道9,000円／光熱71,000円（24H営業想定）", "-80,000円"],
      ["機械使用料", "ドリンクマシン×2台、決済端末×2台", "-86,000円"],
      ["セキュリティ費", "警備会社への委託料", "-20,000円"],
      ["ロイヤリティ", "一律5万円", "-50,000円"],
      ["雑費", "ゴミ回収・防犯カメラ・クラウド利用料等", "-20,000円"],
    ],
    profit: "270,000円", annual: "3,240,000円",
  },
  {
    label: "40坪（80席）", tag: "LARGE",
    rev: ["売上高", "1杯420円 × 100杯 × 30日", "1,260,000円"],
    rows: [
      ["ドリンク原料", "原料・カップ等（20%）", "-252,000円"],
      ["家賃", "仮定坪単価：8,000円", "-320,000円"],
      ["清掃費", "清掃パートナー様への報酬", "-35,000円"],
      ["水道光熱費", "水道11,000円／光熱74,000円（24H営業想定）", "-85,000円"],
      ["機械使用料", "ドリンクマシン38,000円×2台、決済端末5,000円×2台", "-86,000円"],
      ["セキュリティ費", "警備会社への委託料", "-20,000円"],
      ["ロイヤリティ", "一律5万円", "-50,000円"],
      ["雑費", "ゴミ回収・防犯カメラ・クラウド利用料等", "-25,000円"],
    ],
    profit: "387,000円", annual: "4,644,000円",
  },
].forEach((sim) => {
  const s = pres.addSlide();
  shell(s,  "収益", `収益シミュレーション ／ ${sim.label}`,
    "月次の収支イメージです。売上から諸経費を差し引いた償却前営業利益を示しています。");

  const lw = 7.9;
  // 売上ヘッダー
  s.addShape("roundRect", {
    x: M, y: TOP, w: lw, h: 0.58, rectRadius: 0.05,
    fill: { color: C.greenDeep }, line: { type: "none" },
  });
  s.addText(sim.rev[0], {
    x: M + 0.26, y: TOP, w: 1.6, h: 0.58,
    fontFace: F.jp, fontSize: 12, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  s.addText(sim.rev[1], {
    x: M + 1.9, y: TOP, w: 3.6, h: 0.58,
    fontFace: F.jp, fontSize: 9.5, color: C.cvBody, margin: 0, valign: "middle",
  });
  s.addText(sim.rev[2], {
    x: M + lw - 2.4, y: TOP, w: 2.14, h: 0.58,
    fontFace: F.num, fontSize: 18, bold: true, color: C.white, align: "right", margin: 0, valign: "middle",
  });

  sim.rows.forEach((r, i) => {
    const y = TOP + 0.70 + i * 0.42;
    s.addText(r[0], {
      x: M + 0.1, y, w: 1.7, h: 0.38,
      fontFace: F.jp, fontSize: 10.5, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(r[1], {
      x: M + 1.9, y, w: 3.7, h: 0.38,
      fontFace: F.jp, fontSize: 9, color: C.muted, margin: 0, valign: "middle",
    });
    s.addText(r[2], {
      x: M + lw - 2.4, y, w: 2.14, h: 0.38,
      fontFace: F.num, fontSize: 12, color: C.body, align: "right", margin: 0, valign: "middle",
    });
    s.addShape("rect", { x: M + 0.1, y: y + 0.38, w: lw - 0.2, h: 0.011, fill: { color: C.warmLine } });
  });

  const py = TOP + 0.70 + sim.rows.length * 0.42 + 0.14;
  s.addShape("roundRect", {
    x: M, y: py, w: lw, h: 0.64, rectRadius: 0.05,
    fill: { color: C.tint }, line: { color: C.tintLine, width: 0.75 },
  });
  s.addText("償却前営業利益", {
    x: M + 0.26, y: py, w: 3.0, h: 0.64,
    fontFace: F.jp, fontSize: 13, bold: true, color: C.green, margin: 0, valign: "middle",
  });
  s.addText(sim.profit, {
    x: M + lw - 2.6, y: py, w: 2.34, h: 0.64,
    fontFace: F.num, fontSize: 22, bold: true, color: C.green, align: "right", margin: 0, valign: "middle",
  });

  // 右カラム
  const rx = M + lw + 0.42, rw = R - rx;
  panel(s, rx, TOP, rw, 1.9);
  eyebrowIn(s, rx + 0.26, TOP + 0.26, 2.0, sim.tag, C.cvPale);
  s.addText("月間の償却前営業利益", {
    x: rx + 0.26, y: TOP + 0.54, w: rw - 0.52, h: 0.24,
    fontFace: F.jp, fontSize: 10, color: C.cvBody, margin: 0, valign: "middle",
  });
  s.addText(sim.profit, {
    x: rx + 0.26, y: TOP + 0.82, w: rw - 0.52, h: 0.5,
    fontFace: F.num, fontSize: 26, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  s.addText("年間換算 " + sim.annual, {
    x: rx + 0.26, y: TOP + 1.36, w: rw - 0.52, h: 0.3,
    fontFace: F.jp, fontSize: 11, color: C.cvPale, margin: 0, valign: "middle",
  });

  card(s, rx, TOP + 2.1, rw, 2.62);
  s.addText("前提", {
    x: rx + 0.26, y: TOP + 2.3, w: rw - 0.52, h: 0.24,
    fontFace: F.jp, fontSize: 10, bold: true, color: C.muted, margin: 0, valign: "middle",
  });
  ["売上は1杯420円で算出", "家賃は仮定坪単価による試算", "ロイヤリティは月5万円の一律", "減価償却費は含まない"].forEach((t, i) => {
    const y = TOP + 2.62 + i * 0.42;
    icon(s, "LuCheck", "green", rx + 0.28, y + 0.05, 0.18);
    s.addText(t, {
      x: rx + 0.56, y, w: rw - 0.82, h: 0.32,
      fontFace: F.jp, fontSize: 10, color: C.ink, margin: 0, valign: "middle",
    });
  });

  note(s, py + 0.78, "※ シミュレーションは目安であり、売上・利益を保証するものではありません。");
});

/* ===================================================== 投資回収シミュレーション（20坪・50坪） */
{
  const s = pres.addSlide();
  shell(s,  "収益", "投資回収シミュレーション",
    "20坪（初期費用800万円）と40坪（同1,200万円）について、家賃あり／家賃なし（自社物件）の回収期間を試算しました。");

  const cases = [
    {
      no: "CASE 01", size: "20坪（40席）", invest: "800", investNote: "初期費用",
      sales: "945,000円", salesNote: "420円 × 75杯 × 30日",
      cost: "-475,000円", rent: "-200,000円", rentNote: "坪単価 10,000円",
      pats: [
        { name: "家賃あり（賃借）", profit: "27.0万円", pb: "約2年6ヶ月", pbS: "30ヶ月", dark: false },
        { name: "家賃なし（自社物件）", profit: "47.0万円", pb: "約1年5ヶ月", pbS: "17ヶ月", dark: true },
      ],
    },
    {
      no: "CASE 02", size: "40坪（80席）", invest: "1,200", investNote: "初期費用",
      sales: "1,260,000円", salesNote: "420円 × 100杯 × 30日",
      cost: "-553,000円", rent: "-320,000円", rentNote: "坪単価 8,000円",
      pats: [
        { name: "家賃あり（賃借）", profit: "38.7万円", pb: "約2年7ヶ月", pbS: "31ヶ月", dark: false },
        { name: "家賃なし（自社物件）", profit: "70.7万円", pb: "約1年5ヶ月", pbS: "17ヶ月", dark: true },
      ],
    },
  ];

  const cw2 = (CW - 0.42) / 2;
  cases.forEach((cs, ci) => {
    const x = M + ci * (cw2 + 0.42), y = TOP, ch2 = 3.94;
    card(s, x, y, cw2, ch2);
    eyebrowIn(s, x + 0.3, y + 0.22, 1.4, cs.no);
    s.addText(cs.size, {
      x: x + 0.3, y: y + 0.4, w: 3.0, h: 0.36,
      fontFace: F.jp, fontSize: 16, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(
      [
        { text: cs.investNote + " ", options: { fontFace: F.jp, fontSize: 9.5, color: C.muted } },
        { text: cs.invest, options: { fontFace: F.num, fontSize: 20, bold: true, color: C.gold } },
        { text: " 万円", options: { fontFace: F.jp, fontSize: 11, bold: true, color: C.gold } },
      ],
      { x: x + cw2 - 3.0, y: y + 0.36, w: 2.7, h: 0.4, align: "right", margin: 0, valign: "middle" }
    );
    const rows = [
      ["月間売上", cs.sales, cs.salesNote],
      ["月間経費（家賃除く）", cs.cost, "原料・清掃・光熱・機械・警備・ロイヤリティ等"],
      ["家賃", cs.rent, cs.rentNote],
    ];
    rows.forEach((r, i) => {
      const ry = y + 0.9 + i * 0.36;
      s.addText(r[0], {
        x: x + 0.3, y: ry, w: 2.1, h: 0.3,
        fontFace: F.jp, fontSize: 9.5, bold: true, color: C.ink, margin: 0, valign: "middle",
      });
      s.addText(r[2], {
        x: x + 2.1, y: ry, w: cw2 - 3.9, h: 0.3,
        fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
      });
      s.addText(r[1], {
        x: x + cw2 - 1.9, y: ry, w: 1.6, h: 0.3,
        fontFace: F.num, fontSize: 12, bold: true, color: C.body, align: "right", margin: 0, valign: "middle",
      });
      s.addShape("rect", { x: x + 0.3, y: ry + 0.3, w: cw2 - 0.6, h: 0.011, fill: { color: C.warmLine } });
    });
    // 家賃あり／なしの回収ミニパネル
    const pw3 = (cw2 - 0.6 - 0.2) / 2;
    cs.pats.forEach((pt, pi) => {
      const px = x + 0.3 + pi * (pw3 + 0.2), py = y + 2.14, ph = 1.56;
      if (pt.dark) panel(s, px, py, pw3, ph); else tintCard(s, px, py, pw3, ph);
      const fg = pt.dark ? C.white : C.ink;
      const sub = pt.dark ? C.cvBody : C.muted;
      s.addText(pt.name, {
        x: px + 0.2, y: py + 0.14, w: pw3 - 0.4, h: 0.24,
        fontFace: F.jp, fontSize: 9.5, bold: true, color: pt.dark ? C.cvPale : C.green, margin: 0, valign: "middle",
      });
      s.addText(
        [
          { text: "月間利益 ", options: { fontFace: F.jp, fontSize: 8.5, color: sub } },
          { text: pt.profit, options: { fontFace: F.num, fontSize: 12, bold: true, color: fg } },
        ],
        { x: px + 0.2, y: py + 0.4, w: pw3 - 0.4, h: 0.26, margin: 0, valign: "middle" }
      );
      s.addText("回収期間", {
        x: px + 0.2, y: py + 0.7, w: pw3 - 0.4, h: 0.2,
        fontFace: F.jp, fontSize: 8.5, color: sub, margin: 0, valign: "middle",
      });
      s.addText(
        [
          { text: pt.pb, options: { fontFace: F.jp, fontSize: 14.5, bold: true, color: pt.dark ? C.white : C.green } },
          { text: "（" + pt.pbS + "）", options: { fontFace: F.num, fontSize: 8.5, color: sub } },
        ],
        { x: px + 0.14, y: py + 0.9, w: pw3 - 0.26, h: 0.44, margin: 0, valign: "middle" }
      );
    });
  });

  // 前提
  const byr = TOP + 4.14;
  tintCard(s, M, byr, CW, 0.56);
  s.addText("前提：1杯420円／販売数 20坪=75杯・40坪=100杯／日（収益シミュレーションと同一）／ロイヤリティ 月5万円／減価償却費・税は含まない", {
    x: M + 0.26, y: byr, w: CW - 0.52, h: 0.56,
    fontFace: F.jp, fontSize: 10, color: C.ink, margin: 0, valign: "middle",
  });

  note(s, byr + 0.72, "※ 収益シミュレーション（p20〜p21）と同一の前提による目安です。回収期間・利益を保証するものではありません。");
}

/* ===================================================== p21 出店事例（FC固有） */
{
  const s = pres.addSlide();
  shell(s,  "出店事例", "出店事例：岩手県 盛岡駅前店",
    "地方都市の駅前立地。1日80杯ペースで稼働し、家賃を吸収して利益を確保しています。");

  chipRow(s, TOP, [
    { v: "960,000", u: "円", l: "月間売上高（20坪／40席）" },
    { v: "280,000", u: "円", l: "家賃" },
    { v: "232,000", u: "円", l: "償却前営業利益" },
    { v: "80", u: "杯／日", l: "1日あたりの販売杯数" },
  ], { vSize: 18 });

  const y2 = TOP + 0.92;
  const pw = 6.6;
  // 外観を全幅の1枚で大きく（元画像比率 4.15 ≒ 枠比率のため、ほぼトリミングなしの引き）
  photoSlot(s, M, y2, CW, 2.82, "盛岡駅前店の外観", { img: "p02-storefront.jpg", capSize: 10 });

  // 下段：CASE STUDY 横帯
  const y3 = y2 + 2.98, bh3 = 1.06;
  card(s, M, y3, CW, bh3);
  eyebrowIn(s, M + 0.3, y3 + 0.16, 2.0, "CASE STUDY");
  s.addText("1日80杯ペースでの稼働", {
    x: M + 0.3, y: y3 + 0.36, w: 3.4, h: 0.34,
    fontFace: F.jp, fontSize: 15, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
  s.addText("1回の利用で2〜3時間、2杯程度購入。学生からビジネスマンまで幅広い層を獲得。", {
    x: M + 0.3, y: y3 + 0.72, w: 4.3, h: 0.26,
    fontFace: F.jp, fontSize: 8.5, color: C.muted, margin: 0, valign: "middle",
  });
  s.addShape("rect", { x: M + 4.8, y: y3 + 0.18, w: 0.011, h: bh3 - 0.36, fill: { color: C.warmLine } });
  const facts = [["開業", "2025年4月"], ["面積・席数", "20坪／40席"], ["立地", "岩手県盛岡市・駅前"], ["営業時間", "24時間"]];
  const fw = (CW - 5.2) / 4;
  facts.forEach((f, i) => {
    const fx = M + 5.1 + i * fw;
    s.addText(f[0], {
      x: fx, y: y3 + 0.2, w: fw - 0.1, h: 0.24,
      fontFace: F.jp, fontSize: 9, color: C.muted, margin: 0, valign: "middle",
    });
    s.addText(f[1], {
      x: fx, y: y3 + 0.48, w: fw - 0.1, h: 0.36,
      fontFace: F.jp, fontSize: 12.5, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
  });
  note(s, y3 + bh3 + 0.14, "※ 実績値であり、他店舗での同等の売上・利益を保証するものではありません。");
}

/* ===================================================== p22 有料オプション */
{
  const s = pres.addSlide();
  shell(s,  "オプション", "有料オプション一覧",
    "必要に応じて追加できるオプションです（金額は税抜き）。");

  const colW = [2.5, 1.5, 1.5, 6.389];
  const yen = (t, free) => ({
    text: t,
    options: {
      fill: { color: C.white }, color: free ? C.muted : C.green, bold: !free,
      fontSize: 11.5, fontFace: F.num, align: "right", valign: "middle", margin: [0.05, 0.14, 0.05, 0.14],
    },
  });
  const b = (t) => ({ text: t, options: { color: C.green, bold: true, fontSize: 10, breakLine: true } });
  const p = (t, last) => ({ text: t, options: { color: C.body, fontSize: 9.5, breakLine: !last } });
  const rem = (runs) => ({
    text: runs,
    options: { fill: { color: C.white }, fontFace: F.jp, valign: "middle", margin: [0.06, 0.16, 0.06, 0.16] },
  });

  s.addTable(
    [
      [th("項目", { fill: C.grayBand, color: C.ink }), th("初期費用", { align: "right" }),
        th("月額費用", { align: "right" }), th("内容・備考")],
      [tl("公式LINE作成"), yen("15,000円"), yen("0円", true), rem([
        b("FC店の公式LINEを新規作成できます。"),
        p("メリット：独自の発信や告知、広告が打てます。"),
        p("デメリット：セルフカフェ公式LINEから切り離されるため問い合わせが直に届き、本部情報の発信も都度パートナー様側で行う必要があります。"),
        p("※ 貴社にて契約（配信数によって別途費用が発生します）", true),
      ])],
      [tl("顧客用コピー機\n（A3対応）"), yen("50,000円"), yen("30,000円"), rem([
        b("PayPay決済のみ"), p("※ 本体代は別途"), p("※ グローバルIPの取得が必要です", true),
      ])],
      [tl("営業サポートプラン\n（フルサポート）"), yen("0円", true), yen("50,000円"), rem([
        b("顧客チャット対応"), p("セルフカフェ公式LINEやHPからの問い合わせに対応"),
        b("マシンメンテナンス"), p("マシン清掃・原料補充等。店内清掃は含まれず、マシンのみの対応です。", true),
      ])],
      [tl("営業サポートプラン\n（ミニマムサポート）"), yen("0円", true), yen("15,000円"), rem([
        b("顧客チャット対応"), p("セルフカフェ公式LINEやHPからの問い合わせに対応", true),
      ])],
    ],
    {
      x: M, y: TOP, w: CW, colW,
      rowH: [0.44, 1.3, 0.86, 1.0, 0.72],
      border: { type: "solid", color: C.warmLine, pt: 0.75 },
      autoPage: false,
    }
  );
  note(s, TOP + 4.6, "※ SEO対策・MEO対策・PayPayポイント付与は費用なしの標準提供のため、本表からは除いています。");
}

/* ===================================================== p23 ドリンクマシン */
{
  const s = pres.addSlide();
  shell(s,  "設備", "ドリンクマシンに関して",
    "マシン・備品・メンテナンスまで含めたトータルサポートパックでご提供します。");

  const lw = 4.6;
  card(s, M, TOP, lw, 4.72);
  s.addShape("roundRect", {
    x: M + 0.3, y: TOP + 0.28, w: lw - 0.6, h: 0.5, rectRadius: 0.25,
    fill: { color: C.greenDeep }, line: { type: "none" },
  });
  icon(s, "LuCoffee", "white", M + 0.52, TOP + 0.37, 0.32);
  s.addText("トータルサポートパック", {
    x: M + 0.92, y: TOP + 0.28, w: lw - 1.22, h: 0.5,
    fontFace: F.jp, fontSize: 13, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  [["マシン型番", "100RS"], ["付属", "コーヒーサーバー、備品ラック"], ["メンテナンス", "月1回"]].forEach((sp, i) => {
    const y = TOP + 0.98 + i * 0.6;
    s.addText(sp[0], {
      x: M + 0.3, y, w: 1.3, h: 0.32,
      fontFace: F.jp, fontSize: 9.5, color: C.muted, margin: 0, valign: "middle",
    });
    s.addText(sp[1], {
      x: M + 1.64, y, w: lw - 1.94, h: 0.32,
      fontFace: F.jp, fontSize: 11.5, color: C.ink, margin: 0, valign: "middle",
    });
    s.addShape("rect", { x: M + 0.3, y: y + 0.38, w: lw - 0.6, h: 0.011, fill: { color: C.warmLine } });
  });
  tintCard(s, M + 0.3, TOP + 2.86, lw - 0.6, 1.6);
  [["LuCoffee", "挽きたての豆から1杯ずつ抽出"], ["LuTimer", "月1回の定期メンテナンス込み"], ["LuLifeBuoy", "故障・不具合時は本部がサポート"]].forEach((r2, i) => {
    const ry2 = TOP + 3.02 + i * 0.44;
    icon(s, r2[0], "green", M + 0.52, ry2 + 0.03, 0.26);
    s.addText(r2[1], {
      x: M + 0.9, y: ry2, w: lw - 1.3, h: 0.32,
      fontFace: F.jp, fontSize: 10.5, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
  });

  const rx = M + lw + 0.42, rw = R - rx;
  s.addText("＜設置イメージ＞", {
    x: rx, y: TOP, w: rw, h: 0.3,
    fontFace: F.jp, fontSize: 12, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
  photoSlot(s, rx, TOP + 0.4, rw, 3.66, "店内に設置したドリンクマシン", { img: "p24-install.jpg", capSize: 10.5 });
  s.addShape("roundRect", {
    x: rx, y: TOP + 4.2, w: rw, h: 0.52, rectRadius: 0.05,
    fill: { color: C.tint }, line: { type: "none" },
  });
  s.addText("ドリンクマシンは2台設置が基本。既存スペースの一角にも収まります。", {
    x: rx + 0.26, y: TOP + 4.2, w: rw - 0.52, h: 0.52,
    fontFace: F.jp, fontSize: 11, bold: true, color: C.green, margin: 0, valign: "middle",
  });
}

/* ===================================================== FAQ */
{
  const s = pres.addSlide();
  shell(s, "FAQ", "よくある質問",
    "導入をご検討いただく際に、よくいただくご質問をまとめました。");

  const qa = [
    ["準備", "準備は難しいですか？",
      "内装、システム導入まで本部がフルサポートするため、初めての方でも安心です。"],
    ["トラブル対応", "トラブル時の対応は？",
      "24時間体制の警備会社と遠隔監視システムにより、オーナー様が現場に駆けつける必要はほぼありません。また顧客とのLINEチャットを設けており、そこで質問対応・意見収集を行っています。"],
    ["人材", "人材採用については？",
      "新しく人材を採用する必要はありません。清掃管理・原料補充はご自身でもご対応可能ですが、業務委託として委託することも可能です。"],
    ["集客", "集客はどうすれば？",
      "数々のメディア取材や、本部によるSEO・MEO・AIO（LLMO）対策をはじめとしたWebマーケティングによって集客を支援します。"],
    ["許認可", "許認可の申請は必要ですか？",
      "営業届の提出と、食品衛生責任者の設置が必要です。食品衛生責任者の資格は貴社にて取得いただき、営業届の申請は本部が行います。"],
  ];
  const w = (CW - 0.42) / 2, h = 1.5, gy = 0.22;
  qa.forEach(([cat, q, a], i) => {
    const x = M + (i % 2) * (w + 0.42);
    const y = TOP + Math.floor(i / 2) * (h + gy);
    card(s, x, y, w, h);
    eyebrowIn(s, x + 0.3, y + 0.2, 2.6, cat, C.greenLabel);
    s.addText("Q.", {
      x: x + 0.3, y: y + 0.42, w: 0.36, h: 0.32,
      fontFace: F.num, fontSize: 13, bold: true, color: C.gold, margin: 0, valign: "middle",
    });
    s.addText(q, {
      x: x + 0.7, y: y + 0.4, w: w - 1.0, h: 0.34,
      fontFace: F.jp, fontSize: 14.5, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addShape("rect", { x: x + 0.3, y: y + 0.82, w: w - 0.6, h: 0.011, fill: { color: C.warmLine } });
    s.addText("A.", {
      x: x + 0.3, y: y + 0.92, w: 0.36, h: 0.26,
      fontFace: F.num, fontSize: 11.5, bold: true, color: C.green, margin: 0, valign: "middle",
    });
    s.addText(a, {
      x: x + 0.7, y: y + 0.9, w: w - 1.0, h: 0.56,
      fontFace: F.jp, fontSize: 10, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.36,
    });
  });

  // 6枠目：CTA
  const cx = M + (w + 0.42), cy = TOP + 2 * (h + gy);
  panel(s, cx, cy, w, h);
  icon(s, "LuMessageCircle", "white", cx + 0.34, cy + 0.34, 0.36);
  s.addText("その他のご質問もお気軽に", {
    x: cx + 0.88, y: cy + 0.3, w: w - 1.2, h: 0.36,
    fontFace: F.jp, fontSize: 15, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  s.addText("初回面談で、立地条件やご希望に合わせて個別にご説明します。スペースの写真や図面をお持ちいただけると、収益試算までご案内できます。", {
    x: cx + 0.88, y: cy + 0.72, w: w - 1.2, h: 0.6,
    fontFace: F.jp, fontSize: 10, color: C.cvBody, margin: 0, valign: "top", lineSpacingMultiple: 1.36,
  });
}

/* ===================================================== p25 サポート体制 */
{
  const s = pres.addSlide();
  shell(s,  "サポート", "万全のサポート体制",
    "新たに人を採用・配置する必要はありません。開業前から開業後まで本部が伴走します。");

  const cols = [
    ["LuFileCheck", "開業前サポート", "BEFORE OPENING",
      ["パートナー制度のご説明", "収益試算表の作成", "現地調査・工事業者手配", "店舗デザイン・設計", "運営研修"]],
    ["LuLifeBuoy", "開業後サポート", "AFTER OPENING",
      ["店舗運営相談", "マシンの定期メンテナンス", "WEBサイト作成・更新", "SEO・MEO・AIO（LLMO）対策", "改修工事サポート"]],
  ];
  const w = (CW - 0.42) / 2, h = 3.5;
  cols.forEach(([ic, title, eb, list], i) => {
    const x = M + i * (w + 0.42), y = TOP;
    card(s, x, y, w, h);
    icon(s, ic, "green", x + 0.3, y + 0.3, 0.42);
    eyebrowIn(s, x + 0.9, y + 0.3, 3.0, eb);
    s.addText(title, {
      x: x + 0.9, y: y + 0.5, w: w - 1.2, h: 0.36,
      fontFace: F.jp, fontSize: 18, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addShape("rect", { x: x + 0.3, y: y + 1.0, w: w - 0.6, h: 0.011, fill: { color: C.warmLine } });
    list.forEach((t, j) => {
      const ry = y + 1.16 + j * 0.44;
      icon(s, "LuCheck", "green", x + 0.32, ry + 0.06, 0.2);
      s.addText(t, {
        x: x + 0.66, y: ry, w: w - 0.96, h: 0.32,
        fontFace: F.jp, fontSize: 11.5, color: C.ink, margin: 0, valign: "middle",
      });
    });
  });

  panel(s, M, TOP + 3.72, CW, 1.0);
  icon(s, "LuUsers", "white", M + 0.36, TOP + 4.06, 0.32);
  s.addText("専任スタッフを置かずに、店舗運営を続けられる体制をつくります。", {
    x: M + 0.84, y: TOP + 3.86, w: CW - 1.2, h: 0.36,
    fontFace: F.jp, fontSize: 16, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  s.addText("開業前の試算から、開業後のWEB集客・メンテナンスまで本部が担当します。", {
    x: M + 0.84, y: TOP + 4.24, w: CW - 1.2, h: 0.3,
    fontFace: F.jp, fontSize: 10.5, color: C.cvBody, margin: 0, valign: "middle",
  });
}

/* ===================================================== p26 開業までの流れ */
{
  const s = pres.addSlide();
  shell(s,  "導入プロセス", "開業までの流れ",
    "お問い合わせから運営開始まで、2〜3か月程度が目安です。");

  const steps = [
    ["LuUsers", "面談", "制度のご説明と、ご希望条件のヒアリング。"],
    ["LuFileSearch", "加盟審査", "立地・スペースの現地調査と収益試算。"],
    ["LuPenLine", "加盟申込み", "契約締結。契約期間は最低3年〜。"],
    ["LuHammer", "工事", "内装・機器設置。工事業者は本部が手配。"],
    ["LuStore", "運営開始", "運営研修を経て、オープン。"],
  ];
  const gap = 0.3;
  const w = (CW - 4 * gap) / 5, h = 2.94;
  const cy = TOP + 0.86;

  // 「2〜3か月程度」ブラケット（加盟申込み〜運営開始）
  const bx = M + 2 * (w + gap);
  const bw = R - bx;
  s.addShape("roundRect", {
    x: bx + bw / 2 - 1.0, y: TOP, w: 2.0, h: 0.46, rectRadius: 0.23,
    fill: { color: C.goldTint }, line: { color: "E7D3A8", width: 0.75 },
  });
  icon(s, "LuClock", "gold", bx + bw / 2 - 0.78, TOP + 0.11, 0.24);
  s.addText("2〜3か月程度", {
    x: bx + bw / 2 - 0.48, y: TOP, w: 1.5, h: 0.46,
    fontFace: F.jp, fontSize: 11.5, bold: true, color: C.gold, margin: 0, valign: "middle",
  });
  s.addShape("rect", { x: bx, y: TOP + 0.46, w: 0.014, h: 0.3, fill: { color: "DCC79B" } });
  s.addShape("rect", { x: R - 0.014, y: TOP + 0.46, w: 0.014, h: 0.3, fill: { color: "DCC79B" } });
  s.addShape("rect", { x: bx, y: TOP + 0.75, w: bw, h: 0.014, fill: { color: "DCC79B" } });

  steps.forEach(([ic, title, body], i) => {
    const x = M + i * (w + gap);
    const last = i === 4;
    if (last) panel(s, x, cy, w, h); else card(s, x, cy, w, h);
    s.addText(`STEP ${pad2(i + 1)}`, {
      x: x + 0.24, y: cy + 0.26, w: w - 0.48, h: 0.2,
      fontFace: F.num, fontSize: 9, bold: true, color: last ? C.cvPale : C.gold, charSpacing: 1.2, margin: 0, valign: "middle",
    });
    s.addShape("ellipse", {
      x: x + w / 2 - 0.4, y: cy + 0.66, w: 0.8, h: 0.8,
      fill: { color: last ? "17553A" : C.tint }, line: { type: "none" },
    });
    icon(s, ic, last ? "white" : "green", x + w / 2 - 0.22, cy + 0.84, 0.44);
    s.addText(title, {
      x: x + 0.14, y: cy + 1.62, w: w - 0.28, h: 0.4,
      fontFace: F.jp, fontSize: 16, bold: true, color: last ? C.white : C.ink,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(body, {
      x: x + 0.14, y: cy + 2.06, w: w - 0.28, h: 0.78,
      fontFace: F.jp, fontSize: 9.5, color: last ? C.cvBody : C.body,
      align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.35,
    });
    if (!last) {
      s.addShape("triangle", {
        x: x + w + 0.06, y: cy + h / 2 - 0.09, w: 0.18, h: 0.18,
        fill: { color: "CFC9BB" }, line: { type: "none" }, rotate: 90,
      });
    }
  });

  note(s, 6.62, "※ 工事内容・立地条件により期間は変動します。");
}

/* ===================================================== p27 お問い合わせ */
{
  const s = pres.addSlide();
  s.addImage({ path: path.join(A, "cover-bg.png"), x: 0, y: 0, w: 13.333, h: 7.5 });
  s.addImage({ path: LOGO_W, x: 0.889, y: 1.11, w: 2.861, h: 0.407 });
  s.addText(pad2(bare()), {
    x: 11.222, y: 7.069, w: 1.389, h: 0.167,
    fontFace: F.num, fontSize: 10, bold: true, color: C.cvSmall, align: "right", margin: 0, valign: "middle",
  });

  s.addShape("rect", { x: 0.889, y: 2.05, w: 0.778, h: 0.028, fill: { color: C.goldLine } });
  s.addText("お問い合わせ", {
    x: 0.889, y: 2.28, w: 8.0, h: 0.86,
    fontFace: F.jp, fontSize: 42, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  s.addText("会社名・氏名・連絡先をご記入の上ご連絡ください。\nスペースの写真や図面をお持ちいただけると、初回面談で収益試算までご案内できます。", {
    x: 0.889, y: 3.24, w: 7.6, h: 0.72,
    fontFace: F.jp, fontSize: 12.5, color: C.cvBody, margin: 0, valign: "top", lineSpacingMultiple: 1.45,
  });

  [
    ["LuPhone", "TEL", "XXX-XXXX-XXXX", "担当：XXX"],
    ["LuMail", "E-MAIL", "aaaaa@aaaa.aaa", ""],
    ["LuGlobe", "WEB", "https://selfcafe.jp/", ""],
  ].forEach(([ic, label, value, sub], i) => {
    const y = 4.3 + i * 0.78;
    s.addShape("ellipse", { x: 0.889, y, w: 0.5, h: 0.5, fill: { color: "17553A" }, line: { type: "none" } });
    icon(s, ic, "white", 1.03, y + 0.14, 0.22);
    s.addText(label, {
      x: 1.58, y, w: 1.1, h: 0.5,
      fontFace: F.num, fontSize: 9.5, bold: true, color: C.cvPale, charSpacing: 1, margin: 0, valign: "middle",
    });
    s.addText(value, {
      x: 2.72, y, w: 4.6, h: 0.5,
      fontFace: F.num, fontSize: 17, bold: true, color: C.white, margin: 0, valign: "middle",
    });
    if (sub) {
      s.addText(sub, {
        x: 7.4, y, w: 2.0, h: 0.5,
        fontFace: F.jp, fontSize: 10.5, color: C.cvSmall, margin: 0, valign: "middle",
      });
    }
  });

  s.addText("お気軽にご連絡ください", {
    x: 0.889, y: 6.8, w: 5.0, h: 0.26,
    fontFace: F.jp, fontSize: 11, color: C.cvSmall, margin: 0, valign: "middle",
  });
  s.addText("セルフカフェ株式会社", {
    x: 9.0, y: 6.8, w: 3.611, h: 0.26,
    fontFace: F.jp, fontSize: 10.5, color: C.cvSmall, align: "right", margin: 0, valign: "middle",
  });
}

const out = path.join(__dirname, "selfcafe-partner-deck.pptx");
const sharp = require("sharp");
(async () => {
  for (const [dst, { src, aspect, focus }] of CROPS) {
    const meta = await sharp(src).metadata();
    let cw = meta.width, ch = Math.round(cw / aspect);
    if (ch > meta.height) { ch = meta.height; cw = Math.round(ch * aspect); }
    const left = Math.max(0, Math.min(meta.width - cw, Math.round(meta.width * focus - cw / 2)));
    await sharp(src)
      .extract({ left, top: Math.round((meta.height - ch) / 2), width: cw, height: ch })
      .jpeg({ quality: 88 })
      .toFile(dst);
  }
  await pres.writeFile({ fileName: out });
  console.log("written:", out, fs.statSync(out).size, "bytes");
})();
