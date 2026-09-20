/**
 * 未来屋書店様 トライアル導入 特別条件のご提案 — 全6ページ
 *
 * 業態転換提案資料（../selfcafe-miraiya-deck）と同じデザイン基準。
 * メール返信でご提示したトライアル条件のみを扱う独立資料。
 *
 * 実行前に ../selfcafe-miraiya-deck/build-assets.js でアセットを生成しておくこと
 * （assets は ../selfcafe-miraiya-deck/assets へのシンボリックリンク）。
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
  grayText: "6E6A61", // 比較列（可読性のため濃いめ）
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
/** 数字用書体（Arial）は日本語グリフを持たないため、PowerPointが中国語フォントで
 *  代替描画してしまう。日本語を含む文字列は自動で日本語書体に切り替える。 */
const HAS_JP = /[ぁ-んァ-ヶ一-龥々〜ー]/;
const numFace = (s) => (HAS_JP.test(String(s)) ? F.jp : F.num);

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
const FOOTER = "セルフカフェ × 未来屋書店 トライアル導入 特別条件";

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
        { text: it.v, options: { fontFace: numFace(it.v), fontSize: opts.vSize || 20, bold: true, color: gold ? C.gold : C.green } },
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
/** 画像の実寸をヘッダから同期取得（contain 配置の計算用） */
function imgSize(file) {
  const b = fs.readFileSync(file);
  if (b[0] === 0x89 && b[1] === 0x50) return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
  let i = 2;
  while (i < b.length) {
    if (b[i] !== 0xff) { i++; continue; }
    const m = b[i + 1];
    if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) {
      return { h: b.readUInt16BE(i + 5), w: b.readUInt16BE(i + 7) };
    }
    i += 2 + b.readUInt16BE(i + 2);
  }
  return { w: 1, h: 1 };
}
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
    if (opts.fit === "contain") {
      // 図面・フロアマップなど、切らずに枠内へ全体を収める
      const meta = imgSize(img);
      const sc = Math.min(w / meta.w, h / meta.h);
      const dw = meta.w * sc, dh = meta.h * sc;
      s.addImage({ path: img, x: x + (w - dw) / 2, y: y + (h - dh) / 2, w: dw, h: dh });
      return;
    }
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
    fontSize: 11.5, fontFace: F.jp, align: opts.align || "left", valign: "middle",
    margin: [0.05, 0.14, 0.05, 0.14],
  },
});
const tl = (t) => ({
  text: t,
  options: {
    fill: { color: C.grayBand }, color: C.ink, bold: true, fontSize: 12, fontFace: F.jp,
    valign: "middle", margin: [0.05, 0.14, 0.05, 0.14],
  },
});
const td = (t, opts = {}) => ({
  text: t,
  options: {
    fill: { color: opts.fill || C.white }, color: opts.color || C.body,
    bold: !!opts.bold, fontSize: opts.size || 11.5, fontFace: opts.face || F.jp,
    align: opts.align || "left", valign: "middle", margin: [0.05, 0.14, 0.05, 0.14],
  },
});



// ================================================================ deck
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "セルフカフェ株式会社";
pres.title = "セルフカフェ × 未来屋書店 トライアル導入 特別条件のご提案";

/* ===================================================== p1 表紙 */
{
  const s = pres.addSlide();
  bare();
  s.addImage({ path: path.join(A, "cover-bg.png"), x: 0, y: 0, w: 13.333, h: 7.5 });
  s.addImage({ path: LOGO_W, x: 0.889, y: 1.94, w: 2.861, h: 0.407 });

  s.addShape("rect", { x: 0.889, y: 2.94, w: 0.278, h: 0.028, fill: { color: C.goldLine } });
  s.addText("TRIAL — SPECIAL TERMS", {
    x: 1.306, y: 2.843, w: 5.5, h: 0.194,
    fontFace: F.num, fontSize: 10.5, bold: true, color: C.cvPale, charSpacing: 2, margin: 0, valign: "middle",
  });
  s.addText("未来屋書店 秋田店・土浦店", {
    x: 0.889, y: 3.14, w: 7.778, h: 0.32,
    fontFace: F.jp, fontSize: 15, color: C.cvSub, margin: 0, valign: "middle",
  });
  s.addText("トライアル導入 特別条件のご提案", {
    x: 0.889, y: 3.44, w: 10.5, h: 1.0,
    fontFace: F.jp, fontSize: 42, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  s.addText("内装工事は当社が負担し、原状回復義務もいただきません。\nまずは実際の数字をご確認いただいたうえで、その先をご判断ください。", {
    x: 0.889, y: 4.62, w: 8.6, h: 0.72,
    fontFace: F.jp, fontSize: 12.5, color: C.cvBody, margin: 0, valign: "top", lineSpacingMultiple: 1.35,
  });

  s.addShape("rect", { x: 0.889, y: 5.98, w: 0.778, h: 0.028, fill: { color: C.goldLine } });
  const stats = [
    { x: 0.889, w: 3.2, v: "0", u: "円", l: "内装工事費のご負担" },
    { x: 4.5, w: 3.2, v: "25", u: "％", l: "セルフカフェ売上の歩合" },
    { x: 8.1, w: 3.2, v: "なし", u: "", l: "原状回復義務" },
  ];
  stats.forEach((st) => {
    s.addText(
      [
        { text: st.v, options: { fontFace: numFace(st.v), fontSize: 26, bold: true, color: C.white } },
        { text: st.u ? " " + st.u : "", options: { fontFace: F.jp, fontSize: 12, bold: true, color: C.white } },
      ],
      { x: st.x, y: 6.28, w: st.w, h: 0.417, margin: 0, valign: "middle" }
    );
    s.addText(st.l, {
      x: st.x, y: 6.72, w: st.w + 0.4, h: 0.181,
      fontFace: F.jp, fontSize: 9.5, color: C.cvSmall, margin: 0, valign: "middle",
    });
  });
  [4.06, 7.66].forEach((x) =>
    s.addShape("rect", { x, y: 6.31, w: 0.013, h: 0.583, fill: { color: "3C5F4B" } })
  );
  s.addText("セルフカフェ株式会社", {
    x: 9.722, y: 6.9, w: 2.889, h: 0.194,
    fontFace: F.jp, fontSize: 9.5, color: C.cvSmall, align: "right", margin: 0, valign: "middle",
  });
}

/* ===================================================== p2 トライアル特別条件 */
{
  const s = pres.addSlide();
  shell(s, "トライアル特別条件", "通常条件とは別の、トライアル条件です。",
    "メールにてご提示した内容を、あらためて整理したものです。");

  const lw = 7.5;
  const conds = [
    ["LuSparkles", "内装仕様", "スタンダード仕様（既存の書店併設店と同じ設え）"],
    ["LuHammer", "工事部分", "セルフカフェが負担します（御社のご負担なし）"],
    ["LuLandPlot", "家具・什器", "御社にてご手配いただきます（御社の所有となります）"],
    ["LuHandCoins", "歩合", "セルフカフェ売上の25％を御社へお支払い"],
    ["LuKeyRound", "原状回復義務", "なし（当社のマシン・カメラ等のみ撤去します）"],
  ];
  conds.forEach((c2, i) => {
    const y = TOP + i * 0.94;
    tintCard(s, M, y, lw, 0.82);
    icon(s, c2[0], "green", M + 0.28, y + 0.23, 0.36);
    s.addText(c2[1], {
      x: M + 0.78, y: y + 0.1, w: lw - 1.06, h: 0.28,
      fontFace: F.jp, fontSize: 11.5, bold: true, color: C.green, margin: 0, valign: "middle",
    });
    s.addText(c2[2], {
      x: M + 0.78, y: y + 0.4, w: lw - 1.06, h: 0.32,
      fontFace: F.jp, fontSize: 10.5, color: C.ink, margin: 0, valign: "middle",
    });
  });

  const rx = M + lw + 0.42, rw = R - rx;
  panel(s, rx, TOP, rw, 1.9);
  eyebrowIn(s, rx + 0.28, TOP + 0.24, 2.6, "REVENUE SHARE", C.cvPale);
  s.addText("御社へのお支払い", {
    x: rx + 0.28, y: TOP + 0.5, w: rw - 0.56, h: 0.28,
    fontFace: F.jp, fontSize: 11, color: C.cvBody, margin: 0, valign: "middle",
  });
  s.addText(
    [
      { text: "売上の ", options: { fontFace: F.jp, fontSize: 13, color: C.white } },
      { text: "25", options: { fontFace: F.num, fontSize: 34, bold: true, color: "F0C05A" } },
      { text: " ％", options: { fontFace: F.jp, fontSize: 15, bold: true, color: "F0C05A" } },
    ],
    { x: rx + 0.28, y: TOP + 0.82, w: rw - 0.56, h: 0.7, margin: 0, valign: "middle" }
  );
  s.addText("毎月お支払いします（税抜）", {
    x: rx + 0.28, y: TOP + 1.5, w: rw - 0.56, h: 0.28,
    fontFace: F.jp, fontSize: 9.5, color: C.cvBody, margin: 0, valign: "middle",
  });

  tintCard(s, rx, TOP + 2.06, rw, 1.3);
  s.addText("御社のご負担", {
    x: rx + 0.28, y: TOP + 2.24, w: rw - 0.56, h: 0.26,
    fontFace: F.jp, fontSize: 10.5, bold: true, color: C.green, margin: 0, valign: "middle",
  });
  s.addText("家具・什器のご手配のみです。\n内装工事費のご負担はありません。", {
    x: rx + 0.28, y: TOP + 2.54, w: rw - 0.56, h: 0.7,
    fontFace: F.jp, fontSize: 10.5, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
  });

  card(s, rx, TOP + 3.5, rw, 1.08);
  icon(s, "LuShieldCheck", "green", rx + 0.26, TOP + 3.74, 0.34);
  s.addText("原状回復義務なし", {
    x: rx + 0.72, y: TOP + 3.66, w: rw - 0.98, h: 0.28,
    fontFace: F.jp, fontSize: 11, bold: true, color: C.gold, margin: 0, valign: "middle",
  });
  s.addText("終了時も内装を元に戻す必要はありません。", {
    x: rx + 0.72, y: TOP + 3.96, w: rw - 0.98, h: 0.46,
    fontFace: F.jp, fontSize: 9.5, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
  });

  note(s, 6.62, "※ 金額はすべて税抜。トライアルの期間・対象店舗はご相談のうえ決定させていただきます。");
}

/* ===================================================== p3 費用と所有の区分 */
{
  const s = pres.addSlide();
  shell(s, "費用と所有の区分", "どちらが何を用意し、終了後どうなるか。",
    "トライアル期間中の費用負担と、終了後にそれがどう扱われるかを整理しました。");

  const us = (t2) => ({
    text: t2,
    options: { fill: { color: C.white }, color: C.grayText, fontSize: 11, fontFace: F.jp, valign: "middle", margin: [0.05, 0.16, 0.05, 0.16] },
  });
  const you = (t2) => ({
    text: t2,
    options: { fill: { color: C.goldTint }, color: C.gold, bold: true, fontSize: 11, fontFace: F.jp, valign: "middle", margin: [0.05, 0.16, 0.05, 0.16] },
  });
  const after = (t2, hi) => ({
    text: t2,
    options: { fill: { color: hi ? C.tint : C.white }, color: hi ? C.green : C.body, bold: !!hi, fontSize: 11, fontFace: F.jp, valign: "middle", margin: [0.05, 0.16, 0.05, 0.16] },
  });

  s.addTable(
    [
      [th("項目"), th("ご手配・ご負担"), th("トライアル終了後")],
      [tl("内装工事（造作・床・サイン）"), us("セルフカフェ"), after("そのまま残ります", true)],
      [tl("家具・什器（テーブル・椅子・植栽 等）"), you("未来屋書店様"), after("御社の資産として残ります", true)],
      [tl("ドリンクマシン・決済端末"), us("セルフカフェ"), after("当社が撤去します")],
      [tl("防犯カメラ・通信機器"), us("セルフカフェ"), after("当社が撤去します")],
      [tl("原料・機器メンテナンス"), us("セルフカフェ"), after("—")],
      [tl("水道光熱費・通信費・消耗品"), you("未来屋書店様（実費）"), after("—")],
      [tl("原状回復"), after("不要（義務なし）", true), after("—")],
    ],
    {
      x: M, y: TOP, w: CW, colW: [4.4, 3.4, 4.089],
      rowH: [0.4, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5],
      border: { type: "solid", color: C.warmLine, pt: 0.75 },
      autoPage: false,
    }
  );

  const py = 5.80;
  panel(s, M, py, CW, 0.88);
  s.addText(
    [
      { text: "御社にご手配いただくのは、家具・什器だけです。", options: { fontFace: F.jp, fontSize: 13, bold: true, color: C.white } },
      { text: "  内装工事費のご負担はなく、終了時の原状回復も発生しません。", options: { fontFace: F.jp, fontSize: 11, color: C.cvBody } },
    ],
    { x: M + 0.34, y: py, w: CW - 0.68, h: 0.88, margin: 0, valign: "middle" }
  );

  note(s, 6.80, "※ 工事範囲・什器の仕様は現地確認のうえ確定します。金額はすべて税抜。");
}

/* ===================================================== p4 トライアル終了時の扱い */
{
  const s = pres.addSlide();
  shell(s, "終了時の扱い", "原状回復は不要。撤去するのは当社の機器だけです。",
    "トライアルを終了される場合も、内装を元に戻していただく必要はありません。");

  const cw = (CW - 0.42) / 2;
  const blocks = [
    {
      x: M, tint: true, ic: "LuHouse", title: "そのまま残るもの",
      items: ["内装工事（造作・床・サイン）", "家具・什器（テーブル・椅子・植栽 等）", "書店側の売場・動線"],
      foot: "工事部分は当社負担で施工しますが、撤去はいたしません。",
    },
    {
      x: M + cw + 0.42, tint: false, ic: "LuMinus", title: "当社が撤去するもの",
      items: ["ドリンクマシン", "決済端末", "防犯カメラ・通信機器"],
      foot: "撤去費用は当社が負担します。跡の補修も当社で対応します。",
    },
  ];
  blocks.forEach((b) => {
    if (b.tint) tintCard(s, b.x, TOP, cw, 3.2);
    else card(s, b.x, TOP, cw, 3.2);
    icon(s, b.ic, "green", b.x + 0.3, TOP + 0.26, 0.36);
    s.addText(b.title, {
      x: b.x + 0.8, y: TOP + 0.2, w: cw - 1.1, h: 0.42,
      fontFace: F.jp, fontSize: 14, bold: true, color: b.tint ? C.green : C.ink, margin: 0, valign: "middle",
    });
    s.addShape("rect", { x: b.x + 0.3, y: TOP + 0.78, w: cw - 0.6, h: 0.011, fill: { color: b.tint ? C.tintLine : C.warmLine } });
    b.items.forEach((t2, i) => {
      const y = TOP + 0.96 + i * 0.46;
      icon(s, b.tint ? "LuCheck" : "LuX", b.tint ? "green" : "muted", b.x + 0.32, y + 0.07, 0.2);
      s.addText(t2, {
        x: b.x + 0.68, y, w: cw - 1.0, h: 0.34,
        fontFace: F.jp, fontSize: 11, color: C.ink, margin: 0, valign: "middle",
      });
    });
    s.addText(b.foot, {
      x: b.x + 0.3, y: TOP + 2.48, w: cw - 0.6, h: 0.56,
      fontFace: F.jp, fontSize: 9.5, color: C.muted, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
    });
  });

  const py = 5.16;
  panel(s, M, py, CW, 1.1);
  icon(s, "LuBadgeCheck", "pale", M + 0.38, py + 0.3, 0.44);
  s.addText(
    [
      { text: "トライアル後は、基本そのまま残す形を想定しています。", options: { fontFace: F.jp, fontSize: 14, bold: true, color: C.white, breakLine: true } },
      { text: "内装も什器も残るため、継続いただく場合は追加の工事なくそのまま運営を続けられます。", options: { fontFace: F.jp, fontSize: 11, color: C.cvBody } },
    ],
    { x: M + 1.0, y: py, w: CW - 1.4, h: 1.1, margin: 0, valign: "middle", lineSpacingMultiple: 1.25 }
  );

  note(s, 6.44, "※ 撤去の範囲・時期は事前にお打ち合わせのうえ決定します。什器の再利用可否は現地確認時にご相談させてください。");
}

/* ===================================================== p5 家具を御社所有とする理由 */
{
  const s = pres.addSlide();
  shell(s, "家具・什器の考え方", "家具は御社ご所有のほうがよい、と判断しました。",
    "トライアル後もそのまま残す前提で考えると、御社の資産にしておくのが最も無駄がありません。");

  const cw = (CW - 0.6) / 3;
  const reasons = [
    ["LuStore", "そのまま使い続けられる", "トライアル後に継続される場合、什器の入れ替えも買い直しも発生しません。設えを変えずに運営を続けられます。"],
    ["LuShieldCheck", "御社の資産として残る", "当社所有にすると、契約が終わる際に引き上げか買取かの整理が必要になります。はじめから御社ご所有であれば、その論点自体がなくなります。"],
    ["LuPenLine", "仕様を御社で選べる", "テーブル・椅子・植栽は店舗の雰囲気を大きく左右します。書店の売場に合わせて御社側でお選びいただけます。"],
  ];
  reasons.forEach((r, i) => {
    const x = M + i * (cw + 0.3);
    card(s, x, TOP, cw, 3.3);
    icon(s, r[0], "green", x + 0.3, TOP + 0.3, 0.4);
    s.addText(r[1], {
      x: x + 0.3, y: TOP + 0.88, w: cw - 0.6, h: 0.6,
      fontFace: F.jp, fontSize: 13, bold: true, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.2,
    });
    s.addShape("rect", { x: x + 0.3, y: TOP + 1.58, w: cw - 0.6, h: 0.011, fill: { color: C.warmLine } });
    s.addText(r[2], {
      x: x + 0.3, y: TOP + 1.74, w: cw - 0.6, h: 1.4,
      fontFace: F.jp, fontSize: 10, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.4,
    });
  });

  const py = 5.26;
  tintCard(s, M, py, CW, 1.0);
  icon(s, "LuMessagesSquare", "green", M + 0.34, py + 0.3, 0.4);
  s.addText("当社側で家具・什器を手配することも可能です。その場合は当社の資産となるため、トライアル終了時に引き上げ・買取のご相談が必要になります。ご希望があればその形でもお受けします。", {
    x: M + 0.92, y: py, w: CW - 1.3, h: 1.0,
    fontFace: F.jp, fontSize: 10.5, color: C.ink, margin: 0, valign: "middle", lineSpacingMultiple: 1.35,
  });

  note(s, 6.44, "※ 什器の概算費用は、レイアウト確定後にお見積りをお出しします。既存什器を再利用できる場合は、その分ご負担が減ります。");
}

/* ===================================================== p6 今後の進め方／お問い合わせ */
{
  const s = pres.addSlide();
  shell(s, "今後の進め方", "現地確認から、最短2〜3ヶ月で開始できます。",
    "まずは対象店舗と期間のすり合わせ、現地確認からお願いできればと存じます。");

  const steps = [
    ["LuMessagesSquare", "STEP 01", "条件のすり合わせ", "対象店舗・トライアル期間・開始時期を決めます。"],
    ["LuFileSearch", "STEP 02", "現地確認", "区画・電源・動線を確認し、レイアウトと什器の仕様を設計。"],
    ["LuPenLine", "STEP 03", "契約締結", "トライアル条件を明記した契約を締結します。"],
    ["LuHammer", "STEP 04", "工事・設置", "内装工事は当社が施工。家具・什器は御社にてご手配。"],
    ["LuStore", "STEP 05", "トライアル開始", "運営・集客・お問い合わせ対応は当社が担当します。"],
  ];
  const w = (CW - 4 * 0.22) / 5;
  steps.forEach((st, i) => {
    const x = M + i * (w + 0.22);
    card(s, x, TOP, w, 2.72);
    eyebrowIn(s, x + 0.24, TOP + 0.24, 1.6, st[1]);
    icon(s, st[0], "green", x + w - 0.72, TOP + 0.2, 0.36);
    s.addText(st[2], {
      x: x + 0.24, y: TOP + 0.62, w: w - 0.48, h: 0.5,
      fontFace: F.jp, fontSize: 12.5, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addShape("rect", { x: x + 0.24, y: TOP + 1.18, w: w - 0.48, h: 0.011, fill: { color: C.warmLine } });
    s.addText(st[3], {
      x: x + 0.24, y: TOP + 1.32, w: w - 0.48, h: 1.2,
      fontFace: F.jp, fontSize: 9.5, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.4,
    });
  });

  panel(s, M, TOP + 2.98, CW, 1.9);
  s.addImage({ path: LOGO_W, x: M + 0.44, y: TOP + 3.26, w: 2.29, h: 0.326 });
  s.addText("ご不明な点は、お気軽にお問い合わせください。", {
    x: M + 0.44, y: TOP + 3.74, w: 6.4, h: 0.34,
    fontFace: F.jp, fontSize: 13, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  const cx = M + 7.4;
  [["LuPhone", "TEL", "052-879-3557"], ["LuMail", "E-MAIL", "info@selfcafe.jp"], ["LuGlobe", "WEB", "https://selfcafe.jp/"]].forEach(([ic, lb, v], i) => {
    const y = TOP + 3.22 + i * 0.5;
    icon(s, ic, "pale", cx, y + 0.06, 0.24);
    s.addText(lb, {
      x: cx + 0.4, y, w: 1.0, h: 0.36,
      fontFace: F.num, fontSize: 9, bold: true, color: C.cvPale, charSpacing: 1, margin: 0, valign: "middle",
    });
    s.addText(v, {
      x: cx + 1.3, y, w: 3.2, h: 0.36,
      fontFace: numFace(v), fontSize: 13, bold: true, color: C.white, margin: 0, valign: "middle",
    });
  });
}

// ---------------------------------------------------------------- write
const out = path.join(__dirname, "selfcafe-miraiya-trial-deck.pptx");
const sharp = require("sharp");
(async () => {
  for (const [dst, { src, aspect, focus }] of CROPS) {
    const meta = await sharp(src).metadata();
    const srcAspect = meta.width / meta.height;
    let w = meta.width, h = meta.height, left = 0, top = 0;
    if (srcAspect > aspect) { w = Math.round(meta.height * aspect); left = Math.round((meta.width - w) * focus); }
    else { h = Math.round(meta.width / aspect); top = Math.round((meta.height - h) / 2); }
    await sharp(src).extract({ left, top, width: w, height: h }).jpeg({ quality: 88 }).toFile(dst);
  }
  await pres.writeFile({ fileName: out });
  console.log("written:", out);
})();
