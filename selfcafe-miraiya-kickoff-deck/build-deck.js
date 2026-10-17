/**
 * 未来屋書店様 トライアル導入のご提案 — 全7ページ
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
const FOOTER = "セルフカフェ × 未来屋書店 秋田店 ご確認事項とスケジュール";

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
pres.title = "セルフカフェ × 未来屋書店 秋田店 ご確認事項とスケジュール";

/* ===================================================== p1 表紙 */
{
  const s = pres.addSlide();
  bare();
  s.addImage({ path: path.join(A, "cover-bg.png"), x: 0, y: 0, w: 13.333, h: 7.5 });

  s.addText("セルフカフェ株式会社", {
    x: 0.889, y: 0.83, w: 4.0, h: 0.194,
    fontFace: F.jp, fontSize: 9.5, color: C.cvSmall, margin: 0, valign: "middle",
  });
  s.addImage({ path: LOGO_W, x: 0.889, y: 1.94, w: 2.861, h: 0.407 });

  s.addShape("rect", { x: 0.889, y: 2.94, w: 0.278, h: 0.028, fill: { color: C.goldLine } });
  s.addText("PROJECT KICK-OFF", {
    x: 1.306, y: 2.843, w: 5.5, h: 0.194,
    fontFace: F.num, fontSize: 10.5, bold: true, color: C.cvPale, charSpacing: 2, margin: 0, valign: "middle",
  });
  s.addText("未来屋書店 秋田店　カフェ区画 業態転換", {
    x: 0.889, y: 3.14, w: 8.6, h: 0.32,
    fontFace: F.jp, fontSize: 15, color: C.cvSub, margin: 0, valign: "middle",
  });
  s.addText("スケジュールとご確認事項", {
    x: 0.889, y: 3.44, w: 10.5, h: 1.0,
    fontFace: F.jp, fontSize: 42, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  s.addText("2027年2月末のオープンに向けた工程と、事前にご確認・ご準備いただきたい事項を\n整理しました。工期の目安と、仕様確定・発注の期限もあわせてご覧ください。", {
    x: 0.889, y: 4.62, w: 9.2, h: 0.72,
    fontFace: F.jp, fontSize: 12.5, color: C.cvBody, margin: 0, valign: "top", lineSpacingMultiple: 1.35,
  });

  s.addShape("rect", { x: 0.889, y: 5.98, w: 0.778, h: 0.028, fill: { color: C.goldLine } });
  const stats = [
    { x: 0.889, w: 3.4, v: "2027", u: "年2月末", l: "オープン目標" },
    { x: 4.7, w: 3.4, v: "3〜4", u: "週間", l: "閉店後の工事期間" },
    { x: 8.5, w: 3.4, v: "12", u: "月上旬", l: "仕様確定・発注の期限" },
  ];
  stats.forEach((st) => {
    s.addText(
      [
        { text: st.v, options: { fontFace: numFace(st.v), fontSize: 24, bold: true, color: "F0C05A" } },
        { text: " " + st.u, options: { fontFace: F.jp, fontSize: 12, bold: true, color: "F0C05A" } },
      ],
      { x: st.x, y: 6.28, w: st.w, h: 0.417, margin: 0, valign: "middle" }
    );
    s.addText(st.l, {
      x: st.x, y: 6.72, w: st.w + 0.3, h: 0.181,
      fontFace: F.jp, fontSize: 9.5, color: C.cvSmall, margin: 0, valign: "middle",
    });
  });
  [4.42, 8.22].forEach((x) =>
    s.addShape("rect", { x, y: 6.31, w: 0.013, h: 0.583, fill: { color: "3C5F4B" } })
  );
}

/* ===================================================== p2 スケジュール案 */
{
  const s = pres.addSlide();
  shell(s, "スケジュール案", "2月末オープンは、十分に可能です。",
    "ご希望の「1月末クリエ閉店・2月末オープン」から逆算した工程です。");

  const lw = 3.35, months = ["2026年10月", "11月", "12月", "2027年1月", "2月"];
  const cw = (CW - lw) / months.length;

  // 月見出し
  months.forEach((m, i) => {
    const x = M + lw + i * cw;
    s.addShape("rect", { x, y: TOP, w: cw - 0.02, h: 0.3, fill: { color: i % 2 ? C.grayBand : C.tint } });
    s.addText(m, {
      x, y: TOP, w: cw, h: 0.3,
      fontFace: F.jp, fontSize: 9.5, bold: true, color: C.ink, align: "center", margin: 0, valign: "middle",
    });
  });

  const tasks = [
    { t: "現地確認・実測（電源・給排水）", s0: 0.65, len: 0.35, who: "us" },
    { t: "レイアウト・仕様の確定（30坪／60坪）", s0: 1.0, len: 0.65, who: "both" },
    { t: "業務委託契約の締結", s0: 1.5, len: 0.45, who: "both" },
    { t: "什器・機器の発注 〜 納品", s0: 1.9, len: 1.75, who: "us", note: "リードタイム 4〜6週間" },
    { t: "カフェ・ド・クリエ 閉店・退去", s0: 3.55, len: 0.45, who: "you" },
    { t: "内装工事", s0: 4.0, len: 0.55, who: "us" },
    { t: "機器設置・試運転・什器搬入", s0: 4.5, len: 0.35, who: "us" },
    { t: "オープン準備・研修", s0: 4.75, len: 0.25, who: "both" },
  ];
  const color = { us: C.greenDeep, you: C.gold, both: "6FA98A" };
  tasks.forEach((tk, i) => {
    const y = TOP + 0.42 + i * 0.52;
    if (i % 2 === 0) s.addShape("rect", { x: M, y: y - 0.06, w: CW, h: 0.46, fill: { color: "FAF8F2" } });
    s.addText(tk.t, {
      x: M, y, w: lw - 0.12, h: 0.34,
      fontFace: F.jp, fontSize: 10, color: C.ink, margin: 0, valign: "middle",
    });
    s.addShape("roundRect", {
      x: M + lw + tk.s0 * cw, y: y + 0.04, w: tk.len * cw, h: 0.26, rectRadius: 0.05,
      fill: { color: color[tk.who] }, line: { type: "none" },
    });
    if (tk.note) {
      s.addText(tk.note, {
        x: M + lw + (tk.s0 + tk.len) * cw + 0.08, y, w: 1.8, h: 0.34,
        fontFace: F.jp, fontSize: 8.5, color: C.gold, margin: 0, valign: "middle",
      });
    }
  });

  // オープン位置のマーカー
  const openX = M + lw + 5 * cw;
  s.addShape("rect", { x: openX - 0.012, y: TOP + 0.3, w: 0.024, h: 4.3, fill: { color: C.gold } });
  s.addShape("roundRect", {
    x: openX - 0.95, y: TOP + 4.64, w: 0.95, h: 0.34, rectRadius: 0.05,
    fill: { color: C.gold }, line: { type: "none" },
  });
  s.addText("OPEN", {
    x: openX - 0.95, y: TOP + 4.64, w: 0.95, h: 0.34,
    fontFace: F.num, fontSize: 10, bold: true, color: C.white, align: "center", charSpacing: 1, margin: 0, valign: "middle",
  });

  // 凡例
  [["セルフカフェ", color.us], ["未来屋書店様", color.you], ["両社", color.both]].forEach((lg, i) => {
    const x = M + i * 1.9;
    s.addShape("roundRect", { x, y: TOP + 4.72, w: 0.26, h: 0.18, rectRadius: 0.04, fill: { color: lg[1] }, line: { type: "none" } });
    s.addText(lg[0], {
      x: x + 0.34, y: TOP + 4.64, w: 1.5, h: 0.34,
      fontFace: F.jp, fontSize: 9, color: C.muted, margin: 0, valign: "middle",
    });
  });

  note(s, 6.80, "※ 目安の工程です。クリエ様の原状回復範囲、館への工事申請、什器の納期により前後します。");
}

/* ===================================================== p3 工期の考え方 */
{
  const s = pres.addSlide();
  shell(s, "工期の考え方", "工事は3〜4週間。ただし発注は先行して必要です。",
    "閉店後に必要な期間と、閉店前から動かしておく必要がある期間を分けて整理しました。");

  // 上段：閉店後に必要な期間
  tintCard(s, M, TOP, CW, 2.1);
  s.addText("① 閉店後に必要な期間　合計 3〜4週間", {
    x: M + 0.32, y: TOP + 0.2, w: CW - 0.64, h: 0.3,
    fontFace: F.jp, fontSize: 12, bold: true, color: C.green, margin: 0, valign: "middle",
  });
  const after = [
    ["LuHammer", "残置物の撤去・下地調整", "3〜5日", "クリエ様の原状回復範囲により変動します"],
    ["LuRuler", "内装工事（造作・床・サイン）", "2〜3週間", "スタンダード仕様の場合の目安です"],
    ["LuPlug", "機器設置・試運転・什器搬入", "3〜5日", "電源・給排水が既存で使える前提です"],
  ];
  const aw = (CW - 0.64 - 0.4) / 3;
  after.forEach((a, i) => {
    const x = M + 0.32 + i * (aw + 0.2);
    card(s, x, TOP + 0.58, aw, 1.34);
    icon(s, a[0], "green", x + 0.22, TOP + 0.76, 0.3);
    s.addText(a[2], {
      x: x + aw - 1.3, y: TOP + 0.72, w: 1.1, h: 0.3,
      fontFace: F.jp, fontSize: 13, bold: true, color: C.gold, align: "right", margin: 0, valign: "middle",
    });
    s.addText(a[1], {
      x: x + 0.22, y: TOP + 1.1, w: aw - 0.44, h: 0.28,
      fontFace: F.jp, fontSize: 10.5, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(a[3], {
      x: x + 0.22, y: TOP + 1.38, w: aw - 0.44, h: 0.44,
      fontFace: F.jp, fontSize: 8.5, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
    });
  });

  // 下段：閉店前に動かす必要がある期間
  card(s, M, TOP + 2.28, CW, 1.5, { fill: C.goldTint, line: "E7D3A8" });
  s.addText("② 閉店前から動かしておく期間　什器・機器の発注リードタイム 4〜6週間", {
    x: M + 0.32, y: TOP + 2.46, w: CW - 0.64, h: 0.3,
    fontFace: F.jp, fontSize: 12, bold: true, color: C.gold, margin: 0, valign: "middle",
  });
  s.addText("ドリンクマシン・決済端末・什器は受注生産・取り寄せとなるため、発注から納品まで4〜6週間を要します。\n工事は閉店後ですが、この発注だけは閉店前に済ませておく必要があります。発注にはレイアウトと坪数（30坪／60坪）の確定が前提です。", {
    x: M + 0.32, y: TOP + 2.8, w: CW - 0.64, h: 0.86,
    fontFace: F.jp, fontSize: 10.5, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.4,
  });

  const py = 5.78;
  panel(s, M, py, CW, 0.94);
  icon(s, "LuBellRing", "pale", M + 0.38, py + 0.26, 0.44);
  s.addText(
    [
      { text: "1月末閉店・2月末オープンは可能です。ただし12月上旬までに仕様確定と発注が必要です。", options: { fontFace: F.jp, fontSize: 12.5, bold: true, color: C.white, breakLine: true } },
      { text: "逆算すると、坪数とレイアウトのご決定は11月中旬まで、現地確認は10月中にお願いできれば確実です。", options: { fontFace: F.jp, fontSize: 10.5, color: C.cvBody } },
    ],
    { x: M + 1.0, y: py, w: CW - 1.4, h: 0.94, margin: 0, valign: "middle", lineSpacingMultiple: 1.25 }
  );

  note(s, 6.84, "※ 館（イオンモール様）の工事申請に要する期間は含んでいません。申請期限が判明次第、工程に織り込みます。");
}

/* ===================================================== p4 ご確認事項 */
{
  const s = pres.addSlide();
  shell(s, "ご確認事項", "先に確定させたい事項を、期限つきで整理しました。",
    "★印の3点は工期とコストに直結するため、優先してご確認いただけますと助かります。");

  const colW = [3.5, 6.1, 2.289];
  const item = (t2, star) => ({
    text: (star ? "★ " : "") + t2,
    options: {
      fill: { color: star ? C.goldTint : C.grayBand }, color: star ? C.gold : C.ink,
      bold: true, fontSize: 10.5, fontFace: F.jp, valign: "middle", margin: [0.04, 0.14, 0.04, 0.14],
    },
  });
  const desc = (t2) => ({
    text: t2,
    options: { fill: { color: C.white }, color: C.body, fontSize: 9.5, fontFace: F.jp, valign: "middle", margin: [0.04, 0.14, 0.04, 0.14] },
  });
  const when = (t2, hi) => ({
    text: t2,
    options: {
      fill: { color: hi ? C.tint : C.white }, color: hi ? C.green : C.muted,
      bold: !!hi, fontSize: 10, fontFace: F.jp, align: "center", valign: "middle", margin: [0.04, 0.08, 0.04, 0.08],
    },
  });

  s.addTable(
    [
      [th("ご確認事項"), th("内容"), th("ご回答の目安", { align: "center" })],
      [item("クリエ様の解約条件・閉店日", 1), desc("6か月前通告の取扱い、違約金の見通し、閉店日の確定"), when("11月上旬", 1)],
      [item("原状回復の範囲", 1), desc("スケルトン戻しか、造作・床・設備を残せるか。工期とコストに最も影響します"), when("11月上旬", 1)],
      [item("残置物・既存設備の扱い", 1), desc("什器・給排水・電気容量・ダクトを引き継げるか。使えれば工期が短縮できます"), when("11月上旬", 1)],
      [item("区画の確定（30坪／60坪）"), desc("レイアウトと什器数量が決まります。発注前の確定が必要です"), when("11月中旬", 1)],
      [item("館への工事申請"), desc("イオンモール様の申請様式・提出期限、工事可能時間帯、搬入ルート"), when("11月中旬")],
      [item("サイン表記の確定"), desc("「MIRAIYA Book & Cafe」の最終表記と、館の看板規定の確認"), when("11月中旬")],
      [item("社内稟議の進捗"), desc("決裁完了の目処。発注判断の期限と合わせて共有させてください"), when("11月中旬")],
      [item("運営体制"), desc("営業時間、清掃・補充の担当、書店スタッフとの役割分担"), when("12月上旬")],
    ],
    {
      x: M, y: TOP, w: CW, colW,
      rowH: [0.38, 0.42, 0.46, 0.46, 0.42, 0.42, 0.42, 0.42, 0.42],
      border: { type: "solid", color: C.warmLine, pt: 0.75 },
      autoPage: false,
    }
  );

  const py = 5.72;
  panel(s, M, py, CW, 0.9);
  s.addText(
    [
      { text: "まずは★印の3点から。", options: { fontFace: F.jp, fontSize: 13, bold: true, color: C.white } },
      { text: "　原状回復の範囲と残置設備の扱いが決まれば、工期と費用の精度が一気に上がります。", options: { fontFace: F.jp, fontSize: 11, color: C.cvBody } },
    ],
    { x: M + 0.34, y: py, w: CW - 0.68, h: 0.9, margin: 0, valign: "middle" }
  );

  note(s, 6.76, "※ ご回答の目安は、2月末オープンから逆算した日程です。前後する場合は工程をあらためて調整いたします。");
}

/* ===================================================== p5 ご準備／弊社対応 */
{
  const s = pres.addSlide();
  shell(s, "資料と役割分担", "ご提供いただきたい資料と、弊社で進めることです。",
    "現地確認までにお手元の資料をご共有いただけますと、設計と見積りを先行して進められます。");

  const cw = (CW - 0.42) / 2;

  tintCard(s, M, TOP, cw, 4.3);
  icon(s, "LuClipboardList", "green", M + 0.3, TOP + 0.26, 0.36);
  s.addText("ご提供いただきたい資料", {
    x: M + 0.8, y: TOP + 0.2, w: cw - 1.1, h: 0.42,
    fontFace: F.jp, fontSize: 14, bold: true, color: C.green, margin: 0, valign: "middle",
  });
  s.addShape("rect", { x: M + 0.3, y: TOP + 0.78, w: cw - 0.6, h: 0.011, fill: { color: C.tintLine } });
  [
    ["区画の図面（平面図・展開図）", "CADデータがあれば実測の手間が省けます"],
    ["電気容量・分電盤の位置", "増設の要否がこれで判断できます"],
    ["給排水の位置・系統", "製氷・給湯の可否に関わります"],
    ["現況写真（既存設備・厨房まわり）", "残置物の判断材料になります"],
    ["館の工事規定・申請様式", "申請期限から逆算して工程を確定します"],
    ["クリエ様の原状回復条件がわかる資料", "契約書の該当箇所の抜粋で結構です"],
  ].forEach((t2, i) => {
    const y = TOP + 0.94 + i * 0.56;
    icon(s, "LuCheck", "green", M + 0.32, y + 0.06, 0.2);
    s.addText(t2[0], {
      x: M + 0.68, y, w: cw - 1.0, h: 0.26,
      fontFace: F.jp, fontSize: 10.5, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(t2[1], {
      x: M + 0.68, y: y + 0.24, w: cw - 1.0, h: 0.24,
      fontFace: F.jp, fontSize: 8.5, color: C.muted, margin: 0, valign: "middle",
    });
  });

  const rx = M + cw + 0.42;
  card(s, rx, TOP, cw, 4.3);
  icon(s, "LuBadgeCheck", "green", rx + 0.3, TOP + 0.26, 0.36);
  s.addText("弊社で進めること", {
    x: rx + 0.8, y: TOP + 0.2, w: cw - 1.1, h: 0.42,
    fontFace: F.jp, fontSize: 14, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
  s.addShape("rect", { x: rx + 0.3, y: TOP + 0.78, w: cw - 0.6, h: 0.011, fill: { color: C.warmLine } });
  [
    ["現地確認の日程調整", "10月中にお伺いできればと存じます"],
    ["レイアウト図・什器リストの作成", "30坪・60坪の両案でご用意します"],
    ["工事費・初期費用のお見積り", "原状回復の範囲が判明次第お出しします"],
    ["業務委託契約書のドラフト送付", "11月上旬を目途にお送りします"],
    ["館への工事申請のサポート", "図面・仕様書の作成を弊社で行います"],
    ["開業スケジュールの確定版の提示", "現地確認後に確定版をお送りします"],
  ].forEach((t2, i) => {
    const y = TOP + 0.94 + i * 0.56;
    icon(s, "LuCheck", "green", rx + 0.32, y + 0.06, 0.2);
    s.addText(t2[0], {
      x: rx + 0.68, y, w: cw - 1.0, h: 0.26,
      fontFace: F.jp, fontSize: 10.5, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(t2[1], {
      x: rx + 0.68, y: y + 0.24, w: cw - 1.0, h: 0.24,
      fontFace: F.jp, fontSize: 8.5, color: C.muted, margin: 0, valign: "middle",
    });
  });

  note(s, 6.4, "※ すべてが揃っていなくても構いません。お手元にあるものから順次ご共有いただければ、こちらで補いながら進めます。");
}

/* ===================================================== p6 直近のお願い／お問い合わせ */
{
  const s = pres.addSlide();
  shell(s, "直近のお願い", "まずは現地確認の日程をいただけますでしょうか。",
    "現地を拝見できれば、残置設備の活用可否と工期・費用の精度が大きく上がります。");

  const steps = [
    ["LuCalendarCheck", "STEP 01", "現地確認（10月中）", "区画・電源・給排水・動線と、既存設備の活用可否を確認します。"],
    ["LuRuler", "STEP 02", "レイアウト・見積り", "30坪・60坪の両案で図面と費用をお出しします。"],
    ["LuPenLine", "STEP 03", "区画確定・契約締結", "坪数をご決定いただき、業務委託契約を締結します。"],
    ["LuShoppingBag", "STEP 04", "発注（12月上旬）", "什器・機器を発注します。ここが工程上の分岐点です。"],
    ["LuStore", "STEP 05", "工事・オープン", "1月末の閉店後に着工し、2月末のオープンを目指します。"],
  ];
  const w = (CW - 4 * 0.22) / 5;
  steps.forEach((st, i) => {
    const x = M + i * (w + 0.22);
    if (i === 3) card(s, x, TOP, w, 2.72, { fill: C.goldTint, line: "E7D3A8" });
    else card(s, x, TOP, w, 2.72);
    eyebrowIn(s, x + 0.24, TOP + 0.24, 1.6, st[1], i === 3 ? C.gold : undefined);
    icon(s, st[0], i === 3 ? "gold" : "green", x + w - 0.72, TOP + 0.2, 0.36);
    s.addText(st[2], {
      x: x + 0.24, y: TOP + 0.62, w: w - 0.48, h: 0.5,
      fontFace: F.jp, fontSize: 12, bold: true, color: i === 3 ? C.gold : C.ink, margin: 0, valign: "middle",
    });
    s.addShape("rect", { x: x + 0.24, y: TOP + 1.18, w: w - 0.48, h: 0.011, fill: { color: i === 3 ? "E7D3A8" : C.warmLine } });
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
const out = path.join(__dirname, "selfcafe-miraiya-kickoff-deck.pptx");
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
