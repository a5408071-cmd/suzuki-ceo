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
const FOOTER = "セルフカフェ × 未来屋書店 業態転換のご提案";

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
pres.title = "セルフカフェ × 未来屋書店 業態転換のご提案";

/* ===================================================== p1 表紙 */
{
  const s = pres.addSlide();
  bare();
  s.addImage({ path: path.join(A, "cover-bg.png"), x: 0, y: 0, w: 13.333, h: 7.5 });
  s.addImage({ path: LOGO_W, x: 0.889, y: 1.94, w: 2.861, h: 0.407 });

  s.addShape("rect", { x: 0.889, y: 2.94, w: 0.278, h: 0.028, fill: { color: C.goldLine } });
  s.addText("BUSINESS FORMAT CONVERSION", {
    x: 1.306, y: 2.843, w: 5.5, h: 0.194,
    fontFace: F.num, fontSize: 10.5, bold: true, color: C.cvPale, charSpacing: 2, margin: 0, valign: "middle",
  });
  s.addText("未来屋書店 秋田店・土浦店", {
    x: 0.889, y: 3.14, w: 7.778, h: 0.32,
    fontFace: F.jp, fontSize: 15, color: C.cvSub, margin: 0, valign: "middle",
  });
  s.addText("カフェ区画 業態転換のご提案", {
    x: 0.889, y: 3.44, w: 9.5, h: 1.0,
    fontFace: F.jp, fontSize: 44, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  s.addText("既存のカフェ区画を無人型「セルフカフェ」へ。\n人員配置なしで運営し、売上に応じた収入が毎月入る形へ切り替えます。", {
    x: 0.889, y: 4.62, w: 8.2, h: 0.72,
    fontFace: F.jp, fontSize: 12.5, color: C.cvBody, margin: 0, valign: "top", lineSpacingMultiple: 1.35,
  });

  s.addShape("rect", { x: 0.889, y: 5.98, w: 0.778, h: 0.028, fill: { color: C.goldLine } });
  const stats = [
    { x: 0.889, w: 3.2, v: "0", u: "人", l: "カフェ運営に必要な人員" },
    { x: 4.5, w: 3.2, v: "27.0", u: "万円／月", l: "秋田店 想定お受取額" },
    { x: 8.1, w: 3.2, v: "23.9", u: "万円／月", l: "土浦店 想定お受取額" },
  ];
  stats.forEach((st) => {
    s.addText(
      [
        { text: st.v, options: { fontFace: numFace(st.v), fontSize: 26, bold: true, color: C.white } },
        { text: " " + st.u, options: { fontFace: F.jp, fontSize: 12, bold: true, color: C.white } },
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

/* ===================================================== p2 ご提案の要旨 */
{
  const s = pres.addSlide();
  shell(s, "ご提案の要旨", "人を置かずに、カフェ区画を収益区画へ。",
    "フルサービス型のカフェ業態から無人型セルフカフェへ転換し、運営はセルフカフェ本部が担います。");

  const cards = [
    { dark: false, label: "運営体制", v: "0", u: "人", sub: "カフェ専任の人員配置は不要。\n運営主体はセルフカフェ本部です。" },
    { dark: true, label: "秋田店 お受取額（想定）", v: "27.0", u: "万円／月", sub: "年間 324.6万円\n売上の25％＋月5万円" },
    { dark: false, label: "初期費用", v: "30〜50", u: "万円", sub: "機器設置費のみ。加盟金・\n研修費はいただきません。" },
    { dark: true, label: "土浦店 お受取額（想定）", v: "23.9", u: "万円／月", sub: "年間 286.8万円\n売上の25％＋月5万円" },
  ];
  const cw = (CW - 0.42) / 2, chh = 1.86;
  cards.forEach((k, i) => {
    const x = M + (i % 2) * (cw + 0.42);
    const y = TOP + Math.floor(i / 2) * (chh + 0.26);
    if (k.dark) panel(s, x, y, cw, chh); else tintCard(s, x, y, cw, chh);
    s.addText(k.label, {
      x: x + 0.34, y: y + 0.22, w: cw - 0.68, h: 0.28,
      fontFace: F.jp, fontSize: 11.5, bold: true, color: k.dark ? C.cvPale : C.green, margin: 0, valign: "middle",
    });
    s.addText(
      [
        { text: k.v, options: { fontFace: numFace(k.v), fontSize: 32, bold: true, color: k.dark ? "F0C05A" : C.green } },
        { text: " " + k.u, options: { fontFace: F.jp, fontSize: 15, bold: true, color: k.dark ? "F0C05A" : C.green } },
      ],
      { x: x + 0.34, y: y + 0.56, w: cw - 0.68, h: 0.72, margin: 0, valign: "middle" }
    );
    s.addText(k.sub, {
      x: x + 0.34, y: y + 1.32, w: cw - 0.68, h: 0.6,
      fontFace: F.jp, fontSize: 10.5, color: k.dark ? C.cvBody : C.muted, margin: 0, valign: "top", lineSpacingMultiple: 1.4,
    });
  });

  const byr = TOP + 2 * chh + 0.42;
  card(s, M, byr, CW, 0.8);
  icon(s, "LuBadgeJapaneseYen", "green", M + 0.32, byr + 0.21, 0.38);
  s.addText(
    [
      { text: "今回の特別条件：", options: { fontFace: F.jp, fontSize: 12.5, bold: true, color: C.ink } },
      { text: "セルフカフェ売上の25％ ＋ 月額5万円", options: { fontFace: F.jp, fontSize: 14, bold: true, color: C.gold } },
      { text: " を毎月お支払いします。", options: { fontFace: F.jp, fontSize: 12.5, color: C.ink } },
    ],
    { x: M + 0.9, y: byr, w: CW - 1.2, h: 0.8, margin: 0, valign: "middle" }
  );

  note(s, byr + 0.9, "※ お受取額は本資料の想定販売杯数に基づく試算です（税抜）。売上を保証するものではありません。");
}

/* ===================================================== p3 セルフカフェとは */
{
  const s = pres.addSlide();
  shell(s, "セルフカフェとは", "スタッフのいない、ドリンク1杯のワークスペース。",
    "Wi-Fi・電源完備。1杯のご購入で時間制限なくご利用いただける無人カフェです。");

  const lw = 6.6;
  const stats = [
    ["利用料", "420", "円〜", "1杯で滞在時間の制限なし"],
    ["会員登録", "不要", "", "登録手続き・月会費なし"],
    ["運営人員", "0", "人", "無人運営。開閉店も自動"],
    ["日常業務", "15", "分／日", "清掃・原料補充のみ"],
  ];
  const cw = (lw - 0.3) / 2, chh = 1.52;
  stats.forEach((st, i) => {
    const x = M + (i % 2) * (cw + 0.3);
    const y = TOP + Math.floor(i / 2) * (chh + 0.24);
    card(s, x, y, cw, chh);
    s.addText(st[0], {
      x: x + 0.28, y: y + 0.16, w: cw - 0.56, h: 0.26,
      fontFace: F.jp, fontSize: 10.5, bold: true, color: C.muted, margin: 0, valign: "middle",
    });
    s.addText(
      [
        { text: st[1], options: { fontFace: numFace(st[1]), fontSize: 30, bold: true, color: C.green } },
        { text: st[2] ? " " + st[2] : "", options: { fontFace: F.jp, fontSize: 14, bold: true, color: C.green } },
      ],
      { x: x + 0.28, y: y + 0.44, w: cw - 0.56, h: 0.58, margin: 0, valign: "middle" }
    );
    s.addText(st[3], {
      x: x + 0.28, y: y + 1.06, w: cw - 0.56, h: 0.3,
      fontFace: F.jp, fontSize: 9.5, color: C.muted, margin: 0, valign: "middle",
    });
  });

  const fy = TOP + 2 * (chh + 0.24) + 0.06;
  const feats = [["LuWifi", "高速Wi-Fi"], ["LuPlug", "電源完備"], ["LuClock", "席時間無制限"], ["LuCalendarCheck", "年中無休"]];
  const fw = (lw - 3 * 0.2) / 4;
  feats.forEach(([ic, label], i) => {
    const x = M + i * (fw + 0.2);
    tintCard(s, x, fy, fw, 1.0);
    icon(s, ic, "green", x + fw / 2 - 0.19, fy + 0.18, 0.38);
    s.addText(label, {
      x: x + 0.04, y: fy + 0.62, w: fw - 0.08, h: 0.28,
      fontFace: F.jp, fontSize: 10, bold: true, color: C.green, align: "center", margin: 0, valign: "middle",
    });
  });

  const rx = M + lw + 0.42, rw = R - rx;
  photoSlot(s, rx, TOP, rw, 3.3, null, { img: "store-sasashima.jpg" });
  tintCard(s, rx, TOP + 3.48, rw, 1.04);
  s.addText("ささしまライブ店（名古屋）の店内。商業施設内でも同じ仕様で展開しています。", {
    x: rx + 0.26, y: TOP + 3.48, w: rw - 0.52, h: 1.04,
    fontFace: F.jp, fontSize: 10, color: C.ink, margin: 0, valign: "middle", lineSpacingMultiple: 1.35,
  });

  note(s, 6.78, "※ ドリンクは1杯420円〜（税抜）。サブスクプランをご利用の場合は1杯あたり147円〜となります。");
}

/* ===================================================== p4 立地適性 */
{
  const s = pres.addSlide();
  shell(s, "立地適性", "映画館のとなり。待ち時間が、そのまま需要になる。",
    "両店とも同一フロアに映画館があり、上映前後の待ち時間需要を取り込める立地です。");

  const cw = (CW - 0.42) / 2;
  const stores = [
    { name: "秋田店", floor: "2F", img: "map-akita-floor.png", cine: "TOHOシネマズ", h: 2.3,
      facts: [["面積・席数", "62.97坪／73席"], ["月間来客数", "3,760人"], ["営業時間", "10:00〜21:00"]] },
    { name: "土浦店", floor: "3F", img: "map-tsuchiura-floor.png", cine: "シネマサンシャイン土浦", h: 2.3,
      facts: [["面積・席数", "30坪／59席"], ["月間来客数", "4,027人"], ["営業時間", "10:00〜21:00"]] },
  ];
  stores.forEach((st, i) => {
    const x = M + i * (cw + 0.42);
    card(s, x, TOP, cw, 4.16);
    s.addText(
      [
        { text: st.name, options: { fontFace: F.jp, fontSize: 17, bold: true, color: C.ink } },
        { text: "   " + st.floor + " ／ " + st.cine + " と同一フロア", options: { fontFace: F.jp, fontSize: 9.5, color: C.muted } },
      ],
      { x: x + 0.3, y: TOP + 0.18, w: cw - 0.6, h: 0.36, margin: 0, valign: "middle" }
    );
    photoSlot(s, x + 0.3, TOP + 0.58, cw - 0.6, st.h, null, { img: st.img, fit: "contain" });
    const fy = TOP + 0.58 + st.h + 0.1;
    st.facts.forEach((f, j) => {
      const y = fy + j * 0.4;
      s.addText(f[0], {
        x: x + 0.3, y, w: 1.9, h: 0.32,
        fontFace: F.jp, fontSize: 10, color: C.muted, margin: 0, valign: "middle",
      });
      s.addText(f[1], {
        x: x + 2.2, y, w: cw - 2.5, h: 0.32,
        fontFace: F.jp, fontSize: 12, bold: true, color: C.ink, margin: 0, valign: "middle",
      });
      if (j < st.facts.length - 1) s.addShape("rect", { x: x + 0.3, y: y + 0.36, w: cw - 0.6, h: 0.011, fill: { color: C.warmLine } });
    });
  });

  panel(s, M, TOP + 4.34, CW, 0.72);
  icon(s, "LuHourglass", "pale", M + 0.32, TOP + 4.52, 0.36);
  s.addText("上映前後の待ち時間や、書店で本を選んだあとの休憩など「少しだけ座りたい」需要が館内で自然に発生します。1杯で時間制限がないため、その受け皿として機能します。", {
    x: M + 0.88, y: TOP + 4.34, w: CW - 1.18, h: 0.72,
    fontFace: F.jp, fontSize: 10, color: C.white, margin: 0, valign: "middle", lineSpacingMultiple: 1.3,
  });

  note(s, 6.92, "出典：面積・席数・月間来客数は貴社ご提供資料（2025年平均）。フロアマップは各館の館内図より。");
}

/* ===================================================== p5 参考モデル（書店併設型の実績） */
{
  const s = pres.addSlide();
  shell(s, "参考モデル", "書店併設型は、すでに2店舗で稼働しています。",
    "同じ「書店の中のセルフカフェ」として運営している2店舗の実績です。");

  const cw = (CW - 2 * 0.3) / 3;
  const refs = [
    { name: "印西牧の原店", sub: "千葉県印西市／書店併設", img: "store-inzai-real.jpg", cups: "40.6", dark: false },
    { name: "新守山店", sub: "愛知県名古屋市／書店併設", img: "store-shinmoriyama.jpg", cups: "45.2", dark: false },
    { name: "ささしまライブ店", sub: "愛知県名古屋市／商業施設内", img: "store-sasashima.jpg", cups: "122.7", dark: true },
  ];
  refs.forEach((r, i) => {
    const x = M + i * (cw + 0.3);
    if (r.dark) panel(s, x, TOP, cw, 4.2); else card(s, x, TOP, cw, 4.2);
    photoSlot(s, x + 0.24, TOP + 0.24, cw - 0.48, 1.72, null, { img: r.img });
    s.addText(r.name, {
      x: x + 0.24, y: TOP + 2.02, w: cw - 0.48, h: 0.34,
      fontFace: F.jp, fontSize: 15, bold: true, color: r.dark ? C.white : C.ink, margin: 0, valign: "middle",
    });
    s.addText(r.sub, {
      x: x + 0.24, y: TOP + 2.36, w: cw - 0.48, h: 0.26,
      fontFace: F.jp, fontSize: 9.5, color: r.dark ? C.cvBody : C.muted, margin: 0, valign: "middle",
    });
    s.addShape("rect", { x: x + 0.24, y: TOP + 2.74, w: cw - 0.48, h: 0.011, fill: { color: r.dark ? "2A7A4E" : C.warmLine } });
    s.addText("1日あたりの販売杯数", {
      x: x + 0.24, y: TOP + 2.88, w: cw - 0.48, h: 0.26,
      fontFace: F.jp, fontSize: 9.5, color: r.dark ? C.cvBody : C.muted, margin: 0, valign: "middle",
    });
    s.addText(
      [
        { text: r.cups, options: { fontFace: F.num, fontSize: 30, bold: true, color: r.dark ? "F0C05A" : C.green } },
        { text: " 杯", options: { fontFace: F.jp, fontSize: 14, bold: true, color: r.dark ? "F0C05A" : C.green } },
      ],
      { x: x + 0.24, y: TOP + 3.14, w: cw - 0.48, h: 0.66, margin: 0, valign: "middle" }
    );
    s.addText(r.dark ? "商業施設内の大型店。上限の目安" : "書店併設の標準的な水準", {
      x: x + 0.24, y: TOP + 3.82, w: cw - 0.48, h: 0.26,
      fontFace: F.jp, fontSize: 9, color: r.dark ? C.cvPale : C.muted, margin: 0, valign: "middle",
    });
  });

  tintCard(s, M, TOP + 4.38, CW, 0.8);
  icon(s, "LuChartColumn", "green", M + 0.32, TOP + 4.59, 0.38);
  s.addText(
    [
      { text: "書店併設2店の平均は ", options: { fontFace: F.jp, fontSize: 11.5, color: C.ink } },
      { text: "1日42.9杯", options: { fontFace: F.jp, fontSize: 14, bold: true, color: C.green } },
      { text: "。商業施設内のささしまライブ店は ", options: { fontFace: F.jp, fontSize: 11.5, color: C.ink } },
      { text: "1日122.7杯", options: { fontFace: F.jp, fontSize: 14, bold: true, color: C.green } },
      { text: " で稼働しています。", options: { fontFace: F.jp, fontSize: 11.5, color: C.ink } },
    ],
    { x: M + 0.9, y: TOP + 4.38, w: CW - 1.2, h: 0.8, margin: 0, valign: "middle" }
  );

  note(s, 6.94, "出典：社内管理台帳（2026年5月〜8月の月間実杯数の平均）。");
}

/* ===================================================== p6 想定販売杯数の考え方 */
{
  const s = pres.addSlide();
  shell(s, "前提の考え方", "実績を基準に、控えめに置いた想定です。",
    "商業施設内・映画館隣接という条件を踏まえ、実績平均の1.5倍程度を想定しました。");

  // 杯数の比較バー
  const bx = M, bw = 8.1, bh = 3.5;
  card(s, bx, TOP, bw, bh);
  s.addText("1日あたりの販売杯数の比較", {
    x: bx + 0.3, y: TOP + 0.2, w: bw - 0.6, h: 0.28,
    fontFace: F.jp, fontSize: 11, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
  const bars = [
    { label: "印西牧の原店（実績）", v: 40.6, c: "9CC3AA", tag: "実績" },
    { label: "新守山店（実績）", v: 45.2, c: "9CC3AA", tag: "実績" },
    { label: "土浦店（想定）", v: 60, c: C.gold, tag: "想定" },
    { label: "秋田店（想定）", v: 70, c: C.gold, tag: "想定" },
    { label: "ささしまライブ店（実績）", v: 122.7, c: "5FA57C", tag: "実績" },
  ];
  const maxV = 130, lx = bx + 2.7, lw2 = bw - 3.5;
  bars.forEach((b, i) => {
    const y = TOP + 0.66 + i * 0.55;
    s.addText(b.label, {
      x: bx + 0.3, y, w: 2.36, h: 0.34,
      fontFace: F.jp, fontSize: 9.5, bold: b.tag === "想定", color: b.tag === "想定" ? C.ink : C.muted, margin: 0, valign: "middle",
    });
    const w2 = (b.v / maxV) * lw2;
    s.addShape("roundRect", {
      x: lx, y: y + 0.05, w: w2, h: 0.24, rectRadius: 0.05,
      fill: { color: b.c }, line: { type: "none" },
    });
    s.addText(b.v + " 杯", {
      x: lx + w2 + 0.1, y, w: 1.0, h: 0.34,
      fontFace: F.jp, fontSize: 10, bold: true, color: b.tag === "想定" ? C.gold : C.muted, margin: 0, valign: "middle",
    });
  });
  s.addText("※ 帯の色：金＝今回の想定／緑＝既存店の実績", {
    x: bx + 0.3, y: TOP + 3.14, w: bw - 0.6, h: 0.26,
    fontFace: F.jp, fontSize: 8.5, color: C.footer, margin: 0, valign: "middle",
  });

  // 右：根拠
  const rx = M + bw + 0.42, rw = R - rx;
  const pts = [
    ["LuChartColumn", "基準", "書店併設2店の実績平均 1日42.9杯"],
    ["LuTrendingUp", "係数", "館内の通行量・映画館隣接を踏まえ 約1.5倍"],
    ["LuUsers", "席数で配分", "秋田73席＝70杯／土浦59席＝60杯"],
    ["LuShieldCheck", "妥当性", "ささしまライブ店の実績の約半分の水準"],
  ];
  pts.forEach((p2, i) => {
    const y = TOP + i * 0.9;
    tintCard(s, rx, y, rw, 0.8);
    icon(s, p2[0], "green", rx + 0.24, y + 0.22, 0.34);
    s.addText(p2[1], {
      x: rx + 0.7, y: y + 0.1, w: rw - 0.94, h: 0.26,
      fontFace: F.jp, fontSize: 10, bold: true, color: C.green, margin: 0, valign: "middle",
    });
    s.addText(p2[2], {
      x: rx + 0.7, y: y + 0.36, w: rw - 0.94, h: 0.36,
      fontFace: F.jp, fontSize: 9.5, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.2,
    });
  });

  panel(s, M, TOP + 3.72, CW, 0.84);
  s.addText(
    [
      { text: "来客に対する購入率でみると、秋田店は約56％・土浦店は約45％。", options: { fontFace: F.jp, fontSize: 11.5, bold: true, color: C.white } },
      { text: "  ご来店のおよそ半数が1杯ご購入いただく水準です。", options: { fontFace: F.jp, fontSize: 11, color: C.cvBody } },
    ],
    { x: M + 0.34, y: TOP + 3.72, w: CW - 0.68, h: 0.84, margin: 0, valign: "middle" }
  );

  note(s, 6.86, "※ 想定値であり、売上を保証するものではありません。実際の販売杯数は立地・季節・館内の動線により変動します。");
}

/* ===================================================== p7 / p8 収益シミュレーション */
[
  {
    name: "秋田店", spec: "62.97坪／73席", cups: 70, guests: "3,760",
    sales: "882,000", fee25: "220,500", total: "270,500", year: "3,246,000", rate: "56",
  },
  {
    name: "土浦店", spec: "30坪／59席", cups: 60, guests: "4,027",
    sales: "756,000", fee25: "189,000", total: "239,000", year: "2,868,000", rate: "45",
  },
].forEach((sim) => {
  const s = pres.addSlide();
  shell(s, "収益シミュレーション", `${sim.name} のお受取イメージ`,
    `${sim.spec}／想定 1日${sim.cups}杯・1杯420円（税抜）で試算しています。`);

  const lw = 8.1;
  const cell = (t2) => ({
    text: t2,
    options: {
      fill: { color: C.white }, color: C.body, fontSize: 11.5, fontFace: F.num,
      align: "right", valign: "middle", margin: [0.05, 0.16, 0.05, 0.16],
    },
  });
  const hi = (t2, size) => ({
    text: t2,
    options: {
      fill: { color: C.tint }, color: C.green, bold: true, fontSize: size, fontFace: F.num,
      align: "right", valign: "middle", margin: [0.05, 0.16, 0.05, 0.16],
    },
  });
  const hl = (t2) => ({
    text: t2,
    options: {
      fill: { color: C.tint }, color: C.ink, bold: true, fontSize: 11.5, fontFace: F.jp,
      valign: "middle", margin: [0.05, 0.16, 0.05, 0.16],
    },
  });
  s.addTable(
    [
      [th("項目（月次・円／税抜）"), th("金額", { align: "right" })],
      [tl("1日あたりの販売杯数"), cell(sim.cups + "杯")],
      [tl("月間販売杯数（30日）"), cell((sim.cups * 30).toLocaleString() + "杯")],
      [hl("セルフカフェ売上（1杯420円）"), hi(sim.sales, 13)],
      [tl("　うち 売上の25％"), cell(sim.fee25)],
      [tl("　＋ 月額固定"), cell("50,000")],
      [
        { text: "貴社お受取額（月）", options: { fill: { color: C.goldTint }, color: C.gold, bold: true, fontSize: 12.5, fontFace: F.jp, valign: "middle", margin: [0.05, 0.16, 0.05, 0.16] } },
        { text: sim.total, options: { fill: { color: C.goldTint }, color: C.gold, bold: true, fontSize: 17, fontFace: F.num, align: "right", valign: "middle", margin: [0.05, 0.16, 0.05, 0.16] } },
      ],
      [tl("年間換算"), cell(sim.year)],
    ],
    {
      x: M, y: TOP, w: lw, colW: [5.5, 2.6],
      rowH: [0.4, 0.44, 0.44, 0.54, 0.44, 0.44, 0.66, 0.44],
      border: { type: "solid", color: C.warmLine, pt: 0.75 },
      autoPage: false,
    }
  );

  // 右：お受取額パネル
  const rx = M + lw + 0.42, rw = R - rx;
  panel(s, rx, TOP, rw, 2.1);
  eyebrowIn(s, rx + 0.28, TOP + 0.22, 2.6, "YOUR REVENUE", C.cvPale);
  s.addText("貴社お受取額（想定）", {
    x: rx + 0.28, y: TOP + 0.48, w: rw - 0.56, h: 0.28,
    fontFace: F.jp, fontSize: 11, color: C.cvBody, margin: 0, valign: "middle",
  });
  s.addText(
    [
      { text: "月 ", options: { fontFace: F.jp, fontSize: 13, color: C.white } },
      { text: sim.total === "270,500" ? "27.0" : "23.9", options: { fontFace: F.num, fontSize: 32, bold: true, color: "F0C05A" } },
      { text: " 万円", options: { fontFace: F.jp, fontSize: 14, bold: true, color: "F0C05A" } },
    ],
    { x: rx + 0.28, y: TOP + 0.8, w: rw - 0.56, h: 0.7, margin: 0, valign: "middle" }
  );
  s.addShape("rect", { x: rx + 0.28, y: TOP + 1.58, w: rw - 0.56, h: 0.011, fill: { color: "2A7A4E" } });
  s.addText("年 " + (sim.year === "3,246,000" ? "324.6" : "286.8") + " 万円", {
    x: rx + 0.28, y: TOP + 1.7, w: rw - 0.56, h: 0.36,
    fontFace: F.jp, fontSize: 14, bold: true, color: C.white, margin: 0, valign: "middle",
  });

  // 右下：前提
  tintCard(s, rx, TOP + 2.3, rw, 2.24);
  s.addText("この試算の前提", {
    x: rx + 0.28, y: TOP + 2.48, w: rw - 0.56, h: 0.26,
    fontFace: F.jp, fontSize: 11, bold: true, color: C.green, margin: 0, valign: "middle",
  });
  [
    "月間来客数 " + sim.guests + "人",
    "うち約" + sim.rate + "％が1杯ご購入",
    "書店併設2店の実績の約1.5倍",
    "原価・水道光熱費は本部負担",
  ].forEach((t2, i) => {
    const y = TOP + 2.8 + i * 0.4;
    icon(s, "LuCheck", "green", rx + 0.3, y + 0.05, 0.18);
    s.addText(t2, {
      x: rx + 0.6, y, w: rw - 0.9, h: 0.32,
      fontFace: F.jp, fontSize: 9.5, color: C.ink, margin: 0, valign: "middle",
    });
  });

  note(s, 6.78, "※ 想定値であり、売上・お受取額を保証するものではありません。金額はすべて税抜。実際の販売杯数により変動します。");
});

/* ===================================================== p9 運営と導入条件 */
{
  const s = pres.addSlide();
  shell(s, "運営と導入条件", "お願いするのは1日15分の清掃と補充だけ。",
    "運営主体はセルフカフェ本部です。カフェ専任のスタッフを置く必要はありません。");

  const lw = 6.5;
  const tasks = [
    ["① 店内清掃", "テーブル・床・トイレの清掃、ゴミ回収。営業前後の都合の良い時間で結構です。"],
    ["② 原料・備品の補充", "コーヒー豆・カップ等の補充。在庫は本部から定期配送します。"],
    ["③ 発注・棚卸（月1〜2回）", "定期配送に合わせた発注と在庫の棚卸のみです。"],
  ];
  tasks.forEach((tk, i) => {
    const y = TOP + i * 0.94;
    tintCard(s, M, y, lw, 0.82);
    s.addText(tk[0], {
      x: M + 0.3, y: y + 0.1, w: lw - 0.6, h: 0.28,
      fontFace: F.jp, fontSize: 12, bold: true, color: C.green, margin: 0, valign: "middle",
    });
    s.addText(tk[1], {
      x: M + 0.3, y: y + 0.4, w: lw - 0.6, h: 0.34,
      fontFace: F.jp, fontSize: 9.5, color: C.body, margin: 0, valign: "middle",
    });
  });
  panel(s, M, TOP + 2.9, lw, 0.9);
  icon(s, "LuLifeBuoy", "pale", M + 0.3, TOP + 3.16, 0.36);
  s.addText("清掃はパートナーへの委託も可能です（月3万円程度）。お客様からのお問い合わせは本部が対応します。", {
    x: M + 0.82, y: TOP + 2.9, w: lw - 1.12, h: 0.9,
    fontFace: F.jp, fontSize: 10, bold: true, color: C.white, margin: 0, valign: "middle", lineSpacingMultiple: 1.3,
  });

  const rx = M + lw + 0.42, rw = R - rx;
  const conds = [
    ["初期費用", "30万〜50万円程度（機器設置費のみ）"],
    ["加盟金・研修費", "なし"],
    ["原料・水道光熱費", "本部負担"],
    ["お支払い", "セルフカフェ売上の25％ ＋ 月額5万円"],
    ["契約期間", "最低3年〜（本部審査あり）"],
  ];
  card(s, rx, TOP, rw, 3.8);
  eyebrowIn(s, rx + 0.3, TOP + 0.24, 2.4, "CONDITIONS");
  conds.forEach((c2, i) => {
    const y = TOP + 0.58 + i * 0.62;
    s.addText(c2[0], {
      x: rx + 0.3, y, w: rw - 0.6, h: 0.24,
      fontFace: F.jp, fontSize: 9.5, color: C.muted, margin: 0, valign: "middle",
    });
    s.addText(c2[1], {
      x: rx + 0.3, y: y + 0.22, w: rw - 0.6, h: 0.3,
      fontFace: F.jp, fontSize: 11.5, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    if (i < conds.length - 1) s.addShape("rect", { x: rx + 0.3, y: y + 0.54, w: rw - 0.6, h: 0.011, fill: { color: C.warmLine } });
  });

  photoSlot(s, M, TOP + 3.96, CW, 0.96, null, { img: "p24-install.jpg" });
  note(s, 6.94, "※ 既存什器の再利用可否・工事範囲は現地確認のうえ個別にご相談させてください。金額は税抜・目安です。");
}

/* ===================================================== p10 導入までの流れ／お問い合わせ */
{
  const s = pres.addSlide();
  shell(s, "今後の進め方", "現地確認から、最短2〜3ヶ月で切り替えできます。",
    "まずは2店舗の現地確認と、条件のすり合わせからお願いできればと存じます。");

  const steps = [
    ["LuMessagesSquare", "STEP 01", "ご検討・ご相談", "本ご提案へのご質問・条件のすり合わせ。"],
    ["LuFileSearch", "STEP 02", "現地確認", "区画・電源・給排水・動線を確認し、レイアウトを設計。"],
    ["LuPenLine", "STEP 03", "契約締結", "条件確定のうえ業務委託契約を締結。"],
    ["LuHammer", "STEP 04", "設置工事", "機器設置・什器調整。既存什器は可能な範囲で再利用。"],
    ["LuStore", "STEP 05", "運営開始", "オープン後の集客・お問い合わせ対応は本部が担当。"],
  ];
  const w = (CW - 4 * 0.22) / 5;
  steps.forEach((st, i) => {
    const x = M + i * (w + 0.22);
    card(s, x, TOP, w, 2.72);
    eyebrowIn(s, x + 0.24, y0(TOP), 1.6, st[1]);
    icon(s, st[0], "green", x + w - 0.72, TOP + 0.2, 0.36);
    s.addText(st[2], {
      x: x + 0.24, y: TOP + 0.62, w: w - 0.48, h: 0.5,
      fontFace: F.jp, fontSize: 13, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addShape("rect", { x: x + 0.24, y: TOP + 1.18, w: w - 0.48, h: 0.011, fill: { color: C.warmLine } });
    s.addText(st[3], {
      x: x + 0.24, y: TOP + 1.32, w: w - 0.48, h: 1.2,
      fontFace: F.jp, fontSize: 9.5, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.4,
    });
  });
  function y0(t) { return t + 0.24; }

  panel(s, M, TOP + 2.98, CW, 1.9);
  s.addImage({ path: LOGO_W, x: M + 0.44, y: TOP + 3.26, w: 2.29, h: 0.326 });
  s.addText("ご不明な点は、お気軽にお問い合わせください。", {
    x: M + 0.44, y: TOP + 3.74, w: 6.4, h: 0.34,
    fontFace: F.jp, fontSize: 13, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  s.addText("現地確認は日程を合わせて伺います。収益試算の条件変更もその場でご相談いただけます。", {
    x: M + 0.44, y: TOP + 4.08, w: 6.4, h: 0.5,
    fontFace: F.jp, fontSize: 10, color: C.cvBody, margin: 0, valign: "top", lineSpacingMultiple: 1.35,
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

const out = path.join(__dirname, "selfcafe-miraiya-deck.pptx");
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
