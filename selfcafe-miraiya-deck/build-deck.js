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
    { x: 4.5, w: 3.2, v: "31.8", u: "万円／月", l: "秋田店 想定お受取額" },
    { x: 8.1, w: 3.2, v: "28.6", u: "万円／月", l: "土浦店 想定お受取額" },
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
    { dark: true, label: "秋田店 お受取額（想定）", v: "31.8", u: "万円／月", sub: "年間 381.3万円\n売上の25％＋月5万円" },
    { dark: false, label: "初期費用", v: "30〜50", u: "万円", sub: "機器設置費のみ。加盟金・\n研修費はいただきません。" },
    { dark: true, label: "土浦店 お受取額（想定）", v: "28.6", u: "万円／月", sub: "年間 343.5万円\n売上の25％＋月5万円" },
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

/* ===================================================== スキーム比較（業務委託型とFC） */
{
  const s = pres.addSlide();
  shell(s, "スキーム", "場所をご提供いただくだけ。運営主体は本部です。",
    "通常のFC加盟とは異なり、貴社が独立オーナーとして運営する必要はありません。");

  const colW = [2.1, 4.5, 5.289];
  const mine2 = (t2) => ({
    text: t2,
    options: {
      fill: { color: C.tint }, color: C.ink, bold: true, fontSize: 11, fontFace: F.jp,
      valign: "middle", margin: [0.06, 0.16, 0.06, 0.16],
    },
  });
  const oth2 = (t2) => ({
    text: t2,
    options: {
      fill: { color: C.white }, color: C.grayText, fontSize: 10.5, fontFace: F.jp,
      valign: "middle", margin: [0.06, 0.16, 0.06, 0.16],
    },
  });
  s.addShape("roundRect", {
    x: M + colW[0] + colW[1], y: TOP + 0.08, w: colW[2], h: 0.3, rectRadius: 0.05,
    fill: { color: C.gold }, line: { type: "none" },
  });
  s.addText("貴社にご提案するのはこちら", {
    x: M + colW[0] + colW[1], y: TOP + 0.08, w: colW[2], h: 0.3,
    fontFace: F.jp, fontSize: 9.5, bold: true, color: C.white, align: "center", margin: 0, valign: "middle",
  });
  s.addTable(
    [
      [th("項目", { fill: C.grayBand, color: C.ink }), th("通常のFC加盟", { fill: "8E8B84" }), th("業務委託型")],
      [tl("運営主体"), oth2("加盟者様が独立オーナーとして運営"), mine2("本部が運営を主導。貴社は場所のご提供のみ")],
      [tl("初期費用"), oth2("加盟金100万円＋工事費用等\n（開業資金の目安：居抜き750万円〜／スケルトン1,050万円〜）"), mine2("30万〜50万円程度（機器設置費のみ）")],
      [tl("出店スペース"), oth2("独立した店舗区画（20坪〜）が必要"), mine2("既存スペースの一角に間借り設置が可能")],
      [tl("日常業務"), oth2("加盟者様が運営全般を担当"), mine2("清掃・補充のみ、1日15分程度（既存スタッフで対応可）")],
      [tl("売上金の流れ"), oth2("本部が売上金を管理し、ロイヤリティ等を控除した残額を加盟者様へ振込"), mine2("本部が売上金を管理し、売上の25％を業務委託料として貴社へお支払い")],
      [tl("契約・撤退"), oth2("独立出店が前提の長期契約"), mine2("契約期間の縛りあり（最低3年〜）／本部審査あり。原状回復の負担も小さい")],
    ],
    {
      x: M, y: TOP + 0.5, w: CW, colW,
      rowH: [0.4, 0.5, 0.72, 0.5, 0.56, 0.72, 0.66],
      border: { type: "solid", color: C.warmLine, pt: 0.75 },
      autoPage: false,
    }
  );
  note(s, TOP + 4.84, "※ 金額はすべて税抜・目安です。物件条件により変動します。今回は業務委託型でのご提案です。");
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
  photoSlot(s, rx, TOP, rw, 3.3, null, { img: "sasashima-interior.jpg" });
  tintCard(s, rx, TOP + 3.48, rw, 1.04);
  s.addText("ささしまライブ店（名古屋）。商業施設の1階路面で、朝8時から22時まで無人運営しています。", {
    x: rx + 0.26, y: TOP + 3.48, w: rw - 0.52, h: 1.04,
    fontFace: F.jp, fontSize: 10, color: C.ink, margin: 0, valign: "middle", lineSpacingMultiple: 1.35,
  });

  note(s, 6.78, "※ ドリンクは1杯420円〜（税抜）。サブスクプランをご利用の場合は1杯あたり147円〜となります。");
}

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
  const rx = M + lw + 0.42, rw = R - rx;
  const conds = [
    ["初期費用", "30万〜50万円程度（機器設置費のみ）"],
    ["加盟金・研修費", "なし"],
    ["原料・機器・メンテナンス", "本部負担"],
    ["水道光熱費・通信費・消耗品", "貴社ご負担（実費）"],
    ["お支払い", "セルフカフェ売上の25％ ＋ 月額5万円"],
    ["契約期間", "最低3年〜（本部審査あり）"],
  ];
  card(s, rx, TOP, rw, 3.72);
  eyebrowIn(s, rx + 0.3, TOP + 0.22, 2.4, "CONDITIONS");
  conds.forEach((c2, i) => {
    const y = TOP + 0.52 + i * 0.55;
    s.addText(c2[0], {
      x: rx + 0.3, y, w: rw - 0.6, h: 0.22,
      fontFace: F.jp, fontSize: 9.5, color: C.muted, margin: 0, valign: "middle",
    });
    s.addText(c2[1], {
      x: rx + 0.3, y: y + 0.2, w: rw - 0.6, h: 0.28,
      fontFace: F.jp, fontSize: 11.5, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    if (i < conds.length - 1) s.addShape("rect", { x: rx + 0.3, y: y + 0.49, w: rw - 0.6, h: 0.011, fill: { color: C.warmLine } });
  });

  panel(s, rx, TOP + 3.86, rw, 0.72);
  icon(s, "LuLifeBuoy", "pale", rx + 0.28, TOP + 4.04, 0.34);
  s.addText("清掃はパートナーへの委託も可能です（月3万円程度）。お客様からのお問い合わせは本部が対応します。", {
    x: rx + 0.74, y: TOP + 3.86, w: rw - 1.04, h: 0.72,
    fontFace: F.jp, fontSize: 10, bold: true, color: C.white, margin: 0, valign: "middle", lineSpacingMultiple: 1.25,
  });

  // 写真は元比率に近い16:9で左右に配置（潰れないサイズ）
  const phw = 3.1, phh = phw * 9 / 16;
  photoSlot(s, M, TOP + 2.86, phw, phh, null, { img: "inzai-counter.jpg" });
  photoSlot(s, M + phw + 0.3, TOP + 2.86, phw, phh, null, { img: "sasashima-interior.jpg" });
  note(s, 6.62, "※ 水道光熱費・通信費・消耗品は貴社のご負担（実費）となります。既存什器の再利用可否・工事範囲は現地確認のうえ個別にご相談させてください。金額は税抜・目安です。");
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
    { name: "印西牧の原店", sub: "千葉県印西市／書店併設", img: "inzai-hall.jpg", cups: "40.6", dark: false },
    { name: "新守山店", sub: "愛知県名古屋市／書店併設", img: "shinmoriyama-hall.jpg", cups: "45.2", dark: false },
    { name: "ささしまライブ店", sub: "愛知県名古屋市／商業施設内", img: "sasashima-exterior.jpg", cups: "122.7", dark: true },
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

/* ===================================================== 書店併設の実例（印西牧の原店） */
{
  const s = pres.addSlide();
  shell(s, "導入イメージ", "書店の中に、こう入ります。",
    "印西牧の原店の実際の様子です。書架のとなりに席を並べ、書店の動線をそのまま活かしています。");

  const bigW = 7.5, bigH = 3.9;
  photoSlot(s, M, TOP, bigW, bigH, null, { img: "inzai-hall.jpg" });

  const rx = M + bigW + 0.3, rw = R - rx, sh = (bigH - 0.3) / 2;
  photoSlot(s, rx, TOP, rw, sh, null, { img: "inzai-books.jpg" });
  photoSlot(s, rx, TOP + sh + 0.3, rw, sh, null, { img: "inzai-counter.jpg" });

  const pts = [
    ["LuRuler", "書架の間・壁面沿いに設置", "独立区画は不要。既存の什器配置を大きく変えずに席をつくれます。"],
    ["LuPlug", "各席に電源とWi-Fi", "長時間の滞在に耐える設えで、書店の滞在時間もあわせて伸びます。"],
    ["LuUsers", "レジ・接客は発生しない", "ドリンクはセルフのマシン。書店スタッフの手は一切かかりません。"],
  ];
  const by = TOP + bigH + 0.24, pw = (CW - 2 * 0.24) / 3;
  pts.forEach((p2, i) => {
    const x = M + i * (pw + 0.24);
    tintCard(s, x, by, pw, 0.94);
    icon(s, p2[0], "green", x + 0.24, by + 0.28, 0.34);
    s.addText(p2[1], {
      x: x + 0.68, y: by + 0.12, w: pw - 0.92, h: 0.28,
      fontFace: F.jp, fontSize: 10.5, bold: true, color: C.green, margin: 0, valign: "middle",
    });
    s.addText(p2[2], {
      x: x + 0.68, y: by + 0.4, w: pw - 0.92, h: 0.44,
      fontFace: F.jp, fontSize: 9, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
    });
  });

  note(s, by + 1.04, "写真：セルフカフェ 印西牧の原店（千葉県印西市／未来屋書店様 併設）。");
}

/* ===================================================== 業態転換イメージ（土浦店・パース） */
{
  const s = pres.addSlide();
  shell(s, "業態転換イメージ", "土浦店のカフェ区画は、こう変わります。",
    "現況のフルサービス型カフェ区画に、セルフカフェを入れた場合の完成イメージです。");

  const lw = 3.2, gap = 0.5;
  const rx = M + lw + gap, rw = R - rx;

  // 列ラベル
  s.addShape("roundRect", {
    x: M, y: TOP, w: lw, h: 0.3, rectRadius: 0.05,
    fill: { color: "8E8B84" }, line: { type: "none" },
  });
  s.addText("現況", {
    x: M, y: TOP, w: lw, h: 0.3,
    fontFace: F.jp, fontSize: 10, bold: true, color: C.white, align: "center", margin: 0, valign: "middle",
  });
  s.addShape("roundRect", {
    x: rx, y: TOP, w: rw, h: 0.3, rectRadius: 0.05,
    fill: { color: C.green }, line: { type: "none" },
  });
  s.addText("業態転換後のイメージ（完成予想パース）", {
    x: rx, y: TOP, w: rw, h: 0.3,
    fontFace: F.jp, fontSize: 10, bold: true, color: C.white, align: "center", margin: 0, valign: "middle",
  });

  // 左：現況 2枚
  const cy = TOP + 0.34, ch = 2.13;
  photoSlot(s, M, cy, lw, ch, null, { img: "tsuchiura-current-exterior.jpg" });
  photoSlot(s, M, cy + ch + 0.13, lw, ch, null, { img: "tsuchiura-current-interior.jpg" });

  // 中央：変換の矢印
  s.addShape("roundRect", {
    x: M + lw + 0.06, y: 4.16, w: 0.38, h: 0.38, rectRadius: 0.19,
    fill: { color: C.gold }, line: { type: "none" },
  });
  s.addText("▶", {
    x: M + lw + 0.06, y: 4.16, w: 0.38, h: 0.38,
    fontFace: F.jp, fontSize: 11, bold: true, color: C.white, align: "center", margin: 0, valign: "middle",
  });

  // 右：パース 外観大1枚＋内観2枚
  const ah = 2.55;
  photoSlot(s, rx, cy, rw, ah, null, { img: "tsuchiura-after-exterior.jpg" });
  const iw = (rw - 0.3) / 2, ih2 = 1.75, iy = cy + ah + 0.13;
  photoSlot(s, rx, iy, iw, ih2, null, { img: "tsuchiura-after-interior-01.jpg" });
  photoSlot(s, rx + iw + 0.3, iy, iw, ih2, null, { img: "tsuchiura-after-interior-02.jpg" });

  note(s, 6.68, "※ パースは完成予想図です。什器の再利用可否・工事範囲は現地確認のうえ確定します。実際の仕上がりは設計内容により異なります。");
}

/* ===================================================== p6 想定販売杯数の考え方 */
{
  const s = pres.addSlide();
  shell(s, "前提の考え方", "既存店の実績から積み上げた想定です。",
    "商業施設内・映画館隣接・館内の通行量を踏まえ、書店併設2店の実績平均の約1.9倍を想定しました。");

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
    { label: "土浦店（想定）", v: 75, c: C.gold, tag: "想定" },
    { label: "秋田店（想定）", v: 85, c: C.gold, tag: "想定" },
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
    ["LuTrendingUp", "係数", "館内の通行量・映画館隣接を踏まえ 約1.9倍"],
    ["LuUsers", "席数で配分", "秋田73席＝85杯／土浦59席＝75杯"],
    ["LuShieldCheck", "妥当性", "ささしまライブ店の実績の6〜7割の水準"],
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
      { text: "来客に対する購入率でみると、秋田店は約68％・土浦店は約56％。", options: { fontFace: F.jp, fontSize: 11.5, bold: true, color: C.white } },
      { text: "  ご来店の半数以上が1杯ご購入いただく水準です。", options: { fontFace: F.jp, fontSize: 11, color: C.cvBody } },
    ],
    { x: M + 0.34, y: TOP + 3.72, w: CW - 0.68, h: 0.84, margin: 0, valign: "middle" }
  );

  note(s, 6.86, "※ 想定値であり、売上を保証するものではありません。実際の販売杯数は立地・季節・館内の動線により変動します。");
}

/* ===================================================== p7 / p8 収益シミュレーション */
[
  {
    name: "秋田店", spec: "62.97坪／73席", cups: 85, guests: "3,760",
    sales: "1,071,000", fee25: "267,750", total: "317,750", year: "3,813,000", rate: "68",
    man: "31.8", yearMan: "381.3",
  },
  {
    name: "土浦店", spec: "30坪／59席", cups: 75, guests: "4,027",
    sales: "945,000", fee25: "236,250", total: "286,250", year: "3,435,000", rate: "56",
    man: "28.6", yearMan: "343.5",
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
      [tl("　＋ 月額固定（貴社のみの特別条件）"), cell("50,000")],
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
      { text: sim.man, options: { fontFace: F.num, fontSize: 32, bold: true, color: "F0C05A" } },
      { text: " 万円", options: { fontFace: F.jp, fontSize: 14, bold: true, color: "F0C05A" } },
    ],
    { x: rx + 0.28, y: TOP + 0.8, w: rw - 0.56, h: 0.7, margin: 0, valign: "middle" }
  );
  s.addShape("rect", { x: rx + 0.28, y: TOP + 1.58, w: rw - 0.56, h: 0.011, fill: { color: "2A7A4E" } });
  s.addText("年 " + sim.yearMan + " 万円", {
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
    "書店併設2店の実績の約1.9倍",
    "原料・機器の費用は本部負担",
  ].forEach((t2, i) => {
    const y = TOP + 2.8 + i * 0.4;
    icon(s, "LuCheck", "green", rx + 0.3, y + 0.05, 0.18);
    s.addText(t2, {
      x: rx + 0.6, y, w: rw - 0.9, h: 0.32,
      fontFace: F.jp, fontSize: 9.5, color: C.ink, margin: 0, valign: "middle",
    });
  });

  note(s, 6.78, "※ 想定値であり、売上・お受取額を保証するものではありません。水道光熱費・通信費・消耗品は貴社のご負担です。金額はすべて税抜。実際の販売杯数により変動します。");
});

/* ===================================================== 坪数別シミュレーション（業務委託型） */
{
  const s = pres.addSlide();
  shell(s, "坪数別シミュレーション", "30坪でも、お受取額はほとんど変わりません。",
    "60坪・40坪・30坪の3パターンで試算しました。販売杯数を決めるのは坪数ではなく、館内の来客数だからです。");

  const cellR = (t2, opts = {}) => ({
    text: t2,
    options: {
      fill: { color: opts.fill || C.white }, color: opts.color || C.body,
      bold: !!opts.bold, fontSize: opts.size || 11.5, fontFace: numFace(t2),
      align: "right", valign: "middle", margin: [0.04, 0.16, 0.04, 0.16],
    },
  });
  const rowLabel = (t2, opts = {}) => ({
    text: t2,
    options: {
      fill: { color: opts.fill || C.white }, color: opts.color || C.ink,
      bold: !!opts.bold, fontSize: opts.size || 11.5, fontFace: F.jp,
      valign: "middle", margin: [0.04, 0.16, 0.04, 0.16],
    },
  });

  const TSUBO = [
    { t: "60坪", cups: 85, sales: "1,071,000", fee: "267,750", tot: "317,750", yr: "3,813,000" },
    { t: "40坪", cups: 80, sales: "1,008,000", fee: "252,000", tot: "302,000", yr: "3,624,000" },
    { t: "30坪", cups: 75, sales: "945,000", fee: "236,250", tot: "286,250", yr: "3,435,000" },
  ];
  const gcell = (t2, size) => ({
    text: t2,
    options: {
      fill: { color: C.goldTint }, color: C.gold, bold: true, fontSize: size, fontFace: numFace(t2),
      align: "right", valign: "middle", margin: [0.04, 0.16, 0.04, 0.16],
    },
  });
  const hcell = (t2) => ({
    text: t2,
    options: {
      fill: { color: C.tint }, color: C.green, bold: true, fontSize: 12, fontFace: numFace(t2),
      align: "right", valign: "middle", margin: [0.04, 0.16, 0.04, 0.16],
    },
  });

  s.addTable(
    [
      [th("項目（月次・円／税抜）"), ...TSUBO.map((t2) => th(t2.t, { align: "right" }))],
      [rowLabel("1日あたりの販売杯数"), ...TSUBO.map((t2) => cellR(t2.cups + "杯"))],
      [rowLabel("月間販売杯数（30日）"), ...TSUBO.map((t2) => cellR((t2.cups * 30).toLocaleString() + "杯"))],
      [rowLabel("セルフカフェ売上（1杯420円）", { fill: C.tint, bold: true }), ...TSUBO.map((t2) => hcell(t2.sales))],
      [rowLabel("　うち 売上の25％"), ...TSUBO.map((t2) => cellR(t2.fee))],
      [rowLabel("　＋ 月額固定（貴社のみの特別条件）"), ...TSUBO.map(() => cellR("50,000"))],
      [rowLabel("貴社お受取額（月）", { fill: C.goldTint, color: C.gold, bold: true, size: 12.5 }), ...TSUBO.map((t2) => gcell(t2.tot, 15))],
      [rowLabel("年間換算"), ...TSUBO.map((t2) => cellR(t2.yr, { bold: true }))],
    ],
    {
      x: M, y: TOP, w: CW, colW: [4.689, 2.4, 2.4, 2.4],
      rowH: [0.38, 0.4, 0.4, 0.48, 0.4, 0.4, 0.58, 0.4],
      border: { type: "solid", color: C.warmLine, pt: 0.75 },
      autoPage: false,
    }
  );

  const py = 5.36;
  panel(s, M, py, CW, 0.72);
  s.addText(
    [
      { text: "この試算からは、30坪でのご提案をお薦めします。", options: { fontFace: F.jp, fontSize: 13, bold: true, color: C.white, breakLine: true } },
      { text: "60坪との月々のお受取額の差は3.2万円。残りの区画は売場など他の用途にご活用いただけます。", options: { fontFace: F.jp, fontSize: 10.5, color: C.cvBody } },
    ],
    { x: M + 0.34, y: py, w: CW - 0.68, h: 0.72, margin: 0, valign: "middle", lineSpacingMultiple: 1.15 }
  );

  const rw2 = (CW - 0.3) / 2;
  [
    ["LuUsers", "杯数は坪数ではなく来客数で決まります", "館内の通行量と映画館の待ち時間需要が販売杯数を左右するため、区画を広げても杯数は大きく増えません。"],
    ["LuLandPlot", "平日ベースなら30坪・約59席で十分です", "土日は高稼働が見込めますが、平日の来店ペースであれば30坪の席数で受けきれると想定しています。"],
  ].forEach((r, i) => {
    const x = M + i * (rw2 + 0.3), y = 6.16;
    tintCard(s, x, y, rw2, 0.62);
    icon(s, r[0], "green", x + 0.22, y + 0.14, 0.32);
    s.addText(r[1], {
      x: x + 0.64, y: y + 0.06, w: rw2 - 0.86, h: 0.24,
      fontFace: F.jp, fontSize: 9.5, bold: true, color: C.green, margin: 0, valign: "middle",
    });
    s.addText(r[2], {
      x: x + 0.64, y: y + 0.29, w: rw2 - 0.86, h: 0.3,
      fontFace: F.jp, fontSize: 8.5, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.15,
    });
  });

  note(s, 6.86, "※ 杯数の置き方：60坪は秋田店（62.97坪）の想定85杯、30坪は土浦店の想定75杯と同水準とし、40坪はその中間の80杯としました。想定値であり保証するものではありません。金額は税抜。");
}

/* ===================================================== ご参考：FCプランの場合 */
{
  const s = pres.addSlide();
  shell(s, "ご参考：FCプラン", "貴社が運営主体となる場合の試算です。",
    "初期投資は必要ですが、売上・利益はすべて貴社のものになります（パートナー制度＝FC加盟）。");

  const RED = "B0483A";
  const FCT = [
    { t: "60坪", cups: 85, sales: "1,071,000", mat: "214,200", cln: "30,000", utl: "70,000", etc: "30,000", pf: "570,800", inv: "1,200万円", pay: "21ヶ月" },
    { t: "40坪", cups: 80, sales: "1,008,000", mat: "201,600", cln: "30,000", utl: "65,000", etc: "25,000", pf: "530,400", inv: "1,000万円", pay: "19ヶ月" },
    { t: "30坪", cups: 75, sales: "945,000", mat: "189,000", cln: "30,000", utl: "60,000", etc: "20,000", pf: "490,000", inv: "900万円", pay: "18ヶ月" },
  ];
  const lw = 8.1;
  const num2 = (v, red) => ({
    text: red ? "▲" + v : v,
    options: {
      fill: { color: C.white }, color: red ? RED : C.body, fontSize: 11, fontFace: numFace(v),
      align: "right", valign: "middle", margin: [0.03, 0.12, 0.03, 0.12],
    },
  });
  const item2 = (t2) => ({
    text: t2,
    options: {
      fill: { color: C.white }, color: C.ink, fontSize: 10.5, fontFace: F.jp,
      valign: "middle", margin: [0.03, 0.12, 0.03, 0.12],
    },
  });
  const pfL = {
    text: "償却前営業利益",
    options: { fill: { color: C.tint }, color: C.ink, bold: true, fontSize: 11.5, fontFace: F.jp, valign: "middle", margin: [0.03, 0.12, 0.03, 0.12] },
  };
  const pfN = (v) => ({
    text: v,
    options: { fill: { color: C.tint }, color: C.green, bold: true, fontSize: 12.5, fontFace: F.num, align: "right", valign: "middle", margin: [0.03, 0.12, 0.03, 0.12] },
  });

  s.addTable(
    [
      [th("項目（月次・円／税抜）"), ...FCT.map((f) => th(f.t, { align: "right" }))],
      [item2("売上高（1杯420円 × 杯数 × 30日）"), ...FCT.map((f) => num2(f.sales))],
      [item2("　想定販売杯数（1日）"), ...FCT.map((f) => num2(f.cups + "杯"))],
      [item2("ドリンク原料（20％）"), ...FCT.map((f) => num2(f.mat, true))],
      [item2("清掃業務委託費"), ...FCT.map((f) => num2(f.cln, true))],
      [item2("水道光熱費（書店営業時間ベース）"), ...FCT.map((f) => num2(f.utl, true))],
      [item2("機械使用料（マシン38,000円×2台・決済端末5,000円×2台）"), ...FCT.map(() => num2("86,000", true))],
      [item2("セキュリティ費"), ...FCT.map(() => num2("20,000", true))],
      [item2("ロイヤリティ（一律・売上連動なし）"), ...FCT.map(() => num2("50,000", true))],
      [item2("雑費"), ...FCT.map((f) => num2(f.etc, true))],
      [item2("家賃（既存区画のため）"), ...FCT.map(() => num2("0"))],
      [pfL, ...FCT.map((f) => pfN(f.pf))],
    ],
    {
      x: M, y: TOP, w: lw, colW: [3.9, 1.4, 1.4, 1.4],
      rowH: [0.38, 0.38, 0.36, 0.36, 0.36, 0.36, 0.46, 0.36, 0.36, 0.36, 0.36, 0.48],
      border: { type: "solid", color: C.warmLine, pt: 0.75 },
      autoPage: false,
    }
  );

  const rx = M + lw + 0.42, rw = R - rx;
  panel(s, rx, TOP, rw, 1.62);
  eyebrowIn(s, rx + 0.28, TOP + 0.22, 2.6, "INVESTMENT", C.cvPale);
  s.addText("初期投資（目安）", {
    x: rx + 0.28, y: TOP + 0.46, w: rw - 0.56, h: 0.26,
    fontFace: F.jp, fontSize: 11, color: C.cvBody, margin: 0, valign: "middle",
  });
  s.addText(
    [
      { text: "900", options: { fontFace: F.num, fontSize: 28, bold: true, color: "F0C05A" } },
      { text: " 〜 ", options: { fontFace: F.jp, fontSize: 15, color: "F0C05A" } },
      { text: "1,200", options: { fontFace: F.num, fontSize: 28, bold: true, color: "F0C05A" } },
      { text: " 万円", options: { fontFace: F.jp, fontSize: 13, bold: true, color: "F0C05A" } },
    ],
    { x: rx + 0.28, y: TOP + 0.74, w: rw - 0.56, h: 0.56, margin: 0, valign: "middle" }
  );
  s.addText("加盟金100万円・工事費・出店準備金・保証金・諸経費を含む総投資予算", {
    x: rx + 0.28, y: TOP + 1.22, w: rw - 0.56, h: 0.34,
    fontFace: F.jp, fontSize: 8.5, color: C.cvBody, margin: 0, valign: "top", lineSpacingMultiple: 1.2,
  });

  tintCard(s, rx, TOP + 1.78, rw, 1.72);
  s.addText("投資回収の目安", {
    x: rx + 0.28, y: TOP + 1.96, w: rw - 0.56, h: 0.26,
    fontFace: F.jp, fontSize: 11, bold: true, color: C.green, margin: 0, valign: "middle",
  });
  FCT.forEach((f, i) => {
    const y = TOP + 2.3 + i * 0.38;
    s.addText(f.t, {
      x: rx + 0.3, y, w: 0.62, h: 0.3,
      fontFace: F.jp, fontSize: 10, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(f.inv, {
      x: rx + 0.85, y, w: 1.05, h: 0.3,
      fontFace: F.jp, fontSize: 9.5, color: C.muted, align: "right", margin: 0, valign: "middle",
    });
    s.addText(f.pay, {
      x: rx + 1.98, y, w: rw - 2.28, h: 0.3,
      fontFace: F.jp, fontSize: 10.5, bold: true, color: C.green, align: "right", margin: 0, valign: "middle",
    });
  });

  card(s, rx, TOP + 3.62, rw, 0.76);
  icon(s, "LuUsers", "green", rx + 0.26, TOP + 3.81, 0.34);
  s.addText("清掃を貴社スタッフで行う場合は、委託費3万円が不要となり、その分利益が増えます。", {
    x: rx + 0.7, y: TOP + 3.62, w: rw - 0.96, h: 0.76,
    fontFace: F.jp, fontSize: 9, color: C.ink, margin: 0, valign: "middle", lineSpacingMultiple: 1.15,
  });

  note(s, 6.62, "※ 償却前営業利益（減価償却前）。人件費は未計上です。家賃は既存区画をご利用いただく前提のため0円で試算しています。水道光熱費は書店営業時間ベース（空調・照明は書店と共用のため24H営業想定より低く置いています）。投資回収は坪数別の総投資予算を月次利益で除した目安。金額はすべて税抜。");
}

/* ===================================================== 2方式の比較 */
{
  const s = pres.addSlide();
  shell(s, "2方式の比較", "同じ条件で、2つの方式を並べました。",
    "同じ販売杯数の想定で業務委託型とFCプランを比較しています。ご検討の材料としてご覧ください。");

  const lw = 7.3;
  card(s, M, TOP, lw, 3.62);
  s.addText("月次のお受取額・営業利益の比較（万円）", {
    x: M + 0.3, y: TOP + 0.2, w: lw - 0.6, h: 0.28,
    fontFace: F.jp, fontSize: 11, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
  const groups = [
    { t: "60坪（85杯）", i: 31.8, f: 57.1 },
    { t: "40坪（80杯）", i: 30.2, f: 53.0 },
    { t: "30坪（75杯）", i: 28.6, f: 49.0 },
  ];
  const maxV = 60, bx = M + 1.86, bw = lw - 2.9;
  groups.forEach((g, gi) => {
    const gy = TOP + 0.62 + gi * 0.96;
    s.addText(g.t, {
      x: M + 0.3, y: gy + 0.06, w: 1.5, h: 0.5,
      fontFace: F.jp, fontSize: 10.5, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    [["業務委託型", g.i, C.gold], ["FCプラン", g.f, "4C9A6E"]].forEach((b, bi) => {
      const y = gy + bi * 0.32;
      const w2 = (b[1] / maxV) * bw;
      s.addShape("roundRect", {
        x: bx, y: y + 0.03, w: w2, h: 0.22, rectRadius: 0.04,
        fill: { color: b[2] }, line: { type: "none" },
      });
      s.addText(b[0], {
        x: bx - 0.02, y, w: w2 - 0.1, h: 0.28,
        fontFace: F.jp, fontSize: 8.5, bold: true, color: C.white, margin: 0, valign: "middle",
        align: "left",
      });
      s.addText(b[1].toFixed(1) + " 万円", {
        x: bx + w2 + 0.08, y, w: 1.1, h: 0.28,
        fontFace: F.jp, fontSize: 10, bold: true, color: b[2], margin: 0, valign: "middle",
      });
    });
  });
  s.addText("※ 業務委託型＝未来屋書店様のお受取額（売上の25％＋月額固定5万円）／FCプラン＝未来屋書店様の償却前営業利益（人件費未計上・家賃0円）", {
    x: M + 0.3, y: TOP + 3.16, w: lw - 0.6, h: 0.3,
    fontFace: F.jp, fontSize: 8.5, color: C.footer, margin: 0, valign: "middle", lineSpacingMultiple: 1.15,
  });

  const rx = M + lw + 0.42, rw = R - rx;
  const cmpL = (t2) => ({
    text: t2,
    options: { fill: { color: C.grayBand }, color: C.ink, bold: true, fontSize: 9.5, fontFace: F.jp, valign: "middle", margin: [0.04, 0.1, 0.04, 0.1] },
  });
  const cmpA = (t2) => ({
    text: t2,
    options: { fill: { color: C.goldTint }, color: C.gold, bold: true, fontSize: 9.5, fontFace: F.jp, valign: "middle", margin: [0.04, 0.1, 0.04, 0.1] },
  });
  const cmpB = (t2) => ({
    text: t2,
    options: { fill: { color: C.white }, color: C.grayText, fontSize: 9.5, fontFace: F.jp, valign: "middle", margin: [0.04, 0.1, 0.04, 0.1] },
  });
  s.addTable(
    [
      [th("", { fill: C.grayBand, color: C.ink }), th("業務委託型", { fill: C.gold }), th("FCプラン", { fill: "8E8B84" })],
      [cmpL("初期費用"), cmpA("30〜50万円"), cmpB("900〜1,200万円")],
      [cmpL("運営主体"), cmpA("セルフカフェ"), cmpB("未来屋書店様")],
      [cmpL("日常業務"), cmpA("清掃・補充・発注"), cmpB("清掃・補充・発注")],
      [cmpL("原料・機器"), cmpA("セルフカフェ"), cmpB("未来屋書店様")],
      [cmpL("水道光熱費"), cmpA("未来屋書店様"), cmpB("未来屋書店様")],
      [cmpL("通信費"), cmpA("未来屋書店様"), cmpB("未来屋書店様")],
      [cmpL("消耗品（清掃用品）"), cmpA("未来屋書店様"), cmpB("未来屋書店様")],
      [cmpL("売上変動リスク"), cmpA("セルフカフェ"), cmpB("未来屋書店様")],
      [cmpL("投資回収"), cmpA("約2ヶ月"), cmpB("18〜21ヶ月")],
    ],
    {
      x: rx, y: TOP, w: rw, colW: [1.5, 1.42, 1.35],
      rowH: [0.32, 0.34, 0.34, 0.34, 0.34, 0.34, 0.34, 0.34, 0.34, 0.34],
      border: { type: "solid", color: C.warmLine, pt: 0.75 },
      autoPage: false,
    }
  );

  const py = TOP + 3.8;
  panel(s, M, py, CW, 0.86);
  s.addText(
    [
      { text: "まずは業務委託型での切り替えをご提案しております。", options: { fontFace: F.jp, fontSize: 12.5, bold: true, color: C.white } },
      { text: "  初期費用を抑えて実績をご確認いただいたうえで、FCプランへの切り替えもご相談いただけます。", options: { fontFace: F.jp, fontSize: 10.5, color: C.cvBody } },
    ],
    { x: M + 0.34, y: py, w: CW - 0.68, h: 0.86, margin: 0, valign: "middle" }
  );

  note(s, py + 1.0, "※ いずれも想定値であり、売上・利益を保証するものではありません。金額はすべて税抜・目安です。");
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
