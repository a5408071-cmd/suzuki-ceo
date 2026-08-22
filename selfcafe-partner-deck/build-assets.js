/**
 * 素材生成：ロゴ（緑／白）、表紙の背景画像、Lucideアイコン。
 * 配色は業務委託型資料の実測値に合わせている。
 */
const React = require("react");
const { renderToStaticMarkup } = require("react-dom/server");
const Lu = require("react-icons/lu");
const sharp = require("sharp");
const fs = require("fs");
const path = require("path");

const OUT = path.join(__dirname, "assets");
const ICONS = path.join(OUT, "icons");
fs.mkdirSync(ICONS, { recursive: true });

// 業務委託型資料の実測パレット
const C = {
  green: "#106838",
  ink: "#0A3D23",
  gold: "#A8761F",
  goldLine: "#C79A44",
  white: "#FFFFFF",
  pale: "#8CC5A5",
  muted: "#A8B0AA",
};

// ---------------------------------------------------------------- ロゴ
// 元資料から抽出したアルファのみのロゴを、任意色で塗り直す
const SRC_LOGO = path.join(__dirname, "..", "selfcafe-fc-deck", "assets", "logo_white.png");

async function logo(name, hex, alpha = 1) {
  const src = sharp(SRC_LOGO);
  const { width, height } = await src.metadata();
  const rgb = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16));
  const tint = await sharp({ create: { width, height, channels: 3, background: { r: rgb[0], g: rgb[1], b: rgb[2] } } })
    .png().toBuffer();
  // 元画像のアルファをマスクとして使う
  const mask = await sharp(SRC_LOGO).extractChannel("alpha").toBuffer();
  let img = sharp(tint).joinChannel(mask);
  if (alpha < 1) {
    img = sharp(await img.png().toBuffer()).composite([{
      input: Buffer.from([255, 255, 255, Math.round(255 * alpha)]),
      raw: { width: 1, height: 1, channels: 4 }, tile: true, blend: "dest-in",
    }]);
  }
  await img.png().toFile(path.join(OUT, name));
}

// ---------------------------------------------------------------- 表紙背景
/** 濃緑グラデーション＋薄いグリッド＋巨大なロゴ透かし */
async function coverBg() {
  const W = 2666, H = 1500;
  // 斜めグラデーション（左下が最も暗い）を SVG で作り、ラスタライズ
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}">
    <defs>
      <linearGradient id="g" x1="0" y1="1" x2="1" y2="0">
        <stop offset="0%"   stop-color="#06301B"/>
        <stop offset="45%"  stop-color="#0B4526"/>
        <stop offset="100%" stop-color="#1B6B3C"/>
      </linearGradient>
      <linearGradient id="top" x1="0" y1="0" x2="1" y2="0">
        <stop offset="0%"   stop-color="#C79A44"/>
        <stop offset="38%"  stop-color="#8CC5A5"/>
        <stop offset="100%" stop-color="#0B4526"/>
      </linearGradient>
      <pattern id="grid" width="${W / 12}" height="${W / 12}" patternUnits="userSpaceOnUse">
        <path d="M ${W / 12} 0 L 0 0 0 ${W / 12}" fill="none" stroke="#FFFFFF" stroke-opacity="0.045" stroke-width="2"/>
      </pattern>
    </defs>
    <rect width="${W}" height="${H}" fill="url(#g)"/>
    <rect width="${W}" height="${H}" fill="url(#grid)"/>
    <rect width="${W}" height="10" fill="url(#top)"/>
  </svg>`;
  const base = await sharp(Buffer.from(svg)).png().toBuffer();

  // ロゴ透かし（右側に大きく、ごく薄く）
  const markW = Math.round(W * 0.66);
  const mark = await sharp(SRC_LOGO).resize({ width: markW }).png().toBuffer();
  const mm = await sharp(mark).metadata();
  const faded = await sharp(mark).composite([{
    input: Buffer.from([255, 255, 255, 12]),
    raw: { width: 1, height: 1, channels: 4 }, tile: true, blend: "dest-in",
  }]).png().toBuffer();

  await sharp(base)
    .composite([{ input: faded, left: Math.round(W * 0.52), top: Math.round(H * 0.30) }])
    .png()
    .toFile(path.join(OUT, "cover-bg.png"));
}

/** 濃緑パネル用の無地グラデ（左→右にわずかに明るく） */
async function panelBg() {
  const W = 1400, H = 1400;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}">
    <defs><linearGradient id="g" x1="0" y1="1" x2="1" y2="0">
      <stop offset="0%" stop-color="#0A4526"/><stop offset="100%" stop-color="#12633A"/>
    </linearGradient></defs>
    <rect width="${W}" height="${H}" fill="url(#g)"/>
  </svg>`;
  await sharp(Buffer.from(svg)).png().toFile(path.join(OUT, "panel-bg.png"));
}

// ---------------------------------------------------------------- アイコン
const ICON_NAMES = [
  "LuBuilding2", "LuUsers", "LuLaptop", "LuCoins", "LuTimer", "LuMapPin", "LuShieldCheck",
  "LuHouse", "LuCoffee", "LuBookOpen", "LuWifi", "LuPlug", "LuMessagesSquare", "LuCalendarCheck",
  "LuBellRing", "LuCctv", "LuFlame", "LuKeyRound", "LuVideo", "LuMic", "LuMessageCircle",
  "LuWallet", "LuFileCheck", "LuLifeBuoy", "LuFileSearch", "LuPenLine", "LuHammer", "LuStore",
  "LuPhone", "LuMail", "LuGlobe", "LuGraduationCap", "LuCalculator", "LuTrendingUp", "LuSparkles",
  "LuCheck", "LuX", "LuMinus", "LuImage", "LuClock", "LuBadgeCheck", "LuSearch", "LuShoppingBag",
  "LuHourglass", "LuPiggyBank", "LuClipboardList", "LuSmartphone", "LuZap", "LuRuler", "LuLandPlot",
  "LuHandCoins", "LuReceipt", "LuChartColumn",
];
const ICON_COLORS = { green: C.green, white: C.white, gold: C.gold, pale: C.pale, muted: "#B9B2A4", ink: C.ink };

async function icon(name, key) {
  const Comp = Lu[name];
  if (!Comp) { console.warn("skip unknown icon", name); return; }
  const svg = renderToStaticMarkup(React.createElement(Comp, { color: ICON_COLORS[key], size: 512, strokeWidth: 1.7 }));
  await sharp(Buffer.from(svg))
    .resize(512, 512, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } })
    .png()
    .toFile(path.join(ICONS, `${name}-${key}.png`));
}

(async () => {
  await logo("logo-green.png", C.green);
  await logo("logo-white.png", C.white);
  await coverBg();
  await panelBg();
  let n = 0;
  for (const key of Object.keys(ICON_COLORS)) {
    for (const nm of ICON_NAMES) { await icon(nm, key); n++; }
  }
  console.log(`assets ready: logo x2, cover-bg, panel-bg, icons ~${n}`);
})();
