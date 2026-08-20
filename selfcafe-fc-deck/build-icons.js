/**
 * react-icons（Lucide）をブランド色のPNGに書き出す。
 * build-deck.js より先に実行する。出力: assets/icons/<name>-<color>.png
 */
const React = require("react");
const { renderToStaticMarkup } = require("react-dom/server");
const Lu = require("react-icons/lu");
const sharp = require("sharp");
const fs = require("fs");
const path = require("path");

const OUT = path.join(__dirname, "assets", "icons");
fs.mkdirSync(OUT, { recursive: true });

// 使う色（build-deck.js のパレットと合わせる）
const COLORS = { green: "#1B5E20", white: "#FFFFFF", gold: "#B8893B", pale: "#A9C6AC", muted: "#B9B2A4" };

// 使うアイコン名（全色ぶん書き出す）
const NAMES = [
  "LuCoins", "LuClock", "LuMapPin", "LuShieldCheck", "LuUsers", "LuLaptop", "LuBuilding2",
  "LuClipboardList", "LuPiggyBank", "LuBadgeCheck", "LuHourglass", "LuSearch", "LuShoppingBag",
  "LuRuler", "LuHouse", "LuCoffee", "LuBookOpen", "LuWifi", "LuPlug", "LuMessagesSquare",
  "LuBellRing", "LuCctv", "LuFlame", "LuKeyRound", "LuVideo", "LuMic", "LuMessageCircle",
  "LuWallet", "LuTimer", "LuFileCheck", "LuLifeBuoy", "LuFileSearch", "LuPenLine", "LuHammer",
  "LuStore", "LuPhone", "LuMail", "LuGlobe", "LuGraduationCap", "LuCalculator", "LuListPlus",
  "LuCalendarCheck", "LuMonitor", "LuSmartphone", "LuZap", "LuCircleHelp", "LuTrendingUp",
  "LuSparkles", "LuCheck", "LuX", "LuMinus", "LuImage",
];

async function render(name, colorKey) {
  const Comp = Lu[name];
  if (!Comp) throw new Error("unknown icon: " + name);
  const svg = renderToStaticMarkup(
    React.createElement(Comp, { color: COLORS[colorKey], size: 512, strokeWidth: 1.8 })
  );
  const file = path.join(OUT, `${name}-${colorKey}.png`);
  await sharp(Buffer.from(svg))
    .resize(512, 512, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } })
    .png()
    .toFile(file);
}

(async () => {
  let n = 0;
  for (const colorKey of Object.keys(COLORS)) {
    for (const name of NAMES) { await render(name, colorKey); n++; }
  }
  console.log(`rendered ${n} icons -> ${OUT}`);
})();
