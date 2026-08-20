/**
 * セルフカフェ 業務委託型FC資料 — 2〜10ページ 再設計版
 *
 * 既存資料（Googleスライド）の 2〜10 ページを差し替えるための 9 枚。
 * ヘッダー（タイトル位置・罫線・ロゴ・ページ番号）は既存 11〜25 ページと
 * 同一の座標・配色に合わせてあるため、そのまま差し替えられる。
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
  greenTint: "F0F5EF", // ごく薄い緑の面
  greenPale: "A9C6AC", // 濃色パネル上のサブテキスト
  brown: "5D4037", // 既存の見出し補助色
  brownLt: "8D6E63", // 既存ヘッダー罫線色
  ink: "3B3833", // 本文
  inkSub: "6E6A61", // 副本文
  muted: "9A948A", // 注記
  cream: "FAF3E3",
  creamDk: "E9DEC4",
  line: "E6DFCD",
  white: "FFFFFF",
  gold: "B8893B",
};

const F = { jp: "Noto Sans JP", num: "Arial" };

// ---------------------------------------------------------------- geometry
const SW = 13.333;
const M = 0.729; // 左右マージン（既存 52.5pt）
const CW = 11.875; // コンテンツ幅（既存 52.5〜907.5pt）
const RIGHT = M + CW; // 12.604
const TOP = 1.45; // ヘッダー罫線下からのコンテンツ開始
const BOTTOM = 6.9;

const ASSETS = path.join(__dirname, "assets");
const shadow = () => ({ type: "outer", color: "6B6250", blur: 9, offset: 1.4, angle: 90, opacity: 0.14 });

// ---------------------------------------------------------------- helpers
function shell(slide, title, pageNo) {
  slide.background = { color: C.bg };
  slide.addText(title, {
    x: M, y: 0.554, w: 9.3, h: 0.42,
    fontFace: F.jp, fontSize: 30, color: C.green, bold: false,
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

// 数値＋単位のスタットタイル
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

/* =============================================================== p.2
   業務委託型とは — 通常のFC加盟との比較表（ネイティブ表） */
{
  const s = pres.addSlide();
  shell(s, "業務委託型とは", 2);
  kicker(s, "通常のFC加盟とは異なり、貴社は場所を提供いただくだけ／運営主体は本部", 1.36);

  const colW = [2.30, 4.55, 5.02];
  const tx = M, ty = 2.00;

  // 「業務委託型」列を面で強調（表の下に敷く）
  s.addShape("roundRect", {
    x: tx + colW[0] + colW[1] - 0.02, y: ty - 0.30, w: colW[2] + 0.04, h: 4.85 + 0.30,
    rectRadius: 0.07, fill: { color: C.greenTint }, line: { type: "none" },
  });
  s.addText("▼ 貴社にご提案するのはこちら", {
    x: tx + colW[0] + colW[1] + 0.16, y: ty - 0.30, w: colW[2] - 0.32, h: 0.28,
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

  const rows = [
    [hdr("項目", C.creamDk, C.brown), hdr("通常のFC加盟", "8A8272", C.white), hdr("業務委託型", C.green, C.white)],
    [lab("運営主体"), old("加盟者様が独立オーナーとして運営"), nw("本部が運営を主導、貴社は場所提供のみ")],
    [lab("初期費用"), old("加盟金100万円＋工事費用等\n（開業資金目安：居抜き700万円〜／スケルトン工事1,000万円〜）"), nw("30万〜50万円程度（機器設置費のみ）")],
    [lab("出店スペース"), old("独立した店舗区画（25坪〜）が必要"), nw("既存スペースの一角に間借り設置可能")],
    [lab("日常業務"), old("加盟者様が運営全般を担当"), nw("清掃・補充のみ、1日15分程度\n（既存スタッフで対応可）")],
    [lab("売上金の流れ"), old("本部が売上金を管理し、ロイヤリティなどの費用を控除した残額を加盟者様へ振込"), nw("本部が売上金を管理し、25%を業務委託料として貴社へお支払い")],
    [lab("契約・撤退"), old("独立出店が前提の長期契約"), nw("契約期間の縛りあり（最低3年〜）／本部審査あり。原状回復の負担も小さい")],
  ];

  s.addTable(rows, {
    x: tx, y: ty, w: CW, colW,
    rowH: [0.50, 0.60, 0.86, 0.58, 0.72, 0.82, 0.77],
    border: { type: "solid", color: C.line, pt: 0.75 },
    autoPage: false,
  });
}

/* =============================================================== p.3
   業務委託型のメリット — 2×2 カード */
{
  const s = pres.addSlide();
  shell(s, "業務委託型のメリット", 3);
  kicker(s, "貴社にとってのメリット／既存事業の一角を、負担をかけずに収益資産へ", 1.36);

  const data = [
    ["01", "遊休スペースが収益化", "レジ横・待合スペースなどの", "デッドスペースが毎月の収益源に。",
      ["加盟金なし、機器設置費（30万〜50万円程度）のみで開始", "業務委託料として売上の25%を毎月お支払い（収益イメージは次ページ参照）"]],
    ["02", "運営負担はほぼゼロ", "日常業務は清掃・補充のみ、", "1日15分程度。",
      ["既存スタッフの稼働内で対応可能、新規採用は不要", "トラブル対応も24時間警備会社・遠隔監視で完結、現場対応はほぼ不要"]],
    ["03", "新しい来店動機に", "Wi-Fi・電源完備のカフェが", "新しい接点に。",
      ["Googleマップ・カフェ検索アプリで「近くのカフェ」を探す人と出会える", "書店・美容室・整体院・病院の待合など、待ち時間が生まれる業態と特に好相性"]],
    ["04", "低リスクで導入できる", "通常のFC加盟（700万円〜）と比べ、", "意思決定のハードルが低い。",
      ["独立出店ではないため原状回復の負担も小さい", "内装・システム導入まで本部がフルサポート、店舗運営のノウハウをそのまま活用"]],
  ];

  const cw = 5.77, ch = 2.42, gx = 0.325, gy = 0.26;
  data.forEach((d, i) => {
    const x = M + (i % 2) * (cw + gx);
    const y = 1.80 + Math.floor(i / 2) * (ch + gy);
    card(s, x, y, cw, ch);
    chip(s, x + 0.30, y + 0.28, d[0], { size: 0.44, fontSize: 13.5 });
    s.addText(d[1], {
      x: x + 0.88, y: y + 0.28, w: cw - 1.18, h: 0.44,
      fontFace: F.jp, fontSize: 15, bold: true, color: C.green, valign: "middle", margin: 0,
    });
    s.addText(
      [
        { text: d[2], options: { color: C.inkSub } },
        { text: d[3], options: { color: C.green, bold: true } },
      ],
      { x: x + 0.30, y: y + 0.86, w: cw - 0.60, h: 0.58, fontFace: F.jp, fontSize: 11, margin: 0, valign: "top", lineSpacingMultiple: 1.2 }
    );
    s.addShape("rect", { x: x + 0.30, y: y + 1.50, w: cw - 0.60, h: 0.011, fill: { color: C.line } });
    s.addText(
      d[4].map((t, j) => ({ text: t, options: { bullet: { characterCode: "25AA", indent: 11 }, breakLine: j < d[4].length - 1 } })),
      { x: x + 0.30, y: y + 1.62, w: cw - 0.60, h: 0.68, fontFace: F.jp, fontSize: 10, color: C.ink, margin: 0, valign: "top", paraSpaceAfter: 5 }
    );
  });
}

/* =============================================================== p.4
   会社概要 — ブランドパネル＋定義リスト */
{
  const s = pres.addSlide();
  shell(s, "会社概要", 4);

  // 左：ブランドパネル
  const px = M, py = 1.55, pw = 4.55, ph = 5.30;
  s.addShape("roundRect", { x: px, y: py, w: pw, h: ph, rectRadius: 0.09, fill: { color: C.greenDeep }, line: { type: "none" }, shadow: shadow() });
  s.addImage({ path: path.join(ASSETS, "logo_white.png"), x: px + 0.52, y: py + 1.35, w: 3.51, h: 0.50 });
  s.addText("セルフカフェ株式会社", {
    x: px + 0.52, y: py + 2.20, w: pw - 1.04, h: 0.42,
    fontFace: F.jp, fontSize: 19, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  s.addText("SELF CAFE Inc.", {
    x: px + 0.52, y: py + 2.66, w: pw - 1.04, h: 0.30,
    fontFace: F.num, fontSize: 11.5, color: C.greenPale, charSpacing: 2, margin: 0, valign: "middle",
  });
  s.addShape("rect", { x: px + 0.52, y: py + 3.16, w: 1.10, h: 0.021, fill: { color: C.greenPale } });
  s.addText("Wi-Fi・電源完備の無人カフェを\n企画・開発・運営しています。", {
    x: px + 0.52, y: py + 3.42, w: pw - 1.04, h: 0.80,
    fontFace: F.jp, fontSize: 12, color: C.greenPale, margin: 0, valign: "top", lineSpacingMultiple: 1.35,
  });

  // 右：定義リスト
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

/* =============================================================== p.5
   社会背景と市場性 — 課題3カード＋結論バンド */
{
  const s = pres.addSlide();
  shell(s, "社会背景と市場性", 5);
  kicker(s, "現代社会が抱える3つの課題", 1.36);

  const items = [
    ["01", "深刻な労働力不足", "少子高齢化と働き手の価値観の変化でサービス業の人手確保が難しくなり、人件費が上昇して経営負担が増している。", "サービス業の人手確保が困難に"],
    ["02", "サードプレイスの枯渇", "リモートワークの普及で自宅以外の手頃な作業場所が不足し、カフェやコワーキングは価格や混雑の課題が残り、学習スペースも制限が多く気軽に利用できない。", "自宅以外の作業場所が足りない"],
    ["03", "遊休資産の有効活用", "ビル空室やデッドスペースが増え、従来のテナント誘致だけでは埋まらず、新たな活用モデルの開発が求められている。", "空室・デッドスペースが埋まらない"],
  ];
  const cw = (CW - 2 * 0.28) / 3, ch = 3.10;
  items.forEach((it, i) => {
    const x = M + i * (cw + 0.28), y = 1.86;
    card(s, x, y, cw, ch);
    chip(s, x + 0.30, y + 0.30, it[0], { size: 0.44, fontSize: 13.5 });
    s.addText(it[1], {
      x: x + 0.30, y: y + 0.92, w: cw - 0.60, h: 0.40,
      fontFace: F.jp, fontSize: 15.5, bold: true, color: C.green, margin: 0, valign: "middle",
    });
    s.addText(it[3], {
      x: x + 0.30, y: y + 1.36, w: cw - 0.60, h: 0.28,
      fontFace: F.jp, fontSize: 10.5, bold: true, color: C.gold, margin: 0, valign: "middle",
    });
    s.addShape("rect", { x: x + 0.30, y: y + 1.74, w: cw - 0.60, h: 0.011, fill: { color: C.line } });
    s.addText(it[2], {
      x: x + 0.30, y: y + 1.88, w: cw - 0.60, h: 1.44,
      fontFace: F.jp, fontSize: 11, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.42,
    });
  });

  s.addShape("roundRect", { x: M, y: 5.28, w: CW, h: 1.56, rectRadius: 0.08, fill: { color: C.greenDeep }, line: { type: "none" }, shadow: shadow() });
  s.addText("この3つの課題に同時に応えるのが、「無人カフェ × 遊休スペース」というモデルです。", {
    x: M + 0.42, y: 5.54, w: CW - 0.84, h: 0.46,
    fontFace: F.jp, fontSize: 19, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  s.addText("人手をかけずに運営でき、働く・学ぶ場所を求める需要を受け止め、空いたスペースを収益資産に変えられます。", {
    x: M + 0.42, y: 6.06, w: CW - 0.84, h: 0.44,
    fontFace: F.jp, fontSize: 11.5, color: C.greenPale, margin: 0, valign: "middle",
  });
}

/* =============================================================== p.6
   人手不足対策：無人管理の優位性 — 主張パネル＋3カード */
{
  const s = pres.addSlide();
  shell(s, "人手不足対策：無人管理の優位性", 6);

  const px = M, py = 1.55, pw = 4.35, ph = 5.30;
  s.addShape("roundRect", { x: px, y: py, w: pw, h: ph, rectRadius: 0.09, fill: { color: C.greenDeep }, line: { type: "none" }, shadow: shadow() });
  s.addText("THE ADVANTAGE", {
    x: px + 0.46, y: py + 0.60, w: pw - 0.92, h: 0.28,
    fontFace: F.num, fontSize: 10.5, bold: true, color: C.greenPale, charSpacing: 3, margin: 0, valign: "middle",
  });
  s.addText("「人」に\n頼らない\n経営へ", {
    x: px + 0.46, y: py + 1.08, w: pw - 0.92, h: 2.20,
    fontFace: F.jp, fontSize: 31, bold: true, color: C.white, margin: 0, valign: "top", lineSpacingMultiple: 1.24,
  });
  s.addShape("rect", { x: px + 0.46, y: py + 3.48, w: 1.10, h: 0.021, fill: { color: C.greenPale } });
  s.addText("無人運営がもたらす、3つの構造的な強み。人に依存しないから、規模を広げても運営品質とコストが崩れません。", {
    x: px + 0.46, y: py + 3.74, w: pw - 0.92, h: 1.20,
    fontFace: F.jp, fontSize: 11.5, color: C.greenPale, margin: 0, valign: "top", lineSpacingMultiple: 1.4,
  });

  const cx = 5.35, cw = RIGHT - cx, chh = 1.62;
  const items = [
    ["01", "管理・教育コストの大幅削減", "スタッフ管理に伴う工数や人的リスクを抑え、運営にかかる負担を軽減します。"],
    ["02", "省人力運営で、無駄な人件費をカット", "必要最小限の作業で回る仕組みにより、人件費を抑えながら安定した収益を確保できます。"],
    ["03", "安定した収益と品質", "属人性を排除することで、常に均一したサービス品質を維持し、運営の安定性を高めます。"],
  ];
  items.forEach((it, i) => {
    const y = 1.55 + i * (chh + 0.22);
    card(s, cx, y, cw, chh);
    chip(s, cx + 0.34, y + 0.36, it[0], { size: 0.46, fontSize: 14 });
    s.addText(it[1], {
      x: cx + 0.98, y: y + 0.32, w: cw - 1.32, h: 0.42,
      fontFace: F.jp, fontSize: 16, bold: true, color: C.green, margin: 0, valign: "middle",
    });
    s.addText(it[2], {
      x: cx + 0.98, y: y + 0.80, w: cw - 1.32, h: 0.62,
      fontFace: F.jp, fontSize: 11.5, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
    });
  });
}

/* =============================================================== p.7
   既存事業とのシナジー — 主張バンド＋2×2 カード */
{
  const s = pres.addSlide();
  shell(s, "既存事業とのシナジー", 7);

  s.addShape("roundRect", { x: M, y: 1.40, w: CW, h: 1.28, rectRadius: 0.08, fill: { color: C.greenDeep }, line: { type: "none" }, shadow: shadow() });
  s.addText("カフェ併設がもたらす、既存事業へのプラス効果", {
    x: M + 0.42, y: 1.52, w: CW - 0.84, h: 0.28,
    fontFace: F.jp, fontSize: 10.5, bold: true, color: C.greenPale, margin: 0, valign: "middle",
  });
  s.addText("「待ち時間」の解消や新規集客のフックとして、カフェは最適の組み合わせです。", {
    x: M + 0.42, y: 1.86, w: CW - 0.84, h: 0.62,
    fontFace: F.jp, fontSize: 21, bold: true, color: C.white, margin: 0, valign: "middle",
  });

  const items = [
    ["待ち時間の解消", "「待つ時間」を「くつろぐ時間」へ", "書店・美容室・整体院・病院の待合・コインランドリーなど、待ち時間が生まれる場所にカフェがあるだけで顧客満足度が上がります。"],
    ["新規集客の新しい入り口", "「カフェ」で検索されて見つかる", "Googleマップやカフェ検索アプリ・SNSで「近くのカフェ」を探す人にも見つけてもらえる。カフェ目的で立ち寄った人が、既存事業を初めて知る新しいお客様になります。"],
    ["滞在時間・回遊率アップ", "「ついで買い」のきっかけに", "コーヒー片手に店内をゆっくり見て回れるようになり、滞在時間が伸びて追加購入・再来店のきっかけが増えます。"],
    ["省スペース・低コスト", "セルフ型だから始めやすい", "専任スタッフ不要。レジ横や待合の小さなスペースに設置でき、大掛かりな工事や人員追加なしで導入できます。"],
  ];
  const cw = 5.77, chh = 1.82, gx = 0.325, gy = 0.22;
  items.forEach((it, i) => {
    const x = M + (i % 2) * (cw + gx);
    const y = 2.90 + Math.floor(i / 2) * (chh + gy);
    card(s, x, y, cw, chh);
    // タグピル
    const tagW = 0.19 + it[0].length * 0.135;
    s.addShape("roundRect", { x: x + 0.28, y: y + 0.24, w: tagW, h: 0.30, rectRadius: 0.15, fill: { color: C.green }, line: { type: "none" } });
    s.addText(it[0], {
      x: x + 0.28, y: y + 0.24, w: tagW, h: 0.30,
      fontFace: F.jp, fontSize: 9.5, bold: true, color: C.white, align: "center", valign: "middle", margin: 0,
    });
    s.addText(it[1], {
      x: x + 0.28, y: y + 0.62, w: cw - 0.56, h: 0.36,
      fontFace: F.jp, fontSize: 14.5, bold: true, color: C.green, margin: 0, valign: "middle",
    });
    s.addText(it[2], {
      x: x + 0.28, y: y + 1.02, w: cw - 0.56, h: 0.66,
      fontFace: F.jp, fontSize: 10.5, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.26,
    });
  });

  s.addText("※待合スペース・レジ横カウンターなど、限られたスペースでも設置可能です。", {
    x: M, y: 6.80, w: CW, h: 0.24,
    fontFace: F.jp, fontSize: 9.5, color: C.muted, margin: 0, valign: "middle",
  });
}

/* =============================================================== p.8
   セルフカフェとは — 3つの「〜のような」＋店舗数パネル */
{
  const s = pres.addSlide();
  shell(s, "セルフカフェとは", 8);

  const lw2 = 6.55;
  const rows = [
    ["お家のような", "寛ぎすぎないが、設備が整った集中環境。"],
    ["カフェのような", "食事はないが、時間を気にせず美味しいコーヒーが飲める。"],
    ["図書館のような", "静かすぎない、適度な賑やかさがある。"],
  ];
  rows.forEach((r, i) => {
    const y = 1.55 + i * 1.26;
    card(s, M, y, lw2, 1.10);
    chip(s, M + 0.30, y + 0.32, String(i + 1).padStart(2, "0"), { size: 0.46, fontSize: 14 });
    s.addText(r[0], {
      x: M + 0.94, y: y + 0.20, w: lw2 - 1.24, h: 0.36,
      fontFace: F.jp, fontSize: 16, bold: true, color: C.green, margin: 0, valign: "middle",
    });
    s.addText(r[1], {
      x: M + 0.94, y: y + 0.58, w: lw2 - 1.24, h: 0.34,
      fontFace: F.jp, fontSize: 11.5, color: C.ink, margin: 0, valign: "middle",
    });
  });

  s.addShape("roundRect", { x: M, y: 5.44, w: lw2, h: 1.40, rectRadius: 0.08, fill: { color: C.greenDeep }, line: { type: "none" }, shadow: shadow() });
  s.addText("「誰もが気軽に使える」", {
    x: M + 0.40, y: 5.60, w: lw2 - 0.80, h: 0.40,
    fontFace: F.jp, fontSize: 14, color: C.greenPale, margin: 0, valign: "middle",
  });
  s.addText("第三の居場所", {
    x: M + 0.40, y: 6.02, w: lw2 - 0.80, h: 0.62,
    fontFace: F.jp, fontSize: 27, bold: true, color: C.white, margin: 0, valign: "middle",
  });

  // 右：展開店舗パネル
  const rx = 7.60, rw = RIGHT - rx;
  card(s, rx, 1.55, rw, 5.29);
  s.addText("展開エリア", {
    x: rx + 0.34, y: 1.78, w: rw - 0.68, h: 0.26,
    fontFace: F.jp, fontSize: 10.5, bold: true, color: C.brown, margin: 0, valign: "middle",
  });
  s.addText(
    [
      { text: "65", options: { fontFace: F.num, fontSize: 46, bold: true, color: C.green } },
      { text: " 店舗以上", options: { fontFace: F.jp, fontSize: 15, bold: true, color: C.green } },
    ],
    { x: rx + 0.34, y: 2.06, w: rw - 0.68, h: 0.66, margin: 0, valign: "middle", align: "left" }
  );
  s.addShape("rect", { x: rx + 0.34, y: 2.86, w: rw - 0.68, h: 0.011, fill: { color: C.line } });

  const prefs = [["愛知県", 35], ["大阪府", 18], ["東京都", 5], ["岩手県", 2], ["岐阜県", 2], ["埼玉県", 1], ["千葉県", 1], ["静岡県", 1], ["滋賀県", 1]];
  prefs.forEach((p, i) => {
    const y = 3.00 + i * 0.30;
    s.addText(p[0], {
      x: rx + 0.34, y: y, w: 2.0, h: 0.28,
      fontFace: F.jp, fontSize: 11.5, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(String(p[1]) + " 店舗", {
      x: rx + rw - 1.90, y: y, w: 1.56, h: 0.28,
      fontFace: F.jp, fontSize: 11.5, bold: true, color: C.green, align: "right", margin: 0, valign: "middle",
    });
  });

  s.addShape("roundRect", { x: rx + 0.34, y: 5.86, w: rw - 0.68, h: 0.78, rectRadius: 0.06, fill: { color: C.cream }, line: { type: "none" } });
  s.addText("※平均滞在時間は2〜3時間、1回のご利用で2杯程度購入される傾向があります。", {
    x: rx + 0.52, y: 5.86, w: rw - 1.04, h: 0.78,
    fontFace: F.jp, fontSize: 10, color: C.brown, margin: 0, valign: "middle", lineSpacingMultiple: 1.25,
  });
}

/* =============================================================== p.9
   実績データ：店舗展開数の推移 */
{
  const s = pres.addSlide();
  shell(s, "実績データ：店舗展開数の推移", 9);
  kicker(s, "営業中店舗の開店月ベース累計（2022.09〜2026.08／単位：店舗）", 1.36);

  const series = [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 4, 8, 8, 9, 9, 12, 17, 19, 21, 22, 23,
    23, 25, 28, 31, 32, 34, 39, 42, 44, 44, 46, 47, 48, 48, 50, 50, 51, 53, 57, 58, 60, 62, 64, 65];
  const labels = series.map((_, i) => {
    if (i === 0) return "2022.09";
    if (i === 12) return "2023.09";
    if (i === 24) return "2024.09";
    if (i === 36) return "2025.09";
    if (i === 47) return "2026.08";
    return ""; // 目盛りは5点のみ表示
  });

  const cx = M, cy = 1.80, cwid = 8.52, chh = 4.52;
  s.addChart(
    pres.ChartType.area,
    [{ name: "累計店舗数", labels, values: series }],
    {
      x: cx, y: cy, w: cwid, h: chh,
      chartColors: [C.greenMid], chartColorsOpacity: 22,
      showLegend: false, showTitle: false, showValue: false,
      lineSize: 2.5, lineSmooth: false,
      catAxisLabelColor: C.inkSub, catAxisLabelFontFace: F.num, catAxisLabelFontSize: 10,
      catAxisLabelFrequency: 1, catAxisLabelRotate: 0, catAxisMultiLevelLabels: false,
      catAxisLineShow: true, catAxisLineColor: C.creamDk,
      catGridLine: { style: "none" },
      valAxisLabelColor: C.muted, valAxisLabelFontFace: F.num, valAxisLabelFontSize: 10,
      valAxisMinVal: 0, valAxisMaxVal: 80, valAxisMajorUnit: 20,
      valAxisLineShow: false,
      valGridLine: { color: C.line, size: 0.75 },
      dataBorder: { pt: 0, color: C.greenMid },
    }
  );

  const tx = 9.55, tw = RIGHT - tx;
  statTile(s, tx, 1.80, tw, 1.42, "65", "店舗以上", "2026年8月時点の営業中店舗数", { dark: true, valueSize: 34 });
  statTile(s, tx, 3.36, tw, 1.42, "65", "倍", "1号店OPENからの店舗数の伸び");
  statTile(s, tx, 4.92, tw, 1.42, "47", "ヶ月", "1号店OPENからの経過期間");

  s.addText("2022年9月の1号店OPENから、驚異的なスピードで出店エリアを拡大。店舗ネットワークは着実に成長を続けています。", {
    x: M, y: 6.42, w: cwid, h: 0.46,
    fontFace: F.jp, fontSize: 11.5, color: C.ink, margin: 0, valign: "middle", lineSpacingMultiple: 1.25,
  });
}

/* =============================================================== p.10
   実績データ：月間利用者数の推移（複合グラフ） */
{
  const s = pres.addSlide();
  shell(s, "実績データ：月間利用者数の推移", 10);
  kicker(s, "全社月間利用者数（各月の合計）／店舗数の拡大に比例して拡大（2023.01〜2026.07）", 1.36);

  const labels = ["2023.01", "2023.07", "2024.01", "2024.07", "2025.01", "2025.07", "2026.01", "2026.07"];
  const users = [895, 1419, 9722, 26383, 46017, 74260, 80293, 101343];
  const stores = [1, 1, 9, 22, 32, 46, 51, 64];

  const cx = M, cy = 1.80, cwid = 8.52, chh = 4.52;
  s.addChart(
    [
      {
        type: pres.ChartType.bar,
        data: [{ name: "月間利用者数（人・左軸）", labels, values: users }],
        options: {
          chartColors: [C.greenMid], barGapWidthPct: 52,
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
      x: cx, y: cy, w: cwid, h: chh,
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
          valAxisMinVal: 0, valAxisMaxVal: 150000, valAxisMajorUnit: 50000,
          valAxisLineShow: false, valGridLine: { color: C.line, size: 0.75 },
        },
        {
          valAxisLabelColor: C.muted, valAxisLabelFontFace: F.num, valAxisLabelFontSize: 10,
          valAxisMinVal: 0, valAxisMaxVal: 70, valAxisMajorUnit: 20,
          valAxisLineShow: false, valGridLine: { style: "none" },
        },
      ],
    }
  );

  const tx = 9.55, tw = RIGHT - tx;
  statTile(s, tx, 1.80, tw, 1.42, "101,343", "人", "2026年7月の月間利用者数", { dark: true, valueSize: 26 });
  statTile(s, tx, 3.36, tw, 1.42, "104", "倍", "立ち上げ当初比の月間利用者数", { valueSize: 30 });
  statTile(s, tx, 4.92, tw, 1.42, "100,000", "杯", "月間販売数を突破", { valueSize: 26 });

  s.addText("立ち上げ当初から、月間利用者数の規模は着実に拡大を続けています。新規出店直後の一時的な希薄化を挟みながらも、ほぼ店舗数に比例するペースで成長しています。", {
    x: M, y: 6.42, w: cwid, h: 0.46,
    fontFace: F.jp, fontSize: 11.5, color: C.ink, margin: 0, valign: "middle", lineSpacingMultiple: 1.25,
  });
}

const out = path.join(__dirname, "selfcafe-fc-p2-10.pptx");
pres.writeFile({ fileName: out }).then(() => console.log("written:", out, fs.statSync(out).size, "bytes"));
