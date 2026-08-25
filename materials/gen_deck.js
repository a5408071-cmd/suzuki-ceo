const pptxgen = require("pptxgenjs");
const p = new pptxgen();
p.layout = "LAYOUT_WIDE"; // 13.333 x 7.5

// ---- palette ----
const GREEN = "0D4828";   // brand deep green
const GREEN2 = "1D6B44";  // mid green
const TINT  = "EDF4EE";   // light green tint
const TINT2 = "F7FAF7";
const ACCENT = "F2B705";  // chair yellow
const INK   = "1F2A22";
const MUTED = "5F6E62";
const LINE  = "D8E2DA";
const NEG   = "B5433C";
const WHITE = "FFFFFF";
const F = "Yu Gothic";

const W = 13.333, H = 7.5;
const A = "assets/";

function bg(s, color){ s.background = { color }; }
function eyebrow(s, txt, color){
  s.addText(txt, { x:0.6, y:0.32, w:9.5, h:0.32, fontFace:F, fontSize:11, bold:true,
    color: color||GREEN2, charSpacing:2, margin:0 });
}
function title(s, txt, opts){
  s.addText(txt, Object.assign({ x:0.6, y:0.62, w:12.1, h:0.75, fontFace:F, fontSize:27, bold:true,
    color: INK, margin:0 }, opts||{}));
}
function foot(s, txt, color){
  s.addText(txt, { x:0.6, y:7.02, w:12.1, h:0.4, fontFace:F, fontSize:8, color: color||MUTED,
    margin:0, valign:"top" });
}
function pageno(s, n){
  s.addText(String(n).padStart(2,"0"), { x:12.55, y:7.06, w:0.6, h:0.3, fontFace:F, fontSize:9,
    color:MUTED, align:"right", margin:0 });
}
function chip(s, x, y, w, h, label, value, sub, dark){
  s.addShape("roundRect", { x, y, w, h, rectRadius:0.06,
    fill:{ color: dark?GREEN:WHITE }, line:{ color: dark?GREEN:LINE, width:1 },
    shadow:{ type:"outer", color:"3A4A3E", opacity:0.18, blur:6, offset:2, angle:90 } });
  s.addText(label, { x:x+0.18, y:y+0.14, w:w-0.36, h:0.3, fontFace:F, fontSize:10.5, bold:true,
    color: dark?"BFD8C6":MUTED, margin:0 });
  s.addText(value, { x:x+0.18, y:y+0.40, w:w-0.36, h:h-0.9, fontFace:F, fontSize:30, bold:true,
    color: dark?WHITE:GREEN, margin:0, valign:"middle" });
  if(sub) s.addText(sub, { x:x+0.18, y:y+h-0.44, w:w-0.36, h:0.32, fontFace:F, fontSize:9.5,
    color: dark?"D9E8DE":MUTED, margin:0 });
}

// ============================================================ S1 表紙
{
const s = p.addSlide(); bg(s, GREEN);
s.addImage({ path:A+"cover.jpg", x:7.93, y:0, w:5.4, h:7.5 });
s.addShape("rect", { x:7.93, y:0, w:5.4, h:7.5, fill:{ color:GREEN, transparency:82 }, line:{ type:"none" } });
s.addImage({ path:A+"logo.png", x:0.62, y:0.5, w:0.62, h:0.57 });
s.addText("SELF CAFE  PARTNER PROGRAM", { x:1.38, y:0.62, w:6, h:0.34, fontFace:F, fontSize:12,
  bold:true, color:"9FC7AC", charSpacing:3, margin:0 });
s.addText("今ある空間を、\n人を増やさず\n収益資産に変える。", { x:0.6, y:1.7, w:7.1, h:2.9,
  fontFace:F, fontSize:41, bold:true, color:WHITE, lineSpacing:56, margin:0 });
s.addText("自社物件・遊休スペースを活かす、無人カフェFC「セルフカフェ パートナー制度」", {
  x:0.62, y:4.62, w:6.9, h:0.4, fontFace:F, fontSize:14, color:"D9E8DE", margin:0 });
const chips = [
  ["初期費用", "650万円〜", "標準総投資予算 800万円程度"],
  ["オーナー業務", "1日 約15分", "清掃・補充のみ／新規採用不要"],
  ["全国", "71店舗", "月間利用者 約10万人"],
];
chips.forEach((c,i)=>{
  const x = 0.62 + i*2.42;
  s.addShape("roundRect", { x, y:5.35, w:2.26, h:1.28, rectRadius:0.06,
    fill:{ color:"0A3A20" }, line:{ color:"2C6B4A", width:0.75 } });
  s.addText(c[0], { x:x+0.14, y:5.47, w:2.0, h:0.26, fontFace:F, fontSize:9.5, bold:true, color:"9FC7AC", margin:0 });
  s.addText(c[1], { x:x+0.14, y:5.72, w:2.0, h:0.5, fontFace:F, fontSize:19, bold:true, color:ACCENT, margin:0 });
  s.addText(c[2], { x:x+0.14, y:6.24, w:2.02, h:0.34, fontFace:F, fontSize:8, color:"D9E8DE", margin:0 });
});
s.addText("セルフカフェ株式会社", { x:0.62, y:6.85, w:4, h:0.3, fontFace:F, fontSize:11, color:"D9E8DE", margin:0 });
s.addText("※ 店舗数・利用者数は2026年7月時点。初期費用は居抜き・20坪想定の目安。", {
  x:0.62, y:7.16, w:7.0, h:0.28, fontFace:F, fontSize:7.5, color:"9FC7AC", margin:0 });
}

// ============================================================ S2 セルフカフェFCとは
{
const s = p.addSlide(); bg(s, WHITE);
eyebrow(s, "01 ｜ セルフカフェFCとは");
title(s, "貸すのではなく、自分で持つ「省人型テナント」。");
s.addText([
  { text:"セルフカフェは、ドリンク1杯で時間制限なく使える無人のカフェ型ワークスペースです。", options:{ breakLine:true } },
  { text:"空きテナントに入居者を探すのではなく、オーナー自身が低投資で運営する「収益事業」として、", options:{ breakLine:true } },
  { text:"遊休スペースを稼働資産に変えます。", options:{} },
], { x:0.6, y:1.55, w:6.7, h:1.05, fontFace:F, fontSize:12, color:INK, lineSpacing:20, margin:0 });
s.addText("労働力不足 ／ 手頃な作業場所の不足 ／ 遊休不動産の増加 —— 3つの社会課題が需要をつくっています。", {
  x:0.6, y:2.78, w:6.7, h:0.4, fontFace:F, fontSize:10, color:MUTED, margin:0 });
const items = [
  ["必要面積", "20坪〜", "既存施設の一角でも可"],
  ["利用料", "420円〜", "1杯で滞在無制限・登録不要"],
  ["営業時間", "24時間可", "無人・年中無休も設定可能"],
  ["新規採用", "0人", "既存スタッフの稼働内で運営"],
];
items.forEach((c,i)=>{
  const x = 0.6 + (i%2)*3.42, y = 3.3 + Math.floor(i/2)*1.62;
  chip(s, x, y, 3.22, 1.42, c[0], c[1], c[2], false);
});
s.addImage({ path:A+"concept.jpg", x:7.6, y:1.55, w:5.13, h:4.63 });
s.addText("店内利用の様子（実店舗）", { x:7.6, y:6.24, w:5.13, h:0.3, fontFace:F, fontSize:9, color:MUTED, align:"right", margin:0 });
foot(s, "※ 予約不要・会員登録不要。二次利用（貸しスペース・物販等）の可否は立地・契約条件により異なります。");
pageno(s,2);
}

// ============================================================ S3 まず数字から
{
const s = p.addSlide(); bg(s, WHITE);
eyebrow(s, "02 ｜ まず、数字から");
title(s, "650万円から始めて、自社物件なら月47万円モデル・回収約17か月。");
const cards = [
  ["初期投資", "650万円〜", "最低開業費（居抜き・20坪）\n標準総投資予算：800万円程度〜", false],
  ["月間利益モデル（75杯/日）", "自社物件 47万円", "賃貸物件の場合：27万円\n（償却前営業利益・月）", true],
  ["オーナーの運営時間", "1日 約15分", "清掃・原料補充のみ\n委託運営も可能", false],
  ["投資回収期間モデル", "約17か月", "自社物件・75杯/日・投資800万円\n賃貸の場合：約30か月", true],
];
cards.forEach((c,i)=>{
  const x = 0.6 + (i%2)*6.28, y = 1.75 + Math.floor(i/2)*2.45;
  const dark = c[3];
  s.addShape("roundRect", { x, y, w:5.85, h:2.1, rectRadius:0.07,
    fill:{ color: dark?GREEN:TINT }, line:{ color: dark?GREEN:LINE, width:1 } });
  s.addText(c[0], { x:x+0.28, y:y+0.2, w:5.3, h:0.32, fontFace:F, fontSize:12, bold:true,
    color: dark?"9FC7AC":GREEN2, margin:0 });
  s.addText(c[1], { x:x+0.28, y:y+0.52, w:5.3, h:0.85, fontFace:F, fontSize:34, bold:true,
    color: dark?WHITE:GREEN, margin:0, valign:"middle" });
  s.addText(c[2], { x:x+0.28, y:y+1.42, w:5.3, h:0.6, fontFace:F, fontSize:10,
    color: dark?"D9E8DE":MUTED, lineSpacing:14, margin:0 });
});
foot(s, "※ 利益・回収は1杯420円・原価20%・20坪モデルの試算（償却前・税引前）。収益・回収期間を保証するものではありません。詳細はp.11〜15。");
pageno(s,3);
}

// ============================================================ S4 実績: 規模
{
const s = p.addSlide(); bg(s, WHITE);
eyebrow(s, "03 ｜ 実績：店舗網と利用者");
title(s, "47か月で71店舗。月間利用者は約10万人へ。");
s.addChart(p.ChartType.bar, [{
  name:"店舗数",
  labels:["2022.09","2023.09","2024.09","2025.09","2026.07"],
  values:[1,2,23,48,71],
}], { x:0.6, y:1.75, w:7.2, h:4.7, barDir:"col",
  chartColors:[GREEN2], showLegend:false, showTitle:false,
  showValue:true, dataLabelPosition:"outEnd", dataLabelColor:INK, dataLabelFontFace:F, dataLabelFontSize:12, dataLabelFormatCode:"0\"店\"",
  catAxisLabelColor:MUTED, catAxisLabelFontFace:F, catAxisLabelFontSize:10,
  valAxisHidden:true, valGridLine:{ style:"none" }, catGridLine:{ style:"none" },
  valAxisMaxVal:80 });
const st = [
  ["月間利用者数（2026年7月）", "約10万人", "3年半で約100倍に拡大"],
  ["展開エリア", "11都府県", "年内にさらに10店舗以上出店予定"],
  ["1店舗あたり月間利用者", "約1,400人", "平均滞在2〜3時間・1人2杯程度"],
];
st.forEach((c,i)=>{
  chip(s, 8.2, 1.75+i*1.62, 4.5, 1.42, c[0], c[1], c[2], i===0);
});
foot(s, "出典：社内管理台帳（営業中店舗の開店月ベース累計・各月の全社利用者数）。2026年7月時点、利用者数は概数。");
pageno(s,4);
}

// ============================================================ S5 黒字率（新設・核）
{
const s = p.addSlide(); bg(s, GREEN);
eyebrow(s, "04 ｜ 実績：既存店の黒字率", "9FC7AC");
s.addText("開業6か月以上の店舗の93.8%が、営業黒字。", { x:0.6, y:0.62, w:12.1, h:0.7,
  fontFace:F, fontSize:27, bold:true, color:WHITE, margin:0 });
s.addText("93.8%", { x:0.6, y:1.6, w:5.6, h:1.7, fontFace:F, fontSize:96, bold:true, color:ACCENT, margin:0 });
s.addText("2026年7月単月・営業利益ベース\n開業6か月以上の48店中45店が黒字", { x:0.66, y:3.35, w:5.5, h:0.7,
  fontFace:F, fontSize:13, bold:true, color:WHITE, lineSpacing:19, margin:0 });
const rows = [
  ["開業12か月以上でも", "93.0%（43店中40店）"],
  ["5〜7月 3か月連続黒字", "79.2%"],
  ["平均営業利益 ／ 中央値", "21.6万円 ／ 13.2万円"],
  ["平均月商 ／ 中央値", "74.2万円 ／ 66.8万円"],
];
rows.forEach((r,i)=>{
  const y = 4.35 + i*0.56;
  s.addText(r[0], { x:0.66, y, w:2.9, h:0.5, fontFace:F, fontSize:11.5, color:"BFD8C6", margin:0 });
  s.addText(r[1], { x:3.6, y, w:2.9, h:0.5, fontFace:F, fontSize:12.5, bold:true, color:WHITE, margin:0 });
});
// 分布チャート
s.addShape("roundRect", { x:6.85, y:1.6, w:5.9, h:4.9, rectRadius:0.08, fill:{ color:WHITE }, line:{ type:"none" } });
s.addText("7月単月・営業利益の店舗分布（開業6か月以上・48店）", { x:7.1, y:1.78, w:5.4, h:0.3,
  fontFace:F, fontSize:11, bold:true, color:INK, margin:0 });
s.addChart(p.ChartType.bar, [{
  name:"店舗数",
  labels:["赤字","0〜10万円","10〜20万円","20〜30万円","30万円以上"],
  values:[3,13,15,5,12],
}], { x:7.0, y:2.15, w:5.6, h:4.1, barDir:"bar",
  chartColors:[NEG,GREEN2,GREEN2,GREEN2,GREEN2], chartColorsOpacity:100,
  showLegend:false, showTitle:false, showValue:true, dataLabelPosition:"outEnd",
  dataLabelColor:INK, dataLabelFontFace:F, dataLabelFontSize:11, dataLabelFormatCode:"0\"店\"",
  catAxisLabelColor:INK, catAxisLabelFontFace:F, catAxisLabelFontSize:10,
  valAxisHidden:true, valGridLine:{ style:"none" }, catGridLine:{ style:"none" }, valAxisMaxVal:17 });
s.addText("参考：家賃を除いた営業利益では 48店すべてが黒字（平均45.5万円）＝ 自社物件オーナー様の参考値", {
  x:0.66, y:6.62, w:12.0, h:0.3, fontFace:F, fontSize:12, bold:true, color:ACCENT, margin:0 });
foot(s, "※ 2026年7月速報値・単月実績。営業利益（減価償却費込み）ベース。開業6か月以上経過し月次店舗損益を集計している48店が対象（内訳・店舗別数値は非開示）。店舗により賃料・立地・面積・運営条件等は異なり、将来の収益を保証するものではありません。", "9FC7AC");
pageno(s,5);
}

// ============================================================ S6 事例
{
const s = p.addSlide(); bg(s, WHITE);
eyebrow(s, "05 ｜ 実店舗事例");
title(s, "「真ん中の店」でも回る。上振れすれば、こうなる。");
// 標準店
s.addShape("roundRect", { x:0.6, y:1.7, w:6.05, h:4.9, rectRadius:0.08, fill:{ color:TINT }, line:{ color:LINE, width:1 } });
s.addText("STANDARD MODEL", { x:0.9, y:1.95, w:5.4, h:0.3, fontFace:F, fontSize:10, bold:true, color:GREEN2, charSpacing:2, margin:0 });
s.addText("標準的な店舗（実績中央値ベース）", { x:0.9, y:2.25, w:5.4, h:0.44, fontFace:F, fontSize:17, bold:true, color:INK, margin:0 });
const std = [["月商","約67万円"],["営業利益","約13万円"],["販売ペース","約50杯/日 相当"]];
std.forEach((r,i)=>{
  const y = 2.9 + i*0.94;
  s.addText(r[0], { x:0.9, y, w:2.2, h:0.5, fontFace:F, fontSize:12, color:MUTED, margin:0, valign:"middle" });
  s.addText(r[1], { x:3.0, y, w:3.4, h:0.5, fontFace:F, fontSize:24, bold:true, color:GREEN, margin:0, valign:"middle" });
});
s.addText("開業6か月以上店舗の月商・営業利益の中央値から作成した匿名モデル。特別に成功した店ではなく「分布の真ん中」です。", {
  x:0.9, y:5.75, w:5.45, h:0.7, fontFace:F, fontSize:10, color:MUTED, lineSpacing:14, margin:0 });
// 盛岡
s.addShape("roundRect", { x:7.05, y:1.7, w:5.68, h:4.9, rectRadius:0.08, fill:{ color:WHITE }, line:{ color:LINE, width:1 },
  shadow:{ type:"outer", color:"3A4A3E", opacity:0.15, blur:6, offset:2, angle:90 } });
s.addImage({ path:A+"morioka.jpg", x:7.05, y:1.7, w:5.68, h:1.37 });
s.addText("CASE STUDY ｜ 岩手県・盛岡駅前店（2025年4月開業）", { x:7.3, y:3.2, w:5.2, h:0.3,
  fontFace:F, fontSize:10.5, bold:true, color:GREEN2, margin:0 });
const mor = [["販売ペース","約80杯/日"],["月間売上高","96万円"],["家賃","28万円"],["償却前営業利益","23.2万円"]];
mor.forEach((r,i)=>{
  const x = 7.3 + (i%2)*2.75, y = 3.6 + Math.floor(i/2)*1.0;
  s.addText(r[0], { x, y, w:2.6, h:0.28, fontFace:F, fontSize:9.5, color:MUTED, margin:0 });
  s.addText(r[1], { x, y:y+0.27, w:2.6, h:0.5, fontFace:F, fontSize:21, bold:true, color:GREEN, margin:0 });
});
s.addText("地方都市の駅前立地。月28万円の家賃を払った上で利益を確保。20坪・40席・24時間営業。", {
  x:7.3, y:5.75, w:5.2, h:0.66, fontFace:F, fontSize:10, color:MUTED, lineSpacing:14, margin:0 });
foot(s, "※ 標準店は2026年7月速報の中央値による匿名モデル（営業利益は減価償却費込み）。盛岡駅前店は実績値・償却前・概数。他店舗での同等の売上・利益を保証するものではありません。参考：月商180万円超の店舗もあります。");
pageno(s,6);
}

// ============================================================ S7 自社物件
{
const s = p.addSlide(); bg(s, WHITE);
eyebrow(s, "06 ｜ 自社物件の優位性");
title(s, "最大の固定費「家賃」がゼロ。利益は1.7倍、回収期間は約半分。");
// 比較 2カード
const cmp = [
  ["賃貸物件で開業", "月間利益 27万円", "回収期間 約30か月", TINT, INK],
  ["自社物件で開業", "月間利益 47万円", "回収期間 約17か月", GREEN, WHITE],
];
cmp.forEach((c,i)=>{
  const x = 0.6 + i*3.7;
  s.addShape("roundRect", { x, y:1.8, w:3.4, h:3.3, rectRadius:0.08,
    fill:{ color:c[3] }, line:{ color: i?GREEN:LINE, width:1 } });
  s.addText(c[0], { x:x+0.25, y:2.05, w:2.9, h:0.35, fontFace:F, fontSize:13, bold:true,
    color: i? "9FC7AC":MUTED, margin:0 });
  s.addText(c[1].replace("月間利益 ",""), { x:x+0.25, y:2.5, w:2.9, h:0.9, fontFace:F, fontSize:40, bold:true,
    color: i?ACCENT:GREEN, margin:0, valign:"middle" });
  s.addText("月間利益（75杯/日・償却前）", { x:x+0.25, y:3.45, w:2.9, h:0.3, fontFace:F, fontSize:9.5,
    color: i?"D9E8DE":MUTED, margin:0 });
  s.addText(c[2], { x:x+0.25, y:4.0, w:2.9, h:0.5, fontFace:F, fontSize:16, bold:true,
    color: i?WHITE:INK, margin:0 });
});
s.addText("＋20万円/月（+74%）", { x:2.9, y:5.25, w:4.4, h:0.45, fontFace:F, fontSize:16, bold:true, color:GREEN2, margin:0 });
// 右: 理由
const why = [
  ["家賃という最大の固定費が不要", "売上がそのまま利益層に乗る構造。"],
  ["空室の機会損失を収益に転換", "テナント誘致を待たず、自社で稼働させる。"],
  ["人件費の追加もゼロ", "既存スタッフの稼働内で運営。新規採用不要。"],
  ["実績（参考値）", "既存店から家賃を除くと48店すべてが営業黒字・平均45.5万円。"],
];
why.forEach((c,i)=>{
  const y = 1.8 + i*1.22;
  s.addShape("ellipse", { x:8.0, y:y+0.05, w:0.34, h:0.34, fill:{ color:GREEN2 }, line:{ type:"none" } });
  s.addText(String(i+1), { x:8.0, y:y+0.05, w:0.34, h:0.34, fontFace:F, fontSize:12, bold:true, color:WHITE, align:"center", valign:"middle", margin:0 });
  s.addText(c[0], { x:8.5, y, w:4.3, h:0.4, fontFace:F, fontSize:13.5, bold:true, color:INK, margin:0 });
  s.addText(c[1], { x:8.5, y:y+0.4, w:4.3, h:0.6, fontFace:F, fontSize:10.5, color:MUTED, lineSpacing:15, margin:0 });
});
foot(s, "※ 利益・回収は1杯420円・20坪モデル（総投資800万円）の試算。実績参考値は2026年7月速報の家賃控除前営業利益（開業6か月以上・48店）。収益を保証するものではありません。");
pageno(s,7);
}

// ============================================================ S8 無人運営の仕組み
{
const s = p.addSlide(); bg(s, WHITE);
eyebrow(s, "07 ｜ 無人運営の仕組み");
title(s, "開閉店・監視・緊急対応まで、仕組みが店番をする。");
const sys = [
  ["スマートロック", "開錠・施錠を遠隔で完全自動化。営業時間を分単位で制御、24時間営業も容易。"],
  ["カメラ6台以上＋双方向音声", "死角のない常時モニタリング。カメラ越しにお客様へ直接声がけ・サポート。"],
  ["警備会社と連携（ALSOK・SECOM等）", "非常ボタン・火災センサー・侵入検知と連動し、非常時は警備員が駆けつけ。"],
  ["LINEチャットサポート", "お客様からの問い合わせに即時対応（営業サポートプラン）。"],
];
sys.forEach((c,i)=>{
  const y = 1.75 + i*1.22;
  s.addShape("ellipse", { x:0.6, y:y+0.03, w:0.36, h:0.36, fill:{ color:GREEN }, line:{ type:"none" } });
  s.addText(String(i+1), { x:0.6, y:y+0.03, w:0.36, h:0.36, fontFace:F, fontSize:12, bold:true, color:ACCENT, align:"center", valign:"middle", margin:0 });
  s.addText(c[0], { x:1.12, y, w:5.9, h:0.4, fontFace:F, fontSize:14.5, bold:true, color:INK, margin:0 });
  s.addText(c[1], { x:1.12, y:y+0.4, w:5.9, h:0.66, fontFace:F, fontSize:10.5, color:MUTED, lineSpacing:15, margin:0 });
});
s.addImage({ path:A+"camera1.jpg", x:7.5, y:1.75, w:5.23, h:2.83 });
s.addImage({ path:A+"camera2.jpg", x:7.5, y:4.72, w:5.23, h:1.46 });
s.addText("店内カメラの実際の映像（平日日中）", { x:7.5, y:6.22, w:5.23, h:0.28, fontFace:F, fontSize:9, color:MUTED, align:"right", margin:0 });
foot(s, "※ 連携する警備会社・設備構成は物件・環境により異なります。LINEチャット対応は有料オプション（営業サポートプラン）です。");
pageno(s,8);
}

// ============================================================ S9 オーナーの仕事
{
const s = p.addSlide(); bg(s, WHITE);
eyebrow(s, "08 ｜ オーナー様の業務");
title(s, "やることは清掃と補充だけ。1日約15分、それすら委託できる。");
s.addText("約15分/日", { x:0.6, y:1.7, w:5.8, h:1.1, fontFace:F, fontSize:58, bold:true, color:GREEN, margin:0 });
const tasks = [
  ["① 店内清掃", "テーブル・床の簡易清掃、ゴミ回収。営業前後の都合の良い時間でOK。"],
  ["② 原料・備品の補充", "コーヒー豆・カップ等の補充。在庫は本部から定期配送。"],
];
tasks.forEach((c,i)=>{
  const y = 3.0 + i*1.3;
  s.addShape("roundRect", { x:0.6, y, w:6.2, h:1.1, rectRadius:0.07, fill:{ color:TINT }, line:{ color:LINE, width:1 } });
  s.addText(c[0], { x:0.85, y:y+0.13, w:5.7, h:0.36, fontFace:F, fontSize:13.5, bold:true, color:GREEN2, margin:0 });
  s.addText(c[1], { x:0.85, y:y+0.5, w:5.7, h:0.5, fontFace:F, fontSize:10.5, color:MUTED, margin:0 });
});
s.addShape("roundRect", { x:0.6, y:5.7, w:6.2, h:0.95, rectRadius:0.07, fill:{ color:GREEN }, line:{ type:"none" } });
s.addText("自分でやらない選択肢も。清掃パートナーへの委託 月3万円程度（収支モデルに計上済み）", {
  x:0.85, y:5.7, w:5.75, h:0.95, fontFace:F, fontSize:11.5, bold:true, color:WHITE, valign:"middle", lineSpacing:17, margin:0 });
s.addImage({ path:A+"hands.jpg", x:7.6, y:1.7, w:5.13, h:4.69 });
foot(s, "※ 所要時間は店舗規模・営業時間により変動します。運営に必要な営業届の申請は本部が代行、食品衛生責任者の資格は貴社にて取得いただきます。");
pageno(s,9);
}

// ============================================================ S10 本部の仕事
{
const s = p.addSlide(); bg(s, WHITE);
eyebrow(s, "09 ｜ 本部のサポート");
title(s, "立地診断から集客まで本部が伴走。オーナー様は「場所」を出すだけ。");
const cols = [
  ["開業前", ["パートナー制度のご説明・収益試算表の作成","現地調査・工事業者の手配","店舗デザイン・設計","運営研修","営業届の申請代行"]],
  ["開業後", ["店舗運営相談","マシンの定期メンテナンス（月1回）","WEBサイト作成・更新","SEO・MEO・AIO（LLMO）による集客","改修工事サポート"]],
];
cols.forEach((c,i)=>{
  const x = 0.6 + i*4.55;
  s.addShape("roundRect", { x, y:1.75, w:4.25, h:4.55, rectRadius:0.08, fill:{ color:TINT2 }, line:{ color:LINE, width:1 } });
  s.addText(c[0], { x:x+0.25, y:1.95, w:3.7, h:0.4, fontFace:F, fontSize:16, bold:true, color:GREEN, margin:0 });
  s.addText(c[1].map((t,j)=>({ text:t, options:{ bullet:{ code:"2713", indent:14 }, color:INK,
    breakLine: j<c[1].length-1, paraSpaceAfter:10 } })),
    { x:x+0.25, y:2.5, w:3.8, h:3.6, fontFace:F, fontSize:11.5, margin:0, valign:"top" });
});
s.addShape("roundRect", { x:9.75, y:1.75, w:2.98, h:4.55, rectRadius:0.08, fill:{ color:GREEN }, line:{ type:"none" } });
s.addText("ロイヤリティ", { x:9.98, y:2.1, w:2.5, h:0.35, fontFace:F, fontSize:12, bold:true, color:"9FC7AC", margin:0 });
s.addText("月5万円\n一律", { x:9.98, y:2.5, w:2.55, h:1.5, fontFace:F, fontSize:30, bold:true, color:WHITE, lineSpacing:40, margin:0 });
s.addText("売上連動ではありません。売上が伸びた分は、すべてオーナー様の利益になります。", {
  x:9.98, y:4.2, w:2.55, h:1.9, fontFace:F, fontSize:11, color:"D9E8DE", lineSpacing:17, margin:0 });
foot(s, "※ サポート内容の詳細・有料オプション（公式LINE作成、営業サポートプラン等）は別紙をご参照ください。");
pageno(s,10);
}

// ============================================================ S11 初期投資
{
const s = p.addSlide(); bg(s, WHITE);
eyebrow(s, "10 ｜ 初期投資の詳細");
title(s, "最低開業費650万円〜。予備費まで見た標準予算は800万円程度。");
// 積み上げ図（シェイプ）
const parts = [
  ["加盟金", 100, GREEN],
  ["出店工事費用", 400, GREEN2],
  ["出店準備金（備品類）", 100, "4E8B67"],
  ["保証金（預り金）", 50, "7FAE92"],
  ["諸経費・予備費", 150, "C9A227"],
];
let cx = 0.6; const total = 800; const barW = 8.6;
parts.forEach((pt,i)=>{
  const w = barW*pt[1]/total;
  s.addShape("rect", { x:cx, y:2.0, w, h:1.05, fill:{ color:pt[2] }, line:{ color:WHITE, width:1.5 } });
  s.addText(pt[1]+"万", { x:cx, y:2.0, w, h:1.05, fontFace:F, fontSize: w>1?13:10, bold:true, color:WHITE, align:"center", valign:"middle", margin:0 });
  s.addText(pt[0], { x:cx-0.1, y: (i%2? 3.15:1.55), w:w+0.6, h:0.35, fontFace:F, fontSize:9.5, color:MUTED, align:"center", margin:0 });
  cx += w;
});
s.addShape("rightBrace", { x:9.28, y:2.0, w:0.18, h:1.05, line:{ color:MUTED, width:1.2 } });
s.addText("標準総投資予算\n800万円程度", { x:9.55, y:1.95, w:3.2, h:1.15, fontFace:F, fontSize:14, bold:true, color:INK, lineSpacing:20, margin:0, valign:"middle" });
s.addText("最低開業費 650万円〜（居抜き・20坪）", { x:0.6, y:3.55, w:8.6, h:0.35, fontFace:F, fontSize:12, bold:true, color:GREEN, margin:0 });
const notes = [
  ["ロイヤリティ", "月5万円 一律（売上連動なし）"],
  ["研修費", "なし"],
  ["スケルトン物件の場合", "工事費が増え950万円〜（総予算1,200万円程度）"],
  ["保証金50万円", "預り金。原料・ロイヤリティの支払確認後、契約条件に基づき精算"],
];
notes.forEach((n,i)=>{
  const x = 0.6 + (i%2)*6.28, y = 4.35 + Math.floor(i/2)*1.05;
  s.addShape("roundRect", { x, y, w:5.85, h:0.88, rectRadius:0.06, fill:{ color:TINT2 }, line:{ color:LINE, width:1 } });
  s.addText(n[0], { x:x+0.22, y:y+0.1, w:5.4, h:0.3, fontFace:F, fontSize:10.5, bold:true, color:GREEN2, margin:0 });
  s.addText(n[1], { x:x+0.22, y:y+0.4, w:5.4, h:0.4, fontFace:F, fontSize:11.5, color:INK, margin:0 });
});
foot(s, "※ 金額はすべて税抜・目安。物件条件・工事範囲により変動します。他社カフェ業態（開業資金3,000万〜8,400万円）の1/4〜1/12の水準です。");
pageno(s,11);
}

// ============================================================ S12 収益モデル
{
const s = p.addSlide(); bg(s, WHITE);
eyebrow(s, "11 ｜ 収益モデル（20坪・40席・75杯/日）");
title(s, "同じ売上でも、自社物件なら月間利益47万円。");
const rows = [
  ["売上高（420円×75杯×30日）", "945,000", "945,000"],
  ["ドリンク原料（20%）", "▲189,000", "▲189,000"],
  ["清掃費", "▲30,000", "▲30,000"],
  ["水道光熱費（24H営業想定）", "▲80,000", "▲80,000"],
  ["機械使用料（マシン2台・決済端末）", "▲86,000", "▲86,000"],
  ["セキュリティ費", "▲20,000", "▲20,000"],
  ["ロイヤリティ（一律）", "▲50,000", "▲50,000"],
  ["雑費", "▲20,000", "▲20,000"],
  ["家賃（坪1万円）", "▲200,000", "0"],
];
const tbl = [[
  { text:"項目（月次・円）", options:{ bold:true, color:WHITE, fill:{color:GREEN}, fontSize:11 } },
  { text:"賃貸物件", options:{ bold:true, color:WHITE, fill:{color:GREEN}, fontSize:11, align:"right" } },
  { text:"自社物件", options:{ bold:true, color:WHITE, fill:{color:GREEN}, fontSize:11, align:"right" } },
]];
rows.forEach(r=>{
  tbl.push([
    { text:r[0], options:{ color:INK, fontSize:10.5 } },
    { text:r[1], options:{ color: r[1].startsWith("▲")?NEG:INK, fontSize:10.5, align:"right" } },
    { text:r[2], options:{ color: r[2].startsWith("▲")?NEG:INK, fontSize:10.5, align:"right" } },
  ]);
});
tbl.push([
  { text:"償却前営業利益", options:{ bold:true, color:GREEN, fill:{color:TINT}, fontSize:12 } },
  { text:"270,000", options:{ bold:true, color:GREEN, fill:{color:TINT}, fontSize:12, align:"right" } },
  { text:"470,000", options:{ bold:true, color:GREEN, fill:{color:TINT}, fontSize:12, align:"right" } },
]);
s.addTable(tbl, { x:0.6, y:1.7, w:7.6, colW:[4.0,1.8,1.8], border:{ pt:0.75, color:LINE },
  fontFace:F, rowH:0.42, valign:"middle", margin:0.06 });
// 右サマリー
s.addShape("roundRect", { x:8.6, y:1.7, w:4.13, h:2.3, rectRadius:0.08, fill:{ color:TINT }, line:{ color:LINE, width:1 } });
s.addText("賃貸物件", { x:8.85, y:1.92, w:3.6, h:0.3, fontFace:F, fontSize:11, bold:true, color:MUTED, margin:0 });
s.addText("月27万円", { x:8.85, y:2.24, w:3.6, h:0.7, fontFace:F, fontSize:30, bold:true, color:INK, margin:0 });
s.addText("年間換算 324万円", { x:8.85, y:3.02, w:3.6, h:0.3, fontFace:F, fontSize:10.5, color:MUTED, margin:0 });
s.addShape("roundRect", { x:8.6, y:4.2, w:4.13, h:2.3, rectRadius:0.08, fill:{ color:GREEN }, line:{ type:"none" } });
s.addText("自社物件（家賃0円）", { x:8.85, y:4.42, w:3.6, h:0.3, fontFace:F, fontSize:11, bold:true, color:"9FC7AC", margin:0 });
s.addText("月47万円", { x:8.85, y:4.74, w:3.6, h:0.7, fontFace:F, fontSize:30, bold:true, color:ACCENT, margin:0 });
s.addText("年間換算 564万円", { x:8.85, y:5.52, w:3.6, h:0.3, fontFace:F, fontSize:10.5, color:"D9E8DE", margin:0 });
foot(s, "※ 減価償却費・税を含みません。家賃は仮定坪単価による試算。目安であり売上・利益を保証するものではありません。40坪（80席）モデルは別紙参照。");
pageno(s,12);
}

// ============================================================ S13 杯数別
{
const s = p.addSlide(); bg(s, WHITE);
eyebrow(s, "12 ｜ 杯数別シミュレーション（5段階）");
title(s, "下振れも含めて、先にすべてお見せします。");
const head = ["杯数/日","月商","自社物件\n利益","利益率","回収\n(投資800万)","賃貸\n利益","利益率","回収\n(投資800万)"];
const data = [
  ["30杯","37.8万","1.6万","4.3%","—","▲18.4万","—","回収不可"],
  ["50杯","63.0万","21.8万","34.6%","約37か月","1.8万","2.9%","—"],
  ["75杯（標準）","94.5万","47.0万","49.7%","約17か月","27.0万","28.6%","約30か月"],
  ["100杯","126.0万","72.2万","57.3%","約12か月","52.2万","41.4%","約16か月"],
  ["125杯","157.5万","97.4万","61.8%","約9か月","77.4万","49.1%","約11か月"],
];
const tbl = [ head.map(h=>({ text:h, options:{ bold:true, color:WHITE, fill:{color:GREEN}, fontSize:10.5, align:"center", valign:"middle" } })) ];
data.forEach((r,ri)=>{
  const std = ri===2;
  tbl.push(r.map((c,ci)=>({ text:c, options:{
    bold: std || ci===0,
    color: c.startsWith("▲")||c==="回収不可" ? NEG : (std? GREEN : INK),
    fill: { color: std? "FCF3D4" : (ri%2? TINT2 : WHITE) },
    fontSize: std? 12 : 11, align: ci===0?"left":"right", valign:"middle" } })));
});
s.addTable(tbl, { x:0.6, y:1.8, w:12.13, colW:[1.9,1.45,1.5,1.3,1.75,1.5,1.3,1.43],
  border:{ pt:0.75, color:LINE }, fontFace:F, rowH:0.62, margin:0.06 });
s.addText("実績との接続：既存店（開業6か月以上）の月商中央値66.8万円は、約53杯/日に相当します（単価ミックス込み概算）。", {
  x:0.6, y:6.1, w:12.1, h:0.35, fontFace:F, fontSize:12, bold:true, color:GREEN2, margin:0 });
foot(s, "※ 1杯420円・原価20%・固定費28.6万円/月（清掃・水道光熱・機械・警備・ロイヤリティ・雑費／20坪モデル・杯数によらず一定と仮定）。賃貸は家賃20万円/月を加算。償却前・税引前。「—」は利益僅少で回収が非現実的な水準。収益を保証するものではありません。");
pageno(s,13);
}

// ============================================================ S14 損益分岐点
{
const s = p.addSlide(); bg(s, WHITE);
eyebrow(s, "13 ｜ 損益分岐点");
title(s, "自社物件なら1日約29杯。40席の店で「1席1杯以下」で黒字圏。");
const cups = [10,20,30,40,50,60,70,80,90,100];
const own = cups.map(c=> Math.round((c*30*336-286000)/10000*10)/10);
const rent = cups.map(c=> Math.round((c*30*336-486000)/10000*10)/10);
s.addChart(p.ChartType.line, [
  { name:"自社物件", labels:cups.map(String), values:own },
  { name:"賃貸（家賃20万円）", labels:cups.map(String), values:rent },
], { x:0.6, y:1.7, w:7.4, h:4.8,
  chartColors:[GREEN2, "C9A227"], lineSize:3, lineSmooth:false,
  showLegend:true, legendPos:"b", legendFontFace:F, legendFontSize:10, legendColor:INK,
  showTitle:false, showValue:false,
  catAxisTitle:"杯/日", catAxisTitleColor:MUTED, catAxisTitleFontSize:10, showCatAxisTitle:true,
  catAxisLabelColor:MUTED, catAxisLabelFontFace:F, catAxisLabelFontSize:10,
  valAxisLabelColor:MUTED, valAxisLabelFontFace:F, valAxisLabelFontSize:10,
  valAxisTitle:"月間損益（万円）", showValAxisTitle:true, valAxisTitleColor:MUTED, valAxisTitleFontSize:10,
  valGridLine:{ color:LINE, size:0.5 }, catGridLine:{ style:"none" } });
const bep = [
  ["自社物件（家賃0円）", "約29杯/日", "固定費28.6万円 ÷ 1杯あたり限界利益336円 ≒ 月851杯"],
  ["賃貸物件（家賃20万円）", "約49杯/日", "固定費48.6万円 ÷ 336円 ≒ 月1,447杯"],
];
bep.forEach((c,i)=>{
  const y = 1.9 + i*1.85;
  s.addShape("roundRect", { x:8.45, y, w:4.28, h:1.6, rectRadius:0.07,
    fill:{ color: i? TINT : GREEN }, line:{ color: i? LINE : GREEN, width:1 } });
  s.addText(c[0], { x:8.68, y:y+0.15, w:3.85, h:0.3, fontFace:F, fontSize:11, bold:true, color: i? MUTED:"9FC7AC", margin:0 });
  s.addText(c[1], { x:8.68, y:y+0.45, w:3.85, h:0.6, fontFace:F, fontSize:27, bold:true, color: i? GREEN:ACCENT, margin:0 });
  s.addText(c[2], { x:8.68, y:y+1.08, w:3.9, h:0.45, fontFace:F, fontSize:8.5, color: i? MUTED:"D9E8DE", margin:0 });
});
s.addText("参考：既存店（開業6か月以上）の月商中央値は約53杯/日相当", { x:8.45, y:5.75, w:4.3, h:0.6,
  fontFace:F, fontSize:10.5, bold:true, color:GREEN2, lineSpacing:15, margin:0 });
foot(s, "※ 20坪モデルの前提（1杯420円・原価20%・固定費28.6万円/月）による試算。家賃・光熱費等の条件により分岐点は変動します。");
pageno(s,14);
}

// ============================================================ S15 投資回収
{
const s = p.addSlide(); bg(s, WHITE);
eyebrow(s, "14 ｜ 投資回収シミュレーション");
title(s, "標準予算800万円。自社物件・標準杯数で約17か月。");
const months = []; for(let m=0;m<=36;m+=3) months.push(m);
const own = months.map(m=> -800 + 47*m);
const rent = months.map(m=> -800 + 27*m);
s.addChart(p.ChartType.line, [
  { name:"自社物件（月47万円）", labels:months.map(String), values:own },
  { name:"賃貸（月27万円）", labels:months.map(String), values:rent },
], { x:0.6, y:1.7, w:7.4, h:4.8,
  chartColors:[GREEN2, "C9A227"], lineSize:3, lineSmooth:false,
  showLegend:true, legendPos:"b", legendFontFace:F, legendFontSize:10, legendColor:INK,
  showTitle:false, showValue:false,
  catAxisTitle:"経過月数", showCatAxisTitle:true, catAxisTitleColor:MUTED, catAxisTitleFontSize:10,
  catAxisLabelColor:MUTED, catAxisLabelFontFace:F, catAxisLabelFontSize:10,
  valAxisTitle:"累計損益（万円）", showValAxisTitle:true, valAxisTitleColor:MUTED, valAxisTitleFontSize:10,
  valAxisLabelColor:MUTED, valAxisLabelFontFace:F, valAxisLabelFontSize:10,
  valGridLine:{ color:LINE, size:0.5 }, catGridLine:{ style:"none" } });
const rec = [
  ["自社物件", "約17か月", "75杯/日・月間利益47万円"],
  ["賃貸物件", "約30か月", "75杯/日・月間利益27万円"],
  ["100杯/日なら", "約12か月", "自社物件・月間利益72.2万円"],
];
rec.forEach((c,i)=>{
  chip(s, 8.45, 1.7+i*1.62, 4.28, 1.42, c[0], c[1], c[2], i===0);
});
foot(s, "※ 総投資800万円（開業費650万円＋諸経費・予備費150万円）。償却前営業利益ベース・税引前。回収期間・利益を保証するものではありません。");
pageno(s,15);
}

// ============================================================ S16 リスクと対策
{
const s = p.addSlide(); bg(s, WHITE);
eyebrow(s, "15 ｜ リスクと対策");
title(s, "赤字になる店には、共通のパターンがあります。");
s.addShape("roundRect", { x:0.6, y:1.75, w:4.6, h:4.55, rectRadius:0.08, fill:{ color:TINT2 }, line:{ color:LINE, width:1 } });
s.addText("事実の開示", { x:0.88, y:2.0, w:4.0, h:0.32, fontFace:F, fontSize:11, bold:true, color:MUTED, margin:0 });
s.addText("48店中 3店", { x:0.88, y:2.35, w:4.0, h:0.75, fontFace:F, fontSize:34, bold:true, color:NEG, margin:0 });
s.addText("が2026年7月単月で営業赤字（6.3%）", { x:0.88, y:3.15, w:4.0, h:0.35, fontFace:F, fontSize:11.5, color:INK, margin:0 });
s.addText("赤字の要因は主に2類型：", { x:0.88, y:3.7, w:4.0, h:0.3, fontFace:F, fontSize:11, bold:true, color:INK, margin:0 });
s.addText([
  { text:"売上規模に対して家賃が重い立地", options:{ bullet:{ indent:12 }, breakLine:true, paraSpaceAfter:8 } },
  { text:"小型区画・商業施設内で客数が伸びにくい", options:{ bullet:{ indent:12 } } },
], { x:0.88, y:4.05, w:4.1, h:1.0, fontFace:F, fontSize:11, color:INK, margin:0 });
s.addText("いずれも「家賃と売上のミスマッチ」。家賃を除いた営業利益では全店黒字です。", {
  x:0.88, y:5.3, w:4.05, h:0.85, fontFace:F, fontSize:10.5, color:GREEN2, bold:true, lineSpacing:15, margin:0 });
const meas = [
  ["出店前の立地審査", "現地調査と収益試算を実施し、基準に満たない物件には出店をお勧めしません。加盟審査は「お互いに選ぶ」プロセスです。"],
  ["家賃比率の目安を事前提示", "想定売上に対する適正家賃の目安を契約前にご提示。自社物件はこのリスク自体がありません。"],
  ["契約条件の明確化", "契約期間（最低3年〜）・中途解約・原状回復の条件は契約前にすべて書面でご確認いただきます。"],
];
meas.forEach((c,i)=>{
  const y = 1.75 + i*1.55;
  s.addShape("roundRect", { x:5.5, y, w:7.23, h:1.38, rectRadius:0.07, fill:{ color:TINT }, line:{ color:LINE, width:1 } });
  s.addText("対策 "+(i+1), { x:5.75, y:y+0.12, w:1.2, h:0.3, fontFace:F, fontSize:10, bold:true, color:GREEN2, margin:0 });
  s.addText(c[0], { x:5.75, y:y+0.38, w:6.7, h:0.34, fontFace:F, fontSize:13.5, bold:true, color:INK, margin:0 });
  s.addText(c[1], { x:5.75, y:y+0.74, w:6.75, h:0.6, fontFace:F, fontSize:10, color:MUTED, lineSpacing:14, margin:0 });
});
foot(s, "※ 赤字店数は2026年7月速報・営業利益ベース（開業6か月以上・48店）。個別店舗の状況・契約条件の詳細は個別面談でご説明します。");
pageno(s,16);
}

// ============================================================ S17 開業までの流れ
{
const s = p.addSlide(); bg(s, WHITE);
eyebrow(s, "16 ｜ 開業までの流れ");
title(s, "お問い合わせから、2〜3か月でオープン。");
const steps = [
  ["STEP 1", "面談", "制度のご説明と、ご希望条件のヒアリング"],
  ["STEP 2", "加盟審査", "立地・スペースの現地調査と収益試算"],
  ["STEP 3", "契約締結", "審査通過後に契約（契約期間 最低3年〜）"],
  ["STEP 4", "工事", "内装・機器設置。工事業者は本部が手配"],
  ["STEP 5", "運営開始", "運営研修を経てオープン"],
];
steps.forEach((c,i)=>{
  const x = 0.6 + i*2.5;
  s.addShape("roundRect", { x, y:1.9, w:2.3, h:2.5, rectRadius:0.08,
    fill:{ color: i===4? GREEN : TINT2 }, line:{ color: i===4? GREEN: LINE, width:1 } });
  s.addText(c[0], { x:x+0.16, y:2.08, w:2.0, h:0.3, fontFace:F, fontSize:9.5, bold:true,
    color: i===4? "9FC7AC" : GREEN2, charSpacing:1, margin:0 });
  s.addText(c[1], { x:x+0.16, y:2.4, w:2.0, h:0.45, fontFace:F, fontSize:17, bold:true,
    color: i===4? WHITE : INK, margin:0 });
  s.addText(c[2], { x:x+0.16, y:2.92, w:2.0, h:1.3, fontFace:F, fontSize:9.5,
    color: i===4? "D9E8DE" : MUTED, lineSpacing:14, margin:0 });
});
// 会社概要
s.addShape("roundRect", { x:0.6, y:4.9, w:12.13, h:1.75, rectRadius:0.08, fill:{ color:TINT2 }, line:{ color:LINE, width:1 } });
s.addText("会社概要", { x:0.88, y:5.08, w:3.0, h:0.32, fontFace:F, fontSize:11, bold:true, color:GREEN2, margin:0 });
const co = [
  ["会社名", "セルフカフェ株式会社"],
  ["代表取締役", "鈴木 大基"],
  ["資本金", "2,000万円"],
  ["沿革", "2022年9月 1号店開業／2024年5月 WDPグループより分社化"],
];
co.forEach((c,i)=>{
  const x = 0.88 + (i%2)*6.0, y = 5.45 + Math.floor(i/2)*0.52;
  s.addText(c[0], { x, y, w:1.5, h:0.4, fontFace:F, fontSize:10, color:MUTED, margin:0 });
  s.addText(c[1], { x:x+1.5, y, w:4.5, h:0.4, fontFace:F, fontSize:10.5, bold:true, color:INK, margin:0 });
});
foot(s, "※ 工事内容・立地条件により期間は変動します。");
pageno(s,17);
}

// ============================================================ S18 CTA
{
const s = p.addSlide(); bg(s, GREEN);
s.addImage({ path:A+"talk.jpg", x:8.6, y:0, w:4.73, h:7.5 });
s.addShape("rect", { x:8.6, y:0, w:4.73, h:7.5, fill:{ color:GREEN, transparency:80 }, line:{ type:"none" } });
s.addImage({ path:A+"logo.png", x:0.62, y:0.55, w:0.58, h:0.53 });
s.addText("NEXT STEP", { x:1.34, y:0.68, w:4, h:0.3, fontFace:F, fontSize:11, bold:true, color:"9FC7AC", charSpacing:3, margin:0 });
s.addText("まずは、お持ちの物件で\n収支診断を。", { x:0.6, y:1.6, w:7.8, h:1.9,
  fontFace:F, fontSize:38, bold:true, color:WHITE, lineSpacing:52, margin:0 });
s.addText("スペースの写真と図面をお送りいただくだけで、貴社の物件でいくら残るか、無料で試算します。", {
  x:0.62, y:3.55, w:7.4, h:0.7, fontFace:F, fontSize:14, color:"D9E8DE", lineSpacing:20, margin:0 });
const need = [["必要なもの ①","スペースの写真"],["必要なもの ②","図面（面積が分かるもの）"]];
need.forEach((c,i)=>{
  const x = 0.62 + i*3.0;
  s.addShape("roundRect", { x, y:4.45, w:2.84, h:1.0, rectRadius:0.07, fill:{ color:"0A3A20" }, line:{ color:"2C6B4A", width:0.75 } });
  s.addText(c[0], { x:x+0.16, y:4.58, w:2.5, h:0.28, fontFace:F, fontSize:9, bold:true, color:"9FC7AC", margin:0 });
  s.addText(c[1], { x:x+0.16, y:4.88, w:2.6, h:0.4, fontFace:F, fontSize:13.5, bold:true, color:WHITE, margin:0 });
});
const ct = [["TEL","090-6386-5493（担当：佐藤）"],["E-MAIL","selfcafe001@gmail.com"],["WEB","https://selfcafe.jp/"]];
ct.forEach((c,i)=>{
  const y = 5.85 + i*0.42;
  s.addText(c[0], { x:0.62, y, w:1.0, h:0.36, fontFace:F, fontSize:11, bold:true, color:ACCENT, margin:0 });
  s.addText(c[1], { x:1.7, y, w:6.0, h:0.36, fontFace:F, fontSize:12.5, color:WHITE, margin:0 });
});
s.addText("セルフカフェ株式会社", { x:11.0, y:7.05, w:2.3, h:0.3, fontFace:F, fontSize:10, color:"D9E8DE", align:"right", margin:0 });
}

p.writeFile({ fileName: "selfcafe_fc_v8_draft.pptx" }).then(()=>console.log("written"));
