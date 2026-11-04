/**
 * 2026年11月4日 未来屋書店様 事前打合せ資料 — A4横 4ページ
 */
const pptxgen = require("pptxgenjs");
const path = require("path");

const C = {
  bg: "FFFFFF", ink: "1A1A1A", green: "106838", greenDeep: "0C5230",
  gold: "8A6316", rule: "333333", line: "BFBFBF", lightLine: "DCDCDC",
  head: "EDEDED", band: "F7F7F7", tint: "EFF5F1", goldTint: "FBF5E8",
  body: "333333", muted: "6B6B6B", white: "FFFFFF",
};
const F = { jp: "Yu Gothic", num: "Arial" };
const HAS_JP = /[ぁ-んァ-ヶ一-龥々〜ー]/;
const numFace = (s) => (HAS_JP.test(String(s)) ? F.jp : F.num);
const A = path.join(__dirname, "assets");
const W = 11.69, H = 8.27, M = 0.42, CW = W - M * 2, R = M + CW;
const DOC = "未来屋書店様 事前打合せ資料";

const pres = new pptxgen();
pres.defineLayout({ name: "A4L", width: W, height: H });
pres.layout = "A4L";
pres.author = "セルフカフェ株式会社";
pres.title = "2026年11月4日 未来屋書店様 事前打合せ資料";

let PAGE = 0;
/** 2ページ目以降の簡易ヘッダー */
function head(s, title) {
  PAGE++;
  s.background = { color: C.bg };
  s.addText(DOC, {
    x: M, y: 0.34, w: 6.0, h: 0.18,
    fontFace: F.jp, fontSize: 8, color: C.muted, margin: 0, valign: "middle",
  });
  s.addText(String(PAGE), {
    x: R - 1.0, y: 0.34, w: 1.0, h: 0.18,
    fontFace: F.num, fontSize: 8, color: C.muted, align: "right", margin: 0, valign: "middle",
  });
  s.addShape("rect", { x: M, y: 0.56, w: CW, h: 0.014, fill: { color: C.rule } });
  s.addText(title, {
    x: M, y: 0.62, w: CW, h: 0.34,
    fontFace: F.jp, fontSize: 14, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
  s.addShape("rect", { x: M, y: 0.99, w: CW, h: 0.008, fill: { color: C.line } });
}
const sec = (s, y, t, w) => {
  s.addShape("rect", { x: M, y: y + 0.03, w: 0.055, h: 0.16, fill: { color: C.greenDeep } });
  s.addText(t, {
    x: M + 0.14, y, w: (w || CW) - 0.14, h: 0.22,
    fontFace: F.jp, fontSize: 10, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
};
const secAt = (s, x, y, w, t) => {
  s.addShape("rect", { x, y: y + 0.03, w: 0.055, h: 0.16, fill: { color: C.greenDeep } });
  s.addText(t, {
    x: x + 0.14, y, w: w - 0.14, h: 0.22,
    fontFace: F.jp, fontSize: 10, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
};
const th = (t, al) => ({
  text: t,
  options: { fill: { color: C.head }, color: C.ink, bold: true, fontSize: 8.5, fontFace: F.jp, align: al || "left", valign: "middle", margin: [0.02, 0.1, 0.02, 0.1] },
});
const td = (t, o = {}) => ({
  text: t,
  options: {
    fill: { color: o.fill || C.white }, color: o.color || C.body, bold: !!o.bold,
    fontSize: o.size || 8.5, fontFace: o.face || numFace(t), align: o.align || "left",
    valign: "middle", margin: [0.02, 0.1, 0.02, 0.1],
  },
});
/** 箇条書き */
function bullets(s, x, y, w, items, opt = {}) {
  const pitch = opt.pitch || 0.175, fs = opt.size || 8.5;
  items.forEach((t, i) => {
    const yy = y + i * pitch;
    const mark = typeof t === "object" ? t.m : "・";
    const txt = typeof t === "object" ? t.t : t;
    const hl = typeof t === "object" && t.hl;
    s.addText(mark, {
      x, y: yy, w: 0.22, h: pitch,
      fontFace: F.jp, fontSize: fs - 0.5, bold: !!hl, color: hl ? C.gold : C.muted, margin: 0, valign: "middle",
    });
    s.addText(txt, {
      x: x + 0.24, y: yy, w: w - 0.24, h: pitch,
      fontFace: F.jp, fontSize: fs, bold: !!hl, color: hl ? C.gold : C.body, margin: 0, valign: "middle",
    });
  });
}

/* ========================================================= P1 アジェンダ */
{
  const s = pres.addSlide();
  PAGE++;
  s.background = { color: C.bg };
  s.addText("2026年11月4日", {
    x: R - 3.0, y: 0.30, w: 3.0, h: 0.18,
    fontFace: F.jp, fontSize: 9, color: C.body, align: "right", margin: 0, valign: "middle",
  });
  s.addText("株式会社未来屋書店　御中", {
    x: M, y: 0.50, w: 5.0, h: 0.24,
    fontFace: F.jp, fontSize: 11.5, color: C.ink, margin: 0, valign: "middle",
  });
  s.addImage({ path: path.join(A, "logo-green.png"), x: R - 1.45, y: 0.48, w: 1.45, h: 0.207 });
  s.addText("セルフカフェ株式会社", {
    x: R - 4.6, y: 0.70, w: 4.6, h: 0.18,
    fontFace: F.jp, fontSize: 9.5, color: C.ink, align: "right", margin: 0, valign: "middle",
  });
  s.addShape("rect", { x: M, y: 1.06, w: CW, h: 0.016, fill: { color: C.rule } });
  s.addText("事前打合せ　アジェンダおよび論点整理", {
    x: M, y: 1.11, w: CW, h: 0.32,
    fontFace: F.jp, fontSize: 15, bold: true, color: C.ink, align: "center", margin: 0, valign: "middle",
  });
  s.addShape("rect", { x: M, y: 1.46, w: CW, h: 0.008, fill: { color: C.line } });

  sec(s, 1.72, "本日のアジェンダ");
  const ag = [
    ["１", "秋田店 セルフカフェ候補領域の認識共有", "レイアウト作成に伴うご相談ポイントの事前共有", "P2"],
    ["２", "既存カフェ区画について", "工事着手にあたりカフェ・ド・クリエ様へご確認が必要な情報", "P3"],
    ["３", "仙台新規出店の売上試算について", "売上試算に向けた確認事項の共有（試算は本日時点では未着手）", "P4"],
    ["４", "今後の進め方", "秋田店の工程確認と、仙台の進め方", "P5"],
  ];
  s.addTable(
    [
      [th("", "center"), th("項目"), th("内容"), th("頁", "center")],
      ...ag.map((a) => [
        td(a[0], { align: "center", bold: true, color: C.green, size: 10 }),
        td(a[1], { bold: true, color: C.ink, size: 10 }),
        td(a[2], { size: 9 }),
        td(a[3], { align: "center", color: C.muted, size: 8.5 }),
      ]),
    ],
    { x: M, y: 1.96, w: CW, colW: [0.5, 3.9, 5.85, 0.6], rowH: [0.22, 0.4, 0.4, 0.4, 0.4], border: { type: "solid", color: C.line, pt: 0.75 }, autoPage: false }
  );

  sec(s, 3.90, "本日の到達目標");
  const goals = [
    ["秋田", "セルフカフェ候補領域（30坪の切り出し位置）の認識をそろえ、レイアウト作成の前提を確定する。"],
    ["秋田", "カフェ・ド・クリエ様へご確認いただく事項を確定し、解約協議にのせていただく。"],
    ["仙台", "売上試算に必要な情報を本日で取り切る。特に「来客数の定義」をそろえる（P4参照）。"],
    ["共通", "11月上旬の現地内見（弊社施工担当同行）の日程を決める。"],
  ];
  goals.forEach((g, i) => {
    const y = 4.18 + i * 0.32;
    s.addShape("rect", { x: M, y: y + 0.03, w: 0.62, h: 0.22, fill: { color: i === 3 ? C.head : C.tint } });
    s.addText(g[0], {
      x: M, y: y + 0.03, w: 0.62, h: 0.22,
      fontFace: F.jp, fontSize: 8.5, bold: true, color: i === 3 ? C.ink : C.green, align: "center", margin: 0, valign: "middle",
    });
    s.addText(g[1], {
      x: M + 0.76, y, w: CW - 0.76, h: 0.28,
      fontFace: F.jp, fontSize: 9, color: C.body, margin: 0, valign: "middle",
    });
  });

  sec(s, 5.62, "現在の合意事項（秋田店）");
  s.addTable(
    [
      [th("区画", "center"), th("契約形態", "center"), th("貴社ご負担の初期費用", "center"), th("開業予定", "center"), th("閉店後の工期", "center")],
      [
        td("30坪", { align: "center", bold: true, color: C.ink, size: 9.5 }),
        td("業務委託型・3年", { align: "center", bold: true, color: C.ink, size: 9.5 }),
        td("30〜50万円（税抜・設置工事費）", { align: "center", bold: true, color: C.ink, size: 9.5 }),
        td("2027年3月上旬", { align: "center", bold: true, color: C.green, size: 9.5 }),
        td("約1か月（残置前提）", { align: "center", bold: true, color: C.ink, size: 9.5 }),
      ],
    ],
    { x: M, y: 5.86, w: CW, colW: [1.7, 2.3, 2.75, 2.1, 2.0], rowH: [0.22, 0.28], border: { type: "solid", color: C.line, pt: 0.75 }, autoPage: false }
  );

  s.addShape("rect", { x: M, y: 6.60, w: CW, h: 0.008, fill: { color: C.lightLine } });
  s.addText("本資料は、2026年10月17日付「秋田店 進行スケジュールおよびご確認事項」および、ご提供いただいた仙台泉物件の計画概要・ヒアリングシートに基づき作成しております。", {
    x: M, y: 6.70, w: CW, h: 0.16,
    fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
  });
}

/* ========================================================= P2 秋田 レイアウト */
{
  const s = pres.addSlide();
  head(s, "１．秋田店 セルフカフェ候補領域の認識共有");

  const lw = 4.3, rx = M + lw + 0.35, rw = R - rx;
  secAt(s, M, 1.16, lw, "レイアウト作成に先立ち必要となる情報");
  bullets(s, M, 1.42, lw, [
    "区画図面（平面図・展開図／CADデータ）",
    "既存カフェ区画の実測寸法",
    "電気容量および分電盤の位置",
    "給排水の位置・系統",
    "排気ダクト・空調の系統",
    "現況写真（客席・カウンター・厨房まわり）",
    "天井高および床仕上げの納まり",
    "館の内装制限・工事区分",
  ], { pitch: 0.19, size: 8.5 });

  s.addShape("roundRect", { x: M, y: 3.02, w: lw, h: 0.72, rectRadius: 0.04, fill: { color: C.band }, line: { color: C.lightLine, width: 0.75 } });
  s.addText("いずれも現地内見（11月上旬）までにお手元にあるものをご共有いただければ、内見当日の確認項目を絞り込むことができ、レイアウトの精度が上がります。", {
    x: M + 0.18, y: 3.08, w: lw - 0.36, h: 0.6,
    fontFace: F.jp, fontSize: 8, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
  });

  secAt(s, M, 3.92, lw, "候補領域の認識共有（本日のご相談）");
  bullets(s, M, 4.18, lw, [
    { m: "①", t: "30坪をどこから切り出すか", hl: true },
    "　既存カフェ区画をそのまま使うか、書架側へ広げるか",
    { m: "②", t: "書店売場との境界の作り方", hl: true },
    "　什器で仕切るか、床仕上げのみで区切るか",
    { m: "③", t: "店頭（モール通路側）の見え方", hl: true },
    "　サインの位置と、カフェを前面に出すか書店を前面にするか",
  ], { pitch: 0.2, size: 8.5 });

  s.addShape("rect", { x: M, y: 5.52, w: lw, h: 0.008, fill: { color: C.lightLine } });
  s.addText("※ 席数は30坪・約59席（土浦店と同規模）を目安としております。配置確定後にあらためてご提示いたします。", {
    x: M, y: 5.62, w: lw, h: 0.32,
    fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
  });

  secAt(s, rx, 1.16, rw, "レイアウト作成に伴うご相談ポイント");
  const pts = [
    ["席数と席種の構成", "カウンター席／4人掛け／ソファ席の比率。30坪で約55〜60席を想定しております。"],
    ["マシン・決済端末の位置", "電源および給排水の位置に制約されます。壁面か、島配置かをご相談させてください。"],
    ["書籍の持ち込み運用", "ご購入前の書籍・雑誌の持ち込みを可とするか。可の場合、動線と汚損対策の整理が必要です。"],
    ["レジとの関係", "書店レジとカフェ区画の距離。お客様の動線と、スタッフの目が届く範囲。"],
    ["サイン表記", "「MIRAIYA Book & Cafe」の最終表記と、館の看板規定への適合。"],
    ["既存什器の再利用範囲", "テーブル・椅子・造作のうち、再利用可能なものの選別。"],
    ["バックヤード", "原料・備品のストック場所、ゴミの一時保管場所の確保。"],
    ["清掃・補充の動線", "1日15分の作業を、書店業務の流れの中でどこに組み込むか。"],
  ];
  pts.forEach((p, i) => {
    const y = 1.42 + i * 0.52;
    s.addShape("rect", { x: rx, y, w: rw, h: 0.48, fill: { color: i % 2 ? C.band : C.white }, line: { color: C.lightLine, width: 0.5 } });
    s.addText((i + 1) + "．" + p[0], {
      x: rx + 0.14, y: y + 0.03, w: rw - 0.28, h: 0.2,
      fontFace: F.jp, fontSize: 8.5, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(p[1], {
      x: rx + 0.14, y: y + 0.23, w: rw - 0.28, h: 0.22,
      fontFace: F.jp, fontSize: 8, color: C.body, margin: 0, valign: "middle",
    });
  });

  s.addShape("rect", { x: M, y: 7.12, w: CW, h: 0.008, fill: { color: C.lightLine } });
  s.addText("※ レイアウト図および什器リストは、現地内見後、11月中旬のご提出を予定しております。", {
    x: M, y: 7.22, w: CW, h: 0.16,
    fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
  });
}

/* ========================================================= P3 既存カフェ区画 */
{
  const s = pres.addSlide();
  head(s, "２．既存カフェ区画について");

  const lw = 6.4, rx = M + lw + 0.35, rw = R - rx;
  secAt(s, M, 1.16, lw, "カフェ・ド・クリエ様へご確認をお願いしたい事項");
  s.addText("弊社が工事に取り掛かるにあたり、解約協議の中であわせてご確認いただけますと、工期と費用の精度が上がります。", {
    x: M, y: 1.40, w: lw, h: 0.18,
    fontFace: F.jp, fontSize: 8, color: C.body, margin: 0, valign: "middle",
  });
  const chk = [
    ["★", "原状回復の範囲", "どこまで残置いただけるか。工期と費用に最も影響します。"],
    ["★", "給排水の残置", "経路・口径・止水栓の位置。撤去される場合は新設が必要です。"],
    ["★", "電気容量・分電盤の残置", "容量と回路数。カフェ機器の消費電力に足りるかを確認いたします。"],
    ["", "造作の残置可否", "カウンター・間仕切り・天井・照明・床仕上げ。"],
    ["", "厨房設備の撤去範囲", "撤去後の床・壁の状態（下地が露出するか、仕上げまで復旧されるか）。"],
    ["", "排気ダクト・グリストラップ", "残置か撤去か。残置の場合は閉塞処理の要否。"],
    ["", "看板・サインの撤去範囲", "下地・電源の残置可否。弊社サインの取り付けに影響します。"],
    ["", "閉店日と退去完了日", "弊社の工事着手可能日となります。"],
    ["", "館への原状回復届", "提出の有無と内容。弊社の工事申請との整合が必要です。"],
    ["", "引き渡し時の立会い", "弊社施工担当を同席させていただけますと、認識の齟齬を防げます。"],
  ];
  chk.forEach((c, i) => {
    const y = 1.66 + i * 0.50;
    const hl = c[0] === "★";
    s.addShape("rect", { x: M, y, w: lw, h: 0.46, fill: { color: hl ? C.goldTint : (i % 2 ? C.band : C.white) }, line: { color: C.lightLine, width: 0.5 } });
    s.addText(hl ? "★" : String(i + 1), {
      x: M + 0.06, y: y + 0.04, w: 0.3, h: 0.19,
      fontFace: F.jp, fontSize: 8, bold: true, color: hl ? C.gold : C.muted, align: "center", margin: 0, valign: "middle",
    });
    s.addText(c[1], {
      x: M + 0.42, y: y + 0.04, w: lw - 0.56, h: 0.19,
      fontFace: F.jp, fontSize: 9, bold: true, color: hl ? C.gold : C.ink, margin: 0, valign: "middle",
    });
    s.addText(c[2], {
      x: M + 0.42, y: y + 0.24, w: lw - 0.56, h: 0.19,
      fontFace: F.jp, fontSize: 8.5, color: C.body, margin: 0, valign: "middle",
    });
  });
  s.addText("★印の3点は、工期と工事費に直結いたします。優先してご確認いただけますと幸いです。", {
    x: M, y: 6.74, w: lw, h: 0.18,
    fontFace: F.jp, fontSize: 8.5, bold: true, color: C.gold, margin: 0, valign: "middle",
  });

  secAt(s, rx, 1.16, rw, "残置の範囲が、工期と費用を決めます");
  s.addTable(
    [
      [th("閉店後の工程"), th("期間", "center")],
      [td("残置物の撤去・下地調整", { size: 8.5 }), td("3〜5日", { align: "center", size: 8.5 })],
      [td("内装工事（造作・床・サイン）", { size: 8.5 }), td("2〜3週間", { align: "center", size: 8.5 })],
      [td("機器設置・試運転・什器搬入", { size: 8.5 }), td("3〜5日", { align: "center", size: 8.5 })],
      [td("合計", { bold: true, color: C.ink, size: 9, fill: C.tint }), td("約1か月", { align: "center", bold: true, color: C.green, size: 9, fill: C.tint })],
    ],
    { x: rx, y: 1.42, w: rw, colW: [2.5, 1.24], rowH: [0.2, 0.26, 0.26, 0.26, 0.3], border: { type: "solid", color: C.line, pt: 0.75 }, autoPage: false }
  );

  s.addShape("roundRect", { x: rx, y: 3.02, w: rw, h: 1.08, rectRadius: 0.04, fill: { color: C.goldTint }, line: { color: "E2CFA6", width: 0.75 } });
  s.addText("スケルトン戻しとなる場合", {
    x: rx + 0.16, y: 3.10, w: rw - 0.32, h: 0.2,
    fontFace: F.jp, fontSize: 8.5, bold: true, color: C.gold, margin: 0, valign: "middle",
  });
  s.addText("下地からの復旧が必要となり、工期が2週間程度延びる見込みです。2027年3月上旬の開業が厳しくなります。", {
    x: rx + 0.16, y: 3.32, w: rw - 0.32, h: 0.7,
    fontFace: F.jp, fontSize: 8, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
  });

  s.addShape("roundRect", { x: rx, y: 4.26, w: rw, h: 1.26, rectRadius: 0.04, fill: { color: C.band }, line: { color: C.lightLine, width: 0.75 } });
  s.addText("ご相談のお願い", {
    x: rx + 0.16, y: 4.34, w: rw - 0.32, h: 0.2,
    fontFace: F.jp, fontSize: 8.5, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
  s.addText("カフェ・ド・クリエ様との解約協議の中で、造作・床・給排水・電気を可能な限り現状のまま残していただく方向でご調整いただけますと幸いです。工期の短縮と工事費の抑制の両方につながります。", {
    x: rx + 0.16, y: 4.56, w: rw - 0.32, h: 0.88,
    fontFace: F.jp, fontSize: 8, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
  });

  s.addShape("rect", { x: M, y: 7.12, w: CW, h: 0.008, fill: { color: C.lightLine } });
  s.addText("※ 工期は目安です。館（イオンモール様）の工事申請に要する期間は含んでおりません。", {
    x: M, y: 7.22, w: CW, h: 0.16,
    fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
  });
}

/* ========================================================= P4 仙台 */
{
  const s = pres.addSlide();
  head(s, "３．仙台新規出店について（確認事項）");

  const lw = 4.6, rx = M + lw + 0.35, rw = R - rx;
  secAt(s, M, 1.16, lw, "ご提供いただいた物件情報");
  const info = [
    ["商業施設", "旧アリオ仙台泉（SENDAI IZUMI Reborn PROJECT）"],
    ["所在地", "仙台市泉区泉中央1-5-1／泉中央駅 徒歩2分"],
    ["開業予定", "2028年度内予定"],
    ["館の規模", "総賃貸面積 約20,000㎡／専門店50〜70店舗"],
    ["商業集積", "SELVA等と合わせ約40,000㎡・約180店舗"],
    ["交通", "地下鉄・バス 1日約3.1万人／年間約1,130万人"],
    ["書店フロア", "4F"],
    ["書店の坪数", "446坪（新店全体）"],
    ["書籍売場来店客数", "13,000人／月（想定）"],
    ["セルフカフェ坪数", "30坪（60坪の試算もご希望）"],
    ["水回り", "なし"],
    ["電源・回線", "新規物件のため工事対応"],
    ["営業時間", "書店・カフェとも未定"],
  ];
  s.addTable(
    [
      [th("項目"), th("内容")],
      ...info.map((r) => [td(r[0], { bold: true, color: C.ink, size: 8 }), td(r[1], { size: 8 })]),
    ],
    {
      x: M, y: 1.40, w: lw, colW: [1.45, 3.15],
      rowH: [0.2, ...info.map(() => 0.235)],
      border: { type: "solid", color: C.line, pt: 0.75 }, autoPage: false,
    }
  );

  secAt(s, rx, 1.16, rw, "本日ご確認させていただきたい事項");
  s.addShape("roundRect", { x: rx, y: 1.42, w: rw, h: 1.72, rectRadius: 0.04, fill: { color: C.goldTint }, line: { color: "E2CFA6", width: 0.75 } });
  s.addText("【最重要】来客数の定義をそろえさせてください", {
    x: rx + 0.18, y: 1.50, w: rw - 0.36, h: 0.2,
    fontFace: F.jp, fontSize: 9.5, bold: true, color: C.gold, margin: 0, valign: "middle",
  });
  s.addText("秋田店・土浦店の試算は「カフェ区画の来客数」（秋田3,760人／月）を分母としております。一方、今回の13,000人／月は「書籍売場への来店客数」であり、分母が異なります。同じ購入率を当てはめると過大な試算となってしまいます。\n\nつきましては、13,000人／月の算出方法（書店への入店客数か、レジ通過客数か等）をご教示いただくとともに、同じ定義での秋田店・土浦店の数値をあわせてご提供いただけませんでしょうか。既存店の販売実績と突き合わせることで、仙台の販売杯数に根拠を持たせることができます。", {
    x: rx + 0.18, y: 1.72, w: rw - 0.36, h: 1.34,
    fontFace: F.jp, fontSize: 8, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
  });

  const grp = (y, t) => {
    s.addText(t, {
      x: rx, y, w: rw, h: 0.2,
      fontFace: F.jp, fontSize: 9, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addShape("rect", { x: rx, y: y + 0.21, w: rw, h: 0.008, fill: { color: C.lightLine } });
  };

  grp(3.26, "Ａ．売上試算の前提に関する事項");
  bullets(s, rx, 3.50, rw, [
    { m: "★", t: "13,000人／月の算出方法と、同じ定義での秋田店・土浦店の数値", hl: true },
    "館全体の想定来客数（年間・月間）",
    "4Fのフロア構成と、核となるテナント",
    "カフェ区画の想定位置（通路・エスカレーターとの関係）",
    "書店の想定営業時間／カフェの希望営業時間",
    "想定席数（30坪・60坪それぞれ）",
    "書籍の持ち込み可否など、運用のお考え",
  ], { pitch: 0.185, size: 8.5 });

  grp(4.92, "Ｂ．工事・設備に関する事項");
  bullets(s, rx, 5.16, rw, [
    { m: "★", t: "給排水の新設可否と引込ルート（水回りなしのため）", hl: true },
    { m: "★", t: "電源容量および回線の引込条件・申請先", hl: true },
    "館側の内装制限・工事区分（Ａ工事／Ｂ工事／Ｃ工事）",
    "什器・サインに関する館の規定",
  ], { pitch: 0.185, size: 8.5 });

  grp(6.00, "Ｃ．出店条件・スケジュールに関する事項");
  bullets(s, rx, 6.24, rw, [
    "開業時期の確度（2028年度内のいつ頃か）",
    "出店条件（賃料・共益費等）の概算",
    "ご想定の契約形態（業務委託型／FC）",
    "貴社内のご検討スケジュール",
  ], { pitch: 0.185, size: 8.5 });

  s.addShape("rect", { x: M, y: 7.08, w: CW, h: 0.008, fill: { color: C.lightLine } });
  s.addText("★印の3点は、売上試算の根拠および工事費・出店可否そのものに影響いたします。本日時点で未定の場合は、館のご担当者へのご確認をお願いできますと幸いです。", {
    x: M, y: 7.18, w: CW, h: 0.16,
    fontFace: F.jp, fontSize: 7.5, color: C.gold, margin: 0, valign: "middle",
  });
  s.addText("※ 売上試算は、上記の確認事項が揃いしだい、30坪・60坪の2案で作成のうえあらためてご提出いたします。本日時点では未着手でございます。", {
    x: M, y: 7.34, w: CW, h: 0.16,
    fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
  });
  s.addText("※ 書店併設型の既存店の販売実績（2026年5〜8月平均）：1日40〜45杯程度。商業施設内の大型店は1日約123杯。", {
    x: M, y: 7.50, w: CW, h: 0.16,
    fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
  });
}

/* ========================================================= P5 今後の進め方 */
{
  const s = pres.addSlide();
  head(s, "４．今後の進め方");

  const lw = 6.4, rx = M + lw + 0.35, rw = R - rx;
  secAt(s, M, 1.16, lw, "秋田店の工程");
  const sch = [
    ["11月上旬", "現地内見（弊社施工担当同行）", "両社"],
    ["11月中旬", "レイアウト図・什器リストのご提出", "弊社"],
    ["11月中旬", "解約条件・原状回復範囲の確定", "貴社"],
    ["11月下旬", "仕様確定・業務委託契約の締結（カメラ映像の取扱いを含む）", "両社"],
    ["11月下旬〜12月", "館（イオンモール様）への工事申請", "両社"],
    ["12月上旬", "インターネット回線の新規お申込み（開通まで1〜2か月）", "貴社"],
    ["12月上旬", "什器・機器の発注（リードタイム4〜6週間）", "弊社"],
    ["1月末", "カフェ・ド・クリエ様 閉店・退去（造作・床・設備は残置）", "貴社"],
    ["2月（約1か月）", "内装工事、機器設置・試運転、什器搬入、回線開通の確認", "弊社"],
    ["2月下旬", "開業準備・研修", "両社"],
    ["3月上旬", "開業", "—"],
  ];
  const who = (t) => {
    const col = t === "弊社" ? C.green : t === "貴社" ? C.gold : C.muted;
    return td(t, { align: "center", bold: true, color: col, size: 8.5 });
  };
  s.addTable(
    [
      [th("時期"), th("内容"), th("担当", "center")],
      ...sch.map((r, i) => {
        const last = i === sch.length - 1;
        const f = last ? C.tint : (i % 2 ? C.band : C.white);
        return [
          td(r[0], { bold: true, color: last ? C.green : C.ink, size: 8.5, fill: f }),
          td(r[1], { size: 8.5, fill: f, bold: last, color: last ? C.green : C.body }),
          last ? td("—", { align: "center", color: C.muted, fill: f, size: 8.5 }) : who(r[2]),
        ];
      }),
    ],
    {
      x: M, y: 1.40, w: lw, colW: [1.35, 4.1, 0.95],
      rowH: [0.22, ...sch.map(() => 0.27)],
      border: { type: "solid", color: C.line, pt: 0.75 }, autoPage: false,
    }
  );

  secAt(s, M, 4.80, lw, "仙台の進め方");
  bullets(s, M, 5.06, lw, [
    "本日　物件情報・確認事項のすり合わせ",
    "後日　確認事項が揃いしだい、30坪・60坪の2案で売上試算をご提出",
    "その後　出店条件のご提示・ご協議",
  ], { pitch: 0.21, size: 8.5 });

  secAt(s, rx, 1.16, rw, "本日お決めいただきたいこと");
  const dec = [
    ["①", "11月上旬の現地内見の日程", "2〜3候補をいただけますと幸いです。"],
    ["②", "秋田店 30坪の候補領域の方向性", "既存区画のままか、書架側へ広げるか。"],
    ["③", "クリエ様へご確認いただく事項", "P3の10項目のうち、協議にのせていただく範囲。"],
    ["④", "仙台の確認事項の回答期限の目安", "いつ頃までにご回答いただけるか。"],
  ];
  dec.forEach((d, i) => {
    const y = 1.42 + i * 0.66;
    s.addShape("rect", { x: rx, y, w: rw, h: 0.6, fill: { color: i % 2 ? C.band : C.white }, line: { color: C.lightLine, width: 0.5 } });
    s.addText(d[0], {
      x: rx + 0.1, y: y + 0.06, w: 0.3, h: 0.2,
      fontFace: F.jp, fontSize: 9, bold: true, color: C.green, margin: 0, valign: "middle",
    });
    s.addText(d[1], {
      x: rx + 0.42, y: y + 0.06, w: rw - 0.56, h: 0.2,
      fontFace: F.jp, fontSize: 9, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(d[2], {
      x: rx + 0.42, y: y + 0.28, w: rw - 0.56, h: 0.26,
      fontFace: F.jp, fontSize: 8, color: C.body, margin: 0, valign: "middle",
    });
  });

  s.addShape("roundRect", { x: rx, y: 4.18, w: rw, h: 1.34, rectRadius: 0.04, fill: { color: C.goldTint }, line: { color: "E2CFA6", width: 0.75 } });
  s.addText("直近で動かしていただきたい2件", {
    x: rx + 0.16, y: 4.26, w: rw - 0.32, h: 0.2,
    fontFace: F.jp, fontSize: 8.5, bold: true, color: C.gold, margin: 0, valign: "middle",
  });
  s.addText("・インターネット回線のお申込み（12月中）\n　開通まで1〜2か月を要します。\n・カフェ・ド・クリエ様との解約協議（残置の範囲）\n　工期と工事費に直結いたします。", {
    x: rx + 0.16, y: 4.48, w: rw - 0.32, h: 0.96,
    fontFace: F.jp, fontSize: 8, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
  });

  s.addShape("rect", { x: M, y: 6.34, w: CW, h: 0.008, fill: { color: C.lightLine } });
  [
    "※1　工期約1か月は、造作・床・給排水・電気を残置いただくことを前提としております。スケルトン戻しとなる場合は工期が延長いたします。",
    "※2　業務委託型のため、工事費のお見積りのご提出はございません。貴社にご負担いただく費用は、設置工事費（30〜50万円・税抜）のみでございます。",
    "※3　インターネット回線はお申込みから開通まで1〜2か月を要します。12月中のお手続きをお願いいたします。",
    "※4　防犯カメラの映像（個人情報）の取扱いにつきましては、保存期間・閲覧権限等の取り決め案を弊社よりご提示いたします。",
    "※5　館（イオンモール様）の工事申請に要する期間は本工程に織り込んでおりません。判明次第あらためて調整いたします。",
  ].forEach((t, i) => {
    s.addText(t, {
      x: M, y: 6.44 + i * 0.145, w: CW, h: 0.14,
      fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
    });
  });
  s.addText("セルフカフェ株式会社　052-879-3557 ／ info@selfcafe.jp", {
    x: R - 4.0, y: 7.26, w: 4.0, h: 0.15,
    fontFace: F.jp, fontSize: 7.5, color: C.muted, align: "right", margin: 0, valign: "middle",
  });
}

pres.writeFile({ fileName: path.join(__dirname, "mt-20261104.pptx") }).then((f) => console.log("written:", f));
