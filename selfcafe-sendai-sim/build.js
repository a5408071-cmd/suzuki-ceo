/**
 * 未来屋書店 仙台泉（新規出店）セルフカフェ 売上試算 ─ 業務委託型
 * A4横 3ページ／白地・単色・見出し明朝の報告書体裁
 */
const pptxgen = require("pptxgenjs");
const path = require("path");

const C = {
  bg: "FFFFFF",
  ink: "1A1A1A",
  body: "333333",
  muted: "767676",
  ruleS: "1A1A1A",   // 太罫
  ruleT: "808080",   // 表罫
  hair: "C8C8C8",
  headFill: "E4E4E4",
  band: "F2F2F2",
  bandD: "E9F1EC",
  barD: "106838",
  gr: "106838",
  grPale: "E9F1EC",
  barL: "BFBFBF",
  white: "FFFFFF",
};
const FM = "Yu Mincho";   // 見出し
const FG = "Yu Gothic";   // 本文・表・数字

const W = 11.69, H = 8.27, M = 0.58, CW = W - M * 2, R = M + CW;

const pres = new pptxgen();
pres.defineLayout({ name: "A4L", width: W, height: H });
pres.layout = "A4L";
pres.author = "セルフカフェ株式会社";
pres.title = "仙台泉 セルフカフェ 売上試算（業務委託型）";

let PAGE = 0;
const TOTAL = 3;

/* ---------------- 体裁 ---------------- */
function head(s, title, sub) {
  PAGE++;
  s.background = { color: C.bg };
  s.addText(title, {
    x: M, y: 0.36, w: CW - 3.3, h: 0.36,
    fontFace: FM, fontSize: 17, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
  if (sub)
    s.addText(sub, {
      x: M, y: 0.72, w: CW - 3.3, h: 0.2,
      fontFace: FG, fontSize: 8.5, color: C.body, margin: 0, valign: "middle",
    });
  s.addText("未来屋書店株式会社　御中", {
    x: R - 3.3, y: 0.34, w: 3.3, h: 0.2,
    fontFace: FG, fontSize: 9, color: C.ink, align: "right", margin: 0, valign: "middle",
  });
  s.addText("2026年11月20日　セルフカフェ株式会社", {
    x: R - 3.3, y: 0.54, w: 3.3, h: 0.18,
    fontFace: FG, fontSize: 8, color: C.muted, align: "right", margin: 0, valign: "middle",
  });
  s.addText("― " + PAGE + " / " + TOTAL + " ―", {
    x: R - 3.3, y: 0.73, w: 3.3, h: 0.18,
    fontFace: FG, fontSize: 8, color: C.muted, align: "right", margin: 0, valign: "middle",
  });
  s.addShape("rect", { x: M, y: 0.98, w: CW, h: 0.019, fill: { color: C.gr } });
  s.addShape("rect", { x: M, y: 1.025, w: CW, h: 0.006, fill: { color: C.ruleS } });
}
const foot = (s, notes) => {
  s.addShape("rect", { x: M, y: H - 0.78, w: CW, h: 0.006, fill: { color: C.hair } });
  (notes || []).forEach((t, i) =>
    s.addText(t, {
      x: M, y: H - 0.70 + i * 0.15, w: CW - 2.6, h: 0.15,
      fontFace: FG, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
    })
  );
  s.addText("セルフカフェ株式会社", {
    x: R - 2.4, y: H - 0.3, w: 2.4, h: 0.16,
    fontFace: FG, fontSize: 7.5, color: C.muted, align: "right", margin: 0, valign: "middle",
  });
};
let SEC = 0;
const sec = (s, x, y, w, t, note) => {
  SEC++;
  s.addText(SEC + "．" + t, {
    x, y, w, h: 0.24,
    fontFace: FM, fontSize: 11.5, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
  if (note)
    s.addText(note, {
      x: x + w - 3.6, y, w: 3.6, h: 0.24,
      fontFace: FG, fontSize: 8, color: C.muted, align: "right", margin: 0, valign: "middle",
    });
  s.addShape("rect", { x, y: y + 0.25, w, h: 0.008, fill: { color: C.gr } });
};

const th = (t, al) => ({
  text: t,
  options: { fill: { color: C.grPale }, color: C.gr, bold: true, fontSize: 9, fontFace: FG, align: al || "left", valign: "middle", margin: [0.03, 0.1, 0.03, 0.1] },
});
const td = (t, o = {}) => ({
  text: t,
  options: {
    fill: { color: o.fill || C.white }, color: o.color || C.body, bold: !!o.bold,
    fontSize: o.size || 9, fontFace: FG, align: o.align || "left",
    valign: "middle", margin: [0.03, 0.1, 0.03, 0.1],
  },
});
const TB = { type: "solid", color: C.ruleT, pt: 0.5 };

/* ---------------- 棒グラフ（単色） ---------------- */
function bars(s, o) {
  const { x, y, w, h, data, max, unit } = o;
  const n = data.length, gap = o.gap === undefined ? 0.16 : o.gap;
  const bw = (w - gap * (n - 1)) / n, base = y + h;
  data.forEach((d, i) => {
    const bx = x + i * (bw + gap), bh = Math.max(0.03, (d.v / max) * h);
    s.addShape("rect", { x: bx, y: base - bh, w: bw, h: bh, fill: { color: d.color || C.barL } });
    s.addText(d.t, {
      x: bx - 0.25, y: base - bh - 0.25, w: bw + 0.5, h: 0.22,
      fontFace: FG, fontSize: d.fs || 9, bold: true, color: C.ink, align: "center", margin: 0, valign: "middle",
    });
    s.addText(d.label, {
      x: bx - 0.25, y: base + 0.06, w: bw + 0.5, h: 0.2,
      fontFace: FG, fontSize: 8, color: C.body, align: "center", margin: 0, valign: "middle",
    });
    if (d.sub)
      s.addText(d.sub, {
        x: bx - 0.25, y: base + 0.24, w: bw + 0.5, h: 0.18,
        fontFace: FG, fontSize: 7, color: C.muted, align: "center", margin: 0, valign: "middle",
      });
  });
  s.addShape("rect", { x, y: base, w, h: 0.008, fill: { color: C.ink } });
  if (unit)
    s.addText(unit, {
      x, y: y - 0.22, w: 2.2, h: 0.18,
      fontFace: FG, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
    });
}

/* ---------------- 試算データ ---------------- */
const UP = 420, DAYS = 30, FIX = 50000;
const CASE = [
  { t: "30坪", pos: "店頭配置", cups: 86, hot: true },
  { t: "30坪", pos: "店奥配置", cups: 60 },
  { t: "60坪", pos: "店頭配置", cups: 98, ref: true },
  { t: "60坪", pos: "店奥配置", cups: 69, ref: true },
].map((c) => {
  const sales = c.cups * UP * DAYS, roy = Math.floor(sales * 0.25);
  return Object.assign({}, c, { month: c.cups * DAYS, sales, roy, fee: roy + FIX, year: (roy + FIX) * 12 });
});
const yen = (v) => v.toLocaleString();

/* ================================================= P1 */
{
  const s = pres.addSlide();
  head(s, "仙台泉（新規出店）　セルフカフェ 売上試算", "業務委託型／館の来客数は秋田店の1.3倍を前提　開業 2028年度内予定");

  // ご留意（枠線のみ）
  s.addShape("rect", { x: M, y: 1.14, w: CW, h: 0.74, fill: { color: C.white }, line: { color: C.ruleS, width: 1 } });
  s.addText("【ご留意】　本資料は、開業前の時点における机上の試算でございます。", {
    x: M + 0.18, y: 1.20, w: CW - 0.36, h: 0.22,
    fontFace: FM, fontSize: 10.5, bold: true, color: C.gr, margin: 0, valign: "middle",
  });
  s.addText("現地での営業実績はございません。ご提示いただいた前提数値から、下記の計算式のとおり機械的に算出したものでございます。館の来客数は計画値であり、区画の配置・営業時間はいずれも未確定でございます。前提が変われば結果も変わりますので、あくまでご検討の出発点としてお取り扱いくださいますようお願い申し上げます。", {
    x: M + 0.18, y: 1.43, w: CW - 0.36, h: 0.40,
    fontFace: FG, fontSize: 8.5, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.2,
  });

  // 計算式
  sec(s, M, 2.02, CW, "計算式", "30坪・店頭配置の場合");
  s.addShape("rect", { x: M, y: 2.36, w: CW, h: 0.82, fill: { color: C.band } });
  [
    ["（１）", "販売杯数（1日）　×　単価 420円　×　30日　＝　セルフカフェ売上（月）", "86杯　×　420円　×　30日　＝　1,083,600円"],
    ["（２）", "セルフカフェ売上　×　25％　＋　月額固定 50,000円　＝　貴社お受取額（月）", "1,083,600円　×　25％　＋　50,000円　＝　320,900円"],
  ].forEach((e, i) => {
    const y = 2.46 + i * 0.32;
    s.addText(e[0], { x: M + 0.18, y, w: 0.56, h: 0.26, fontFace: FG, fontSize: 9, color: C.ink, margin: 0, valign: "middle" });
    s.addText(e[1], { x: M + 0.74, y, w: 5.2, h: 0.26, fontFace: FG, fontSize: 9.5, color: C.ink, margin: 0, valign: "middle" });
    s.addText(e[2], { x: M + 6.1, y, w: CW - 6.28, h: 0.26, fontFace: FG, fontSize: 9.5, bold: true, color: C.gr, align: "right", margin: 0, valign: "middle" });
  });

  // 試算結果
  const lw = 7.0, rx = M + lw + 0.42, rw = R - rx;
  sec(s, M, 3.36, lw, "試算結果（月次・税抜）");
  const rc = (t, o) => td(t, Object.assign({ align: "right" }, o));
  const hf = (i) => (CASE[i].hot ? C.band : C.white);
  s.addTable(
    [
      [th("項目"), ...CASE.map((c) => th(c.t + (c.ref ? "（参考）" : "") + "\n" + c.pos, "right"))],
      [td("1日あたりの販売杯数"), ...CASE.map((c, i) => rc(c.cups + "杯", { bold: true, color: C.ink, fill: hf(i) }))],
      [td("月間販売杯数（30日）"), ...CASE.map((c, i) => rc(yen(c.month) + "杯", { fill: hf(i) }))],
      [td("セルフカフェ売上（1杯420円）"), ...CASE.map((c, i) => rc(yen(c.sales), { bold: true, color: C.ink, fill: hf(i) }))],
      [td("　うち 売上の25％"), ...CASE.map((c, i) => rc(yen(c.roy), { fill: hf(i) }))],
      [td("　月額固定"), ...CASE.map((c, i) => rc(yen(FIX), { fill: hf(i) }))],
      [td("貴社お受取額（月）", { fill: C.bandD, bold: true, color: C.gr }), ...CASE.map((c) => rc(yen(c.fee), { fill: C.bandD, bold: true, color: C.gr, size: 10.5 }))],
      [td("貴社お受取額（年・12か月）", { fill: C.bandD, bold: true, color: C.gr }), ...CASE.map((c) => rc(yen(c.year), { fill: C.bandD, bold: true, color: C.gr, size: 10 }))],
    ],
    { x: M, y: 3.70, w: lw, colW: [2.44, 1.14, 1.14, 1.14, 1.14], rowH: [0.36, 0.26, 0.24, 0.26, 0.24, 0.24, 0.3, 0.3], border: TB, autoPage: false }
  );
  s.addText("秋田店（30坪・1日75杯）のお受取額 286,250円／月に対し、仙台泉30坪・店頭配置は ＋34,650円（＋12.1％）でございます。来客数の増加に対して伸びが小さいのは、月額固定50,000円が売上に連動しないためでございます。", {
    x: M, y: 5.98, w: lw, h: 0.44,
    fontFace: FG, fontSize: 8.5, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
  });

  sec(s, rx, 3.36, rw, "お受取額（月）");
  bars(s, {
    x: rx + 0.08, y: 3.98, w: rw - 0.16, h: 1.46, max: 40, unit: "単位：万円／月", gap: 0.12,
    data: CASE.map((c) => ({
      v: c.fee / 10000, t: (c.fee / 10000).toFixed(1), label: c.t, sub: c.pos,
      color: c.hot ? C.barD : c.ref ? "D9D9D9" : C.barL, fs: 8.5,
    })),
  });

  sec(s, M, 6.44, CW, "この試算を大きく左右する前提");
  [
    "区画の配置　　店頭配置と店奥配置で、1日86杯と60杯。お受取額では月82,000円の差が生じます。",
    "館の来客数　　秋田店の1.3倍は貴社よりご提示いただいた計画値でございます。実績ではございません。",
    "営業時間　　　未定のため書店営業時間ベースで置いております。稼働時間により増減いたします。",
  ].forEach((t, i) => {
    s.addText("（" + (i + 1) + "）", {
      x: M, y: 6.80 + i * 0.21, w: 0.5, h: 0.2,
      fontFace: FG, fontSize: 8.5, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(t, {
      x: M + 0.5, y: 6.80 + i * 0.21, w: CW - 0.5, h: 0.2,
      fontFace: FG, fontSize: 8.5, color: C.body, margin: 0, valign: "middle",
    });
  });

  foot(s, [
    "※　想定値であり、売上およびお受取額を保証するものではございません。金額はすべて税抜でございます。",
    "※　水道光熱費・通信費・消耗品費は貴社ご負担のため、本試算には含めておりません。想定坪数は30坪、60坪はご参考でございます。",
  ]);
}

/* ================================================= P2 */
{
  const s = pres.addSlide();
  head(s, "販売杯数の置き方と、試算の前提", "館の来客数 1.3倍 に対し、販売杯数は 1.15倍 として算出しております");

  sec(s, M, 1.14, CW, "販売杯数の算出", "いずれも開業前の仮定値でございます");
  const fw = 2.70, ga = 1.06;
  [
    { t: "秋田店 30坪", v: "75", s2: "確定済みの想定値", f: C.white },
    { t: "仙台泉 30坪　店頭配置", v: "86", s2: "本試算の基準", f: C.bandD, tc: C.gr },
    { t: "仙台泉 30坪　店奥配置", v: "60", s2: "通路からの直接来店を見込まず", f: C.white },
  ].forEach((f, i) => {
    const x = M + i * (fw + ga);
    s.addShape("rect", { x, y: 1.52, w: fw, h: 1.02, fill: { color: f.f }, line: { color: C.ruleT, width: 0.75 } });
    s.addText(f.t, {
      x: x + 0.14, y: 1.57, w: fw - 0.28, h: 0.22,
      fontFace: FG, fontSize: 8.5, bold: true, color: f.tc || C.ink, margin: 0, valign: "middle",
    });
    s.addText(
      [
        { text: f.v, options: { fontFace: FG, fontSize: 22, bold: true, color: f.tc || C.ink } },
        { text: " 杯／日", options: { fontFace: FG, fontSize: 9, color: C.body } },
      ],
      { x: x + 0.14, y: 1.80, w: fw - 0.28, h: 0.38, margin: 0, valign: "middle" }
    );
    s.addText(f.s2, {
      x: x + 0.14, y: 2.20, w: fw - 0.28, h: 0.26,
      fontFace: FG, fontSize: 7.5, color: C.muted, margin: 0, valign: "top",
    });
  });
  [["来客数 1.3倍", "→　杯数 1.15倍"], ["配置の違い", "→　▲30％"]].forEach((a, i) => {
    const x = M + fw + i * (fw + ga);
    s.addText(a[0], {
      x: x - 0.14, y: 1.80, w: ga + 0.28, h: 0.18,
      fontFace: FG, fontSize: 7.5, color: C.muted, align: "center", margin: 0, valign: "middle",
    });
    s.addText(a[1], {
      x: x - 0.14, y: 1.98, w: ga + 0.28, h: 0.18,
      fontFace: FG, fontSize: 8.5, bold: true, color: C.ink, align: "center", margin: 0, valign: "middle",
    });
  });
  s.addText("60坪の場合は 98杯／日（店頭配置）、69杯／日（店奥配置）でございます。", {
    x: M, y: 2.60, w: CW, h: 0.18,
    fontFace: FG, fontSize: 8, color: C.muted, margin: 0, valign: "middle",
  });

  const lw = 6.1, rx = M + lw + 0.42, rw = R - rx;
  sec(s, M, 2.94, lw, "販売杯数を1.15倍とした理由");
  [
    ["立地特性の違い", "泉中央は地下鉄・バスに直結した駅前型で、通過利用の比率が郊外モール型の秋田店より高くなります。"],
    ["ピーク時間帯の席数", "席数は秋田店と同じ約59席のため、ピーク時には満席となり、増加分の一部は取りこぼしとなります。"],
    ["計画値であること", "新規開業の館のため、来客数は計画値であり実績の裏付けがございません。"],
  ].forEach((w, i) => {
    const y = 3.32 + i * 0.6;
    s.addText("（" + (i + 1) + "）" + w[0], {
      x: M, y, w: lw, h: 0.2,
      fontFace: FG, fontSize: 9.5, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(w[1], {
      x: M + 0.38, y: y + 0.21, w: lw - 0.38, h: 0.34,
      fontFace: FG, fontSize: 8.5, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
    });
  });

  sec(s, rx, 2.94, rw, "弊社既存店の実績との位置づけ");
  bars(s, {
    x: rx + 0.42, y: 3.56, w: rw - 0.6, h: 1.26, max: 140, unit: "単位：杯／日", gap: 0.3,
    data: [
      { v: 42.9, t: "42.9", label: "書店併設", sub: "既存2店平均", color: C.barL, fs: 8.5 },
      { v: 86, t: "86", label: "仙台泉", sub: "30坪・店頭", color: C.barD, fs: 9.5 },
      { v: 122.7, t: "122.7", label: "大型店", sub: "商業施設内", color: C.barL, fs: 8.5 },
    ],
  });
  s.addText("弊社の実績の範囲に収まっております。席の回転で見ますと、秋田店の1.27回転／日に対し1.46回転／日でございます。", {
    x: rx, y: 5.26, w: rw, h: 0.4,
    fontFace: FG, fontSize: 8, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
  });

  sec(s, M, 5.80, CW, "試算の前提");
  const PRE = [
    ["契約形態", "業務委託型（秋田店と同条件）", "想定席数", "30坪＝約59席／60坪＝約73席"],
    ["お支払い", "セルフカフェ売上の25％ ＋ 月額固定 50,000円", "営業時間", "未定（書店営業時間ベースで試算）"],
    ["販売単価", "1杯 420円（税抜・定価ベース）", "館の来客数", "秋田店の1.3倍（貴社ご提示・計画値）"],
    ["営業日数", "月30日・年12か月で算出", "開業時期", "2028年度内予定"],
    ["想定坪数", "30坪（60坪はご参考）", "", ""],
  ];
  s.addTable(
    PRE.map((r) => [
      td(r[0], { bold: true, color: C.ink, size: 8.5, fill: C.band }),
      td(r[1], { size: 8.5 }),
      td(r[2], { bold: true, color: C.ink, size: 8.5, fill: r[2] ? C.band : C.white }),
      td(r[3], { size: 8.5 }),
    ]),
    { x: M, y: 6.16, w: CW, colW: [1.42, 3.62, 1.42, CW - 6.46], rowH: PRE.map(() => 0.27), border: TB, autoPage: false }
  );

  foot(s, [
    "※　既存店の実績は社内管理台帳（2026年5月〜8月の月間実杯数の平均）によるものでございます。",
    "※　上記の倍率は弊社にて置いた仮定であり、実績の裏付けがあるものではございません。",
  ]);
}

/* ================================================= P3 */
{
  const s = pres.addSlide();
  head(s, "本試算の位置づけと、ご確認事項", "開業前の机上計算であることを、あらためてご確認ください");

  sec(s, M, 1.14, CW, "本試算の位置づけ");
  s.addTable(
    [
      [th("", "center"), th("項目"), th("内容")],
      ...[
        ["開業前の机上計算でございます", "現地での営業実績はございません。ご提示いただいた前提数値から、計算式のとおり機械的に算出したものでございます。"],
        ["来客数は計画値でございます", "仙台泉は2028年度内の開業予定でございます。館の来客数は計画段階の数値であり、実績ではございません。"],
        ["未確定の前提がございます", "区画の配置（店頭・店奥）、営業時間、坪数はいずれも未確定でございます。前提が変われば結果も変わります。"],
        ["お約束するものではございません", "売上およびお受取額を保証するものではございません。条件が固まりしだい、あらためて更新しご提出いたします。"],
      ].map((p, i) => [
        td("（" + (i + 1) + "）", { align: "center", color: C.ink, size: 9 }),
        td(p[0], { bold: true, color: C.ink, size: 9 }),
        td(p[1], { size: 9 }),
      ]),
    ],
    { x: M, y: 1.50, w: CW, colW: [0.62, 2.72, CW - 3.34], rowH: [0.24, 0.34, 0.34, 0.34, 0.34], border: TB, autoPage: false }
  );
  s.addText("本試算は、ご検討の出発点としてお使いいただくことを想定しております。数字だけを機械的に計算した場合にどうなるかをお示ししたものでございますので、実際の売上を示すものではない点を、何卒ご了承くださいませ。", {
    x: M, y: 3.14, w: CW, h: 0.24,
    fontFace: FG, fontSize: 9, bold: true, color: C.ink, margin: 0, valign: "middle",
  });

  sec(s, M, 3.66, CW, "ご確認をお願いしたい事項");
  [
    ["カフェ区画の配置", "店頭配置・店奥配置のいずれを想定されるか。本試算で最も結果を左右する項目でございます。"],
    ["給排水", "水回りなしのため新設が必要でございます。引込ルートと費用区分のご確認をお願いいたします。"],
    ["営業時間", "館および書店の営業時間。稼働時間により販売杯数が変動いたします。"],
    ["区画の坪数", "30坪・60坪のいずれを前提とされるか。"],
  ].forEach((c, i) => {
    const y = 4.06 + i * 0.52;
    s.addText("□", { x: M, y, w: 0.26, h: 0.22, fontFace: FG, fontSize: 9.5, color: C.ink, margin: 0, valign: "middle" });
    s.addText(c[0], {
      x: M + 0.32, y, w: 2.5, h: 0.22,
      fontFace: FG, fontSize: 9.5, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(c[1], {
      x: M + 2.9, y, w: CW - 2.9, h: 0.22,
      fontFace: FG, fontSize: 9, color: C.body, margin: 0, valign: "middle",
    });
  });

  s.addShape("rect", { x: M, y: 6.40, w: CW, h: 0.6, fill: { color: C.band } });
  s.addText("FCプランでの試算につきましては、本試算と同じ販売杯数を前提に、別途ご提出いたします。", {
    x: M + 0.2, y: 6.46, w: CW - 0.4, h: 0.48,
    fontFace: FG, fontSize: 9, bold: true, color: C.ink, margin: 0, valign: "middle",
  });

  foot(s, [
    "※　金額はすべて税抜でございます。水道光熱費・通信費・消耗品費は貴社ご負担のため、本試算には含めておりません。",
  ]);
}

pres.writeFile({ fileName: path.join(__dirname, "sendai-sim.pptx") }).then((f) => console.log("written:", f));
