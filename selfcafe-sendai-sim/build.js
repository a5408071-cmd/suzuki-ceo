/**
 * 未来屋書店 仙台泉（新規出店）セルフカフェ 売上試算 ─ 業務委託型
 * A4横 3ページ／白地・計算書スタイル（ネイビー基調）
 */
const pptxgen = require("pptxgenjs");
const path = require("path");

const C = {
  bg: "FFFFFF",
  ink: "1C1C1C",
  navy: "27486B",
  navyMid: "4A6E96",
  navyPale: "E8EEF4",
  navyFaint: "F4F7FA",
  accent: "A05A2C",
  accentPale: "F7EFE7",
  body: "3C3C3C",
  muted: "7A8189",
  grid: "D6DADF",
  hair: "E6E9EC",
  soft: "F7F8F9",
  white: "FFFFFF",
};
const F = { jp: "Yu Gothic", num: "Arial" };
const HAS_JP = /[ぁ-んァ-ヶ一-龥々〜ー％円杯年月日席]/;
const nf = (s) => (HAS_JP.test(String(s)) ? F.jp : F.num);

const W = 11.69, H = 8.27, M = 0.56, CW = W - M * 2, R = M + CW;

const pres = new pptxgen();
pres.defineLayout({ name: "A4L", width: W, height: H });
pres.layout = "A4L";
pres.author = "セルフカフェ株式会社";
pres.title = "仙台泉 セルフカフェ 売上試算（業務委託型）";

let PAGE = 0;
const TOTAL = 3;

/* ---------------- chrome：白地＋細罫 ---------------- */
function head(s, title, sub) {
  PAGE++;
  s.background = { color: C.bg };
  s.addShape("rect", { x: 0, y: 0, w: W, h: 0.075, fill: { color: C.navy } });
  s.addText(title, {
    x: M, y: 0.34, w: CW - 3.2, h: 0.34,
    fontFace: F.jp, fontSize: 16, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
  if (sub)
    s.addText(sub, {
      x: M, y: 0.69, w: CW - 3.2, h: 0.2,
      fontFace: F.jp, fontSize: 8.5, color: C.muted, margin: 0, valign: "middle",
    });
  s.addText("未来屋書店株式会社　御中", {
    x: R - 3.2, y: 0.36, w: 3.2, h: 0.2,
    fontFace: F.jp, fontSize: 8.5, color: C.muted, align: "right", margin: 0, valign: "middle",
  });
  s.addText("セルフカフェ株式会社　2026年11月20日", {
    x: R - 3.2, y: 0.55, w: 3.2, h: 0.18,
    fontFace: F.jp, fontSize: 7.5, color: C.muted, align: "right", margin: 0, valign: "middle",
  });
  s.addText(PAGE + " / " + TOTAL, {
    x: R - 3.2, y: 0.73, w: 3.2, h: 0.18,
    fontFace: F.num, fontSize: 8, color: C.navy, align: "right", margin: 0, valign: "middle",
  });
  s.addShape("rect", { x: M, y: 0.99, w: CW, h: 0.018, fill: { color: C.navy } });
}
const foot = (s, notes) => {
  s.addShape("rect", { x: M, y: H - 0.80, w: CW, h: 0.008, fill: { color: C.grid } });
  (notes || []).forEach((t, i) =>
    s.addText(t, {
      x: M, y: H - 0.72 + i * 0.145, w: CW - 3.2, h: 0.14,
      fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
    })
  );
  s.addText("セルフカフェ株式会社", {
    x: R - 3.0, y: H - 0.32, w: 3.0, h: 0.16,
    fontFace: F.jp, fontSize: 7.5, color: C.muted, align: "right", margin: 0, valign: "middle",
  });
};
const sec = (s, x, y, w, t, note) => {
  s.addText(t, {
    x, y, w: w, h: 0.22,
    fontFace: F.jp, fontSize: 10.5, bold: true, color: C.navy, margin: 0, valign: "middle",
  });
  if (note)
    s.addText(note, {
      x: x + w - 3.6, y, w: 3.6, h: 0.22,
      fontFace: F.jp, fontSize: 7.5, color: C.muted, align: "right", margin: 0, valign: "middle",
    });
  s.addShape("rect", { x, y: y + 0.23, w, h: 0.01, fill: { color: C.navyMid } });
};

const th = (t, al) => ({
  text: t,
  options: { fill: { color: C.navyPale }, color: C.navy, bold: true, fontSize: 8.5, fontFace: F.jp, align: al || "left", valign: "middle", margin: [0.03, 0.1, 0.03, 0.1] },
});
const td = (t, o = {}) => ({
  text: t,
  options: {
    fill: { color: o.fill || C.white }, color: o.color || C.body, bold: !!o.bold,
    fontSize: o.size || 8.5, fontFace: o.face || nf(t), align: o.align || "left",
    valign: "middle", margin: [0.03, 0.1, 0.03, 0.1],
  },
});

/* ---------------- 棒グラフ ---------------- */
function bars(s, o) {
  const { x, y, w, h, data, max, unit } = o;
  const n = data.length, gap = o.gap === undefined ? 0.16 : o.gap;
  const bw = (w - gap * (n - 1)) / n, base = y + h;
  (o.grid || []).forEach((g) => {
    const gy = base - (g.v / max) * h;
    s.addShape("rect", { x, y: gy, w, h: 0.008, fill: { color: g.strong ? C.accent : C.grid } });
    if (g.label)
      s.addText(g.label, {
        x: x - 1.3, y: gy - 0.09, w: 1.2, h: 0.18,
        fontFace: F.jp, fontSize: 7.5, bold: !!g.strong, color: g.strong ? C.accent : C.muted, align: "right", margin: 0, valign: "middle",
      });
  });
  data.forEach((d, i) => {
    const bx = x + i * (bw + gap), bh = Math.max(0.03, (d.v / max) * h);
    s.addShape("rect", { x: bx, y: base - bh, w: bw, h: bh, fill: { color: d.color || C.navyMid } });
    s.addText(d.t || String(d.v), {
      x: bx - 0.25, y: base - bh - 0.25, w: bw + 0.5, h: 0.22,
      fontFace: F.jp, fontSize: d.fs || 9, bold: true, color: d.lc || C.ink, align: "center", margin: 0, valign: "middle",
    });
    s.addText(d.label, {
      x: bx - 0.25, y: base + 0.05, w: bw + 0.5, h: 0.2,
      fontFace: F.jp, fontSize: 8, color: C.body, align: "center", margin: 0, valign: "middle",
    });
    if (d.sub)
      s.addText(d.sub, {
        x: bx - 0.25, y: base + 0.23, w: bw + 0.5, h: 0.18,
        fontFace: F.jp, fontSize: 7, color: C.muted, align: "center", margin: 0, valign: "middle",
      });
  });
  s.addShape("rect", { x, y: base, w, h: 0.01, fill: { color: C.ink } });
  if (unit)
    s.addText(unit, {
      x, y: y - 0.22, w: 2.0, h: 0.18,
      fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
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

/* ================================================= P1 試算結果 */
{
  const s = pres.addSlide();
  head(s, "仙台泉（新規出店）　セルフカフェ 売上試算", "業務委託型／館の来客数は秋田店の1.3倍を前提　開業 2028年度内予定");

  // 冒頭の注記
  s.addShape("rect", { x: M, y: 1.12, w: CW, h: 0.70, fill: { color: C.accentPale }, line: { color: C.accent, width: 1 } });
  s.addShape("rect", { x: M, y: 1.12, w: 0.05, h: 0.70, fill: { color: C.accent } });
  s.addText("本資料は、開業前の時点における机上の試算でございます。", {
    x: M + 0.2, y: 1.18, w: CW - 0.4, h: 0.22,
    fontFace: F.jp, fontSize: 10, bold: true, color: C.accent, margin: 0, valign: "middle",
  });
  s.addText("現地での営業実績はございません。ご提示いただいた前提数値から、下記の計算式のとおり機械的に算出したものでございます。館の来客数は計画値であり、区画の配置・営業時間はいずれも未確定でございます。前提が変われば結果も変わりますので、あくまでご検討の出発点としてお取り扱いください。", {
    x: M + 0.2, y: 1.40, w: CW - 0.4, h: 0.38,
    fontFace: F.jp, fontSize: 8, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.2,
  });

  // 計算式
  sec(s, M, 1.98, CW, "計算式", "30坪・店頭配置の場合");
  s.addShape("rect", { x: M, y: 2.30, w: CW, h: 0.86, fill: { color: C.navyFaint } });
  const EQ = [
    ["①", "販売杯数（1日）", "×", "単価 420円", "×", "30日", "＝", "セルフカフェ売上（月）"],
    ["②", "セルフカフェ売上", "×", "25％", "＋", "月額固定 50,000円", "＝", "貴社お受取額（月）"],
  ];
  EQ.forEach((e, i) => {
    const y = 2.40 + i * 0.34;
    s.addText(e[0], {
      x: M + 0.18, y, w: 0.3, h: 0.26,
      fontFace: F.jp, fontSize: 9, bold: true, color: C.navy, margin: 0, valign: "middle",
    });
    const txt = e.slice(1).map((t, j) => ({
      text: t + "　",
      options: { fontFace: nf(t), fontSize: 9.5, bold: j === 6, color: j === 6 ? C.accent : C.ink },
    }));
    s.addText(txt, { x: M + 0.5, y, w: 5.3, h: 0.26, margin: 0, valign: "middle" });
    const v = i === 0
      ? [["86杯", "×", "420円", "×", "30日", "＝", "1,083,600円"]]
      : [["1,083,600円", "×", "25％", "＋", "50,000円", "＝", "320,900円"]];
    const vt = v[0].map((t, j) => ({
      text: t + "　",
      options: { fontFace: nf(t), fontSize: 9.5, bold: j === 6, color: j === 6 ? C.accent : C.body },
    }));
    s.addText(vt, { x: M + 5.95, y, w: CW - 6.1, h: 0.26, margin: 0, valign: "middle" });
  });

  // 試算表
  const lw = 7.0, rx = M + lw + 0.4, rw = R - rx;
  sec(s, M, 3.34, lw, "試算結果（月次・税抜）");
  const rc = (t, o) => td(t, Object.assign({ align: "right" }, o));
  const hf = (i) => (CASE[i].hot ? C.accentPale : C.white);
  s.addTable(
    [
      [th("項目"), ...CASE.map((c) => th(c.t + (c.ref ? "（参考）" : "") + "\n" + c.pos, "right"))],
      [td("1日あたりの販売杯数"), ...CASE.map((c, i) => rc(c.cups + "杯", { bold: true, color: C.ink, fill: hf(i) }))],
      [td("月間販売杯数（30日）"), ...CASE.map((c, i) => rc(yen(c.month) + "杯", { fill: hf(i) }))],
      [td("セルフカフェ売上（1杯420円）", { fill: C.navyFaint, bold: true, color: C.navy }), ...CASE.map((c) => rc(yen(c.sales), { bold: true, color: C.navy, fill: C.navyFaint }))],
      [td("　うち 売上の25％"), ...CASE.map((c, i) => rc(yen(c.roy), { fill: hf(i) }))],
      [td("　月額固定"), ...CASE.map((c, i) => rc(yen(FIX), { fill: hf(i) }))],
      [td("貴社お受取額（月）", { fill: C.accentPale, bold: true, color: C.accent }), ...CASE.map((c) => rc(yen(c.fee), { fill: C.accentPale, bold: true, color: C.accent, size: 10.5 }))],
      [td("貴社お受取額（年・12か月）", { fill: C.accentPale, bold: true, color: C.accent }), ...CASE.map((c) => rc(yen(c.year), { fill: C.accentPale, bold: true, color: C.accent, size: 10 }))],
    ],
    { x: M, y: 3.66, w: lw, colW: [2.44, 1.14, 1.14, 1.14, 1.14], rowH: [0.36, 0.26, 0.24, 0.28, 0.24, 0.24, 0.3, 0.3], border: { type: "solid", color: C.grid, pt: 0.75 }, autoPage: false }
  );
  s.addText("秋田店（30坪・1日75杯）のお受取額 286,250円／月に対し、仙台泉30坪・店頭配置は ＋34,650円（＋12.1％）でございます。来客数の増加に対して伸びが小さいのは、月額固定50,000円が売上に連動しないためでございます。", {
    x: M, y: 6.02, w: lw, h: 0.44,
    fontFace: F.jp, fontSize: 8, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
  });

  sec(s, rx, 3.34, rw, "お受取額（月）");
  bars(s, {
    x: rx + 0.08, y: 3.92, w: rw - 0.16, h: 1.52, max: 40, unit: "万円／月", gap: 0.12,
    data: CASE.map((c) => ({
      v: c.fee / 10000, t: (c.fee / 10000).toFixed(1), label: c.t, sub: c.pos,
      color: c.hot ? C.accent : c.ref ? C.navyPale : C.navyMid,
      lc: c.hot ? C.accent : C.ink, fs: 8.5,
    })),
  });

  sec(s, M, 6.54, CW, "この試算を大きく左右する前提");
  const SW = [
    ["区画の配置", "店頭配置と店奥配置で、1日86杯と60杯。お受取額で月82,000円の差が生じます。"],
    ["館の来客数", "秋田店の1.3倍は貴社ご提示の計画値でございます。実績ではございません。"],
    ["営業時間", "未定のため書店営業時間ベースで置いております。稼働時間により増減いたします。"],
  ];
  const sw = (CW - 0.2 * 2) / 3;
  SW.forEach((d, i) => {
    const x = M + i * (sw + 0.2);
    s.addShape("rect", { x, y: 6.84, w: 0.04, h: 0.52, fill: { color: C.navyMid } });
    s.addText(d[0], {
      x: x + 0.14, y: 6.84, w: sw - 0.14, h: 0.2,
      fontFace: F.jp, fontSize: 8.5, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(d[1], {
      x: x + 0.14, y: 7.04, w: sw - 0.14, h: 0.34,
      fontFace: F.jp, fontSize: 7.5, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.2,
    });
  });

  foot(s, [
    "※　想定値であり、売上およびお受取額を保証するものではございません。金額はすべて税抜でございます。",
    "※　水道光熱費・通信費・消耗品費は貴社ご負担のため、本試算には含めておりません。想定坪数は30坪、60坪はご参考でございます。",
  ]);
}

/* ================================================= P2 杯数の置き方 */
{
  const s = pres.addSlide();
  head(s, "販売杯数の置き方", "館の来客数 1.3倍 に対し、販売杯数は 1.15倍 として算出しております");

  const fw = 2.46, ga = 0.82;
  const FLOW = [
    { t: "秋田店 30坪", v: "75", s2: "確定済みの想定値", c: C.soft, tc: C.ink },
    { t: "仙台泉 30坪　店頭配置", v: "86", s2: "本試算の基準", c: C.accentPale, tc: C.accent },
    { t: "仙台泉 30坪　店奥配置", v: "60", s2: "通路からの直接来店を見込まず", c: C.navyPale, tc: C.navy },
  ];
  FLOW.forEach((f, i) => {
    const x = M + i * (fw + ga);
    s.addShape("rect", { x, y: 1.14, w: fw, h: 1.02, fill: { color: f.c } });
    s.addText(f.t, {
      x: x + 0.14, y: 1.19, w: fw - 0.28, h: 0.22,
      fontFace: F.jp, fontSize: 8.5, bold: true, color: f.tc, margin: 0, valign: "middle",
    });
    s.addText(
      [
        { text: f.v, options: { fontFace: F.num, fontSize: 22, bold: true, color: f.tc } },
        { text: " 杯／日", options: { fontFace: F.jp, fontSize: 9, bold: true, color: C.body } },
      ],
      { x: x + 0.14, y: 1.42, w: fw - 0.28, h: 0.38, margin: 0, valign: "middle" }
    );
    s.addText(f.s2, {
      x: x + 0.14, y: 1.82, w: fw - 0.28, h: 0.26,
      fontFace: F.jp, fontSize: 7.3, color: C.muted, margin: 0, valign: "top",
    });
  });
  [["館の来客数 1.3倍", "販売杯数は 1.15倍"], ["配置の違い", "▲30％"]].forEach((a, i) => {
    const x = M + fw + i * (fw + ga);
    s.addShape("rect", { x: x + 0.06, y: 1.64, w: ga - 0.12, h: 0.018, fill: { color: C.navyMid } });
    s.addText(a[0], {
      x: x - 0.1, y: 1.26, w: ga + 0.2, h: 0.18,
      fontFace: F.jp, fontSize: 7, color: C.muted, align: "center", margin: 0, valign: "middle",
    });
    s.addText(a[1], {
      x: x - 0.1, y: 1.43, w: ga + 0.2, h: 0.18,
      fontFace: F.jp, fontSize: 8, bold: true, color: C.navy, align: "center", margin: 0, valign: "middle",
    });
  });
  s.addText("60坪の場合は 98杯／日（店頭配置）、69杯／日（店奥配置）でございます。いずれも開業前の仮定値でございます。", {
    x: M, y: 2.24, w: CW, h: 0.18,
    fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
  });

  const lw = 6.5, rx = M + lw + 0.4, rw = R - rx;
  sec(s, M, 2.58, lw, "来客数1.3倍に対し、販売杯数を1.15倍とした理由");
  const WHY = [
    ["立地特性の違い", "泉中央は地下鉄・バス（1日約3.1万人）に直結した駅前型でございます。通過利用や短時間の買い回りの比率が、郊外モール型の秋田店より高くなるため、滞在を伴う需要の増加率は来客数の増加率を下回ると考えております。"],
    ["ピーク時間帯の席数", "30坪・約59席は秋田店と同条件でございます。来客数が増えても席数は変わらないため、土日昼間などのピーク時間帯に満席となる時間が延び、増加分の一部は取りこぼしとなります。"],
    ["計画値であること", "新規開業の館のため、来客数は計画値であり実績の裏付けがございません。開業当初の変動を見込み、安全側に置いております。"],
  ];
  WHY.forEach((w, i) => {
    const y = 2.92 + i * 0.86;
    s.addShape("rect", { x: M, y, w: 0.04, h: 0.78, fill: { color: C.navyMid } });
    s.addText((i + 1) + "．" + w[0], {
      x: M + 0.16, y, w: lw - 0.16, h: 0.2,
      fontFace: F.jp, fontSize: 9, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(w[1], {
      x: M + 0.16, y: y + 0.21, w: lw - 0.16, h: 0.58,
      fontFace: F.jp, fontSize: 8, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
    });
  });
  s.addShape("rect", { x: M, y: 5.56, w: lw, h: 0.60, fill: { color: C.navyFaint }, line: { color: C.navyMid, width: 0.75 } });
  s.addText("来客数の増加分（＋30％）のうち、およそ半分が販売杯数に反映されるものとして置いております。", {
    x: M + 0.16, y: 5.64, w: lw - 0.32, h: 0.44,
    fontFace: F.jp, fontSize: 8.5, bold: true, color: C.navy, margin: 0, valign: "middle", lineSpacingMultiple: 1.2,
  });

  sec(s, rx, 2.58, rw, "弊社既存店の実績との位置づけ");
  bars(s, {
    x: rx + 0.42, y: 3.14, w: rw - 0.6, h: 1.46, max: 140, unit: "杯／日", gap: 0.3,
    data: [
      { v: 42.9, t: "42.9", label: "書店併設", sub: "既存2店平均", color: C.navyPale, lc: C.muted, fs: 8.5 },
      { v: 86, t: "86", label: "仙台泉", sub: "30坪・店頭", color: C.accent, lc: C.accent, fs: 9.5 },
      { v: 122.7, t: "122.7", label: "大型店", sub: "商業施設内", color: C.navyMid, fs: 8.5 },
    ],
  });
  s.addText("弊社の実績の範囲に収まっており、無理のない水準でございます。", {
    x: rx, y: 5.06, w: rw, h: 0.2,
    fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
  });

  sec(s, rx, 5.42, rw, "席の回転で見た場合");
  s.addTable(
    [
      [th("区分"), th("1席あたり／日", "right")],
      [td("秋田店 30坪（75杯）", { size: 8 }), td("1.27回転", { align: "right", size: 8 })],
      [td("仙台泉 30坪（86杯）", { size: 8, fill: C.accentPale, bold: true, color: C.accent }), td("1.46回転", { align: "right", size: 8, fill: C.accentPale, bold: true, color: C.accent })],
      [td("※1.3倍をそのまま適用した場合", { size: 8, color: C.muted }), td("1.65回転", { align: "right", size: 8, color: C.muted })],
    ],
    { x: rx, y: 5.74, w: rw, colW: [2.5, rw - 2.5], rowH: [0.2, 0.24, 0.26, 0.24], border: { type: "solid", color: C.grid, pt: 0.75 }, autoPage: false }
  );

  foot(s, [
    "※　既存店の実績は社内管理台帳（2026年5月〜8月の月間実杯数の平均）によるものでございます。",
    "※　上記の倍率は弊社にて置いた仮定であり、実績の裏付けがあるものではございません。",
  ]);
}

/* ================================================= P3 位置づけ・前提 */
{
  const s = pres.addSlide();
  head(s, "本試算の位置づけと、前提条件", "開業前の机上計算であることを、あらためてご確認ください");

  sec(s, M, 1.12, CW, "本試算の位置づけ");
  const POS = [
    ["開業前の机上計算です", "現地での営業実績はございません。ご提示いただいた前提数値から、計算式のとおり機械的に算出したものでございます。"],
    ["来客数は計画値です", "仙台泉は2028年度内の開業予定でございます。館の来客数は計画段階の数値であり、実績ではございません。"],
    ["未確定の前提があります", "区画の配置（店頭・店奥）、営業時間、坪数はいずれも未確定でございます。前提が変われば結果も変わります。"],
    ["お約束するものではありません", "売上およびお受取額を保証するものではございません。条件が固まりしだい、あらためて更新しご提出いたします。"],
  ];
  const pw = (CW - 0.18 * 3) / 4;
  POS.forEach((p, i) => {
    const x = M + i * (pw + 0.18);
    s.addShape("rect", { x, y: 1.46, w: pw, h: 1.26, fill: { color: C.accentPale } });
    s.addShape("rect", { x, y: 1.46, w: pw, h: 0.04, fill: { color: C.accent } });
    s.addText(String(i + 1), {
      x: x + 0.16, y: 1.56, w: 0.3, h: 0.22,
      fontFace: F.num, fontSize: 11, bold: true, color: C.accent, margin: 0, valign: "middle",
    });
    s.addText(p[0], {
      x: x + 0.44, y: 1.56, w: pw - 0.6, h: 0.22,
      fontFace: F.jp, fontSize: 8.8, bold: true, color: C.accent, margin: 0, valign: "middle",
    });
    s.addText(p[1], {
      x: x + 0.16, y: 1.84, w: pw - 0.32, h: 0.80,
      fontFace: F.jp, fontSize: 8, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
    });
  });
  s.addText("本試算は、ご検討の出発点としてお使いいただくことを想定しております。数字だけを機械的に計算した場合にどうなるかをお示ししたものでございますので、実際の売上を示すものではない点を、何卒ご了承くださいませ。", {
    x: M, y: 2.84, w: CW, h: 0.22,
    fontFace: F.jp, fontSize: 8.5, bold: true, color: C.accent, margin: 0, valign: "middle",
  });

  const lw = 6.1, rx = M + lw + 0.4, rw = R - rx;
  sec(s, M, 3.30, lw, "試算の前提");
  const PRE = [
    ["契約形態", "業務委託型（秋田店と同条件）"],
    ["お支払い", "セルフカフェ売上の25％ ＋ 月額固定 50,000円"],
    ["販売単価", "1杯 420円（税抜・定価ベース）"],
    ["営業日数", "月30日・年12か月で算出"],
    ["想定坪数", "30坪（60坪はご参考）"],
    ["想定席数", "30坪＝約59席／60坪＝約73席"],
    ["営業時間", "未定（書店営業時間ベースで試算）"],
    ["館の来客数", "秋田店の1.3倍（貴社ご提示・計画値）"],
    ["開業時期", "2028年度内予定"],
  ];
  s.addTable(
    PRE.map((p, i) => [
      td(p[0], { bold: true, color: C.ink, size: 8.5, fill: i % 2 ? C.soft : C.white }),
      td(p[1], { size: 8.5, fill: i % 2 ? C.soft : C.white }),
    ]),
    { x: M, y: 3.62, w: lw, colW: [1.6, lw - 1.6], rowH: PRE.map(() => 0.28), border: { type: "solid", color: C.grid, pt: 0.75 }, autoPage: false }
  );

  sec(s, rx, 3.30, rw, "ご確認をお願いしたい事項");
  const CHK = [
    ["カフェ区画の配置", "店頭配置・店奥配置のいずれを想定されるか。本試算で最も結果を左右する項目でございます。"],
    ["給排水", "水回りなしのため新設が必要でございます。引込ルートと費用区分のご確認をお願いいたします。"],
    ["営業時間", "館および書店の営業時間。稼働時間により販売杯数が変動いたします。"],
    ["区画の坪数", "30坪・60坪のいずれを前提とされるか。"],
  ];
  CHK.forEach((c, i) => {
    const y = 3.62 + i * 0.62;
    s.addShape("rect", { x: rx, y, w: rw, h: 0.56, fill: { color: i % 2 ? C.soft : C.white }, line: { color: C.grid, width: 0.5 } });
    s.addText("□", {
      x: rx + 0.1, y: y + 0.04, w: 0.24, h: 0.2,
      fontFace: F.jp, fontSize: 9, color: C.navyMid, margin: 0, valign: "middle",
    });
    s.addText(c[0], {
      x: rx + 0.36, y: y + 0.04, w: rw - 0.5, h: 0.2,
      fontFace: F.jp, fontSize: 8.8, bold: true, color: C.navy, margin: 0, valign: "middle",
    });
    s.addText(c[1], {
      x: rx + 0.36, y: y + 0.24, w: rw - 0.5, h: 0.3,
      fontFace: F.jp, fontSize: 7.6, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.2,
    });
  });
  s.addShape("rect", { x: rx, y: 6.26, w: rw, h: 0.62, fill: { color: C.navyFaint }, line: { color: C.navyMid, width: 0.75 } });
  s.addText("FCプランでの試算につきましては、本試算と同じ販売杯数を前提に、別途ご提出いたします。", {
    x: rx + 0.16, y: 6.34, w: rw - 0.32, h: 0.46,
    fontFace: F.jp, fontSize: 8.3, bold: true, color: C.navy, margin: 0, valign: "middle", lineSpacingMultiple: 1.2,
  });

  foot(s, [
    "※　金額はすべて税抜でございます。水道光熱費・通信費・消耗品費は貴社ご負担のため、本試算には含めておりません。",
  ]);
}

pres.writeFile({ fileName: path.join(__dirname, "sendai-sim.pptx") }).then((f) => console.log("written:", f));
