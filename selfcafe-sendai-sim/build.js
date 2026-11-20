/**
 * 未来屋書店 仙台泉（新規出店）セルフカフェ 売上試算 ─ 業務委託型
 * A4横 2ページ
 */
const pptxgen = require("pptxgenjs");
const path = require("path");

/* ---------------- palette / type ---------------- */
const C = {
  bg: "FFFFFF",
  deep: "0B4A2F",      // ヘッダー帯
  green: "16764B",     // 主グリーン
  mint: "C9E0D4",      // 淡グリーン
  tint: "EEF5F1",
  bronze: "B5812C",    // アクセント
  bronzeT: "F8F1E2",
  ink: "16201B",
  body: "3A423E",
  muted: "7C8580",
  rule: "D9DFDB",
  soft: "F4F7F5",
  white: "FFFFFF",
};
const F = { jp: "Yu Gothic", num: "Arial" };
const HAS_JP = /[ぁ-んァ-ヶ一-龥々〜ー％円杯年月日]/;
const nf = (s) => (HAS_JP.test(String(s)) ? F.jp : F.num);

const W = 11.69, H = 8.27, M = 0.5, CW = W - M * 2, R = M + CW;
const BAND = 0.92;

const pres = new pptxgen();
pres.defineLayout({ name: "A4L", width: W, height: H });
pres.layout = "A4L";
pres.author = "セルフカフェ株式会社";
pres.title = "仙台泉 セルフカフェ 売上試算（業務委託型）";

let PAGE = 0;
const TOTAL = 2;

/* ---------------- chrome ---------------- */
function head(s, title, sub) {
  PAGE++;
  s.background = { color: C.bg };
  s.addShape("rect", { x: 0, y: 0, w: W, h: BAND, fill: { color: C.deep } });
  s.addShape("rect", { x: 0, y: BAND, w: W, h: 0.045, fill: { color: C.bronze } });
  s.addShape("rect", { x: M, y: 0.28, w: 0.075, h: 0.38, fill: { color: C.bronze } });
  s.addText(title, {
    x: M + 0.2, y: 0.24, w: 7.6, h: 0.3,
    fontFace: F.jp, fontSize: 15, bold: true, color: C.white, margin: 0, valign: "middle",
  });
  if (sub)
    s.addText(sub, {
      x: M + 0.2, y: 0.55, w: 7.6, h: 0.2,
      fontFace: F.jp, fontSize: 8.5, color: "AFC6B9", margin: 0, valign: "middle",
    });
  s.addText("未来屋書店株式会社　御中", {
    x: R - 4.0, y: 0.26, w: 4.0, h: 0.2,
    fontFace: F.jp, fontSize: 8.5, color: "AFC6B9", align: "right", margin: 0, valign: "middle",
  });
  s.addText(PAGE + " / " + TOTAL, {
    x: R - 4.0, y: 0.50, w: 4.0, h: 0.22,
    fontFace: F.num, fontSize: 10, bold: true, color: C.bronze, align: "right", margin: 0, valign: "middle",
  });
}
const foot = (s, notes) => {
  s.addShape("rect", { x: M, y: H - 0.80, w: CW, h: 0.01, fill: { color: C.rule } });
  (notes || []).forEach((t, i) =>
    s.addText(t, {
      x: M, y: H - 0.72 + i * 0.145, w: CW - 3.6, h: 0.14,
      fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
    })
  );
  s.addText("セルフカフェ株式会社", {
    x: R - 3.5, y: H - 0.34, w: 3.5, h: 0.16,
    fontFace: F.jp, fontSize: 7.5, color: C.muted, align: "right", margin: 0, valign: "middle",
  });
};
const sec = (s, x, y, w, t, note) => {
  s.addShape("rect", { x, y: y + 0.02, w: 0.06, h: 0.19, fill: { color: C.bronze } });
  s.addText(t, {
    x: x + 0.15, y, w: w - 0.15, h: 0.23,
    fontFace: F.jp, fontSize: 10.5, bold: true, color: C.ink, margin: 0, valign: "middle",
  });
  if (note)
    s.addText(note, {
      x: x + w - 3.4, y, w: 3.4, h: 0.23,
      fontFace: F.jp, fontSize: 7.5, color: C.muted, align: "right", margin: 0, valign: "middle",
    });
};

/* ---------------- table cells ---------------- */
const th = (t, al) => ({
  text: t,
  options: { fill: { color: C.deep }, color: C.white, bold: true, fontSize: 8.5, fontFace: F.jp, align: al || "left", valign: "middle", margin: [0.03, 0.1, 0.03, 0.1] },
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
  const n = data.length;
  const gap = o.gap === undefined ? 0.16 : o.gap;
  const bw = o.bw || (w - gap * (n - 1)) / n;
  const base = y + h;
  // 目盛線
  (o.grid || []).forEach((g) => {
    const gy = base - (g.v / max) * h;
    s.addShape("rect", { x, y: gy, w, h: 0.008, fill: { color: g.strong ? C.bronze : C.rule } });
    if (g.label)
      s.addText(g.label, {
        x: x - 1.35, y: gy - 0.09, w: 1.25, h: 0.18,
        fontFace: F.jp, fontSize: 7.5, bold: !!g.strong, color: g.strong ? C.bronze : C.muted, align: "right", margin: 0, valign: "middle",
      });
  });
  data.forEach((d, i) => {
    const bx = x + i * (bw + gap);
    const bh = Math.max(0.03, (d.v / max) * h);
    s.addShape("rect", { x: bx, y: base - bh, w: bw, h: bh, fill: { color: d.color || C.green } });
    s.addText(d.t || String(d.v), {
      x: bx - 0.25, y: base - bh - 0.26, w: bw + 0.5, h: 0.22,
      fontFace: F.jp, fontSize: d.fs || 9.5, bold: true, color: d.lc || C.ink, align: "center", margin: 0, valign: "middle",
    });
    s.addText(d.label, {
      x: bx - 0.25, y: base + 0.05, w: bw + 0.5, h: 0.2,
      fontFace: F.jp, fontSize: 8, color: C.body, align: "center", margin: 0, valign: "middle",
    });
    if (d.sub)
      s.addText(d.sub, {
        x: bx - 0.25, y: base + 0.24, w: bw + 0.5, h: 0.18,
        fontFace: F.jp, fontSize: 7, color: C.muted, align: "center", margin: 0, valign: "middle",
      });
  });
  s.addShape("rect", { x, y: base, w, h: 0.012, fill: { color: C.ink } });
  if (unit)
    s.addText(unit, {
      x, y: y - 0.22, w: 2.0, h: 0.18,
      fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
    });
}


/* ---- 試算データ ------------------------------------------------ */
const P1C = 420, DAYS = 30, FIX = 50000;
const CASE = [
  { t: "30坪", pos: "店頭配置", cups: 86, hot: true,  col: "B5812C" },
  { t: "30坪", pos: "店奥配置", cups: 60, hot: false, col: "8FBFA6" },
  { t: "60坪", pos: "店頭配置", cups: 98, hot: false, col: "0B4A2F", ref: true },
  { t: "60坪", pos: "店奥配置", cups: 69, hot: false, col: "B9C4BD", ref: true },
].map((c) => {
  const sales = c.cups * P1C * DAYS;
  const roy = Math.floor(sales * 0.25);
  return Object.assign({}, c, {
    month: c.cups * DAYS, sales, roy, fee: roy + FIX, year: (roy + FIX) * 12,
  });
});
const yen = (v) => v.toLocaleString();

/* ================================================= P1 試算 */
{
  const s = pres.addSlide();
  head(s, "仙台泉（新規出店）　セルフカフェ 売上試算", "業務委託型／館の来客数は秋田店の1.3倍を前提　※開業 2028年度内予定");

  const kw = (CW - 0.18 * 3) / 4;
  CASE.forEach((k, i) => {
    const x = M + i * (kw + 0.18), y = 1.22, h = 1.14;
    s.addShape("rect", { x, y, w: kw, h, fill: { color: k.hot ? C.bronzeT : C.soft } });
    s.addShape("rect", { x, y, w: kw, h: 0.045, fill: { color: k.hot ? C.bronze : C.green } });
    s.addText(k.t + "　" + k.pos + (k.ref ? "（ご参考）" : ""), {
      x: x + 0.16, y: y + 0.12, w: kw - 0.32, h: 0.2,
      fontFace: F.jp, fontSize: 8.5, bold: true, color: k.hot ? C.bronze : C.green, margin: 0, valign: "middle",
    });
    s.addText("1日 " + k.cups + "杯 ／ 月 " + yen(k.month) + "杯", {
      x: x + 0.16, y: y + 0.31, w: kw - 0.32, h: 0.16,
      fontFace: F.jp, fontSize: 6.8, color: C.muted, margin: 0, valign: "middle",
    });
    s.addText(
      [
        { text: yen(k.fee), options: { fontFace: F.num, fontSize: 18, bold: true, color: k.hot ? C.bronze : C.ink } },
        { text: " 円／月", options: { fontFace: F.jp, fontSize: 8.5, bold: true, color: C.body } },
      ],
      { x: x + 0.16, y: y + 0.50, w: kw - 0.32, h: 0.34, margin: 0, valign: "middle" }
    );
    s.addText("年間 " + yen(k.year) + "円", {
      x: x + 0.16, y: y + 0.86, w: kw - 0.32, h: 0.18,
      fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
    });
  });
  s.addText("上記は貴社のお受取額（セルフカフェ売上の25％ ＋ 月額固定50,000円）でございます。すべて税抜。想定坪数は30坪、60坪はご参考でございます。", {
    x: M, y: 2.42, w: CW, h: 0.18,
    fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
  });

  const lw = 6.6, rx = M + lw + 0.38, rw = R - rx;
  sec(s, M, 2.72, lw, "試算の内訳（月次・税抜）");
  const rc = (t, o) => td(t, Object.assign({ align: "right" }, o));
  const hf = (i) => (CASE[i].hot ? C.bronzeT : C.white);
  s.addTable(
    [
      [th("項目"), ...CASE.map((c) => th(c.t + "\n" + c.pos, "right"))],
      [td("1日あたりの販売杯数"), ...CASE.map((c, i) => rc(c.cups + "杯", { bold: true, color: C.ink, fill: hf(i) }))],
      [td("月間販売杯数（30日）"), ...CASE.map((c, i) => rc(yen(c.month) + "杯", { fill: hf(i) }))],
      [td("セルフカフェ売上（1杯420円）", { fill: C.tint }), ...CASE.map((c, i) => rc(yen(c.sales), { bold: true, color: C.green, fill: C.tint }))],
      [td("　うち 売上の25％"), ...CASE.map((c, i) => rc(yen(c.roy), { fill: hf(i) }))],
      [td("　月額固定"), ...CASE.map((c, i) => rc(yen(FIX), { fill: hf(i) }))],
      [td("貴社お受取額（月）", { fill: C.bronzeT, bold: true, color: C.bronze }), ...CASE.map((c) => rc(yen(c.fee), { fill: C.bronzeT, bold: true, color: C.bronze, size: 10 }))],
      [td("貴社お受取額（年）", { fill: C.bronzeT, bold: true, color: C.bronze }), ...CASE.map((c) => rc(yen(c.year), { fill: C.bronzeT, bold: true, color: C.bronze, size: 10 }))],
    ],
    { x: M, y: 3.02, w: lw, colW: [2.2, 1.1, 1.1, 1.1, 1.1], rowH: [0.34, 0.26, 0.24, 0.28, 0.24, 0.24, 0.3, 0.3], border: { type: "solid", color: C.rule, pt: 0.75 }, autoPage: false }
  );

  sec(s, rx, 2.72, rw, "お受取額（月）");
  bars(s, {
    x: rx + 0.1, y: 3.30, w: rw - 0.2, h: 1.46, max: 40, unit: "万円／月", gap: 0.13,
    data: CASE.map((c) => ({ v: c.fee / 10000, t: (c.fee / 10000).toFixed(1), label: c.t, sub: c.pos, color: c.col, lc: c.hot ? C.bronze : C.ink, fs: 8.5 })),
  });

  s.addShape("rect", { x: rx, y: 5.34, w: rw, h: 0.86, fill: { color: C.tint } });
  s.addText("秋田店（30坪・1日75杯）のお受取額 286,250円／月に対し、仙台泉30坪・店頭配置は ＋34,650円（＋12.1％）でございます。", {
    x: rx + 0.16, y: 5.42, w: rw - 0.32, h: 0.70,
    fontFace: F.jp, fontSize: 8, color: C.ink, margin: 0, valign: "middle", lineSpacingMultiple: 1.25,
  });

  sec(s, M, 6.40, CW, "試算の前提");
  const PRE = [
    ["契約形態", "業務委託型（秋田店と同条件）"],
    ["お支払い", "売上の25％ ＋ 月額固定 50,000円"],
    ["販売単価", "1杯 420円（税抜・定価ベース）"],
    ["営業日数", "月30日で算出"],
    ["想定席数", "30坪＝約59席／60坪＝約73席"],
    ["営業時間", "未定（書店営業時間ベース）"],
  ];
  const pw = (CW - 0.16 * 2) / 3;
  PRE.forEach((p, i) => {
    const x = M + (i % 3) * (pw + 0.16), y = 6.70 + Math.floor(i / 3) * 0.34;
    s.addShape("rect", { x, y, w: pw, h: 0.3, fill: { color: i % 2 ? C.soft : C.white }, line: { color: C.rule, width: 0.5 } });
    s.addText(p[0], {
      x: x + 0.12, y, w: 1.3, h: 0.3,
      fontFace: F.jp, fontSize: 8, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(p[1], {
      x: x + 1.42, y, w: pw - 1.54, h: 0.3,
      fontFace: F.jp, fontSize: 8, color: C.body, margin: 0, valign: "middle",
    });
  });

  foot(s, [
    "※　想定値であり、売上およびお受取額を保証するものではございません。金額はすべて税抜でございます。",
    "※　水道光熱費・通信費・消耗品費は貴社ご負担のため、本試算には含めておりません。",
  ]);
}

/* ================================================= P2 杯数の置き方 */
{
  const s = pres.addSlide();
  head(s, "販売杯数の置き方", "館の来客数 1.3倍 に対し、販売杯数は 1.15倍 として算出しております");

  // フロー
  const fw = 2.42, ga = 0.78;
  const FLOW = [
    { t: "秋田店 30坪", v: "75", u: "杯／日", s2: "確定済みの想定値", c: C.soft, tc: C.ink },
    { t: "仙台泉 30坪　店頭配置", v: "86", u: "杯／日", s2: "本試算の基準", c: C.bronzeT, tc: C.bronze },
    { t: "仙台泉 30坪　店奥配置", v: "60", u: "杯／日", s2: "通路からの直接来店を見込まず", c: C.tint, tc: C.green },
  ];
  FLOW.forEach((f, i) => {
    const x = M + i * (fw + ga);
    s.addShape("rect", { x, y: 1.14, w: fw, h: 1.06, fill: { color: f.c } });
    s.addText(f.t, {
      x: x + 0.14, y: 1.20, w: fw - 0.28, h: 0.22,
      fontFace: F.jp, fontSize: 8.5, bold: true, color: f.tc, margin: 0, valign: "middle",
    });
    s.addText(
      [
        { text: f.v, options: { fontFace: F.num, fontSize: 22, bold: true, color: f.tc } },
        { text: " " + f.u, options: { fontFace: F.jp, fontSize: 9, bold: true, color: C.body } },
      ],
      { x: x + 0.14, y: 1.44, w: fw - 0.28, h: 0.4, margin: 0, valign: "middle" }
    );
    s.addText(f.s2, {
      x: x + 0.14, y: 1.86, w: fw - 0.28, h: 0.26,
      fontFace: F.jp, fontSize: 7.3, color: C.muted, margin: 0, valign: "top",
    });
  });
  [["館の来客数 1.3倍", "販売杯数は 1.15倍"], ["配置の違い", "▲30％"]].forEach((a, i) => {
    const x = M + fw + i * (fw + ga);
    s.addShape("rect", { x: x + 0.06, y: 1.66, w: ga - 0.12, h: 0.02, fill: { color: C.bronze } });
    s.addText(a[0], {
      x: x - 0.1, y: 1.30, w: ga + 0.2, h: 0.18,
      fontFace: F.jp, fontSize: 7, color: C.muted, align: "center", margin: 0, valign: "middle",
    });
    s.addText(a[1], {
      x: x - 0.1, y: 1.46, w: ga + 0.2, h: 0.18,
      fontFace: F.jp, fontSize: 8, bold: true, color: C.bronze, align: "center", margin: 0, valign: "middle",
    });
  });
  s.addText("60坪の場合は 98杯／日（店頭配置）、69杯／日（店奥配置）でございます。", {
    x: M, y: 2.28, w: CW, h: 0.18,
    fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
  });

  const lw = 6.6, rx = M + lw + 0.38, rw = R - rx;
  sec(s, M, 2.62, lw, "来客数1.3倍に対し、販売杯数を1.15倍とした理由");
  const WHY = [
    ["立地特性の違い", "泉中央は地下鉄・バス（1日約3.1万人）に直結した駅前型でございます。通過利用や短時間の買い回りの比率が、郊外モール型の秋田店より高くなるため、滞在を伴う需要の増加率は来客数の増加率を下回ると考えております。"],
    ["ピーク時間帯の席数", "30坪・約59席は秋田店と同条件でございます。来客数が増えても席数は変わらないため、土日昼間などのピーク時間帯に満席となる時間が延び、増加分の一部は取りこぼしとなります。"],
    ["計画値であること", "新規開業の館のため、来客数は計画値であり実績の裏付けがございません。開業当初の変動を見込み、安全側に置いております。"],
  ];
  WHY.forEach((w, i) => {
    const y = 2.92 + i * 0.84;
    s.addShape("rect", { x: M, y, w: 0.045, h: 0.76, fill: { color: C.bronze } });
    s.addText((i + 1) + "．" + w[0], {
      x: M + 0.16, y, w: lw - 0.16, h: 0.2,
      fontFace: F.jp, fontSize: 9, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(w[1], {
      x: M + 0.16, y: y + 0.21, w: lw - 0.16, h: 0.56,
      fontFace: F.jp, fontSize: 8, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
    });
  });
  s.addShape("rect", { x: M, y: 5.48, w: lw, h: 0.62, fill: { color: C.bronzeT } });
  s.addText("来客数の増加分（＋30％）のうち、およそ半分が販売杯数に反映されるものとして置いております。", {
    x: M + 0.16, y: 5.56, w: lw - 0.32, h: 0.46,
    fontFace: F.jp, fontSize: 8.5, bold: true, color: C.bronze, margin: 0, valign: "middle", lineSpacingMultiple: 1.2,
  });

  sec(s, rx, 2.62, rw, "既存店の実績との位置づけ");
  bars(s, {
    x: rx + 0.42, y: 3.14, w: rw - 0.62, h: 1.42, max: 140, unit: "杯／日", gap: 0.3,
    data: [
      { v: 42.9, t: "42.9", label: "書店併設", sub: "既存2店平均", color: "B9C4BD", lc: C.muted, fs: 8.5 },
      { v: 86, t: "86", label: "仙台泉", sub: "30坪・店頭", color: C.bronze, fs: 9.5 },
      { v: 122.7, t: "122.7", label: "大型店", sub: "商業施設内", color: C.green, fs: 8.5 },
    ],
  });
  s.addText("弊社の実績の範囲に収まっており、無理のない水準でございます。", {
    x: rx, y: 5.02, w: rw, h: 0.2,
    fontFace: F.jp, fontSize: 7.5, color: C.muted, margin: 0, valign: "middle",
  });

  sec(s, rx, 5.34, rw, "席の回転で見た場合");
  s.addTable(
    [
      [th("区分"), th("1席あたり／日", "right")],
      [td("秋田店 30坪（75杯）", { size: 8 }), td("1.27回転", { align: "right", size: 8 })],
      [td("仙台泉 30坪（86杯）", { size: 8, fill: C.bronzeT, bold: true, color: C.bronze }), td("1.46回転", { align: "right", size: 8, fill: C.bronzeT, bold: true, color: C.bronze })],
      [td("※1.3倍をそのまま適用", { size: 8, color: C.muted }), td("1.65回転", { align: "right", size: 8, color: C.muted })],
    ],
    { x: rx, y: 5.64, w: rw, colW: [2.35, rw - 2.35], rowH: [0.2, 0.22, 0.24, 0.22], border: { type: "solid", color: C.rule, pt: 0.75 }, autoPage: false }
  );

  sec(s, M, 6.46, CW, "ご確認をお願いしたい事項");
  const CHK = [
    ["カフェ区画の配置", "店頭配置・店奥配置のいずれを想定されるか"],
    ["給排水", "水回りなしのため新設が必要。引込の費用区分のご確認"],
    ["営業時間", "館および書店の営業時間（現時点で未定）"],
  ];
  const cw3 = (CW - 0.2 * 2) / 3;
  CHK.forEach((c, i) => {
    const x = M + i * (cw3 + 0.2);
    s.addShape("rect", { x, y: 6.76, w: cw3, h: 0.56, fill: { color: C.soft } });
    s.addShape("rect", { x, y: 6.76, w: 0.04, h: 0.56, fill: { color: C.bronze } });
    s.addText(c[0], {
      x: x + 0.14, y: 6.80, w: cw3 - 0.24, h: 0.2,
      fontFace: F.jp, fontSize: 8.5, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(c[1], {
      x: x + 0.14, y: 6.99, w: cw3 - 0.24, h: 0.3,
      fontFace: F.jp, fontSize: 7.3, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.2,
    });
  });

  foot(s, [
    "※　FCプランでの試算につきましては、本試算と同じ販売杯数を前提に、別途ご提出いたします。",
  ]);
}

pres.writeFile({ fileName: path.join(__dirname, "sendai-sim.pptx") }).then((f) => console.log("written:", f));
