/**
 * 未来屋書店 仙台泉（新規出店）セルフカフェ FCプラン試算
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
pres.title = "仙台泉 セルフカフェ FCプラン試算";

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
const UP = 420, DAYS = 30;
const UTL = { 30: 60000, 60: 70000 };
const ETC = { 30: 20000, 60: 30000 };
const INV = { 30: 9000000, 60: 12000000 };
const FEE = { 86: 320900, 60: 239000, 98: 358700, 69: 267350 };   // 業務委託型のお受取額
const CASE = [
  { t: "30坪", pos: "店頭配置", cups: 86, tb: 30, hot: true },
  { t: "30坪", pos: "店奥配置", cups: 60, tb: 30 },
  { t: "60坪", pos: "店頭配置", cups: 98, tb: 60, ref: true },
  { t: "60坪", pos: "店奥配置", cups: 69, tb: 60, ref: true },
].map((c) => {
  const sales = c.cups * UP * DAYS;
  const mat = Math.floor(sales * 0.2);
  const fix = UTL[c.tb] + 86000 + 50000 + ETC[c.tb];
  const pf = sales - mat - fix;
  return Object.assign({}, c, { sales, mat, cln: 30000, utl: UTL[c.tb], mac: 86000, sec: 20000, roy: 50000, etc: ETC[c.tb], pf, inv: INV[c.tb], pay: INV[c.tb] / pf });
});
const yen = (v) => v.toLocaleString();
const man = (v) => (v / 10000).toLocaleString();

/* ================================================= P1 月次損益 */
{
  const s = pres.addSlide();
  head(s, "仙台泉（新規出店）　セルフカフェ FCプラン試算", "貴社が運営主体となる場合／販売杯数は業務委託型の試算と同一　開業 2028年度内予定");

  s.addShape("rect", { x: M, y: 1.14, w: CW, h: 0.58, fill: { color: C.white }, line: { color: C.ruleS, width: 1 } });
  s.addText("【ご留意】　本資料は、開業前の時点における机上の試算でございます。", {
    x: M + 0.18, y: 1.19, w: CW - 0.36, h: 0.22,
    fontFace: FM, fontSize: 10.5, bold: true, color: C.gr, margin: 0, valign: "middle",
  });
  s.addText("現地での営業実績はございません。館の来客数は計画値であり、区画の配置・営業時間はいずれも未確定でございます。売上・利益を保証するものではございません。", {
    x: M + 0.18, y: 1.42, w: CW - 0.36, h: 0.24,
    fontFace: FG, fontSize: 8.5, color: C.body, margin: 0, valign: "middle",
  });

  const lw = 7.3, rx = M + lw + 0.42, rw = R - rx;
  sec(s, M, 1.94, lw, "月次損益（税抜）");
  const rc = (t, o) => td(t, Object.assign({ align: "right" }, o));
  const hf = (i) => (CASE[i].hot ? C.band : C.white);
  const row = (label, fn, o = {}) => [td(label, o.l || {}), ...CASE.map((c, i) => rc(fn(c), Object.assign({ fill: hf(i) }, o.v || {})))];
  s.addTable(
    [
      [th("項目"), ...CASE.map((c) => th(c.t + (c.ref ? "（参考）" : "") + "\n" + c.pos, "right"))],
      row("売上高（1杯420円×杯数×30日）", (c) => yen(c.sales), { l: { bold: true, color: C.ink }, v: { bold: true, color: C.ink } }),
      row("　想定販売杯数（1日）", (c) => c.cups + "杯", { l: { color: C.muted }, v: { color: C.muted } }),
      row("ドリンク原料（売上の20％）", (c) => "▲ " + yen(c.mat)),
      row("清掃業務委託費（貴社スタッフ対応のため）", () => "（30,000）", { l: { color: C.muted }, v: { color: C.muted } }),
      row("水道光熱費（書店営業時間ベース）", (c) => "▲ " + yen(c.utl)),
      row("機械使用料（マシン2台・決済端末2台）", () => "▲ 86,000"),
      row("ロイヤリティ（一律・売上連動なし）", () => "▲ 50,000"),
      row("雑費", (c) => "▲ " + yen(c.etc)),
      row("家賃（貴社の賃借面積に含まれる前提）", () => "0"),
      [td("償却前営業利益", { fill: C.bandD, bold: true, color: C.gr }), ...CASE.map((c) => rc(yen(c.pf), { fill: C.bandD, bold: true, color: C.gr, size: 11 }))],
      [td("平均荒利率", { fill: C.band, bold: true, color: C.ink }), ...CASE.map(() => rc("80.0％", { fill: C.band, bold: true, color: C.ink }))],
    ],
    { x: M, y: 2.30, w: lw, colW: [3.06, 1.06, 1.06, 1.06, 1.06], rowH: [0.34, 0.26, 0.22, 0.26, 0.24, 0.26, 0.24, 0.24, 0.26, 0.32, 0.28], border: TB, autoPage: false }
  );

  sec(s, rx, 1.94, rw, "償却前営業利益（月）");
  bars(s, {
    x: rx + 0.08, y: 2.56, w: rw - 0.16, h: 1.44, max: 80, unit: "単位：万円／月", gap: 0.12,
    data: CASE.map((c) => ({
      v: c.pf / 10000, t: (c.pf / 10000).toFixed(1), label: c.t, sub: c.pos,
      color: c.hot ? C.gr : c.ref ? "D9D9D9" : C.barL, fs: 8.5,
    })),
  });
  s.addShape("rect", { x: rx, y: 4.52, w: rw, h: 0.72, fill: { color: C.band } });
  s.addText("平均荒利率は 80.0％ でございます。ドリンク原料のみが変動費で、他の費目は固定でございます。", {
    x: rx + 0.18, y: 4.58, w: rw - 0.36, h: 0.6,
    fontFace: FG, fontSize: 8.5, color: C.ink, margin: 0, valign: "middle", lineSpacingMultiple: 1.25,
  });

  sec(s, M, 5.86, CW, "ランニング経費についての補足");
  [
    "清掃業務委託費　貴社スタッフでご対応いただく前提のため、計上しておりません。外部委託とされる場合は月30,000円が加わります。",
    "人件費　未計上でございます。日常業務（清掃・補充・発注）は、業務委託型・FCプランのいずれも貴社スタッフによる運用を想定しております。",
    "セキュリティ費　貴社書店内への併設のため、計上しておりません。",
    "家賃　カフェ区画が貴社の賃借面積に含まれる前提で0円としております。別途ご負担が生じる場合は差し引いてご覧ください。",
  ].forEach((t, i) => {
    s.addText("（" + (i + 1) + "）", {
      x: M, y: 6.22 + i * 0.21, w: 0.5, h: 0.2,
      fontFace: FG, fontSize: 8.5, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(t, {
      x: M + 0.5, y: 6.22 + i * 0.21, w: CW - 0.5, h: 0.2,
      fontFace: FG, fontSize: 8.5, color: C.body, margin: 0, valign: "middle",
    });
  });

  foot(s, [
    "※　償却前営業利益（減価償却前）でございます。想定値であり、売上および利益を保証するものではございません。金額はすべて税抜。",
    "※　販売杯数は業務委託型の試算と同一（館の来客数は秋田店の1.3倍、販売杯数は1.15倍）でございます。60坪はご参考でございます。",
  ]);
}

/* ================================================= P2 投資 */
{
  const s = pres.addSlide();
  head(s, "ご投資額と、投資回収の目安", "基本工事（床・壁・天井・照明）は本表に含めておりません");

  const lw = 6.3, rx = M + lw + 0.42, rw = R - rx;
  sec(s, M, 1.14, lw, "ご投資額の内訳（概算）");
  const IV = [
    ["加盟金", 1000000, 1000000],
    ["保証金", 500000, 500000],
    ["内装造作工事（カウンター・造作・二次側配線配管）", 4200000, 6000000],
    ["家具什器（テーブル・椅子・ソファ）", 1300000, 1900000],
    ["サイン工事", 500000, 600000],
    ["開業準備金（研修・オープン販促・初期在庫）", 500000, 600000],
    ["設計・申請・諸経費", 1000000, 1400000],
  ];
  s.addTable(
    [
      [th("項目"), th("30坪", "right"), th("60坪（参考）", "right")],
      ...IV.map((r) => [td(r[0], { size: 8.5 }), td(yen(r[1]), { align: "right", size: 8.5 }), td(yen(r[2]), { align: "right", size: 8.5, color: C.muted })]),
      [td("合　計", { fill: C.bandD, bold: true, color: C.gr }), td("9,000,000", { align: "right", fill: C.bandD, bold: true, color: C.gr, size: 11 }), td("12,000,000", { align: "right", fill: C.bandD, bold: true, color: C.gr, size: 10 })],
    ],
    { x: M, y: 1.50, w: lw, colW: [3.5, 1.4, 1.4], rowH: [0.22, 0.26, 0.24, 0.3, 0.26, 0.24, 0.3, 0.26, 0.3], border: TB, autoPage: false }
  );
  s.addText("マシンおよび決済端末は、月額の機械使用料（86,000円／月）でのご提供でございます。初期のご購入は不要でございます。", {
    x: M, y: 3.98, w: lw, h: 0.4,
    fontFace: FG, fontSize: 8.5, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
  });

  sec(s, rx, 1.14, rw, "基本工事の扱い");
  s.addShape("rect", { x: rx, y: 1.50, w: rw, h: 1.26, fill: { color: C.white }, line: { color: C.ruleS, width: 1 } });
  s.addText("上表は、業務委託型であれば弊社が負担していた部分でございます。", {
    x: rx + 0.18, y: 1.58, w: rw - 0.36, h: 0.22,
    fontFace: FG, fontSize: 9, bold: true, color: C.gr, margin: 0, valign: "middle",
  });
  s.addText("天井・壁・床・照明等のベース部分（基本工事）は、他の書店区画と同等の仕上がりが必要となるため、いずれの契約形態でも貴社のご負担となります。したがって本表には含めておりません。ご判断にあたっては、上表の金額が「業務委託型と比べて増える分」とお考えいただけますと幸いでございます。", {
    x: rx + 0.18, y: 1.82, w: rw - 0.36, h: 0.86,
    fontFace: FG, fontSize: 8.5, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
  });

  sec(s, rx, 2.94, rw, "ご参考：仙台泉は水回りがございません");
  s.addText("給排水の新設が必要でございます。引込は基本工事側か弊社仕様側か、費用区分のご確認をお願いいたします。区分により上表の内装造作工事が増減いたします。", {
    x: rx, y: 3.30, w: rw, h: 0.6,
    fontFace: FG, fontSize: 8.5, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
  });

  sec(s, M, 4.58, CW, "投資回収の目安");
  s.addTable(
    [
      [th("区分"), th("ご投資額", "right"), th("償却前営業利益（月）", "right"), th("投資回収の目安", "right")],
      ...CASE.map((c) => [
        td(c.t + "　" + c.pos + (c.ref ? "（参考）" : ""), { bold: !!c.hot, color: c.hot ? C.gr : C.body, fill: c.hot ? C.bandD : C.white }),
        td(yen(c.inv), { align: "right", fill: c.hot ? C.bandD : C.white }),
        td(yen(c.pf), { align: "right", bold: !!c.hot, color: c.hot ? C.gr : C.body, fill: c.hot ? C.bandD : C.white }),
        td("約 " + c.pay.toFixed(1) + " か月", { align: "right", bold: true, color: c.hot ? C.gr : C.ink, fill: c.hot ? C.bandD : C.white }),
      ]),
    ],
    { x: M, y: 4.94, w: CW, colW: [3.4, 2.3, 2.5, CW - 8.2], rowH: [0.24, 0.3, 0.28, 0.28, 0.28], border: TB, autoPage: false }
  );
  s.addText("30坪・店頭配置の場合、ご投資額 900万円に対して約13.8か月でございます。清掃を外部委託とされる場合は月30,000円が加わり、回収の目安は約14.5か月となります。", {
    x: M, y: 6.50, w: CW, h: 0.4,
    fontFace: FG, fontSize: 8.5, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
  });

  foot(s, [
    "※　ご投資額は概算でございます。仕様確定後にあらためて正式なお見積りをご提出いたします。金額はすべて税抜。",
    "※　投資回収の目安は、ご投資額を償却前営業利益（月）で除したものでございます。",
  ]);
}

/* ================================================= P3 比較 */
{
  const s = pres.addSlide();
  head(s, "業務委託型とFCプランの比較", "30坪・店頭配置（1日86杯）で比較しております");

  const lw = 5.9, rx = M + lw + 0.42, rw = R - rx;
  sec(s, M, 1.14, lw, "2方式の比較");
  s.addTable(
    [
      [th("観点"), th("業務委託型", "right"), th("FCプラン", "right")],
      [td("運営主体"), td("セルフカフェ", { align: "right" }), td("未来屋書店様", { align: "right", bold: true, color: C.gr })],
      [td("貴社のご投資"), td("設置工事費 30〜50万円", { align: "right" }), td("900万円", { align: "right", bold: true, color: C.gr })],
      [td("貴社の月次収入"), td("320,900円（お受取額）", { align: "right" }), td("650,880円（償却前営業利益）", { align: "right", bold: true, color: C.gr })],
      [td("売上変動のリスク"), td("小（固定5万円を含む）", { align: "right" }), td("貴社がご負担", { align: "right" })],
      [td("食品衛生責任者"), td("不要", { align: "right" }), td("1名のご選任が必要", { align: "right" })],
      [td("設備の保守・更新"), td("弊社", { align: "right" }), td("貴社（弊社がサポート）", { align: "right" })],
      [td("月次の差額", { fill: C.bandD, bold: true, color: C.gr }), td("－", { align: "right", fill: C.bandD }), td("＋329,980円", { align: "right", fill: C.bandD, bold: true, color: C.gr, size: 10.5 })],
    ],
    { x: M, y: 1.50, w: lw, colW: [1.7, 1.9, 2.3], rowH: [0.22, 0.26, 0.26, 0.3, 0.26, 0.26, 0.26, 0.3], border: TB, autoPage: false }
  );

  sec(s, rx, 1.14, rw, "累計損益の推移", "単位：万円／30坪・店頭配置");
  s.addTable(
    [
      [th("経過"), th("業務委託型", "right"), th("FCプラン", "right"), th("差", "right")],
      [td("12か月", { size: 8.5 }), td("335.1", { align: "right", size: 8.5 }), td("▲ 118.9", { align: "right", size: 8.5 }), td("▲ 454.0", { align: "right", size: 8.5 })],
      [td("25.8か月", { color: C.gr, bold: true, size: 8.5 }), td("ほぼ同額", { align: "right", color: C.gr, size: 8.5 }), td("ほぼ同額", { align: "right", color: C.gr, size: 8.5 }), td("均衡点", { align: "right", bold: true, color: C.gr, size: 8.5 })],
      [td("36か月", { size: 8.5 }), td("1,105.2", { align: "right", size: 8.5 }), td("1,443.2", { align: "right", size: 8.5 }), td("＋337.9", { align: "right", bold: true, color: C.gr, size: 8.5 })],
      [td("60か月", { size: 8.5 }), td("1,875.4", { align: "right", size: 8.5 }), td("3,005.3", { align: "right", size: 8.5 }), td("＋1,129.9", { align: "right", bold: true, color: C.gr, size: 8.5 })],
    ],
    { x: rx, y: 1.50, w: rw, colW: [0.88, 1.06, 1.06, rw - 3.0], rowH: [0.22, 0.26, 0.26, 0.26, 0.26], border: TB, autoPage: false }
  );
  s.addText("ご投資額の差（850万円）を月次の差（329,980円）で回収しますと、約25.8か月（2年2か月）で逆転いたします。以降はFCプランが有利となります。", {
    x: rx, y: 2.86, w: rw, h: 0.5,
    fontFace: FG, fontSize: 8.5, color: C.body, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
  });

  sec(s, M, 3.74, lw, "FCプランの利点");
  [
    ["売上・利益がすべて貴社のものになります", "業務委託型の月320,900円に対し、650,880円。約2.0倍でございます。"],
    ["ロイヤリティが売上連動ではございません", "一律 月50,000円のため、売上が伸びるほど利益率が向上いたします。"],
    ["マシンは月額使用料でご提供いたします", "初期のご購入が不要で、更新・故障対応も弊社にて行います。"],
    ["自社業態としての運営知見が蓄積されます", "今後の新規出店のご判断材料となります。"],
  ].forEach((t, i) => {
    const y = 4.12 + i * 0.5;
    s.addText("（" + (i + 1) + "）" + t[0], {
      x: M, y, w: lw, h: 0.2,
      fontFace: FG, fontSize: 9, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(t[1], {
      x: M + 0.38, y: y + 0.2, w: lw - 0.38, h: 0.26,
      fontFace: FG, fontSize: 8.3, color: C.body, margin: 0, valign: "middle",
    });
  });

  sec(s, rx, 3.74, rw, "FCプランの留意点");
  [
    ["初期投資が必要でございます", "900万円。回収の目安は約13.8か月でございます。"],
    ["売上変動のリスクを貴社が負われます", "業務委託型の月額固定50,000円のような下支えがございません。"],
    ["食品衛生責任者の選任が必要でございます", "1名。保健所の申請手続きは弊社にて代行いたします。"],
    ["損益分岐点がございます", "1日約22杯を下回ると赤字でございます（30坪・想定86杯）。"],
  ].forEach((t, i) => {
    const y = 4.12 + i * 0.5;
    s.addText("（" + (i + 1) + "）" + t[0], {
      x: rx, y, w: rw, h: 0.2,
      fontFace: FG, fontSize: 9, bold: true, color: C.ink, margin: 0, valign: "middle",
    });
    s.addText(t[1], {
      x: rx + 0.38, y: y + 0.2, w: rw - 0.38, h: 0.26,
      fontFace: FG, fontSize: 8.3, color: C.body, margin: 0, valign: "middle",
    });
  });

  s.addShape("rect", { x: M, y: 6.28, w: CW, h: 0.64, fill: { color: C.band } });
  s.addText("ご判断の目安　─　2年半を超えてご継続いただける見通しであればFCプラン、まずは負担を抑えて始められる場合は業務委託型が適しております。", {
    x: M + 0.2, y: 6.34, w: CW - 0.4, h: 0.52,
    fontFace: FG, fontSize: 9, bold: true, color: C.ink, margin: 0, valign: "middle",
  });

  foot(s, [
    "※　いずれも開業前の机上試算でございます。想定値であり、売上および利益を保証するものではございません。金額はすべて税抜。",
    "※　人件費は未計上でございます。家賃はカフェ区画が貴社の賃借面積に含まれる前提で0円としております。",
  ]);
}

pres.writeFile({ fileName: path.join(__dirname, "sendai-fc.pptx") }).then((f) => console.log("written:", f));
