#!/usr/bin/env python3
"""
pptxgenjs が生成する PPTX の OOXML スキーマ違反を修正する後処理。
PowerPoint の「プレゼンテーションに問題が見つかりました（修復）」対策。

修正内容:
  1. 段落 <a:p> 内の2つ目以降の <a:pPr> を削除し、最初の <a:pPr> を先頭に移動
  2. lineChart / areaChart に欠落している <c:grouping> を補完
  3. チャート要素（barChart/lineChart/areaChart/pie/doughnut, ser, catAx, valAx,
     scaling, dLbls）の子要素をスキーマ定義の順序に並べ替え
  4. area / line の ser に混入する invertIfNegative を削除

使い方: python3 fix-pptx.py <file.pptx>
"""
import sys, zipfile, shutil, re
from lxml import etree

C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"

ORDERS = {
    f"{{{C}}}barChart": ["barDir", "grouping", "varyColors", "ser", "dLbls", "gapWidth", "gapDepth", "overlap", "serLines", "axId", "extLst"],
    f"{{{C}}}lineChart": ["grouping", "varyColors", "ser", "dLbls", "dropLines", "hiLowLines", "upDownBars", "marker", "smooth", "axId", "extLst"],
    f"{{{C}}}areaChart": ["grouping", "varyColors", "ser", "dLbls", "dropLines", "gapDepth", "axId", "extLst"],
    f"{{{C}}}pieChart": ["varyColors", "ser", "dLbls", "firstSliceAng", "extLst"],
    f"{{{C}}}doughnutChart": ["varyColors", "ser", "dLbls", "firstSliceAng", "holeSize", "extLst"],
    f"{{{C}}}ser": ["idx", "order", "tx", "spPr", "explosion", "invertIfNegative", "marker", "pictureOptions", "dPt", "dLbls", "trendline", "errBars", "cat", "val", "shape", "smooth", "bubble3D", "extLst"],
    f"{{{C}}}catAx": ["axId", "scaling", "delete", "axPos", "majorGridlines", "minorGridlines", "title", "numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossesAt", "auto", "lblAlgn", "lblOffset", "tickLblSkip", "tickMarkSkip", "noMultiLvlLbl", "extLst"],
    f"{{{C}}}valAx": ["axId", "scaling", "delete", "axPos", "majorGridlines", "minorGridlines", "title", "numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossesAt", "crossBetween", "majorUnit", "minorUnit", "dispUnits", "extLst"],
    f"{{{C}}}scaling": ["logBase", "orientation", "max", "min", "extLst"],
    f"{{{C}}}dLbls": ["dLbl", "numFmt", "spPr", "txPr", "dLblPos", "showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize", "separator", "showLeaderLines", "leaderLines", "extLst"],
    f"{{{C}}}dLbl": ["idx", "delete", "layout", "tx", "numFmt", "spPr", "txPr", "dLblPos", "showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize", "separator", "extLst"],
    f"{{{C}}}legend": ["legendPos", "legendEntry", "layout", "overlay", "spPr", "txPr", "extLst"],
}

def local(el):
    return etree.QName(el).localname

def reorder(el, order):
    key = {n: i for i, n in enumerate(order)}
    kids = list(el)
    kids.sort(key=lambda k: key.get(local(k), len(order)))  # 安定ソート
    for k in kids:
        el.append(k)

def fix_chart(root):
    changed = False
    # 定義されていない軸IDへの参照を除去（pptxgenjs は存在しない3本目の axId を書き出す）
    defined = set()
    for axtag in ("catAx", "valAx", "serAx", "dateAx"):
        for ax in root.iter(f"{{{C}}}{axtag}"):
            axid = ax.find(f"{{{C}}}axId")
            if axid is not None:
                defined.add(axid.get("val"))
    for chtag in ("barChart", "lineChart", "areaChart", "pieChart", "doughnutChart", "scatterChart"):
        for ch in root.iter(f"{{{C}}}{chtag}"):
            for ref in ch.findall(f"{{{C}}}axId"):
                if ref.get("val") not in defined:
                    ch.remove(ref); changed = True
    # 整数であるべき線幅（EMU）に小数が入っている場合を補正（ポイント指定とみなして変換）
    for ln in root.iter(f"{{{A}}}ln"):
        w = ln.get("w")
        if w and "." in w:
            v = float(w)
            ln.set("w", str(round(v * 12700) if v < 100 else round(v)))
            changed = True
    # grouping の補完
    for tag, default in ((f"{{{C}}}lineChart", "standard"), (f"{{{C}}}areaChart", "standard"), (f"{{{C}}}barChart", "clustered")):
        for ch in root.iter(tag):
            if ch.find(f"{{{C}}}grouping") is None:
                g = etree.SubElement(ch, f"{{{C}}}grouping"); g.set("val", default)
                changed = True
    # area / line の ser から invertIfNegative を除去
    for tag in (f"{{{C}}}lineChart", f"{{{C}}}areaChart"):
        for ch in root.iter(tag):
            for ser in ch.findall(f"{{{C}}}ser"):
                for bad in ser.findall(f"{{{C}}}invertIfNegative"):
                    ser.remove(bad); changed = True
    # 並べ替え
    for tag, order in ORDERS.items():
        for el in root.iter(tag):
            before = [local(k) for k in el]
            reorder(el, order)
            if [local(k) for k in el] != before:
                changed = True
    return changed

def fix_paragraphs(root):
    changed = False
    for p in root.iter(f"{{{A}}}p"):
        pprs = p.findall(f"{{{A}}}pPr")
        if not pprs:
            continue
        first = pprs[0]
        if list(p).index(first) != 0:
            p.remove(first); p.insert(0, first); changed = True
        for extra in pprs[1:]:
            p.remove(extra); changed = True
    return changed

def main(path):
    src = zipfile.ZipFile(path)
    fixed = {}
    for name in src.namelist():
        if re.match(r"ppt/charts/chart\d+\.xml$", name):
            root = etree.fromstring(src.read(name))
            if fix_chart(root):
                fixed[name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        elif re.match(r"ppt/(slides|notesSlides|slideLayouts|slideMasters)/[^/]+\.xml$", name):
            root = etree.fromstring(src.read(name))
            if fix_paragraphs(root):
                fixed[name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    if not fixed:
        src.close(); print(f"{path}: 修正なし"); return
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for item in src.infolist():
            data = fixed.get(item.filename, src.read(item.filename))
            out.writestr(item, data)
    src.close()
    shutil.move(tmp, path)
    print(f"{path}: {len(fixed)} パートを修正 -> " + ", ".join(sorted(fixed)))

if __name__ == "__main__":
    for f in sys.argv[1:]:
        main(f)
