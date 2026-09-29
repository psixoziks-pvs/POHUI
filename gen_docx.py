# -*- coding: utf-8 -*-
import zipfile, shutil, os, re, math
from lxml import etree

OUT = "/workspace/План школьной лаборатории.docx"

W  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A  = "http://schemas.openxmlformats.org/drawingml/2006/main"
C  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
R  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
V  = "urn:schemas-microsoft-com:vml"
O  = "urn:schemas-microsoft-com:office:office"
X  = "urn:schemas-microsoft-com:office:excel"
W10= "urn:schemas-microsoft-com:office:word"

def w(tag): return "{%s}%s" % (W, tag)
def wp(tag): return "{%s}%s" % (WP, tag)
def a(tag): return "{%s}%s" % (A, tag)

EMU_CM = 360000

# ---------------- doc.xml helpers ----------------
doc = etree.Element(w("document"), nsmap={"w":W,"wp":WP,"a":A,"r":R})
body = etree.SubElement(doc, w("body"))

def add_par(text="", style=None, align=None):
    p = etree.SubElement(body, w("p"))
    if style or align:
        pPr = etree.SubElement(p, w("pPr"))
        if style:
            e = etree.SubElement(pPr, w("pStyle")); e.set(w("val"), style)
        if align:
            e = etree.SubElement(pPr, w("jc")); e.set(w("val"), align)
    if text:
        r = etree.SubElement(p, w("r")); t = etree.SubElement(r, w("t"))
        t.text = text; xml_space="{http://www.w3.org/XML/1998/namespace}space"; t.set(xml_space,"preserve")
    return p

def pagebreak():
    p = add_par()
    r = etree.SubElement(p, w("r"))
    br = etree.SubElement(r, w("br")); br.set(w("type"), "page")

# ---------------- table builder ----------------
def make_table(rows, cols, colw_cm, rowh_cm, cell_texts, borders):
    """cell_texts: dict {(r,c): [lines]}; borders: dict {(r,c): set of edges with styles}"""
    tbl = etree.SubElement(body, w("tbl"))
    tblPr = etree.SubElement(tbl, w("tblPr"))
    tw = etree.SubElement(tblPr, w("tblW")); tw.set(w("w"), str(int(sum(colw_cm)*567))); tw.set(w("type"),"dxa")
    tblBorders = etree.SubElement(tblPr, w("tblBorders"))
    for edge in ("top","left","bottom","right","insideH","insideV"):
        e = etree.SubElement(tblBorders, w(edge)); e.set(w("val"),"none")
    lay = etree.SubElement(tblPr, w("tblLayout")); lay.set(w("type"),"fixed")
    grid = etree.SubElement(tbl, w("tblGrid"))
    for cw in colw_cm:
        gc = etree.SubElement(grid, w("gridCol")); gc.set(w("w"), str(int(cw*567)))
    used = [[False]*cols for _ in range(rows)]
    for ri in range(rows):
        tr = etree.SubElement(tbl, w("tr"))
        trPr = etree.SubElement(tr, w("trPr"))
        th = etree.SubElement(trPr, w("trHeight")); th.set(w("val"), str(int(rowh_cm[ri]*567))); th.set(w("hRule"),"exact")
        for ci in range(cols):
            if used[ri][ci]: continue
            # span detection
            rs, cs = 1, 1
            while ci+cs < cols and (ri, ci+cs) not in cell_texts and (ri, ci+cs) not in borders:
                used[ri][ci+cs] = True; cs += 1
            while ri+rs < rows and all((ri+rs, cj) not in cell_texts and (ri+rs, cj) not in borders for cj in range(ci, ci+cs)):
                for rj in range(ri+rs, rows):
                    pass
                for rj in range(ri+rs, ri+rs+1):
                    for cj in range(ci, ci+cs): used[rj][cj] = True
                rs += 1
            tc = etree.SubElement(tr, w("tc"))
            tcPr = etree.SubElement(tc, w("tcPr"))
            tcW = etree.SubElement(tcPr, w("tcW")); tcW.set(w("w"), str(int(colw_cm[ci]*cs*567))); tcW.set(w("type"),"dxa")
            if cs > 1:
                sp = etree.SubElement(tcPr, w("gridSpan")); sp.set(w("val"), str(cs))
            if rs > 1:
                vm = etree.SubElement(tcPr, w("vMerge")); vm.set(w("val"),"restart")
                for rj in range(ri+1, ri+rs):
                    for cj in range(ci, ci+cs): used[rj][cj] = True
            bset = borders.get((ri,ci), {})
            if bset:
                tb = etree.SubElement(tcPr, w("tcBorders"))
                for edge, style in bset.items():
                    val, sz = style
                    e = etree.SubElement(tb, w(edge)); e.set(w("val"), val); e.set(w("sz"), str(sz)); e.set(w("color"),"000000")
            elif rs > 1:
                tb = etree.SubElement(tcPr, w("tcBorders"))
                e = etree.SubElement(tb, w("bottom")); e.set(w("val"),"nil")
            texts = cell_texts.get((ri,ci), [])
            if not texts:
                tp = etree.SubElement(tc, w("p"))
                pPr = etree.SubElement(tp, w("pPr")); jc = etree.SubElement(pPr, w("jc")); jc.set(w("val"),"center")
            for line in texts:
                p = etree.SubElement(tc, w("p"))
                pPr = etree.SubElement(p, w("pPr"))
                jc = etree.SubElement(pPr, w("jc")); jc.set(w("val"),"center")
                r_ = etree.SubElement(p, w("r")); t_ = etree.SubElement(r_, w("t"))
                t_.text = line
    return tbl

# ---------------- VML canvas drawing ----------------
_drawing_ids = []

def add_canvas(width_cm, height_cm, shapes_xml, anchor_par_text=None):
    """Inserts a w:p containing a v:group with given raw shapes XML string."""
    global body
    p = etree.SubElement(body, w("p"))
    r = etree.SubElement(p, w("r"))
    d = etree.SubElement(r, w("drawing"))
    anchor = etree.SubElement(d, wp("anchor"))
    for tag, attrs in (("distT",{}),("distB",{}),("distL",{}),("distR",{})):
        e = etree.SubElement(anchor, wp(tag)); e.text = "0"
    ext = etree.SubElement(anchor, wp("extent")); ext.set("cx", str(int(width_cm*EMU_CM))); ext.set("cy", str(int(height_cm*EMU_CM)))
    posH = etree.SubElement(anchor, wp("positionH")); etree.SubElement(posH, wp("relativeFrom")).text = "margin"
    etree.SubElement(posH, wp("posOffset")).text = "0"
    posV = etree.SubElement(anchor, wp("positionV")); etree.SubElement(posV, wp("relativeFrom")).text = "paragraph"
    etree.SubElement(posV, wp("posOffset")).text = "0"
    cnv = etree.SubElement(anchor, wp("cnvGrpPicture"))
    grp = etree.fromstring('<v:group xmlns:v="%s" xmlns:o="%s" xmlns:w="%s" coordsize="21600,21600" o:spt="100" style="mso-width-percent:0;mso-height-percent:0;" coordorigin="0,0,%d,%d">'
                           % (V, O, W, int(width_cm*400), int(height_cm*400)) + shapes_xml + "</v:group>")
    cnv.append(grp)
    rPr = etree.SubElement(d, w("docPr")); rPr.set("id", str(len(_drawing_ids)+1)); rPr.set("name", "Group")
    _drawing_ids.append(1)
    return p

def cm(v): return "%d.0pt" % round(v*28.35)  # point string helper (approx via pt)

def S(style, coordsize, spt, inner=""):
    return '<v:shape style="%s" coordsize="%s" o:spt="%d">%s</v:shape>' % (style, coordsize, spt, inner)

def rect(x,y,w_,h, fill=None, stroke="#000000", sw="1.25pt", extra=""):
    st = "position:absolute;left:%d;top:%d;width:%d;height:%d;" % (x,y,w_,h)
    if extra: st += extra
    f = 'fillcolor="%s"' % fill if fill else 'filled="f"'
    sv = ('<v:stroke weight="%s" color="%s"/>' % (sw, stroke)) if stroke else ""
    return '<v:rect o:allowincell="f" style="%s" coordsize="21600,21600" o:spt="1" %s>%s%s</v:rect>' % (st, f, sv, "")

def oval(x,y,w_,h, fill=None, stroke="#000000", sw="1.25pt", adj=None):
    st = "position:absolute;left:%d;top:%d;width:%d;height:%d;" % (x,y,w_,h)
    f = 'fillcolor="%s"' % fill if fill else 'filled="f"'
    sv = ('<v:stroke weight="%s" color="%s"/>' % (sw, stroke)) if stroke else ""
    ad = ' adj="%s"' % adj if adj else ""
    return '<v:oval o:allowincell="f" style="%s" coordsize="21600,21600" o:spt="9" %s%s>%s</v:oval>' % (st, f, ad, sv)

def line(x1,y1,x2,y2, color="#000000", sw="1.25pt", dash=None):
    st = "position:absolute;left:%d;top:%d;width:%d;height:%d;" % (min(x1,x2), min(y1,y2), abs(x2-x1) or 1, abs(y2-y1) or 1)
    d = 'o:dashstyle="%s"' % dash if dash else ""
    return ('<v:line o:allowincell="f" from="%d,%d" to="%d,%d" style="%s" strokeweight="%s" strokecolor="%s"%s><v:stroke %s endarrow="none"/></v:line>'
            % (x1,y1,x2,y2, st, sw, color, d, ('dashstyle="'+dash+'"') if dash else 'color="'+color+'"'))

def txt(x,y,w_,h, text, size=8, bold=False, align="center", color="#000000"):
    st = "position:absolute;left:%d;top:%d;width:%d;height:%d;text-align:%s;" % (x,y,w_,h,align)
    b = "<b>" if bold else ""
    fb = "Arial"
    return ('<v:rect o:allowincell="f" style="%s" coordsize="21600,21600" o:spt="1" filled="f" stroked="f">'
            '<v:textbox inset="0,0,0,0"><w:p><w:pPr><w:jc w:val="%s"/></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"/>%s<w:sz w:val="%d"/><w:color w:val="%s"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p></v:textbox></v:rect>'
            % (st, align, fb, fb, b, size*2, color.lstrip("#"), text))



def cshape(path, x, y, w_, h, fill="#FFFFFF", stroke="#000000", sw="1pt", rot=None):
    st = "position:absolute;left:%d;top:%d;width:%d;height:%d;" % (x,y,w_,h)
    if rot is not None: st += "rotation:%d;" % rot
    toks = re.findall(r"[a-zA-Z]|-?\d+", path.strip())
    parts = []
    i = 0
    while i < len(toks):
        cmd = toks[i]; nums = []
        j = i+1
        while j < len(toks) and not toks[j].isalpha():
            nums.append(toks[j]); j += 1
        parts.append((cmd, [int(v) for v in nums]))
        i = j
    allnums = [v for _, ns in parts for v in ns]
    minx = min(allnums[0::2]) if allnums else 0
    miny = min(allnums[1::2]) if len(allnums) > 1 else 0
    sx = -minx if minx < 0 else 0
    sy = -miny if miny < 0 else 0
    segs = []
    for cmd, ns in parts:
        out = []
        for k, v in enumerate(ns):
            out.append(str(v + (sx if k % 2 == 0 else sy)))
        segs.append(cmd + (" " + " ".join(out) if out else ""))
    fpath = "f:" + " ".join(segs)
    return ('<v:shape o:allowincell="f" type="%s" coordsize="21600,21600" o:spt="100" style="%s" fillcolor="%s" strokeweight="%s" strokecolor="%s"></v:shape>') % (fpath, st, fill, sw, stroke)

# ============================================================
# СТРАНИЦА 1: План школьной лаборатории (таблица)
# ============================================================
add_par("План школьной лаборатории", style="Title", align="center")

# Сетка таблицы плана: 10 колонок x 8 строк
NC, NR = 10, 8
colw = [2.0]*NC          # 2 см каждая -> 20 см ширина
rowh = [1.9]*NR         # высота строк
SOLID = ("single", 18)   # стена
def W_(*edges):
    return {e: SOLID for e in edges}

cells = {}
borders = {}

# Внешние стены: границы крайних ячеек задаём через tblBorders? Лучше по ячейкам:
for ri in range(NR):
    for ci in range(NC):
        b = {}
        if ri == 0: b["top"] = SOLID
        if ri == NR-1: b["bottom"] = SOLID
        if ci == 0: b["left"] = SOLID
        if ci == NC-1: b["right"] = SOLID
        if b: borders[(ri,ci)] = b

# Кабинет химии: строки 0-3, колонки 0-4  (merge обрабатывается span-логикой автоматически,
# если задать текст только в (0,0) и границу внутри). Чтобы слияние сработало, ставим текст в (0,0),
# а внутренние стены — на границах соседних ячеек.
cells[(0,0)] = ["Кабинет химии"]
# правая стена кабинета химии (между кол.4 и 5, строки 0..3) + нижняя стена (строка 3/4, кол 0..4)
for ri in range(0,4):
    borders.setdefault((ri,4), {}).update(W_("right"))
for ci in range(0,5):
    borders.setdefault((3,ci), {}).update(W_("bottom"))

# Кабинет физики: строки 0-3, колонки 6-9
cells[(0,6)] = ["Кабинет физики"]
for ri in range(0,4):
    borders.setdefault((ri,6), {}).update(W_("left"))
for ci in range(6,10):
    borders.setdefault((3,ci), {}).update(W_("bottom"))
# коридор между ними: колонка 5 сверху не закрыт -> входная зона; закроем верх стенкой кроме двери
for ci in [0,1,2,3,6,7,8,9]:
    pass

# Верхняя стена вестибюля (колонка 5, строка 0) закрыта, кроме прохода в коридор (ячейка (1,5))
borders.setdefault((0,5), {}).update(W_("top"))
for ri in range(0,4):
    borders.setdefault((ri,5), {}).update(W_("left","right"))
cells[(2,5)] = ["Вход"]

# Коридор: строка 4 целиком (колонки 0-9)
cells[(4,0)] = ["Коридор"]
# верхняя граница коридора: снизу у строк 3 (уже есть от кабинетов) и вестибюля:
for ci in [0,1,2,3,4,6,7,8,9]:
    borders.setdefault((3,ci), {}).update(W_("bottom"))
borders.setdefault((3,5), {}).update(W_("bottom"))
for ci in range(0,10):
    borders.setdefault((4,ci), {}).update(W_("bottom"))

# Нижние помещения: строки 5-7
# Лаборатория №1: кол 0-2 ; Лаборатория №2: кол 3-5 ; Склад реактивов: кол 6-7 ; Приборная: кол 8-9
cells[(5,0)] = ["Лаборатория №1"]
cells[(5,3)] = ["Лаборатория №2"]
cells[(5,6)] = ["Склад реактивов"]
cells[(5,8)] = ["Приборная"]
for ci in range(0,10):
    borders.setdefault((7,ci), {}).update(W_("bottom"))  # нижняя стена (уже есть), дублируем ок
# вертикальные стены между нижними помещениями
for ri in range(5,8):
    borders.setdefault((ri,2), {}).update(W_("right"))
    borders.setdefault((ri,3), {}).update(W_("left"))
    borders.setdefault((ri,5), {}).update(W_("right"))
    borders.setdefault((ri,6), {}).update(W_("left"))
    borders.setdefault((ri,7), {}).update(W_("right"))
    borders.setdefault((ri,8), {}).update(W_("left"))
# двери из коридора в лаборатории (разрывы стен): уберём bottom у некоторых ячеек коридора
for (r,c) in [(4,1),(4,4),(4,6),(4,8)]:
    borders[(r,c)].pop("bottom", None)

make_table(NR, NC, colw, rowh, cells, borders)
add_par("")

# ---------------- VML-рисунок поверх/рядом: столы и шкафы в лабораториях ----------------
# Простая схема мебели отдельным холстом под таблицей
sh = []
Wp, Hp = 20*400, 6*400  # 20см x 6см
# фон-рамка
sh.append(rect(0,0,Wp,Hp, fill="#FFFFFF", stroke="#444444", sw="1pt"))
sh.append(txt(10,10,3000,300,"Мебель в лабораториях:", size=9, bold=False, align="left"))
y0=700
# столы Лаборатории №1 (три сдвоенных стола)
for i in range(3):
    sh.append(rect(300+i*900, y0, 700, 500, fill="#D2B48C", stroke="#333333", sw="1pt"))
    sh.append(rect(300+i*900, y0+1400, 700, 500, fill="#D2B48C", stroke="#333333", sw="1pt"))
sh.append(txt(300, y0+2100, 2500, 300, "Лаборатория №1 — парты для учеников", size=8, align="left"))
# шкаф реактивов справа
sh.append(rect(6200, y0, 500, 1200, fill="#8FBC8F", stroke="#333333", sw="1pt"))
sh.append(rect(6200, y0+1400, 500, 500, fill="#8FBC8F", stroke="#333333", sw="1pt"))
sh.append(txt(6100, y0+2100, 900, 300, "Шкаф", size=8))
# вытяжной шкаф
sh.append(rect(7400, y0, 900, 1900, fill="#B0C4DE", stroke="#333333", sw="1pt"))
sh.append(txt(7300, y0+2100, 1100, 300, "Вытяжной шкаф", size=8))
# демонстрационный стол
sh.append(rect(8700, y0+300, 1600, 900, fill="#DEB887", stroke="#333333", sw="1pt"))
sh.append(txt(8700, y0+2100, 1600, 300, "Демонстрационный стол", size=8))
add_canvas(20, 6, "".join(sh))

# ============================================================
# СТРАНИЦА 2: Солнечный день на море, пляж (фигуры Word)
# ============================================================
pagebreak()
add_par("Солнечный день на море. Пляж", style="Title", align="center")

CW, CH = 16*400, 11*400   # холст 16 x 11 см
sh = []
# Небо
sh.append(rect(0,0,CW,int(CH*0.58), fill="#87CEEB", stroke=None))
# Море
sh.append(rect(0,int(CH*0.58),CW,int(CH*0.20), fill="#1E90FF", stroke=None))
# Песок (пляж)
sh.append(rect(0,int(CH*0.78),CW,int(CH*0.22), fill="#F4D03F", stroke=None))
# Солнце + лучи
sx, sy = int(CW*0.80), int(CH*0.14)
sh.append(oval(sx-450, sy-450, 900, 900, fill="#FFD700", stroke="#FFA500", sw="2pt"))
for ang in range(0,360,45):
    a1 = math.radians(ang); L1, L2 = 550, 850
    sh.append(line(sx+int(L1*math.cos(a1)), sy+int(L1*math.sin(a1)),
                   sx+int(L2*math.cos(a1)), sy+int(L2*math.sin(a1)), color="#FFA500", sw="2pt"))
# Облака
def cloud(cx, cy, s):
    out = []
    out.append(oval(cx, cy, int(700*s), int(350*s), fill="#FFFFFF", stroke=None))
    out.append(oval(cx+int(350*s), cy-int(150*s), int(700*s), int(420*s), fill="#FFFFFF", stroke=None))
    out.append(oval(cx+int(750*s), cy, int(650*s), int(330*s), fill="#FFFFFF", stroke=None))
    return "".join(out)
sh.append(cloud(int(CW*0.12), int(CH*0.12), 1.0))
sh.append(cloud(int(CW*0.42), int(CH*0.22), 0.8))
# Чайки
def bird(bx, by, s=1):
    return line(bx, by, bx+int(120*s), by-int(80*s), color="#333333", sw="1.5pt") + \
           line(bx+int(120*s), by-int(80*s), bx+int(240*s), by, color="#333333", sw="1.5pt")
sh.append(bird(int(CW*0.30), int(CH*0.30)))
sh.append(bird(int(CW*0.40), int(CH*0.35), 0.8))
# Волны на море
for i in range(5):
    wx = int(CW*(0.08+0.18*i)); wy = int(CH*(0.63+0.03*(i%2)))
    sh.append(oval(wx, wy, 700, 160, fill=None, stroke="#FFFFFF", sw="1.5pt"))
# Парусник
px, py = int(CW*0.55), int(CH*0.62)
sh.append(cshape("m,-60,0 l,60,0 l,30,30 l,-30,30 x e", px-60, py, 120, 30, fill="#8B4513"))
sh.append(cshape("m,0,0 l,0,-120 l,90,0 x e", px, py-120, 90, 120, fill="#FFFFFF"))
sh.append(cshape("m,0,0 l,0,-90 l,-70,0 x e", px-70, py-10, 70, 90, fill="#FF6347"))
# Зонт на пляже
ux, uy = int(CW*0.20), int(CH*0.80)
sh.append(line(ux, uy-600, ux, uy+900, color="#8B4513", sw="3pt"))
sh.append(cshape("m,-500,0 c,-500,-450 500,-450 500,0 x e", ux-500, uy-600, 1000, 450, fill="#FF69B4", sw="1.5pt"))
# Полотенце
sh.append(rect(int(CW*0.33), int(CH*0.86), 1400, 700, fill="#FF7F50", stroke="#E5593F", sw="1pt"))
# Морская звезда
mx, my = int(CW*0.72), int(CH*0.90)
sh.append(cshape("m,0,-260 l,90,-80 l,260,-80 l,120,40 l,170,230 l,0,120 l,-170,230 l,-120,40 l,-260,-80 l,-90,-80 x e", mx-260, my-260, 520, 520, fill="#FF8C00"))
# Пальма слева
tx, ty = int(CW*0.05), int(CH*0.82)
sh.append(cshape("m,0,0 c,30,-250 60,-450 40,-650 l,90,-650 c,110,-450 80,-250 50,0 x e", tx, ty-650, 120, 650, fill="#8B4513"))
for da in (-60,-20,20,60,100,140):
    ar = math.radians(da)
    sh.append(cshape("m,0,0 c,80,-60 160,-40 220,20 c,150,10 70,20 0,0 x e", tx+40, ty-650, 230, 120, fill="#228B22", rot=da))
sh.append(oval(tx+10, ty-680, 120, 120, fill="#6B8E23", stroke=None))
# Рамка вокруг всей картинки
sh.insert(0, rect(0,0,CW,CH, fill=None, stroke="#555555", sw="2pt"))
add_canvas(16, 11, "".join(sh))

# ============================================================
# Сборка docx
# ============================================================
sectPr = etree.SubElement(body, w("sectPr"))
pgSz = etree.SubElement(sectPr, w("pgSz")); pgSz.set(w("w"),"11906"); pgSz.set(w("h"),"16838")
pgMar = etree.SubElement(sectPr, w("pgMar"))
for k,v in (("top","1134"),("right","1134"),("bottom","1134"),("left","1134"),("header","709"),("footer","709"),("gutter","0")):
    pgMar.set(w(k), v)

doc_xml = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + etree.tostring(
    doc, xml_declaration=False)
_ns = (b'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
       b'xmlns:v="urn:schemas-microsoft-com:vml" '
       b'xmlns:o="urn:schemas-microsoft-com:office:office" '
       b'xmlns:w10="urn:schemas-microsoft-com:office:word" '
       b'mc:Ignorable="w14 wp14" ')
doc_xml = doc_xml.replace(b'<w:document ', b'<w:document ' + _ns, 1)

content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''

root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="%s">
<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
 <w:name w:val="Normal"/><w:pPr><w:spacing w:after="120"/></w:pPr>
 <w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr>
</w:style>
<w:style w:type="paragraph" w:styleId="Title">
 <w:name w:val="Title"/><w:basedOn w:val="Normal"/>
 <w:pPr><w:jc w:val="center"/><w:spacing w:after="240"/></w:pPr>
 <w:rPr><w:b/><w:sz w:val="36"/></w:rPr>
</w:style>
</w:styles>''' % W

if os.path.exists(OUT): os.remove(OUT)
zf = zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED)
zf.writestr("[Content_Types].xml", content_types)
zf.writestr("_rels/.rels", root_rels)
zf.writestr("word/_rels/document.xml.rels", doc_rels)
zf.writestr("word/document.xml", doc_xml)
zf.writestr("word/styles.xml", styles_xml)
zf.close()
print("OK:", OUT, os.path.getsize(OUT), "bytes")
