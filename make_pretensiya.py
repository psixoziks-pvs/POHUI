# -*- coding: utf-8 -*-
"""
Выполнение практической работы из PDF "24.09.26 конспект+ практическая работа.pdf"
Создание документа Word "Претензия" согласно заданиям 1-15.
"""
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# ---- Задание 2: поля — верхнее, нижнее, левое 2 см, правое 1 см ----
sec = doc.sections[0]
sec.top_margin = Cm(2)
sec.bottom_margin = Cm(2)
sec.left_margin = Cm(2)
sec.right_margin = Cm(1)

# ---- Задание 8.1: базовый шрифт Times New Roman, 12 pt ----
style = doc.styles['Normal']
style.font.name = 'Times New Roman'
style.font.size = Pt(12)
rpr = style.element.get_or_add_rPr()
rfonts = rpr.find(qn('w:rFonts'))
if rfonts is None:
    rfonts = OxmlElement('w:rFonts')
    rpr.append(rfonts)
for attr in ('w:ascii', 'w:hAnsi', 'w:cs'):
    rfonts.set(qn(attr), 'Times New Roman')

# ---- Задания 3-4: абзац с отступом слева ~10,25 см, межстрочный интервал 1 ----
# (маркер "Отступ слева" на линейке перетянут на 10,25 см)
p_head = doc.add_paragraph()
pf = p_head.paragraph_format
pf.left_indent = Cm(10.25)
pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
pf.space_before = Pt(0)
pf.space_after = Pt(0)

# ---- Задание 5: текст заявления (шапка) ----
head_lines = [
    'Руководителю ООО Рога и копыта',
    'от И.И. Иванова,',
    'проживающему по адресу ул. Московская,',
    'д. 28',
    'г. Кимры Тверской обл.',
]
for i, line in enumerate(head_lines):
    if i == 0:
        run = p_head.add_run(line)
    else:
        p = doc.add_paragraph()
        ppf = p.paragraph_format
        ppf.left_indent = Cm(10.25)
        ppf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        ppf.space_before = Pt(0)
        ppf.space_after = Pt(0)
        p.add_run(line)

# ---- Задание 6: выравнивание По центру; Задание 7: заголовок ----
# (все маркеры линейки на 0 — для данного абзаца отступы не задаём)
p_title = doc.add_paragraph()
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_t = p_title.add_run('ПРЕТЕНЗИЯ')
run_t.bold = True

# ---- Задание 8: первая строка 1,25 см, межстрочный 1,5 строки; Задание 9: текст ----
body_paras = [
    '27 ноября 2017 года я предоставил в вашу фирму автомобиль марки Вольво '
    '(государственный номер №555555, номер двигателя 222, номер кузова, номер шасси) '
    'для проведения ремонта.',

    'В соответствии с договором (номер квитанции №222) оплатил стоимость работ '
    '(две тысячи рублей). Ваша организация обязалась отремонтировать узлы и детали ***.',

    'К моменту получения автомобиля, обнаружил, что узлы отремонтированы не качественно.',

    'В соответствии со статьёй 4 Закона РФ О защите прав потребителей, исполнитель '
    'обязан выполнить работы качественно.',

    'В соответствии со статьёй 30 указанного закона, потребитель может потребовать *****.',

    'На основании вышеизложенного требую: безвозмездно устранить недостатки в '
    'выполненной работе в течении 14 дней.',

    'В случае неудовлетворения претензии я буду вынужден обратиться в суд. Буду '
    'просить суд также взыскать с вашей организации взыскать моральный ущерб '
    '(статья 13 Закона).',

    'Ответ на претензию прошу дать в 20 дневный срок.',
]
for text in body_paras:
    p = doc.add_paragraph(text)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    ppf = p.paragraph_format
    ppf.first_line_indent = Cm(1.25)   # Первая строка – 1,25 см
    ppf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE  # Межстрочный – 1,5 строки

# Пустая строка перед таблицей
sp = doc.add_paragraph()
sp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE

# ---- Задание 10: таблица 5 столбцов x 1 строка ----
table = doc.add_table(rows=1, cols=5)
table.alignment = WD_TABLE_ALIGNMENT.LEFT

# ---- Задание 11: текст в ячейках (выравнивание как на рисунке) ----
cells = table.rows[0].cells
cell_data = [
    ('27', WD_ALIGN_PARAGRAPH.CENTER),
    ('ноября', WD_ALIGN_PARAGRAPH.CENTER),
    ('2017 года', WD_ALIGN_PARAGRAPH.CENTER),
    ('И.И. Иванов', WD_ALIGN_PARAGRAPH.CENTER),
    ('', WD_ALIGN_PARAGRAPH.LEFT),
]
for c, (txt, align) in zip(cells, cell_data):
    cp = c.paragraphs[0]
    cp.alignment = align
    cp.paragraph_format.space_before = Pt(0)
    cp.paragraph_format.space_after = Pt(0)
    cp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    if txt:
        run = cp.add_run(txt)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)

# ---- Задание 13: границы таблицы — Нет границ ----
tbl = table._tbl
tblPr = tbl.tblPr
borders = OxmlElement('w:tblBorders')
for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
    e = OxmlElement(f'w:{edge}')
    e.set(qn('w:val'), 'none')
    e.set(qn('w:sz'), '0')
    e.set(qn('w:space'), '0')
    borders.append(e)
tblPr.append(borders)

# ---- Задание 14: нижняя граница у ячейки с подписью (И.И. Иванов) ----
sig_cell = cells[3]
tc = sig_cell._tc
tcPr = tc.get_or_add_tcPr()
cell_borders = OxmlElement('w:tcBorders')
bottom = OxmlElement('w:bottom')
bottom.set(qn('w:val'), 'single')
bottom.set(qn('w:sz'), '6')      # 0.75 pt
bottom.set(qn('w:space'), '0')
bottom.set(qn('w:color'), '000000')
cell_borders.append(bottom)
tcPr.append(cell_borders)

# ---- Задание 14: Автоподбор — по содержимому и по ширине окна ----
layout_auto = OxmlElement('w:tblLayout')
layout_auto.set(qn('w:type'), 'autofit')
tblPr.append(layout_auto)
tblW = tblPr.find(qn('w:tblW'))
if tblW is None:
    tblW = OxmlElement('w:tblW')
    tblPr.append(tblW)
tblW.set(qn('w:w'), '5000')
tblW.set(qn('w:type'), 'pct')    # ширина 100% окна

# ---- Задание 15: сохранить документ ----
out = '/workspace/Практическая работа Претензия.docx'
doc.save(out)
print('Сохранено:', out)
