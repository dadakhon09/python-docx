from docx import Document
from docx.shared import Inches, Cm, RGBColor

document = Document()


def make_rows_bold(*rows):
    for row in rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True


def set_column_width(column, width):
    column.width = width
    for cell in column.cells:
        cell.width = width


table1 = document.add_table(rows=7, cols=3)
hdr_cells1 = table1.rows[0]
hdr_cells1.cells[0].text = 'ТАЙИНЛАНСИН:'
hdr_cells1.cells[1].text = ' '
hdr_cells1.cells[2].text = 'ИЧКИ ИШЛАР ВАЗИРЛИГИ МАРКАЗИЙ АППАРАТИ БЎЙИЧА:'
row2_cells = table1.rows[2].cells
row2_cells[0].text = ''
row2_cells[1].text = ''
row2_cells[2].text = 'полковник АААА ВВВВВ (Н-000000), штатларни қайта ташкил этилиши муносабати билан, 2018 йилнинг 31' \
                    'декабрь кунидан ЛЛЛЛЛЛ бошлиғи лавозимидан озод этилиб, ЛЛЛЛЛЛ бошлиғи лавозимига;'

row4_cells = table1.rows[4].cells
row4_cells[0].text = ''
row4_cells[1].text = ''
row4_cells[2].text = 'полковник АААА ВВВВВ ББББББ (А-000000), **** лавозимидан озод этилиб, махсус унвони ҳарбий унвонга' \
                     ' тенглаштирилиб, ***** лавозимига;'
row6_cells = table1.rows[6].cells
row6_cells[0].text = ''
row6_cells[1].text = ''
row6_cells[2].text = 'полковник АААА ВВВВВ ББББББ (А-000000), **** лавозимидан озод этилиб, махсус унвони ҳарбий' \
                     ' унвонга тенглаштирилиб, ***** лавозимига;'
table1.autofit = False

make_rows_bold(hdr_cells1)
set_column_width(table1.columns[1], Cm(0.5))
set_column_width(table1.columns[2], Cm(11))

blank = document.add_paragraph(' ')
run_blank = blank.add_run()
run_blank.add_break()

table2 = document.add_table(1, 1)
table2.rows[0].cells[0].text = '«Ўзбекистон Республикаси фуқаролари томонидан ҳарбий хизматни ўташ тартиби тўғрисидаги ' \
                               'Низом»нинг 36-моддаси, «е» бандига мувофиқ, резервдаги бор офицерлик унвони билан, ' \
                               'ихтиёрий равишда шартнома асосидаги ҳарбий хизматга чақирилиб,'

table3 = document.add_table(rows=4, cols=3)


document.save('namuna.docx')
