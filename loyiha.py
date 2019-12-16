from docx import Document
from docx.shared import Cm, Inches

document = Document()



def set_column_width(column, width):
    column.width = width
    for cell in column.cells:
        cell.width = width

        
def resize_table(table):
    set_column_width(table.columns[1], Cm(0.5))
    set_column_width(table.columns[2], Cm(13))


def make_rows_bold(*rows):
    for row in rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True



p = document.add_paragraph()
run = p.add_run()
p.alignment = 1
gerb = run.add_picture('/home/dadakhon/Pictures/gerb.png', width=Inches(1.05))


p = document.add_paragraph('ЎЗБЕКИСТОН РЕСПУБЛИКАСИ ИЧКИ ИШЛАР ВАЗИРИНИНГ\nБУЙРУҒИ')
run = p.runs[0]
p.alignment = 1
run.font.bold = True

p = document.add_paragraph('Шахсий таркиб бўйича')
run = p.runs[0]
p.alignment = 1
run.font.bold = True



table1 = document.add_table(rows=17, cols=3)
table1.cell(0, 0).text = 'ТАЙИНЛАНСИН:'
table1.cell(0, 2).text = 'ИЧКИ ИШЛАР ВАЗИРЛИГИ МАРКАЗИЙ АППАРАТИ БЎЙИЧА:'
table1.cell(2, 2).text = 'подполковник АААА АААА ААА (А-), **** лавозимидан озод этилиб, **** бошлиғи лавозимига;'
table1.cell(4, 0).text = 'ҚОЛДИРИЛСИН:'
table1.cell(4, 2).text = 'ҚОРАҚАЛПОҒИСТОН РЕСПУБЛИКАСИ ИИВ БЎЙИЧА:'
table1.cell(6, 2).text = 'подполковник ААА АААА ААА (А-0), *** лавозимидан озод этилиб, шу вазирлик ихтиёрида;'
table1.cell(8, 2).text = 'Асос:'
table1.rows[10].cells[0].merge(table1.rows[10].cells[2]).text = ' «Ички ишлар органларида хизматни ўташ тартиби тўғрисида»ги Низом талабларига мувофиқ '
table1.cell(12, 2).text = 'ҚУРОЛЛИ КУЧЛАР ЗАХИРАСИГА'
table1.cell(13, 0).text = 'БЎШАТИЛСИН:'
table1.cell(14, 0).text = '144-бандининг «к» кичик бандига мувофиқ (ички ишлар органи ходимининг шаънига путур етказувчи хатти-ҳаракатлар содир этганлиги учун) топширган'
table1.cell(14, 2).text = '**** бошлиғи (Самарқанд шаҳри) подполковник ААА АААА АААА (А-).'
table1.cell(15, 2).text = 'Асос:'



resize_table(table1)

blank = document.add_paragraph(' ')
run_blank = blank.add_run()
run_blank.add_break()


table2 = document.add_table(rows=10, cols=2)
table2.cell(1, 0).text = 'Вазир'
table2.cell(2, 0).text = 'генерал-лейтенант'
table2.cell(2, 1).text = 'П.Р. Бобожонов'
table2.cell(5, 0).text = 'Тошкент шаҳри,'
table2.cell(6, 0).text = '2019  йил «_____» январь'
table2.cell(7, 0).text = '_____-сон'


set_column_width(table2.columns[0], Cm(5))
set_column_width(table2.columns[1], Cm(13))

blank = document.add_paragraph(' ')
run_blank = blank.add_run()
run_blank.add_break()


table3 = document.add_table(1, 2)
table3.cell(0, 0).text = '****\nполковник'
table3.cell(0, 1).text = 'А.А.ААААА'


set_column_width(table3.columns[0], Cm(11))
set_column_width(table3.columns[1], Cm(4))

blank = document.add_paragraph(' ')
run_blank = blank.add_run()
run_blank.add_break()

p = document.add_paragraph('К Е Л И Ш И Л Г А Н :')
run = p.runs[0]
p.alignment = 1
run.font.bold = True


table4 = document.add_table(rows=4, cols=2)
table4.cell(0, 0).text = '*****\nподполковник'
table4.cell(0, 1).text = 'А.А.ААААА'
table4.cell(1, 0).text = '*****\nподполковник'
table4.cell(1, 1).text = 'А.А.ААААА'
table4.cell(2, 0).text = '*****\nподполковник'
table4.cell(2, 1).text = 'А.А.ААААА'
table4.cell(3, 0).text = '*****\nподполковник'
table4.cell(3, 1).text = 'А.А.ААААА'



set_column_width(table3.columns[0], Cm(11))
set_column_width(table3.columns[1], Cm(4))

blank = document.add_paragraph(' ')
run_blank = blank.add_run()
run_blank.add_break()

p = document.add_paragraph('Т А Қ С И М О Т :')
run = p.runs[0]
p.alignment = 1
run.font.bold = True

table5 = document.add_table(rows=16, cols=2)
table5.cell(0, 0).text = 'ТД Котибияти'
table5.cell(1, 0).text = 'КББ'
table5.cell(2, 0).text = 'МваМТТББ'
table5.cell(3, 0).text = 'Қорақалпоғистон Республикаси ИИВ'
table5.cell(4, 0).text = 'Андижон вилояти ИИБ'
table5.cell(5, 0).text = 'Бухоро вилояти ИИБ'
table5.cell(6, 0).text = 'Қашқадарё вилояти ИИБ'
table5.cell(7, 0).text = 'Навоий вилояти ИИБ'
table5.cell(8, 0).text = 'Наманган вилояти ИИБ'
table5.cell(9, 0).text = 'Самарқанд вилояти ИИБ'
table5.cell(10, 0).text = 'Тошкент шаҳар ИИББ '
table5.cell(11, 0).text = 'Фарғона вилояти ИИБ '
table5.cell(15, 0).text = 'Ж  А  М  И '


table5.cell(0, 1).text = '1 нусха'
table5.cell(1, 1).text = '1 нусха'
table5.cell(2, 1).text = '1 нусха'
table5.cell(3, 1).text = '1 нусха'
table5.cell(4, 1).text = '1 нусха'
table5.cell(5, 1).text = '1 нусха'
table5.cell(6, 1).text = '1 нусха'
table5.cell(7, 1).text = '1 нусха'
table5.cell(8, 1).text = '1 нусха'
table5.cell(9, 1).text = '1 нусха'
table5.cell(10, 1).text = '1 нусха'
table5.cell(11, 1).text = '1 нусха'
table5.cell(15, 1).text = 'нусха'


set_column_width(table5.columns[0], Cm(10))
set_column_width(table5.columns[1], Cm(10))

blank = document.add_paragraph(' ')
run_blank = blank.add_run()
run_blank.add_break()


p = document.add_paragraph('*****\nмайор')
p.alignment = 0


p = document.add_paragraph('А.А.ААААА')
p.alignment = 2


for i in range(0, 8):
	blank = document.add_paragraph(' ')
	run_blank = blank.add_run()
	run_blank.add_break()


p = document.add_paragraph('Бажарди: *****\nкапитан А.А.ААААА тел.00-00')


sections = document.sections
for section in sections:
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)


make_rows_bold(table1.rows[0], table1.rows[4], table1.rows[10], table1.rows[12], table1.rows[13], table1.rows[15], table2.rows[1], table2.rows[2], table5.rows[15])



document.save('loyiha.docx')
