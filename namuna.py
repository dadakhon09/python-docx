from docx import Document
from docx.shared import Inches, Cm, RGBColor

document = Document()


def resize_table(table):
    set_column_width(table.columns[1], Cm(0.5))
    set_column_width(table.columns[2], Cm(10))


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
row4_cells[
    2].text = 'полковник АААА ВВВВВ ББББББ (А-000000), **** лавозимидан озод этилиб, махсус унвони ҳарбий унвонга' \
              ' тенглаштирилиб, ***** лавозимига;'
row6_cells = table1.rows[6].cells
row6_cells[0].text = ''
row6_cells[1].text = ''
row6_cells[2].text = 'полковник АААА ВВВВВ ББББББ (А-000000), **** лавозимидан озод этилиб, махсус унвони ҳарбий' \
                     ' унвонга тенглаштирилиб, ***** лавозимига;'
table1.autofit = False

make_rows_bold(hdr_cells1)

resize_table(table1)

blank = document.add_paragraph(' ')
run_blank = blank.add_run()
run_blank.add_break()

table2 = document.add_table(1, 1)
table2.rows[0].cells[0].text = '«Ўзбекистон Республикаси фуқаролари томонидан ҳарбий хизматни ўташ тартиби тўғрисидаги ' \
                               'Низом»нинг 36-моддаси, «е» бандига мувофиқ, резервдаги бор офицерлик унвони билан, ' \
                               'ихтиёрий равишда шартнома асосидаги ҳарбий хизматга чақирилиб,'

table3 = document.add_table(rows=10, cols=3)
row2_cells_t3 = table3.rows[1].cells
row2_cells_t3[0].text = 'ЮБОРИЛСИН:'
row2_cells_t3[2].text = 'Бухоро вилояти Пешку тумани мудофаа ишлари бўйича бўлимининг ҳарбий ҳисобида турган захирадаги' \
                        ' лейтенант ААААА ВВВВВ БББББ (А-000000), Қоровул қўшинлари қўмондони ихтиёрига;'
table3.rows[3].cells[2].text = 'Асос:'

table3.rows[5].cells[0].text = 'ЮБОРИЛСИН:'
table3.rows[5].cells[2].text = 'ИЧКИ ИШЛАР ВАЗИРЛИГИ МАРКАЗИЙ АППАРАТИ БЎЙИЧА:'
table3.rows[7].cells[2].text = 'подполковник ААААА АААА ААААА (А-014296), штатларни қайта ташкил этилиши муносабати' \
                               ' билан, 2018 йилнинг 31 декабрь кунидан **** лавозимидан озод этилиб, келгуси хизмат' \
                               ' фаолиятини давом эттириш учун Жазони ижро этиш бош бошқармаси ихтиёрига;'
table3.rows[9].cells[2].text = 'Асос:'

resize_table(table3)

blank = document.add_paragraph(' ')
run_blank = blank.add_run()
run_blank.add_break()

table4 = document.add_table(rows=3, cols=3)

table4.rows[0].cells[0].text = 'ТАСДИҚЛАНСИН:'
table4.rows[0].cells[2].text = 'САМАРҚАНД ВИЛОЯТИ ИЧКИ ИШЛАР БОШҚАРМАСИ БЎЙИЧА:'
table4.rows[2].cells[2].text = 'подполковник ААА ААА ААА (А-), **** бошлиғи лавозимида;'

resize_table(table4)

blank = document.add_paragraph(' ')
run_blank = blank.add_run()
run_blank.add_break()

table5 = document.add_table(rows=3, cols=3)

table5.rows[0].cells[0].text = 'ҚОЛДИРИЛСИН:'
table5.rows[0].cells[
    2].text = 'подполковник ААА ААА ААА (А-), **** бошлиғи лавозимидан озод этилиб, бошқарма ихтиёрида.'
table5.rows[2].cells[2].text = 'Асос:'

resize_table(table5)

blank = document.add_paragraph(' ')
run_blank = blank.add_run()
run_blank.add_break()

table6 = document.add_table(rows=7, cols=3)
hdr = table6.rows[0].cells[0].merge(table6.rows[0].cells[2])
hdr.text = 'Ўзбекистон Республикаси ИИВнинг 2018 йил 30 декабрдаги 2/4707-сонли кўрсатмаси талабларига мувофиқ, қуйидаги' \
           ' ходимлар 2019 йилнинг 14 январь кунидан 14 апрель кунига қадар бошланғич тайёргарликдан ўтишлари учун '

table6.rows[2].cells[0].text = 'ЮБОРИЛСИН:'
table6.rows[2].cells[2].text = 'ЎЗБЕКИСТОН РЕСПУБЛИКАСИ ИИВ АППАРАТИ БЎЙИЧА:'
row4 = table6.rows[4].cells[0].merge(table6.rows[4].cells[2])
row4.text = 'Вазирлик Ички ишлар органлари ходимларини бошланғич тайёрлаш ва малакасини ошириш маркази (Самарқанд шаҳар)' \
            ' ихтиёрига'
table6.cell(6, 2).text = '**** сафдор ходими сафдор ААА ААА ААА (А-);'

resize_table(table6)

blank = document.add_paragraph(' ')
run_blank = blank.add_run()
run_blank.add_break()

table7 = document.add_table(rows=13, cols=3)
table7.cell(0, 0).text = 'ЎЗГАРТИРИШ КИРИТИЛСИН:'
table7.cell(0, 2).text = '**** сафдор ААА ААА АААнинг (А-) ҳисоб-тавсифловчи ҳужжатларига «АААА АААА АААА »-деб.'
table7.cell(2, 2).text = 'Асос:'
table7.cell(4, 0).text = 'БЕРИЛСИН:'
table7.cell(4, 2).text = 'Мирзо Улуғбек  номидаги  Ўзбекистон Миллий  университетининг сиртқи бўлимида таҳсил олаётган:'
table7.cell(6, 2).text = '**** сафдор ААА ААА ААА (А-), 2019 йилнинг 3 январь кунидан 30 январь кунига қадар тўловли ' \
                         'ўқув таътили;'
table7.cell(8, 2).text = 'Асос:'
table7.cell(10, 2).text = '*** подполковник ААА ААА АААга (А-), 2019 йилнинг 08 январь кунидан 2019 йилнинг 13 май' \
                          ' кунига қадар хомиладорлик ва туғиш таътили. \n2019 йилнинг 14 май кунидан хизмат вазифасини ' \
                          'бажаришга киришиши белгилансин;'
table7.cell(12, 2).text = 'Асос:'

resize_table(table7)

blank = document.add_paragraph(' ')
run_blank = blank.add_run()
run_blank.add_break()

table8 = document.add_table(rows=9, cols=3)
table8.cell(0, 0).text = 'ЮБОРИЛСИН:'
table8.cell(0, 2).text = 'ИЧКИ ИШЛАР ВАЗИРЛИГИ МАРКАЗИЙ АППАРАТИ БЎЙИЧА:'
table8.cell(2, 2).text = 'подполковник ААА ААА ААА (А-00), 2019 йилнинг 9 январь кунидан вазирлик кадрларининг амалдаги' \
                         ' захирасидан чиқарилиб, вазирлик ихтиёрида бўлган деб ҳисобланиб, буйруқ имзоланган кундан' \
                         ' Қорақалпоғистон Республикаси Ички ишлар вазирлиги ихтиёрига;'
table8.cell(4, 2).text = 'Асос:'
table8.cell(6, 2).text = 'подполковник ААА ААА ААА (А-), **** лавозимидан озод этилиб, келгуси хизмат фаолиятини давом' \
                         ' эттириш учун Қашқадарё вилояти ички ишлар бошқармаси ихтиёрига;'
table8.cell(8, 2).text = 'Асос:'

resize_table(table8)

blank = document.add_paragraph(' ')
run_blank = blank.add_run()
run_blank.add_break()

table9 = document.add_table(rows=10, cols=3)
table9.cell(0, 0).text = 'ҚЎЛЛАНИЛСИН:'
table9.cell(0, 2).text = 'сабаблар, **** майор ААА ААА АААга (А-) “Қаттиқ ҳайфсан” интизомий жазоси;'
table9.cell(2, 2).text = 'Асос:'
row4 = table9.rows[4].cells[0].merge(table9.rows[4].cells[2])
row4.text = '«Ички ишлар органларида хизматни ўташ тартиби тўғрисида»ги Низом талабларига мувофиқ'
table9.cell(6, 2).text = 'ҚУРОЛЛИ КУЧЛАР ЗАХИРАСИГА'
table9.cell(7, 0).text = 'БЎШАТИЛСИН:'
table9.cell(8, 0).text = '144-бандининг «к» кичик бандига мувофиқ (ички ишлар органи ходимининг шаънига путур етказувчи' \
                         ' хатти-ҳаракатлар содир этганлиги учун)'
table9.cell(8, 2).text = '**** бошлиғи (Самарқанд шаҳри) подполковник ААА ААА ААА (А-0).'
table9.cell(9, 2).text = 'Асос:'

resize_table(table9)

blank = document.add_paragraph(' ')
run_blank = blank.add_run()
run_blank.add_break()

table10 = document.add_table(rows=10, cols=3)
table10.cell(0, 0).text = 'ТАЙИНЛАНСИН:'
table10.cell(0, 2).text = 'ИЧКИ ИШЛАР ВАЗИРЛИГИ МАРКАЗИЙ АППАРАТИ БЎЙИЧА:'
table10.cell(2, 2).text = 'майор АААА ААА ААА  (А-), *** лавозимидан озод этилиб, *** лавозимига;'
table10.cell(4, 2).text = 'майор ААА ААА ААА (А-), *** бошлиғи лавозимидан озод этилиб, “майор” махсус унвони ҳарбий ' \
                          'унвонга тенглаштирилиб, *** лавозимига;'
table10.cell(6, 2).text = 'капитан АААА ААА А (А-), **** лавозимидан озод этилиб, ўша ернинг ўзида **** лавозимига;'
table10.cell(8, 2).text = 'Жазони ижро этиш бош бошқармасидан келган капитан ААА ААА ААА (А-), *** лавозимига;'

resize_table(table10)

blank = document.add_paragraph(' ')
run_blank = blank.add_run()
run_blank.add_break()

table11 = document.add_table(rows=3, cols=3)
table11.cell(0, 0).text = 'ЮБОРИЛСИН:'
table11.cell(0, 2).text = 'полковник ААА ААА ААА (К-), штатлар қайта ташкил этилиши муносабати билан, 2018 йилнинг ' \
                          '10 сентрябрь кунидан *** лавозимидан озод этилиб, вазирлик ихтиёрида бўлган деб ҳисобланиб, ' \
                          'буйруқ имзоланган кундан келгуси хизмат фаолиятини давом эттириш учун вазирлик Академияси ихтиёрига;'
table11.cell(2, 2).text = 'Асос:'

resize_table(table11)

blank = document.add_paragraph(' ')
run_blank = blank.add_run()
run_blank.add_break()

table12 = document.add_table(rows=9, cols=3)
table12.cell(0, 0).text = 'БЕРИЛСИН:'
table12.cell(0, 2).text = 'Россия Федерациясининг Қозон федерал университети Тошкент шаҳар филиалининг сиртқи бўлимида ' \
                          'таҳсил олаётган:'
table12.cell(2, 2).text = '**** сержант АА ААА ААА (А-0), 2019 йилнинг 21 январь кунидан 11 февраль кунига тўловли ўқув' \
                          ' таътили;'
table12.cell(4, 2).text = 'Асос:'
table12.cell(6, 2).text = '*** сафдор ААА ААА ААА (А-), 2018 йилнинг 12 январь кунидан 2020 йилнинг 20 ноябрь қадар ' \
                          'бола икки ёшга тўлгунча парваришлаш учун иш ҳақи сақланмаган ҳолда таътили.\n2020 йилнинг 21' \
                          ' ноябрь кунидан хизмат вазифасини бажаришга киришиши белгилансин.'
table12.cell(8, 2).text = 'Асос:'

resize_table(table12)

blank = document.add_paragraph(' ')
run_blank = blank.add_run()
run_blank.add_break()

table13 = document.add_table(rows=3, cols=3)
table13.cell(0, 0).text = 'ҲИСОБЛАНСИН:'
table13.cell(0, 2).text = '*** сафдор ААА ААА ААА (А-), бола парваришлаш таътил муддати тугашидан олдин 2019 йилнинг ' \
                          '01 февраль кунидан хизмат вазифасини бажаришга киришган деб. '
table13.cell(2, 2).text = 'Асос:'

resize_table(table13)

blank = document.add_paragraph(' ')
run_blank = blank.add_run()
run_blank.add_break()

table14 = document.add_table(rows=3, cols=3)
table14.cell(0, 0).text = 'ЭЪЛОН ҚИЛИНСИН:'
table14.cell(0, 2).text = 'сабаб **** майор ААА ААА ААА (А-) “Қаттиқ ҳайфсан”;'
table14.cell(2, 2).text = 'Асос:'


document.save('namuna.docx')
