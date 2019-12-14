from docx import Document
from docx.shared import Cm, Inches

document = Document()

document.add_picture('/home/dadakhon/Pictures/gerb.png', width=Inches(3.25))

document.save('loyiha.docx')
