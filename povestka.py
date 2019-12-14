from docx import Document
from docx.shared import Cm, Inches

document = Document()

document.add_paragraph('')

document.save('povestka.docx')
