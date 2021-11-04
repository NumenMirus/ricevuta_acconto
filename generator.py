import docx
from docx import Document
from docx.text.paragraph import Paragraph
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Pt
from docx.shared import RGBColor
from docx.shared import Mm
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn

def insertHR(paragraph, size, space):
    p = paragraph._p  # p is the <w:p> XML element
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    pPr.insert_element_before(pBdr,
        'w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap',
        'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN',
        'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind',
        'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc',
        'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap',
        'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr',
        'w:pPrChange'
    )
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), str(size))
    bottom.set(qn('w:space'), str(space))
    bottom.set(qn('w:color'), 'auto')
    pBdr.append(bottom)

d = Document()

p1 = d.add_paragraph()
p1.paragraph_format.line_spacing = 1
p1.alignment = 0
p1.add_run("Sara Ezelina Vantini\n").bold = True
p1.add_run("via Don Carlo Gnocchi 19\n")
p1.add_run("20148 MILANO\n")
p1.add_run("Codice Fiscale:\n").bold = True
p1.add_run("VNTSZL74E71F205T\n").bold = True

p2 = d.add_paragraph()
p2.paragraph_format.line_spacing = 1
p2.alignment = 2
p2.add_run("Spett.le                                       \n")
p2.add_run("la LUSac srl                              \n").bold = True
p2.add_run("via Defendente Sacchi 13     \n")
p2.add_run("27100 PAVIA                            \n\n")
p2.add_run("Partita IVA 02778180188\n").bold = True

p3 = d.add_paragraph()
p3.paragraph_format.line_spacing = 1
p3.alignment = 0
p3.add_run("Ricevuta n 3/2021\n")
p3.add_run("Milano 15 Giugno 2021\n")
p3.add_run("Ricevo dalla Società la LUSac srl le somme sotto specificate a fronte delle prestazioni rientranti in rapporto di prestazione di lavoro autonomo occasionale*\n")


table = d.add_table(1,1, style="Light List")
table.add_column(Mm(30))
table.alignment = WD_TABLE_ALIGNMENT.CENTER

cell = table.rows[0].cells
cell[0].text = "Tipologia della prestazione"
cell[0].vertical_alignment = WD_TABLE_ALIGNMENT.CENTER
cell[1].text = "costo netto cadauno"
cell[1].horizontal_alignment = WD_TABLE_ALIGNMENT.CENTER
cell[1].vertical_alignment = WD_TABLE_ALIGNMENT.CENTER

row0 = table.rows[0]
row0.height = Mm(10)

###################################################################################

item_list = 

for i in len(item_list):






# table = d.add_table(1, 1)
# table.alignment = WD_TABLE_ALIGNMENT.CENTER
# table.add_column(Mm(30))

# cell = table.rows[0].cells
# cell[0].vertical_alignment = WD_TABLE_ALIGNMENT.CENTER
# p0 = cell[0].add_paragraph()
# p0.alignment = 1
# run = p0.add_run("Tipologia della prestazione")
# run.font.size = Pt(12)
# run.font.name = 'Arial'
# run.bold = True
# insertHR(p0)

# p1 = cell[1].add_paragraph()
# p1.alignment = 1
# run = p1.add_run("costo netto cadauno")
# run.font.size = Pt(12)
# run.font.name = 'Arial'
# run.bold = True

d.save("test.docx")