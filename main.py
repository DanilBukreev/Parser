from copy import deepcopy
from xml.dom.minidom import Document


def copy_table_after(table, paragraph):
    tbl, p = table._tbl, paragraph._p
    new_tbl = deepcopy(tbl)
    p.addnext(new_tbl)


def replaceText(document, search, replace):
    for table in document.tables:
        for row in table.rows:
            for paragraph in row.cells:
                if search in paragraph.text:
                    paragraph.text = replace


name = 'док.docx'
document = Document(name)
template = document.tables[0]
replaceText(document, '<<VALUE_TO_FIND>>', 'New value')
paragraph = document.add_paragraph()
copy_table_after(template, paragraph)
# pip install C:\path\to\downloaded\file\lxml-4.3.5-cp38-cp38-win32.whl
# pip install C:\Users\dbukreev\Downloads\lxml-4.4.2-cp38-cp38-win_amd64.whl   pip install lxml==3.6.0

