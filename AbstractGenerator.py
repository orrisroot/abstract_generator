# -*- coding: utf-8 -*-
"""
Generate abstract document (docx) file from table (xlsx)
by nebula

Dependency: pandas, python-docx
"""
import pandas as pd
from docx import Document
from docx.shared import Pt
import re


class AbstractGenerator:
    def __init__(self):
        self.records = None

    def read_xlsx(self, filename):
        print('Reading: %s' % filename)
        exls = pd.ExcelFile(filename)
        self.records = exls.parse()

    def write_docx(self, filename):
        print('Writing: %s' % filename)

        doc = Document()
        super_match = re.compile(r'(\(\w+\))')
        for i in self.records.index:

            # ID + Title
            p = doc.add_paragraph(('P-%03d ' % i) + self.records.title[i])
            p.runs[0].font.size = Pt(16)
            p.runs[0].bold = True

            # Authors
            p = doc.add_paragraph()
            author_list = super_match.split(self.records.authors[i])
            for j in range(len(author_list)):
                if j & 1:
                    p.add_run(author_list[j]).font.superscript = True
                else:
                    p.add_run(author_list[j])

            # Affiliations
            p = doc.add_paragraph(self.records.affiliations[i])
            p.runs[0].font.size = Pt(9)
            p.runs[0].italic = True

            # Abstract Body
            p = doc.add_paragraph(self.records.abstract[i])

            # keywords
            p = doc.add_paragraph('Keywords: ')
            p.add_run(self.records.keywords[i]).italic = True

            doc.add_page_break()

        doc.save(filename)


if __name__ == '__main__':

    abgen = AbstractGenerator()

    abgen.read_xlsx('input.xlsx')
    abgen.write_docx('output.docx')

