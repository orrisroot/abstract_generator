# -*- coding: utf-8 -*-
"""
Generate abstract document (docx) file from table (xlsx)
by nebula

Dependency: pandas, python-docx
"""
import pandas as pd
import docx
# import docxtpl
import re
import os


class AbstractGenerator:
    def __init__(self, image_dir=''):
        self.records = None
        self.image_dir = image_dir
        self.exreg4super = re.compile(r'(\(\w+\))')
        self.exreg4italic = re.compile(r'(\<i\>\w+\</i\>)')

    def _insert_image(self, filename, image_filename):
        doc = docx.Document(filename)

        for paragraph in doc.paragraphs:
            if '[[FIGURE]]' in paragraph.text:
                #paragraph.text = ''
                run = paragraph.add_run()
                run.add_paragraph()
                inline_shape = run.add_picture(image_filename, width=docx.shared.Pt(300))
                run.add_paragraph()

        doc.save(filename)

    def read_xlsx(self, filename):
        print('Reading: %s' % filename)
        exls = pd.ExcelFile(filename)
        self.records = exls.parse()

    def write_docx(self, filename, template=None):
        print('Writing: %s' % filename)

        if template is not None:
            doc = docx.Document(template)
        else:
            doc = docx.Document()

        for i in self.records.index:
            # self._write_doc_jscpb2016(doc, self.records.loc[i])
            self._write_doc_aini2016(doc, self.records.loc[i])

            doc.add_page_break()

        doc.save(filename)

    def _write_doc_jscpb2016(self, doc, record):
        print('"%s"' % record['title'])

        # ID + Title
        p = doc.add_paragraph(record.title)
        p.runs[0].font.size = docx.shared.Pt(16)
        p.runs[0].bold = True

        # Authors
        p = doc.add_paragraph()
        author_list = self.exreg4super.split(record.authors)
        for j in range(len(author_list)):
            if j & 1:
                p.add_run(author_list[j]).font.superscript = True
            else:
                p.add_run(author_list[j])


        # Affiliations
        p = doc.add_paragraph(record.affiliations)
        p.runs[0].font.size = docx.shared.Pt(9)
        p.runs[0].italic = True

        # Abstract Body
        p = doc.add_paragraph(record.abstract)

        # keywords
        p = doc.add_paragraph('Keywords: ')
        p.add_run(record.keywords).italic = True



    def _write_doc_aini2016(self, doc, record):
        print('"%s"' % record['Title'])

        font = doc.styles['Normal'].font
        font.size = docx.shared.Pt(10)
        font.name = 'Lucida Grande'


        # ID + Title
        p = doc.add_paragraph(record['Title'])
        p.runs[0].font.size = docx.shared.Pt(12)
        p.runs[0].bold = True
        p.runs[0].italic = True
        p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph()

        # Authors
        p = doc.add_paragraph()
        p.add_run(record['Name']).bold = True
        p.add_run(record['Affiliation'] + '\n' + record['email'])
        p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
        paragraph_format = p.paragraph_format
        paragraph_format.line_spacing = docx.shared.Pt(12)
        doc.add_paragraph()

        p = doc.add_paragraph(record['ProgramNo'] + '\n' + record['DOI'])

        # Abstract Body
        p = doc.add_paragraph(record['Abstract'])
        paragraph_format = p.paragraph_format
        paragraph_format.line_spacing = docx.shared.Pt(11)

        # Figure
        doc.add_picture(os.path.join(self.image_dir, record['FigureFileName']))
        p = doc.add_paragraph(record['FigureComment'])
        paragraph_format = p.paragraph_format
        paragraph_format.line_spacing = docx.shared.Pt(11)

        # Citation
        # split author's affiliation number
        author_list = self.exreg4super.split(record['Name'])
        author_tmp = ''
        for j in range(len(author_list)):
            if j & 1:
                pass
            else:
                author_tmp += author_list[j]
        p = doc.add_paragraph('Citation: ' + author_tmp + ', ' + record['Title'] + ', AINI 2016 Abstracts' + ', ' + record['DOI'])
        paragraph_format = p.paragraph_format
        paragraph_format.line_spacing = docx.shared.Pt(11)


if __name__ == '__main__':
    abgen = AbstractGenerator(image_dir='./private')
    abgen.read_xlsx('./private/aini2016_example.xlsx')
    abgen.write_docx('output.docx', './private/aini2016_template2.docx')

