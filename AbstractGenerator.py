# -*- coding: utf-8 -*-
"""
Generate abstract document (docx) file from table (xlsx)
by nebula

Dependency: pandas, python-docx
"""
import pandas as pd
import docx
import docxtpl
import re
import os


class AbstractGenerator:
    def __init__(self):
        self.records = None
        self.exreg4super = re.compile(r'(\(\w+\))')
        self.exreg4italic = re.compile(r'(\<i\>\w+\</i\>)')

    def _insert_image(self, filename, image_filename):
        doc = docx.Document(filename)

        for paragraph in doc.paragraphs:
            if '[[FIGURE]]' in paragraph.text:
                paragraph.text = ''
                run = paragraph.add_run()
                run.add_picture(image_filename)

        doc.save(filename)

    def read_xlsx(self, filename):
        print('Reading: %s' % filename)
        exls = pd.ExcelFile(filename)
        self.records = exls.parse()

    def write_docx(self, filename):
        print('Writing: %s' % filename)

        doc = docx.Document()
        
        for i in self.records.index:

            # ID + Title
            p = doc.add_paragraph(('P-%03d ' % i) + self.records.title[i])
            p.runs[0].font.size = docx.shared.Pt(16)
            p.runs[0].bold = True

            # Authors
            p = doc.add_paragraph()
            author_list = self.exreg4super.split(self.records.authors[i])
            for j in range(len(author_list)):
                if j & 1:
                    p.add_run(author_list[j]).font.superscript = True
                else:
                    p.add_run(author_list[j])

            # Affiliations
            p = doc.add_paragraph(self.records.affiliations[i])
            p.runs[0].font.size = docx.shared.Pt(9)
            p.runs[0].italic = True

            # Abstract Body
            p = doc.add_paragraph(self.records.abstract[i])

            # keywords
            p = doc.add_paragraph('Keywords: ')
            p.add_run(self.records.keywords[i]).italic = True

            doc.add_page_break()

        doc.save(filename)

    def write_docx_with_template(self, filename, templatename, template_words, image_col=None, image_dir=''):
        print('Writing: %s with %s' % (filename, templatename))
        filename_base, filename_ext = os.path.splitext(filename)

        for i in self.records.index:
            doc = docxtpl.DocxTemplate(templatename)

            filename_single = str(i) + '.docx'
            context = {}

            print(self.records.loc[i,:])

            for word in template_words:
                context[word] = self.records.loc[i, word]

            print(context)

            doc.render(context)
            doc.save(filename_single)

            if image_col is not None:
                self._insert_image(filename_single, os.path.join(image_dir, context['FigureFileName']))

        # Joint docx files
        #files = ['0.docx', '1.docx']
        #self.combine_word_documents(files)

        doc = docx.Document()
        for i in self.records.index:
            filename_single = str(i) + '.docx'
            doc_single = docx.Document(filename_single)
            for element in doc_single._body._element:
                doc._body._element.append(element)

            doc.add_page_break()

        doc.save(filename)

    '''
    def combine_word_documents(self, files):
        combined_document = docx.Document()
        for file in files:
            sub_doc = docx.Document(file)

            for element in sub_doc._body._element:
                combined_document._body._element.append(element)

        combined_document.save('combined_word_documents.docx')
    '''

if __name__ == '__main__':

    template_words = [
        u'Title',
        u'Name',
        u'Affiliation',
        u'email',
        u'ProgramNo',
        u'DOI',
        u'Abstract',
        u'FigureFileName',
        u'FigureComment'
    ]


    abgen = AbstractGenerator()
    abgen.read_xlsx('./private/aini2016_example.xlsx')
    # abgen.write_docx('output.docx')
    abgen.write_docx_with_template('output.docx', 'template.docx', template_words, image_col='FigureFileName', image_dir='private')

