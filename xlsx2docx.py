from AbstractGenerator import AbstractGenerator

if __name__ == '__main__':
    import sys

    argvs = sys.argv
    argc = len(argvs)

    if(argc < 3):
        print('  USEAGE:\n   $ python %s [input xlsx] [output docx]' % argvs[0])
        exit()
    else:
        input_xlsx = argvs[1]
        output_docx = argvs[2]
        img_dir = ''
        template_docx = ''

    if(argc >= 4):
        img_dir = argvs[3]
    if(argc >= 5):
        template_docx = argvs[4]
        

    abgen = AbstractGenerator(img_dir)
    abgen.read_xlsx(input_xlsx)
    abgen.write_docx(output_docx, template_docx)
