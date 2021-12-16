from docx2pdf import convert
import time
import sys
import os
import comtypes.client


class Convert:
    def __init__(self, filename):
        self.toPdf(filename)

    def toPdf(self, filename):

        ext = filename.split('.')[-1]

        if ext == 'docx':
            try:
                convert('input/', 'output/')
            except Exception as e:
                print(e)
        elif ext == 'doc':
            try:
                wdFormatPDF = 17
                BASE = os.path.abspath(os.getcwd())
                inFile = os.path.join(BASE, 'input', filename)
                outFile = os.path.join(BASE, 'output', filename)

                print(os.path.abspath(sys.argv[1]))
                print(os.path.abspath(sys.argv[2]))

                word = comtypes.client.CreateObject('Word.Application')
                doc = word.Documents.Open(inFile)
                doc.SaveAs(outFile, FileFormat=wdFormatPDF)

                doc.Close()
                word.Quit()
            except Exception as e:
                print(e)
        else:
            print("Format not supported.")


if __name__ == '__main__':
    now = time.time()
    Convert('sample.doc')
    end = time.time()

    print(end - now)
