from docx2pdf import convert
import time
import sys
import os


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
        else:
            print("Format not supported.")


if __name__ == '__main__':
    now = time.time()
    Convert('sample.doc')
    end = time.time()

    print(end - now)
