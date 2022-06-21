# https://darkcoding.net/software/printing-word-and-pdf-files-from-python/ 

# https://www.stackvidhya.com/run-python-file-in-terminal/

# python winprint.py RS22ZODVAL.docx

from win32com import client
import time
import sys

def printWordDocument(filename):
    word.Documents.Open(filename)
    word.ActiveDocument.PrintOut()
    time.sleep(2)
    word.ActiveDocument.Close()

word = client.Dispatch("Word.Application")

# filename = sys.argv[1]
filename = 'C:/Users/mathlab/Documents/Matlab/20190318_Interface/Input/Templates/FinRep.docx'

printWordDocument(filename)

# print(sys.argv[1])

word.Quit()
print('Finished')


