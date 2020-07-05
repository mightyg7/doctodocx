import os
from win32com import client as wc
for filename in os.listdir('C:\\Users\Mehrag\\Desktop\\samenvattingen'):
    w = wc.Dispatch('Word.Application')
    doc=w.Documents.Open('C:\\Users\\Mehrag\\Desktop\\samenvattingen\\' + filename)
    doc.SaveAs('C:\\Users\\Mehrag\\Desktop\\samenvattingen\\' + filename + "x",16)
    os.remove('C:\\Users\\Mehrag\\Desktop\\samenvattingen\\' + filename)
