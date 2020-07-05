import os
from win32com import client as wc

#be sure your path ends with \\ 
path = 'C:\\your_folder\\'

#path to your folder where all the .doc files are located
for filename in os.listdir(path):
    w = wc.Dispatch('Word.Application')
    doc=w.Documents.Open(path + filename)
    doc.SaveAs(path + filename + "x",16)
#this deletes the .doc files
    os.remove(path + filename)
