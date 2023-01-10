
from tkinter import filedialog
import tkinter.messagebox as mb
import sys

import os.path
import re




#***************************************************************************************************


Path_file_bib=filedialog.askopenfilename(title='Откройте \"scopus...bib\"',filetypes=[('bib file','*.bib')])
if not Path_file_bib: 
  mb.showwarning("Ошибка","Файл *.bib не выбран. \n Программа завершит работу")
  sys.exit()

Bib_file = open(Path_file_bib,'r',encoding='UTF-8')
buff=Bib_file.read()
Bib_file.close()

At=buff.split('@')

sys.exit()