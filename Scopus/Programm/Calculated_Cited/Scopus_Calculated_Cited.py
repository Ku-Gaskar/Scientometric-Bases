
from tkinter import filedialog
import tkinter.messagebox as mb
import sys
import re
import os


#***************************************************************************************************


folder_selected = filedialog.askdirectory(title='Откройте папку с файлами \"scopus...bib\"')
if not folder_selected: 
  mb.showwarning("Ошибка","папка *.bib не выбрана. \n Программа завершит работу")
  sys.exit()
#os.path.dirname(Path_file_bib)
arr = os.listdir(folder_selected)
res=0
for Path_file_bib in arr:
    if Path_file_bib[-3:] != 'bib':continue
    Bib_file = open(folder_selected+'/'+Path_file_bib,'r',encoding='UTF-8')
    buff=Bib_file.read()
    Bib_file.close()

    find_Cited=re.findall(r'{cited By ([\d]*)}', buff)
    res+=sum(int(i) for i in find_Cited )

print('Суммарное количество цитированний: %d' %res)

sys.exit()