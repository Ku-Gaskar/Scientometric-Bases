import psycopg2 
from psycopg2 import Error

from tkinter import filedialog
import tkinter.messagebox as mb
import sys
import re
import os
#********************************************************************
def data_preparation(one_autor):
    data=['']*11
    data[0]=re.findall(r'eid=(2-s2.0-[0-9]{8,15})[&]?',one_autor)[0]    # eid
    data[1]=re.findall(r'title=\{(.*)\}',one_autor)[0]            # title
    data[2]=re.findall(r'journal=\{(.*)\}',one_autor)[0]          # journal
    data[3]=re.findall(r'year=\{(.*)\}',one_autor)[0]             # year
    if 'volume={' in one_autor:
        data[4]=re.findall(r'volume=\{(.*)\}',one_autor)[0]
    if 'number={' in one_autor:
        data[5]=re.findall(r'number=\{(.*)\}',one_autor)[0]
    if 'pages={' in one_autor: 
        data[6]=re.findall(r'pages=\{(.*)\}',one_autor)[0]
    if 'doi={' in one_autor: 
        data[7]=re.findall(r'doi=\{(.*)\}',one_autor)[0]
    if 'note={' in one_autor: 
        data[8]=re.findall(r'note=\{cited By (.*)\}',one_autor)[0]
    if 'publisher={' in one_autor: 
        data[9]=re.findall(r'publisher=\{(.*)\}',one_autor)[0]
    if 'document_type={' in one_autor: 
        data[10]=re.findall(r'document_type=\{(.*)\}',one_autor)[0]
    return data

#********************************************************************
def update_scopus(db_conect,one_autor):
    data=data_preparation(one_autor)
    cursor = db_conect.cursor()
    cursor.execute("""INSERT INTO public.scopus AS t(eid,title,journal,year,volume,number,pages,doi,note,publisher,document_type) 
            SELECT * FROM (values (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)) v(eid,title,journal,year,volume,number,pages,doi,note,publisher,document_type) 
            WHERE NOT EXISTS  (SELECT FROM public.scopus AS d where d.eid = v.eid) 
            on conflict do nothing returning "eid";""",tuple(data))
    res=cursor.fetchone()
    if res: Id_Aticl=res[0]
    else: 
        cursor.execute("""UPDATE public.scopus AS s SET note = %s, data_update = now()
                            WHERE  s.eid = %s 
                            RETURNING  s.eid;""",(data[8],data[0]))
        Id_Aticl=cursor.fetchone()[0]

    print (Id_Aticl)
    return

#********************************************************************
def main():
    try:

        folder_selected = filedialog.askdirectory(title='Откройте папку с файлами \"scopus...bib\"')
        if not folder_selected: 
            mb.showwarning("Ошибка","папка *.bib не выбрана. \n Программа завершит работу")
            sys.exit()
        #os.path.dirname(Path_file_bib)
        arr = os.listdir(folder_selected)

        db_conect = psycopg2.connect(user="postgres",
                            password="postgress",
                                host="127.0.0.1",
                                port="5432")
        db_conect.autocommit = True

        for Path_file_bib in arr:
            if Path_file_bib[-3:] != 'bib':continue
            Bib_file = open(folder_selected+'/'+Path_file_bib,'r',encoding='UTF-8')
            buff=Bib_file.read()
            Bib_file.close()  
            for One_Autor in buff.split("@"):
                if len(One_Autor)<40: continue  
                update_scopus(db_conect,One_Autor)            
        
    except (Exception, Error) as error:
        print("Ошибка при работе :", error)
    finally:
        if db_conect:
            #cursor.close()
            db_conect.close()
            print("Соединение с PostgreSQL закрыто")
            return
#********************************************************************

if __name__=="__main__":
    main()    