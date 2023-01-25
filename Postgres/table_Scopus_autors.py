import psycopg2 
from psycopg2 import Error
import csv

from tkinter import filedialog
import tkinter.messagebox as mb
import sys
import os
#********************************************************************
def update_scopus_autors(db_conect,One_Autor:list[str]):  
    eid=One_Autor[2]
    if not eid: return 0
    for id_autor in One_Autor[0].split(';'):
        if not id_autor: continue 
        cursor = db_conect.cursor()
        cursor.execute("""INSERT INTO public.scopus_autors AS t(id_sc_autor,eid) 
                SELECT * FROM (values (%s,%s)) v(id_sc_autor,eid) 
                WHERE NOT EXISTS  (SELECT FROM public.scopus_autors AS d where d.eid = v.eid and d.id_sc_autor = v.id_sc_autor) 
                on conflict do nothing returning "eid";""",(id_autor,eid))
    res= cursor.fetchone() 
    return res[0] if res else 0
#********************************************************************
def main():
    try:
        folder_selected = filedialog.askdirectory(title='Откройте папку с файлами \"scopus...csv\"')
        if not folder_selected: 
            mb.showwarning("Ошибка","папка *.csv не выбрана. \n Программа завершит работу")
            sys.exit()
        #os.path.dirname(Path_file_bib)
        arr = os.listdir(folder_selected)

        db_conect = psycopg2.connect(user="postgres",
                            password="postgress",
                                host="127.0.0.1",
                                port="5432")
        db_conect.autocommit = True

        for Path_file_bib in arr:
            if Path_file_bib[-3:] != 'csv':continue
            with open(folder_selected+'/'+Path_file_bib,'r',encoding='UTF-8') as csv_file:
                buff = csv.reader(csv_file)
                #csv_file.close()

                for i,One_Autor in enumerate(buff):
                    if i==0: continue  
                    print(update_scopus_autors(db_conect,One_Autor))

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