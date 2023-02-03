import psycopg2 
from psycopg2 import Error


from tkinter import filedialog
import tkinter.messagebox as mb
import sys
import re
import os
#********************************************************************
LIST_columns=('Unique-ID','Title','Journal','Year','Author','Volume','Number','Pages','DOI','Times-Cited',
              'Publisher','Type')
LIST_id=('ResearcherID-Numbers','ORCID-Numbers')
#----------------------------------------------------------------
def data_preparation(one_autor:str):
    def find_one(one_,name:str):
        #request=f"^{name} = "+chr(92)+"{([\s\S]+?*)\}"
        start=one_.rfind('\n'+name)
        f_str=re.findall("^"+name+"\{([\s\S]+?*)\}",one_)
        return f_str[0].replace('\n','').replace('  ',' ') if f_str else ''
    data=[]    
    for name in LIST_columns:
        data.append(find_one(one_autor,name))
    return data, find_one(one_autor,LIST_id[0]), find_one(one_autor,LIST_id[0])

#********************************************************************
def update_wos(db_conect,one_autor):
    data1=data_preparation(one_autor)
    data, r_id, r_orcid = data1
    cursor = db_conect.cursor()
    cursor.execute("""INSERT INTO public.scopus AS t(eid,title,journal,year,volume,number,pages,doi,note,publisher,document_type,author) 
            SELECT * FROM (values (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)) v(eid,title,journal,year,volume,number,pages,doi,note,publisher,document_type,author) 
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

        folder_selected = filedialog.askdirectory(title='Откройте папку с файлами \"savedrecs ...bib\"')
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
            if (Path_file_bib[-3:] != 'bib') and (Path_file_bib[:9] != 'savedrecs'):continue
            Bib_file = open(folder_selected+'/'+Path_file_bib,'r',encoding='UTF-8')
            buff=Bib_file.read()
            Bib_file.close()  
            for One_Autor in buff.split("@"):
                if len(One_Autor)<40: continue  
                update_wos(db_conect,One_Autor)            
        
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