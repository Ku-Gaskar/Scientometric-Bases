import psycopg2 
from psycopg2 import Error

#import app_logger


from tkinter import filedialog
import tkinter.messagebox as mb
import sys
import re
import os
#********************************************************************
#logger
import logging

_log_format = f"%(asctime)s - [%(levelname)s] - %(name)s - (%(filename)s).%(funcName)s(%(lineno)d) - %(message)s"

def get_file_handler():
    file_handler = logging.FileHandler("XXX_XXX.log")
    file_handler.setLevel(logging.WARNING)
    file_handler.setFormatter(logging.Formatter(_log_format))
    return file_handler

def get_stream_handler():
    stream_handler = logging.StreamHandler()
    stream_handler.setLevel(logging.INFO)
    stream_handler.setFormatter(logging.Formatter(_log_format))
    return stream_handler

def get_logger(name):
    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)
    logger.addHandler(get_file_handler())
    logger.addHandler(get_stream_handler())
    return logger


#********************************************************************
LIST_columns=('Unique-ID','Title','Journal','Year','Author','Volume','Number','Pages','DOI','Times-Cited',
              'Publisher','Type')
LIST_id=('ResearcherID-Numbers','ORCID-Numbers')
#----------------------------------------------------------------

def data_preparation(one_autor:str):
    #----------------------------------------------------------------
    def find_one(one_:str,name:str):
        #request=f"^{name} = "+chr(92)+"{([\s\S]+?*)\}"
        patern='\n'+name+' = {'
        start=one_.find(patern)
        if start==-1: return ''
        start+=len(patern)
        end=one_.find('}',start)
        return one_[start:end].replace('   ',' ').replace('  ',' ').replace('\n','')
    #----------------------------------------------------------------
        #f_str=re.findall("^"+name+r"\{([\s\S]+?*)\}",one_)
        #return f_str[0].replace('\n','').replace('  ',' ') if f_str else '
    data=[]    
    for name in LIST_columns:
        data.append(find_one(one_autor,name))

    researcher_dict=dict(re.findall(r'([\w]+, [\w]+).{0,}?/([A-Za-z]{1,}-[0-9]{3,}-[0-9]{4})',find_one(one_autor,LIST_id[0])))
    orcid_dict=dict(re.findall(r'([\w]+, [\w]+).{0,}?/([0-9]{4}-[0-9]{4}-[0-9]{4}-[0-9X]{4})',find_one(one_autor,LIST_id[1])))
    
    return data,researcher_dict,orcid_dict

#********************************************************************
def update_wos(db_conect,one_autor):
    def search_in_table_hnure_for_author(author,cursor,orcid):
        if orcid:
            cursor.execute("""SELECT "id_Sciencer" FROM public."Table_Sсience_HNURE" AS tsh WHERE tsh."ORCID_ID" = %s;""",(orcid,))
            res=cursor.fetchall()
            if res: return res
        
        cursor.execute("""SELECT * FROM public.lat_name_hnure AS lnh WHERE lnh.name_lat = %s;""",(author,))
        return cursor.fetchall()   
    #----------------------------------------------------------------
    data1=data_preparation(one_autor)
    data, r_id, r_orcid = data1
    cursor = db_conect.cursor()
    cursor.execute("""INSERT INTO public.wos AS t(unique_id,title,journal,year,author,volume,number,pages,doi,note,publisher,document_type) 
            SELECT * FROM (values (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)) v(unique_id,title,journal,year,author,volume,number,pages,doi,note,publisher,document_type) 
            WHERE NOT EXISTS  (SELECT FROM public.wos AS d where d.unique_id = v.unique_id) 
            on conflict do nothing returning "unique_id";""",tuple(data))
    res=cursor.fetchone()
    if res: 
        Id_Aticl=res[0]        
        for author in data[4].split(" and "):
            author_orcid,author_r_id='',''
            if author in r_orcid: author_orcid=r_orcid[author]   
            if author in r_id: author_r_id=r_id[author]
            author_id_hnure=search_in_table_hnure_for_author(author,cursor,author_orcid) 
            if  len(author_id_hnure) > 1:
                logger.warning(f"ERROR: more than 1 -> {len(author_id_hnure)} :\n{data[0]} --> {author_id_hnure}")
                #print(f"ERROR: more than 1 -> {len(author_id_hnure)} :\n{author_id_hnure}")
                
                author_id_hnure=int(author_id_hnure[0][0])
            elif len(author_id_hnure)==1:
                author_id_hnure=int(author_id_hnure[0][0])
            else: author_id_hnure=None

            cursor.execute("""INSERT INTO public.wos_autors AS t(unique_id,orcid,researcher_id,author,id_autor) 
            SELECT unique_id,orcid,researcher_id,author,id_autor::int FROM (values (%s,%s,%s,%s,%s)) v(unique_id,orcid,researcher_id,author,id_autor) 
            WHERE NOT EXISTS  (SELECT FROM public.wos_autors AS d where (d.unique_id = v.unique_id) and (d.author = v.author)) 
            on conflict do nothing returning "id_autor";""",(data[0],author_orcid ,author_r_id,author,author_id_hnure))
    else: 
        cursor.execute("""UPDATE public.wos AS s SET note = %s, data_update = now()
                            WHERE  s.unique_id = %s 
                            RETURNING  s.unique_id;""",(data[9],data[0]))
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
        logger.warning("Ошибка при работе : "+ error)
        print("Ошибка при работе :", error)
    finally:
        if db_conect:
            #cursor.close()
            db_conect.close()
            print("Соединение с PostgreSQL закрыто")
            return
#********************************************************************

if __name__=="__main__":
    logger = get_logger(__name__)
    logger.info("Программа стартует")
    main()    