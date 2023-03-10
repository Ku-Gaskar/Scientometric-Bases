import psycopg2 
from psycopg2 import Error

import re
import openpyxl

NOT_DEP_NAME=''

path_xlsx = 'd:\\Work_AVERS\\Python\\Scientometric-Bases\\Data\\Таблица ученых 2023-02-16.xlsx'

#********************************************************************
def open_sheet(path_:str) -> openpyxl.Workbook.active:
  GreenTable = openpyxl.load_workbook(path_)
  return GreenTable.active
#********************************************************************
def create_tabels(conect):
    """создание таблиц latName_Table, Sсience_HNURE, autors_in_departments, department """
    
    List_quert=[
    """CREATE TABLE IF NOT EXISTS public.lat_name_hnure
        (
        id_autor bigint NOT NULL,
        name_lat text COLLATE pg_catalog."default" NOT NULL
        )
    TABLESPACE "T ";""",
    
    """CREATE TABLE IF NOT EXISTS public."Table_Sсience_HNURE"
        (
            "id_Sciencer" bigint NOT NULL GENERATED ALWAYS AS IDENTITY ( INCREMENT 1 START 1 MINVALUE 1 MAXVALUE 9223372036854775807 CACHE 1 ),
            "FIO" text COLLATE pg_catalog."default" NOT NULL,
            "ID_Scopus_Author" character varying COLLATE pg_catalog."default",
            "Researcher_ID" character varying COLLATE pg_catalog."default",
            "ORCID_ID" character varying COLLATE pg_catalog."default",
            "Data_Time" timestamp without time zone NOT NULL DEFAULT now(),
            works bool NOT NULL DEFAULT true,
            CONSTRAINT "Table_Sсience_HNURE_pkey" PRIMARY KEY ("id_Sciencer")
        )
    TABLESPACE "T ";""",
    
    """CREATE TABLE IF NOT EXISTS public.autors_in_departments
        (
            id_autors bigint NOT NULL,
            name_autor character varying COLLATE pg_catalog."default" NOT NULL,
            name_department character varying COLLATE pg_catalog."default" NOT NULL,
            id_depatment bigint NOT NULL
        )
    TABLESPACE "T ";""",

    """CREATE TABLE IF NOT EXISTS public.departments
        (
            id_depat bigint NOT NULL GENERATED ALWAYS AS IDENTITY ( INCREMENT 1 START 10000 MINVALUE 1 MAXVALUE 9223372036854775807 CACHE 1 ),
            name_depat character varying COLLATE pg_catalog."default" NOT NULL,
            CONSTRAINT depart_pkey PRIMARY KEY (id_depat)
        )
    TABLESPACE "T ";""",
    
    f"""INSERT INTO public.departments AS t(name_depat) 
                SELECT * FROM (values ('{NOT_DEP_NAME}')) v(name_depat) 
                WHERE NOT EXISTS  (SELECT FROM public.departments AS d where d.name_depat = v.name_depat) 
                on conflict do nothing returning id_depat;""",
    
    """CREATE TABLE IF NOT EXISTS public.author_in_scopus
        (
            id_author bigint NOT NULL,
            id_scopus character varying COLLATE pg_catalog."default" NOT NULL,
            doc character varying COLLATE pg_catalog."default",
            note character varying COLLATE pg_catalog."default",
            h_ind character varying COLLATE pg_catalog."default"
        )

    TABLESPACE "T ";""",

    """DROP TABLE IF EXISTS public.stamp_tables;""",

    """CREATE TABLE IF NOT EXISTS public.stamp_tables
    (
        dimensions integer[] NOT NULL,
        title character varying[] COLLATE pg_catalog."default" NOT NULL,
        class_col character varying[] COLLATE pg_catalog."default",
        href character varying[] COLLATE pg_catalog."default",
        id_table name COLLATE pg_catalog."C" NOT NULL,
        href_postfix character varying[] COLLATE pg_catalog."default",
        CONSTRAINT stamp_tables_pkey PRIMARY KEY (id_table)
    )

    TABLESPACE "T ";
    """,
    """INSERT INTO public.stamp_tables (dimensions,title,class_col,href,id_table,href_postfix) VALUES
	 ('{7,35,10,18,10,10,10}','{"ID Автора",ФИО,Кафедра,ID_Scopus,Кол.док,Цити-рования,H-индекс}','{text-center,NULL,text-center,text-center,"text-end px-3","text-end px-3","text-end px-3"}','{NULL,NULL,NULL,https://www.scopus.com/authid/detail.uri?authorId=,NULL,NULL,NULL}','author','{NULL,NULL,NULL,NULL,NULL,NULL,NULL}'),
	 ('{10,35,22,8,10,15}','{EID,"Название публикации",Авторы,Год,"Тип докум.",Журнал}','{text-center,NULL,NULL,text-center,text-center,NULL}','{https://www.scopus.com/record/display.uri?eid=,NULL,NULL,NULL,NULL,NULL}','article','{&origin=resultslist&sort=plf-f,NULL,NULL,NULL,NULL,NULL}');
    """
    ]

    conect.autocommit = True
    cursor = conect.cursor()
    for query in List_quert:
        cursor.execute(query)
    
    # conect.commit()
    return True
#********************************************************************
def ready_data_autor(sheet,i):
    
    author={}    # autor=['']*5
    name_autor=sheet['B'+str(i)].value 

    author['name_author']=' '.join(c.capitalize() if ("`" in c) or ("'" in c) else c.title() for c in name_autor.split())
    author['id_scopus']=''
    if sheet['D'+str(i)].value:
        find=re.findall(r'authorid=([\d]*)',sheet['D'+str(i)].value.lower())
        if find:
            author['id_scopus']=find[0]    #ID_Scopus authorId=57200986251
    author['id_orcid']=''
    if sheet['C'+str(i)].value:
        find=re.findall(r'orcid.org/(.{4}-.{4}-.{4}-.{4})',sheet['C'+str(i)].value)
        if find:
            author['id_orcid']=find[0]  #ID_Orcid 0000-0003-0803-7222
    author['dep']=sheet['L'+str(i)].value # Кафедра
    author['lat_name']=sheet['R'+str(i)].value # Латинские варианты фамилии
    author['sc_doc'] = int(sheet['E'+str(i)].value)  if sheet['E'+str(i)].value else 0 # doc количество документов scopus
    author['sc_note'] =int(sheet['F'+str(i)].value)  if sheet['F'+str(i)].value  else 0  # note цитирования документов  scopus
    author['sc_h_ind'] =int(sheet['G'+str(i)].value) if sheet['G'+str(i)].value  else  0 # h-index scopus 

    return author
#********************************************************************
def update_db(conect):
  
    create_tabels(conect)
    sheet=open_sheet(path_xlsx)
    cursor = conect.cursor()
    conect.autocommit = True
    for i in range(3,sheet.max_row+1):
        if not i%20: print('-',end='')
        if not sheet['B'+str(i)].value: continue
        data_author=ready_data_autor(sheet,i)        
        # таблица ученных        
        cursor.execute("""INSERT INTO public."Table_Sсience_HNURE" AS t("FIO","ID_Scopus_Author","ORCID_ID") 
            SELECT * FROM (values (%s,%s,%s)) v("FIO","ID_Scopus_Author","ORCID_ID") 
            WHERE NOT EXISTS  (SELECT FROM public."Table_Sсience_HNURE" AS d where d."FIO" = v."FIO") 
            on conflict do nothing returning "id_Sciencer";""",(data_author['name_author'],data_author['id_scopus'],data_author['id_orcid']))
        res=cursor.fetchone()
        if res: Id_Autor=res[0]
        else: 
            cursor.execute("""SELECT t."id_Sciencer" FROM public."Table_Sсience_HNURE"  AS t
                              WHERE  t."FIO" = %s ;""",(data_author['name_author'],))
            Id_Autor=cursor.fetchone()[0]
        
        if not data_author['dep']: data_author['dep']=NOT_DEP_NAME
        for dp in data_author['dep'].upper().split(','):           
            # заполняем таблицу departments
            dp=dp.strip()
            cursor.execute("""INSERT INTO public.departments AS t(name_depat) 
                SELECT * FROM (values (%s)) v(name_depat) 
                WHERE NOT EXISTS  (SELECT FROM public.departments AS d where d.name_depat = v.name_depat) 
                on conflict do nothing returning id_depat;""",(dp,))
            res=cursor.fetchone()            
            if not res: 
                cursor.execute("""SELECT t.id_depat FROM public.departments  AS t
                                WHERE  t.name_depat = %s ;""",(dp,))                 
                res=cursor.fetchone()
            id_dp=res[0]
            
            # заполняем таблицу Autors_in_Departments
            cursor.execute("""INSERT INTO public.autors_in_departments AS t(id_autors,name_autor,id_depatment,name_department) 
                SELECT * FROM (values (%s,%s,%s,%s)) v(id_autors,name_autor,id_depatment,name_department) 
                WHERE NOT EXISTS  (SELECT FROM public.autors_in_departments AS d where d.id_autors = v.id_autors AND d.id_depatment = v.id_depatment) 
                on conflict do nothing returning id_autors;""",(Id_Autor,data_author['name_author'],id_dp,dp))

        # заполняем таблицу LatName
        if data_author['lat_name']: 
            for f_name in data_author['lat_name'].split(';'):
                f_name=f_name.strip()
                if f_name:
                    cursor.execute("""INSERT INTO public.lat_name_hnure AS t(id_autor,name_lat) 
                        SELECT * FROM (values (%s,%s)) v(id_autor,name_lat) 
                        WHERE NOT EXISTS  (SELECT FROM public.lat_name_hnure AS d where d.id_autor = v.id_autor AND d.name_lat = v.name_lat) 
                        on conflict do nothing returning id_autor;""",(Id_Autor,f_name))

        #заполняем таблицу author_in_scopus
        if data_author['id_scopus']:
            cursor.execute(f"""INSERT INTO public.author_in_scopus AS t(id_author,id_scopus,doc,note,h_ind) 
                SELECT * FROM (values ({Id_Autor},'{data_author['id_scopus']}',{data_author['sc_doc']},{data_author['sc_note']},{data_author['sc_h_ind']})) v(id_author,id_scopus,doc,note,h_ind) 
                WHERE NOT EXISTS  (SELECT FROM public.author_in_scopus AS d where d.id_author = v.id_author AND d.id_scopus = v.id_scopus) 
                on conflict do nothing returning id_author;""")     


    res=cursor.fetchone()
    if not res: return 0
    return res[0]
#********************************************************************
def main():
    try:
        db_conect = psycopg2.connect(user="postgres",
                                 password="postgress",
                                     host="127.0.0.1",
                                     port="5432")
        print (update_db(db_conect))
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