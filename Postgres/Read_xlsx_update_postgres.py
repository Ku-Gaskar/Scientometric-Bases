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
            id_scopus character varying COLLATE pg_catalog."default" NOT NULL
        )

    TABLESPACE "T ";"""
            
    ]
    
    cursor = conect.cursor()
    for query in List_quert:
        cursor.execute(query)
    
    conect.commit()
    return True
#********************************************************************
def ready_data_autor(sheet,i):
    autor=['']*5
    name_autor=sheet['B'+str(i)].value 

    autor[0]=' '.join(c.capitalize() if ("`" in c) or ("'" in c) else c.title() for c in name_autor.split())
    if sheet['D'+str(i)].value:
        find=re.findall(r'authorid=([\d]*)',sheet['D'+str(i)].value.lower())
        if find:
            autor[1]=find[0]    #ID_Scopus authorId=57200986251
    if sheet['C'+str(i)].value:
        find=re.findall(r'orcid.org/(.{4}-.{4}-.{4}-.{4})',sheet['C'+str(i)].value)
        if find:
            autor[2]=find[0]  #ID_Orcid 0000-0003-0803-7222
    autor[3]=sheet['L'+str(i)].value # Кафедра
    autor[4]=sheet['R'+str(i)].value # Латинские варианты фамилии

    return autor
#********************************************************************
def update_db(conect):
  
    create_tabels(conect)
    sheet=open_sheet(path_xlsx)
    cursor = conect.cursor()
    conect.autocommit = True
    for i in range(3,sheet.max_row+1):
        if not i%20: print('-',end='')
        if not sheet['B'+str(i)].value: continue
        data_autor=ready_data_autor(sheet,i)        
        # таблица ученных        
        cursor.execute("""INSERT INTO public."Table_Sсience_HNURE" AS t("FIO","ID_Scopus_Author","ORCID_ID") 
            SELECT * FROM (values (%s,%s,%s)) v("FIO","ID_Scopus_Author","ORCID_ID") 
            WHERE NOT EXISTS  (SELECT FROM public."Table_Sсience_HNURE" AS d where d."FIO" = v."FIO") 
            on conflict do nothing returning "id_Sciencer";""",tuple(data_autor[:3]))
        res=cursor.fetchone()
        if res: Id_Autor=res[0]
        else: 
            cursor.execute("""SELECT t."id_Sciencer" FROM public."Table_Sсience_HNURE"  AS t
                              WHERE  t."FIO" = %s ;""",(data_autor[0],))
            Id_Autor=cursor.fetchone()[0]
        
        if not data_autor[3]: data_autor[3]=NOT_DEP_NAME
        for dp in data_autor[3].upper().split(','):           
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
                on conflict do nothing returning id_autors;""",(Id_Autor,data_autor[0],id_dp,dp))

        # заполняем таблицу LatName
        if data_autor[4]: 
            for f_name in data_autor[4].split(';'):
                f_name=f_name.strip()
                if f_name:
                    cursor.execute("""INSERT INTO public.lat_name_hnure AS t(id_autor,name_lat) 
                        SELECT * FROM (values (%s,%s)) v(id_autor,name_lat) 
                        WHERE NOT EXISTS  (SELECT FROM public.lat_name_hnure AS d where d.id_autor = v.id_autor AND d.name_lat = v.name_lat) 
                        on conflict do nothing returning id_autor;""",(Id_Autor,f_name))
        #заполняем таблицу author_in_scopus
        if data_autor[1]:
            cursor.execute("""INSERT INTO public.author_in_scopus AS t(id_author,id_scopus) 
                SELECT * FROM (values (%s,%s)) v(id_author,id_scopus) 
                WHERE NOT EXISTS  (SELECT FROM public.author_in_scopus AS d where d.id_author = v.id_author AND d.id_scopus = v.id_scopus) 
                on conflict do nothing returning id_author;""",(Id_Autor,data_autor[1]))


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