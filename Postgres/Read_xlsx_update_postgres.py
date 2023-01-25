import psycopg2 
from psycopg2 import Error

import re
import openpyxl

path_xlsx = 'd:\\Work_AVERS\\Python\\Scientometric-Bases\\Data\\Таблица ученых 2023-01-12.xlsx'

#********************************************************************
def open_sheet(path_:str) -> openpyxl.Workbook.active:
  GreenTable = openpyxl.load_workbook(path_)
  return GreenTable.active
#********************************************************************
def create_tabel_latName(conect):
    quert="""CREATE TABLE IF NOT EXISTS public.lat_name_hnure
        (
        id_autor bigint NOT NULL,
        name_lat text COLLATE pg_catalog."default" NOT NULL
        )
    TABLESPACE "T ";"""
    cursor = conect.cursor()
    cursor.execute(quert)
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
  
    create_tabel_latName(conect)
    sheet=open_sheet(path_xlsx)
    cursor = conect.cursor()
    conect.autocommit = True
    for i in range(3,sheet.max_row+1):
        if not sheet['B'+str(i)].value: continue
        data_autor=ready_data_autor(sheet,i)        
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
        
        if not data_autor[3]: data_autor[3]="X????X"
        for dp in data_autor[3].upper().split(','):           
            # заполняем таблицу departments
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