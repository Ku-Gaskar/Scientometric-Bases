## Работа с Postgres
__Read_xlsx_update_postgres (`Таблица_ученных_`) -> таблицы `{Table_Sсience_HNURE, autors_in_departments,departments,lat_name_hnure}`__
*заполнение таблиц базы данных из зеленной таблицы*

__Сreat_table_Scopus_update(`.bib`)->таблица `{scopus}`__ 
*заполнение или обновление таблицы статей scopus базы данных из всех файлов bib в каталоге*

__table_Scopus_autors(`.csv`)->таблица `{scopus_autors}`__
*заполнение таблицы (id_author->id_aticle)   привязки статей к авторам scopus базы данных из всех файлов .csv (наличие полей в csv -> `Author(s) ID` и `EID`) в каталоге*