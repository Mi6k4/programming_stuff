

from sqlalchemy import create_engine
import time
import psycopg2
from typing import List,AnyStr
import yadisk
from openpyxl import Workbook, load_workbook
import pandas as pd
import pyexcel

y = yadisk.YaDisk('406760da5b1345a88999f0acb1ef95bf', '7ecb5577edb643d3a07ce020292db4b6',
                  "AQAEA7qkBXctAAfNlAAXrxr49UN8l0uBHNszaAo")

gp_conn = gp_conn = 'postgresql://zeppelin:R63v5NspNsSEem@c-c9qbht031ah0gtrlftmj.rw.mdb.yandexcloud.net:5432/warehouse' # подлючение под пользаком zeppelin

class DwsConn:
    def __init__(self, conn_string):
        self.conn_string = conn_string

    def select(self, query: str) -> List[tuple]:
        with psycopg2.connect(self.conn_string) as conn:
            with conn.cursor() as cursor:
                cursor.execute(query)
                return cursor.fetchall()

    def execute(self, query: str) -> None:
        with psycopg2.connect(self.conn_string) as conn:
            with conn.cursor() as cursor:
                cursor.execute(query)
                conn.commit()


DWH = DwsConn(conn_string=gp_conn)
#data = DWH.select("select path from yandex_disk.checked_files where root_dir_name = 'clinics' and path like '%morphed%' and name = 'Бестдоктор Реестр (2).xlsx';")
#destination_directory=DWH.select("select path from yandex_disk.checked_dirs where path like '%ООО «СуперМедик»%';")

#list_of_dirs= DWH.select("select name,path from yandex_disk.clinics_files_test where status is null;")
#print(list_of_dirs)

#dir=y.listdir("disk:/Clinics/Clinics_СЗФО/Санкт-Петербург/МЦ 21 ВЕК, ООО/Реестры эл.вид/2023/02")
#dir=y.listdir("disk:/Clinics/Clinics_СЗФО/Санкт-Петербург/МЦ 21 ВЕК, ООО/Реестры эл. вид/2023/02/       февраль 21   век.xlsx")

#files_path={'path': [], 'file_name': [], 'folder_name': []}

#for tuple in list_of_dirs:
#    files_path['path'].append(tuple[1])
#    files_path['file_name'].append(tuple[0])
#    files_path['folder_name'].append(tuple[1].split('/',-1)[3])

#print(files_path)




#for i in files_path['path']:
    #print(y.listdir(i))
#start=time.time()
#file_path='disk:/Clinics/morphed/ООО СК "ТЕСТ"/2023/03/Март     2023 Бестдоктор для теста.xls'

#list_dir=y.listdir(files_path['path'][0])
#list_dir=y.listdir('disk:/Clinics/morphed/АО "Ильинская больница"/2023/03/Реестр от    31.03.2023_ООО Бестдоктор_амб_для_теста.xlsx')
#for i in list_dir:
#    print(i)
#"disk/:Clinics/morphed/ООО «МЦ» XXI ВЕК»/2023/03/   Бестдоктор  (  88).xlsx"

#'disk:/Clinics/morphed/ООО  СК "ТЕСТ" 2023 03 Март 2023 Бестдоктор.xlsx'

#'disk:/Clinics/morphed/АО "Ильинская больница"/2023/03/Реестр от 31.03.2023_ООО    Бестдоктор_амб_для_теста.xlsx'




dir=y.listdir('disk:/Clinics/morphed/ООО «Клиника ОсНова»/2023/03/')
for i in dir:
    print(i)

disk:/Clinics/Clinics_ДФО/Благовещенск/А-СТОМ, ООО/Реестр/2023/01/Реестр/Бестдоктор январь 2023 проверка.xlsx