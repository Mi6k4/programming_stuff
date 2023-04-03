import psycopg2
from typing import List,AnyStr
import yadisk
from openpyxl import Workbook, load_workbook
import pandas as pd


y = yadisk.YaDisk('406760da5b1345a88999f0acb1ef95bf', '7ecb5577edb643d3a07ce020292db4b6',
                  "AQAEA7qkBXctAAfNlAAXrxr49UN8l0uBHNszaAo")

gp_conn = 'postgresql://zeppelin:R63v5NspNsSEem@c-c9qbht031ah0gtrlftmj.rw.mdb.yandexcloud.net:5432/warehouse'
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



def cringe_function(file_name):
    wb=load_workbook(file_name)
    new_name=file_name+"_morphed_test_khitrin.xlsx"
    wb.save(new_name)
    return new_name

DWH = DwsConn(conn_string=gp_conn)
data = DWH.select("select path from yandex_disk.checked_files where root_dir_name = 'clinics' and path like '%morphed%' and name = 'Бестдоктор Реестр (2).xlsx';")
destination_directory=DWH.select("select path from yandex_disk.checked_dirs where path like '%ООО «СуперМедик»%';")
#print(type(data))
#print(data)

#for i in range(len(data)):
#    print(data[i])

#print(data[0])
for i in data:
    path=i[0]

#print(destination_directory[1])
for i in destination_directory:
    dd=i[0]+"/"

#for i in range(len(destination_directory)):
#  print(destination_directory[i])

print(path)
#print(type(path))
print(dd)
print(type(dd))
#print(y.check_token())
#downloaded_file=y.download(path,"reestr.xlsx")
#
#new_file=cringe_function("reestr.xlsx")
#y.upload(new_file,dd)