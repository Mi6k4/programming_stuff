
import psycopg2
from typing import List,AnyStr
import yadisk
from openpyxl import Workbook, load_workbook
import pandas as pd

y = yadisk.YaDisk('406760da5b1345a88999f0acb1ef95bf', '7ecb5577edb643d3a07ce020292db4b6',
                  "AQAEA7qkBXctAAfNlAAXrxr49UN8l0uBHNszaAo")
old_con='postgresql://zeppelin:R63v5NspNsSEem@c-c9qbht031ah0gtrlftmj.rw.mdb.yandexcloud.net:5432/warehouse'
gp_conn = 'postgresql://zeppelin:R63v5NspNsSEem@c-c9qbht031ah0gtrlftmj.rw.mdb.yandexcloud.net:5432/warehouse' # подлючение под пользаком zeppelin
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
data = DWH.select('select type, max(id) from chat2desk.dataset_chat2desk_messages group by type')
print(data)

