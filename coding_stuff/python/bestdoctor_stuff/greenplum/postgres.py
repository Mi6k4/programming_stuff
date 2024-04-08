import psycopg2
from typing import List,AnyStr
bb_conn = 'postgresql://analytics:b2FyKPPR6fWsVr@172.21.16.38:6432/bigbrother'
gp_con = 'postgresql://zeppelin:R63v5NspNsSEem@c-c9qbht031ah0gtrlftmj.rw.mdb.yandexcloud.net:5432/warehouse'
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


DWH = DwsConn(conn_string=bb_conn)
DWH2 = DwsConn(conn_string=gp_con)
#data = DWH.select("select * from instancechangerecord where changed_at::date='2023-05-08'")

data2=DWH2.select("select * from bestdoctor_bb_external_tables.instancechangerecord_columnar_cache limit 10")
row_to_insert=[(000000, '{}', 'update', '72c161b5-c140-4e32-849a-927a34ca3ac0', '2023-05-09 16:01:45.902866', 'modified_at', '2021-09-23 10:18:09.691259+00', '2022-10-25 16:01:45.020538+00', None, 154152, 501, None)]

DWH2.execute(f"insert into bestdoctor_bb_external_tables.instancechangerecord_columnar_cache values {row_to_insert[0]} ;")