import psycopg2
from typing import List

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

def main():
    DWH = DwsConn(conn_string=gp_conn)
    data = DWH.select("select 'revoke ALL PRIVILEGES on bestdoctor_adminka.' || tablename || ' from analytics;' from pg_tables  where schemaname = 'bestdoctor_adminka' and tablename like '%cache%';")
    for query_row in data:
        DWH.execute(query_row[0])
