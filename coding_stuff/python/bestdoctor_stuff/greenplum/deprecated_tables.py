
import psycopg2
from typing import List,AnyStr

gp_conn = 'postgresql://zeppelin:R63v5NspNsSEem@172.21.216.13:5432/warehouse' # подлючение под пользаком zeppelin
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

deprec_tables = DWH.select('select * from khitrin_m.deprecated_tables order by "table"  asc')

wanted_row=[]

for value in deprec_tables:
    print((value[0],value[1]))
    views=DWH.select(f"select * from pg_views where definition like '%{value[0]}.{value[1]}%'")
    for view in views:
        wanted_row.append(value[0])
        wanted_row.append(value[1])
        wanted_row.append(view[0])
        wanted_row.append(view[1])
        wanted_row.append(view[3])
        print(wanted_row)
        DWH.execute(f"""insert into khitrin_m.views_with_deprecated_columns values ('{wanted_row[0]}','{wanted_row[1]}','{wanted_row[2]}','{wanted_row[3]}','{view[3].replace("'",'"')}')""")
        wanted_row.clear()