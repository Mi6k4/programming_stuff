
import psycopg2
from typing import List,AnyStr
gp_con = 'postgresql://zeppelin:R63v5NspNsSEem@172.21.216.13:5432/warehouse'
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




DWH = DwsConn(conn_string=gp_con)
data=DWH.select("select * from a_source_schema.source_metadata_table  order by source_table_name asc")
ddls=[]
table=[]
source=[]
for i in data:
    source.append(i[0])
    ddls.append(i[5])
    table.append((i[2]))

list_of_table_columns=[]
result = []
for ddl in ddls:
    columns = ""
    list_of_columns = []
    for i in range(ddl.find('(')+1,ddl.find('CONSTRAINT')-1):
        columns+=ddl[i]
    for i in columns.split('L,'):
        column= i.replace('NUL',' ')
        column = column.replace('NOT',' ')
        list_of_columns.append(column)
    result.append(list_of_columns)
for i in range(len(table)):
    for j in range(len(result[i])-1):
        col=result[i][j].split(' ')
        print(col[4],col[5])
        DWH.execute(f"""insert into khitrin_m.source_table_columns_with_types_and_source values ('{source[i]}','{table[i]}','{col[4]}','{col[5]}') """)
