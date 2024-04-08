import json
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


data=DWH.select("select * from a_source_schema.source_metadata_table where source_name = 'bestdoctor_replica' order by source_table_name asc")
ddls=[]
table=[]
for i in data:
    ddls.append(i[5])
    table.append((i[2]))

result=""
list_of_table_columns=[]
for ddl in ddls:
    columns = ""
    for i in range(ddl.find('(')+1,ddl.find('CONSTRAINT')-1):
        columns+=ddl[i]
    #print(columns)
    result = ""
    for i in columns.split(','):
       column= i.replace('NULL',' ')
       column = column.replace('NOT',' ')
       result+=column
    #print(result)
    list_of_table_columns.append(result)

#print(list_of_table_columns)



for i in range(len(list_of_table_columns)):
    print(table[i])
    print(list_of_table_columns[i])
    DWH.execute(f"insert into khitrin_m.source_table_columns values ('{table[i]}','{list_of_table_columns[i]}') ")