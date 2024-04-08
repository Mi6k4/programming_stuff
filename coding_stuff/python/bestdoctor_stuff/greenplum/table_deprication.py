
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

deprec_tables = DWH.select("select * from data_quality.views_with_deprecated_columns where schema_name = 'bestdoctor_adminka_external_tables' order by deprecated_table desc" )

data_type = DWH.select("""select data_type from information_schema.columns where 
table_schema ='bestdoctor_adminka_external_tables' and table_name = 'insurance_insurancerequestsetting' and column_name = 'use_short_policy_template' """)

#print (data_type[0][0])


for row in deprec_tables:
    data_type = DWH.select(f"""select data_type from information_schema.columns where 
   table_schema ='bestdoctor_adminka_external_tables' and table_name = '{row[0]}' and column_name = '{row[1]}' """)
   # print(row[0]+ '.'+ row[1])
    column=row[0]+ '.'+ row[1]
    column_to_replace='null::' +data_type[0][0] + ' as ' + row[1]
   # print(row[4])
    final_description="create or replace view bestdoctor_adminka_external_tables."+row[3] + " as "+row[4].replace(column,column_to_replace)
    print(final_description)
    if data_type[0][0] == 'character varying':
        with open('out_with_character.txt', 'a') as f:
            f.write(final_description)
    elif data_type[0][0] == 'numeric':
        with open('out_with_numeric.txt', 'a') as f:
            f.write(final_description)
    else:
        with open('out_usual.txt', 'a') as f:
            f.write(final_description)