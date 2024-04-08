import psycopg2
from typing import List
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


def main():
    DWH = DwsConn(conn_string=gp_con)
    data=DWH.select("select source,table_name from khitrin_m.gp_and_source_tables_extended where flag is false and gp_table is not null and table_name is not null and src_col_sum > gp_col_sum  ")
    list_of_source_tables=[]
    for source_table in data:
        tmp_list=[]
        tmp_list.append(source_table[0])
        tmp_list.append(source_table[1])
        list_of_source_tables.append(tmp_list)

    for table in list_of_source_tables:
        data2=DWH.select(f"  select * from  khitrin_m.changelog_for_development_extended where table_name = '{table[1]}' and source = '{table[0]}'")
        for column in data2:
            if None in column:
                if column[4] == 'bestdoctor_replica':
                    #DWH.execute(f"alter table bestdoctor_adminka_external_tables.{column[5]} add column {column[6]} {column[7]}")
                    print(f"alter table bestdoctor_adminka_external_tables.{column[5]} add column if not exists {column[6]} {column[7]}")
                else:
                    #DWH.execute(f"alter table {column[4]}_external_tables.{column[5]} add column {column[6]} {column[7]}")
                    print(f"alter table {column[4]}_external_tables.{column[5]} add column if not exists {column[6]} {column[7]}")

