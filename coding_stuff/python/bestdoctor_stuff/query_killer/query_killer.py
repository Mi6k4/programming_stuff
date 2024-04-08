import psycopg2
from typing import List

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


def main():
    DWH = DwsConn(conn_string=gp_conn)
    data = DWH.select("""
                    select 	'select mdb_toolkit.gp_terminate_backend('||pid||');' , waiting_reason,duration,usename, query	from	
		            (select *, extract(epoch from (now() - query_start)::interval)/60 as duration
                    from mdb_toolkit.pg_stat_activity() where state = 'active' 
                    and usename not in ('m_khitrin','f_polyakov')) as t  where duration > 30
                        """)
    for i in data:
        if i[1] is None and i[2]>30:
            print(i)
            #DWH.execute(i[0])
            DWH.execute(f"insert into data_quality.killed_queries values('{i[3]}','{i[4]}')")
        elif i[1]=='lock' and i[2]>60:
            pass
            #DWH.execute(i[0])

main()