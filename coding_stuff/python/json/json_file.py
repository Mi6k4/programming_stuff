import json
import psycopg2
from typing import List,AnyStr
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

DWH = DwsConn(conn_string=gp_con)



with open('models_graph.json','r') as read_json:
    data=json.load(read_json)

lst=[]

for key in data.keys():
    #print(key)
    if bool(data[key]['deprecated_fields']):
        print(key)
        #lst.append(key)

        #print(data[key]['deprecated_fields'])
        for field in data[key]['deprecated_fields'].keys():
            lst.append(key)
            print(field)
            lst.append(field)
            print(data[key]['deprecated_fields'][field])
            for key_2 in data[key]['deprecated_fields'][field].keys():
                if key_2=='context' and bool(data[key]['deprecated_fields'][field]['context']):
                    print(data[key]['deprecated_fields'][field][key_2]['hint'])
                    lst.append(str(data[key]['deprecated_fields'][field][key_2]['hint']))
                else:
                    print(data[key]['deprecated_fields'][field][key_2])
                    lst.append(str(data[key]['deprecated_fields'][field][key_2]))
            print(lst)
            tuple_to_insert=tuple(lst)
            print(tuple_to_insert)
            DWH.execute(f'insert into khitrin_m.deprecated_tables values {tuple_to_insert};')
            lst.clear()
