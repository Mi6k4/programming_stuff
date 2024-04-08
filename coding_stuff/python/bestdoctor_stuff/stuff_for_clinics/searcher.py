import pandas as pd
import yadisk
from sqlalchemy import create_engine
import time

y = yadisk.YaDisk('406760da5b1345a88999f0acb1ef95bf', '7ecb5577edb643d3a07ce020292db4b6',
                  "AQAEA7qkBXctAAfNlAAXrxr49UN8l0uBHNszaAo")
#file_path="disk:/Clinics"
#file_path='disk:/Clinics/Clinics_ДФО'
file_path='disk:/Clinics/Clinics_ДФО/Clinics_Москва и МО/Москва'
#file_path='disk:/Clinics/Clinics_ДФО/Clinics_Москва и МО/Москва/"ДОКТОР НА ДОМ" ООО'
columns=['name', 'path']
#conn_string="postgresql://zeppelin:R63v5NspNsSEem@c-c9qbht031ah0gtrlftmj.rw.mdb.yandexcloud.net:5432/warehouse"
#db= create_engine(conn_string)
#conn= db.connect()
root_dirs=["Clinics_ДФО",
"Clinics_Москва и МО",
"Clinics_ПФО",
"Clinics_СЗФО",
"Clinics_СФО",
"Clinics_УФО",
"Clinics_ЦФО",
"Clinics_ЮФО и СКФО",
"morphed",
"Оцифровка",
"Профосмотры"]
dirs_to_skip=["morphed",
"Оцифровка",
"Профосмотры"]

def search_file(path,list_of_files=[]):
    for file in y.listdir(path):
        if file.type == "dir" and file.name in dirs_to_skip:
            pass
        elif file.type == "dir" and file.name not in dirs_to_skip:
            print(file.name)
            search_file(file.path)

        else:
            if file.type == "file" and '.xls' in file.name and '/Реестр' in file.path :
                list_of_files.append(file)
                print(file.name)

    #print(list_of_files)
    return list_of_files


def adding_to_df(list_of_file_pathes,columns):
    df=pd.DataFrame(columns=columns)
    for file in list_of_file_pathes:
        row=[file.name,file.path]
        df.loc[len(df)]=row
    return df

def df_to_gp(df,conn):

    df.to_sql('clinics_files', conn, schema='yandex_disk', if_exists='replace', index=False)



def main():
    start=time.time()
    list_of_file_pathes=search_file(file_path)
    middle=time.time()
    print(middle-start)
    dataframe=adding_to_df(list_of_file_pathes,columns)
    print(dataframe)
    #df_to_gp(dataframe,conn)
    end=time.time()
    print(end-start)


main()

