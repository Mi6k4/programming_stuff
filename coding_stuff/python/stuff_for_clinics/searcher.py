import pandas as pd
import yadisk

y = yadisk.YaDisk('406760da5b1345a88999f0acb1ef95bf', '7ecb5577edb643d3a07ce020292db4b6',
                  "AQAEA7qkBXctAAfNlAAXrxr49UN8l0uBHNszaAo")
path="disk:/Clinics"
file_path='disk:/Clinics/Clinics_ДФО/Clinics_Москва и МО/Москва/"ДОКТОР НА ДОМ" ООО/Реестры/2022/11'

#for i in y.listdir(path):
#   if i.type == "dir":
 #      print("dir")
#       print(i.name)
#   else: pass
dir = y.listdir(file_path)

for file in dir:
    print(file.name, file.modified, file.created, file.path)
