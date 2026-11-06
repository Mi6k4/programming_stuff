"""
Задание
1. Скачать исторические курсы валют по списку за определенные даты из сервиса с помощью обращения к API
2. Сложить результаты в CSV в виде: Date, BaseCurrency, Currency, Rate
3. Прочитать результаты из CSV файла и сохранить результат в файлы
    a. Формат XLSX
    b. Название файла %ISO_WEEK%.xlsx, где iso-week порядковый номер недели от Date
    c. Состав файла: все скаченные курсы валют за выбранную неделю

Примечание:
Сервис с курсами валют: https://fixer.io;
Ключ c36efe3e3bf039ce9777de61e5907540;
Базовая валюта: EUR
Список валют: CAD, THB, USD
Период: с 2024-09-01 по 2024-09-15

"""


first_Date='2024-09-01'
last_date='2024-09-15'



api=f'https://data.fixer.io/api/{first_Date}?access_key=c36efe3e3bf039ce9777de61e5907540'
api=f'https://data.fixer.io/api/{first_Date}?access_key=c36efe3e3bf039ce9777de61e5907540'
api=f'https://data.fixer.io/api/{first_Date}?access_key=c36efe3e3bf039ce9777de61e5907540'
api=f'https://data.fixer.io/api/{first_Date}?access_key=c36efe3e3bf039ce9777de61e5907540'
api=f'https://data.fixer.io/api/{first_Date}?access_key=c36efe3e3bf039ce9777de61e5907540'
api=f'https://data.fixer.io/api/{first_Date}?access_key=c36efe3e3bf039ce9777de61e5907540'
api=f'https://data.fixer.io/api/{first_Date}?access_key=c36efe3e3bf039ce9777de61e5907540'
api=f'https://data.fixer.io/api/{first_Date}?access_key=c36efe3e3bf039ce9777de61e5907540'
api=f'https://data.fixer.io/api/{first_Date}?access_key=c36efe3e3bf039ce9777de61e5907540'
api=f'https://data.fixer.io/api/{first_Date}?access_key=c36efe3e3bf039ce9777de61e5907540'
api=f'https://data.fixer.io/api/{first_Date}?access_key=c36efe3e3bf039ce9777de61e5907540'
api=f'https://data.fixer.io/api/{first_Date}?access_key=c36efe3e3bf039ce9777de61e5907540'
api=f'https://data.fixer.io/api/{first_Date}?access_key=c36efe3e3bf039ce9777de61e5907540'
api=f'https://data.fixer.io/api/{first_Date}?access_key=c36efe3e3bf039ce9777de61e5907540'
api=f'https://data.fixer.io/api/{first_Date}?access_key=c36efe3e3bf039ce9777de61e5907540'
api=f'https://data.fixer.io/api/{first_Date}?access_key=c36efe3e3bf039ce9777de61e5907540'
api=f'https://data.fixer.io/api/{first_Date}?access_key=c36efe3e3bf039ce9777de61e5907540'
api=f'https://data.fixer.io/api/{first_Date}?access_key=c36efe3e3bf039ce9777de61e5907540'
api=f'https://data.fixer.io/api/{first_Date}?access_key=c36efe3e3bf039ce9777de61e5907540'
api=f'https://data.fixer.io/api/{first_Date}?access_key=c36efe3e3bf039ce9777de61e5907540'
api=f'https://data.fixer.io/api/{first_Date}?access_key=c36efe3e3bf039ce9777de61e5907540'
api=f'https://data.fixer.io/api/{first_Date}?access_key=c36efe3e3bf039ce9777de61e5907540'
api=f'https://data.fixer.io/api/{first_Date}?access_key=c36efe3e3bf039ce9777de61e5907540'
api=f'https://data.fixer.io/api/{first_Date}?access_key=c36efe3e3bf039ce9777de61e5907540'





import requests
import pandas as pd
response = requests.get(api)
data = response.json()
print(data)
df=pd.DataFrame(columns=['Date', 'BaseCurrency', 'Currency', 'Rate'])
print(df)
list_of_currencies=['CAD', 'THB', 'USD']
for curr in list_of_currencies:
    data_to_insert=[data['date'],curr,data['rates'][curr]]

    print(data_to_insert)
print(df)
    #pd.to_csv
#data_to_insert=[data['date'],data['base'],data['rates']['CAD'],data['rates']['THB'],data['rates']['USD']]
#print(data_to_insert)