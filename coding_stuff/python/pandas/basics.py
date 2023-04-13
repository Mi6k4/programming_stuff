import pandas as pd

my_series=pd.Series([5,6,7,8,9,10,11])
#print(my_series.index)
#print(my_series.values)
#print(my_series[4])
my_series2=pd.Series([5,6,7,8,9,10,11],index=['a','b','c','d','e','f','g'])
#print(my_series2)


df = pd.DataFrame({
    'country': ['Kazakhstan', 'Russia', 'Belarus', 'Ukraine'],
    'population': [17.04, 143.5, 9.5, 45.5],
    'square': [2724902, 17125191, 207600, 603628]
})
print(df)
print(df['country'])
print(df.columns)