import requests
import json
import pandas as pd


api_url = "https://jsonplaceholder.typicode.com/posts"
response = requests.get(api_url)
data = response.json()


with open("data.json", "w") as f:
    json.dump(data, f, indent=4)


df = pd.json_normalize(data)


df.to_csv("normalized_data.csv", index=False)

print("Данные успешно получены, сохранены и нормализованы!")