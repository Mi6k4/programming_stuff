import requests
import sqlite3
import pandas as pd


api_url = "https://jsonplaceholder.typicode.com/users"
response = requests.get(api_url)
data = response.json()

conn = sqlite3.connect('data.db')
cursor = conn.cursor()


cursor.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY,
        name TEXT,
        username TEXT,
        email TEXT,
        phone TEXT
    )
""")


df = pd.json_normalize(data)


for _, row in df.iterrows():
    cursor.execute("""
        INSERT INTO users (id, name, username, email, phone)
        VALUES (?, ?, ?, ?, ?)
    """, (row['id'], row['name'], row['username'], row['email'], row['phone']))


conn.commit()


conn.close()

print("Данные успешно получены и сохранены в базу данных!")