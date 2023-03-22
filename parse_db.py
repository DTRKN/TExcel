import openpyxl
import sqlite3
import random

wb = openpyxl.load_workbook('Excel.xlsx')
sheet = wb['Лист1']

data = []
for row in sheet.iter_rows(values_only=True):
    data.append(list(row))

conn = sqlite3.connect('data.db')
cur = conn.cursor()
cur.execute('''
    CREATE TABLE IF NOT EXISTS data (
        id INTEGER PRIMARY KEY,
        company TEXT,
        date TEXT,
        fact_qliq INTEGER,
        fact_qoil INTEGER,
        forecast_qliq INTEGER,
        forecast_qoil INTEGER
    )
''')

for row in data[1:]:
    id = row[0]
    if id is not None:
        company = row[1]
        date = "2023-01-%02d" % (id % 10 + random.randrange(0, 20))  # выбираем произвольные даты
        fact_qliq = row[2]
        fact_qoil = row[3]
        forecast_qliq = row[4]
        forecast_qoil = row[5]
        cur.execute('''
            INSERT INTO data (id, company, date, fact_qliq, fact_qoil, forecast_qliq, forecast_qoil)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (id, company, date, fact_qliq, fact_qoil, forecast_qliq, forecast_qoil))

conn.commit()

for row in cur.fetchall():
    date = row[0]
    fact_qliq_total = row[1]
    fact_qoil_total = row[2]
    forecast_qliq_total = row[3]
    forecast_qoil_total = row[4]
    print(date, fact_qliq_total, fact_qoil_total,
          forecast_qliq_total, forecast_qoil_total)

conn.close()