# @title Загрузим xlsm-файл из GitHub и преобразуем его в tsv-формат

import pandas as pd
import requests
import io
import os
from pathlib import Path

print("Загрузка файла c GitHub (используем raw ссылку)...")
url = "https://github.com/dashleb33/DataSets_for_HomeWork/raw/main/Dynamic_Pricing/lenta.xlsm"
response = requests.get(url)
response.raise_for_status()  # Проверяем успешность запроса

# Разбиваем URL по '/' и берем последнюю часть
filename = url.split('/')[-1]
base_name = filename.rsplit('.', 1)[0]

excel_data = io.BytesIO(response.content)  # Создаем байтовый поток для чтения
print("Чтение Excel файла с помощью pandas...")
df = pd.read_excel(excel_data, engine='openpyxl')  # Используем engine='openpyxl' для работы с .xlsm файлами
print(f"Файл успешно загружен!")
print(f"Размер данных: {df.shape}")
print(f"Колонки: {df.columns.tolist()}")

print("\nПреобразование в TSV-формат...")
output_dir = "/content/output"  # Создаем директорию для сохранения (если нужно)
Path(output_dir).mkdir(parents=True, exist_ok=True)
output_filename = f"{base_name}.tsv"
output_path = os.path.join(output_dir, output_filename)
df.to_csv(output_path, sep='\t', index=False)

print(f"Файл сохранен как: {output_path}")
print(f"Размер TSV файла: {os.path.getsize(output_path) / 1024:.2f} KB")
print("\n" + "="*50)
print("Сводка по данным:")
print("="*50)
print(f"Количество строк: {len(df)}")
print(f"Количество колонок: {len(df.columns)}")
print("\nТипы данных:")
print(df.dtypes)
print("\nПроверка на пропущенные значения:")
print(df.isnull().sum())
