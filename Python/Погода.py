import requests
import pandas as pd
from datetime import datetime

# Выполняем GET запрос к API
url = 'https://weather.sakhalin.gov.ru/api/air/now'

# Добавляем заголовки для имитации браузера
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'application/json',
    'Accept-Language': 'ru-RU,ru;q=0.9,en;q=0.8',
    'Accept-Encoding': 'gzip, deflate',
    'Connection': 'keep-alive'
}

print("Отправка запроса с User-Agent...")
response = requests.get(url, headers=headers)

# Проверяем статус ответа
print(f"Статус код: {response.status_code}")

if response.status_code == 200:
    # Получаем данные в формате JSON
    data = response.json()
    print(f"Получено записей: {len(data) if isinstance(data, list) else 'не список'}")
    
    # Функция для извлечения нужных полей с русскими названиями
    def process_weather_data(data):
        processed_data = []
        
        for item in data:
            # Создаем базовый словарь с основной информацией на русском
            row = {
                'ID': item.get('id'),
                'Внешний ID': item.get('ext_id'),
                'Населенный пункт': item.get('name'),
                'Широта': item.get('latitude'),
                'Долгота': item.get('longitude'),
                'Дата и время': item.get('date'),
                'Индекс качества воздуха (AQI)': item.get('aqi')
            }
            
            # Извлекаем значения из вложенных словарей с русскими названиями
            if 'temperature' in item and item['temperature']:
                row['Температура (значение)'] = item['temperature'].get('value')
                row['Температура (единица)'] = item['temperature'].get('unit')
                
            if 'pressure' in item and item['pressure']:
                row['Давление (значение)'] = item['pressure'].get('value')
                row['Давление (единица)'] = item['pressure'].get('unit')
                
            if 'humidity' in item and item['humidity']:
                row['Влажность (значение)'] = item['humidity'].get('value')
                row['Влажность (единица)'] = item['humidity'].get('unit')
                
            if 'pm2' in item and item['pm2']:
                row['PM2.5 (значение)'] = item['pm2'].get('value')
                row['PM2.5 (единица)'] = item['pm2'].get('unit')
                
            if 'pm10' in item and item['pm10']:
                row['PM10 (значение)'] = item['pm10'].get('value')
                row['PM10 (единица)'] = item['pm10'].get('unit')
                
            if 'co' in item and item['co']:
                row['Угарный газ CO (значение)'] = item['co'].get('value')
                row['Угарный газ CO (единица)'] = item['co'].get('unit')
                
            if 'no2' in item and item['no2']:
                row['Диоксид азота NO2 (значение)'] = item['no2'].get('value')
                row['Диоксид азота NO2 (единица)'] = item['no2'].get('unit')
                
            if 'so2' in item and item['so2']:
                row['Диоксид серы SO2 (значение)'] = item['so2'].get('value')
                row['Диоксид серы SO2 (единица)'] = item['so2'].get('unit')
                
            if 'o3' in item and item['o3']:
                row['Озон O3 (значение)'] = item['o3'].get('value')
                row['Озон O3 (единица)'] = item['o3'].get('unit')
                
            if 'h2s' in item and item['h2s']:
                row['Сероводород H2S (значение)'] = item['h2s'].get('value')
                row['Сероводород H2S (единица)'] = item['h2s'].get('unit')
                
            if 'wda' in item and item['wda']:
                row['Направление ветра (значение)'] = item['wda'].get('value')
                row['Направление ветра (единица)'] = item['wda'].get('unit')
                
            if 'wva' in item and item['wva']:
                row['Скорость ветра (значение)'] = item['wva'].get('value')
                row['Скорость ветра (единица)'] = item['wva'].get('unit')
                
            if 'ch2o' in item and item['ch2o']:
                row['Формальдегид CH2O (значение)'] = item['ch2o'].get('value')
                row['Формальдегид CH2O (единица)'] = item['ch2o'].get('unit')
            
            processed_data.append(row)
        
        return processed_data
    
    # Обрабатываем данные
    processed_data = process_weather_data(data)
    
    # Создаем DataFrame
    df = pd.DataFrame(processed_data)
    
    # Сохраняем в Excel файл с русским именем
    filename = f'погодные_данные.xlsx'
    df.to_excel(filename, index=False, engine='openpyxl')
    
    print(f"\n✅ Данные успешно сохранены в файл: {filename}")
    print(f"📊 Всего записей: {len(df)}")
    print(f"📋 Столбцов в файле: {len(df.columns)}")
    print(f"\nСписок столбцов:")
    for col in df.columns:
        print(f"  • {col}")
    
else:
    print(f"❌ Ошибка при получении данных: {response.status_code}")
    print(f"Текст ошибки: {response.text}")       