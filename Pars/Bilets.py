import requests
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import time
import os

token = os.environ.get("AVIASALES_TOKEN")
url = 'https://api.travelpayouts.com/aviasales/v3/prices_for_dates'

# Авторизация в Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)
spreadsheet = client.open("Avia Sexuals")

# Получаем или создаем два листа: туда и обратно
try:
    worksheet_to = spreadsheet.worksheet("Билеты туда")
except:
    worksheet_to = spreadsheet.add_worksheet(title="Билеты туда", rows="100", cols="20")

try:
    worksheet_back = spreadsheet.worksheet("Билеты обратно")
except:
    worksheet_back = spreadsheet.add_worksheet(title="Билеты обратно", rows="100", cols="20")

worksheet_to.clear()
worksheet_back.clear()

# Заголовки
worksheet_to.append_row([
    "Дата", "Цена", "Авиакомпания", "Номер рейса", "Время в пути (ч:м)", "Маршрут"
])
worksheet_back.append_row([
    "Дата", "Цена", "Авиакомпания", "Номер рейса", "Время в пути (ч:м)", "Маршрут"
])

# Карта кодов в города
airport_city_map = {
    'MOW': 'Москва', 'KZN': 'Казань', 'CEK': 'Челябинск', 'SVX': 'Екатеринбург',
    'KIX': 'Осака', 'TYO': 'Токио', 'FUK': 'Фукуока'
}

# Отдельные списки
origins_to = ['MOW', 'KZN', 'CEK', 'SVX']
destinations_to = ['KIX', 'TYO', 'FUK']
origins_back = ['KIX', 'TYO']
destinations_back = ['KZN', 'CEK', 'SVX']

start_date = datetime(2025, 9, 1)
end_date = datetime(2025, 10, 20)

tickets_to = []
tickets_back = []

while start_date <= end_date:
    departure_date = start_date.strftime('%Y-%m-%d')
    return_date = (start_date + timedelta(days=10)).strftime('%Y-%m-%d')

    # Туда
    for origin in origins_to:
        for destination in destinations_to:
            params_to = {
                'origin': origin,
                'destination': destination,
                'currency': 'rub',
                'sorting': 'price',
                'limit': 30,
                'token': token
            }
            response = requests.get(url, params=params_to)
            data_to = response.json()

            if data_to.get('success') and data_to.get('data'):
                for ticket in data_to['data']:
                    if ticket.get('transfers', 2) <= 1:
                        duration_min = ticket.get('duration', 0)
                        hours, minutes = divmod(duration_min, 60)
                        tickets_to.append([
                            departure_date,
                            ticket.get('price', '—'),
                            ticket.get('airline', '—'),
                            ticket.get('flight_number', '—'),
                            f"{hours}ч {minutes}м",
                            f"{airport_city_map.get(origin)} → {airport_city_map.get(destination)}"
                        ])
                        break  # Берем только первый подходящий билет

    # Обратно
    for origin in origins_back:
        for destination in destinations_back:
            params_back = {
                'origin': origin,
                'destination': destination,
                'currency': 'rub',
                'sorting': 'price',
                'limit': 30,
                'token': token
            }
            response = requests.get(url, params=params_back)
            data_back = response.json()

            if data_back.get('success') and data_back.get('data'):
                for ticket in data_back['data']:
                    if ticket.get('transfers', 2) <= 1:
                        duration_min = ticket.get('duration', 0)
                        hours, minutes = divmod(duration_min, 60)
                        tickets_back.append([
                            return_date,
                            ticket.get('price', '—'),
                            ticket.get('airline', '—'),
                            ticket.get('flight_number', '—'),
                            f"{hours}ч {minutes}м",
                            f"{airport_city_map.get(origin)} → {airport_city_map.get(destination)}"
                        ])
                        break

    time.sleep(1)
    start_date += timedelta(days=1)

# Сортировка по цене
tickets_to.sort(key=lambda x: x[1])
tickets_back.sort(key=lambda x: x[1])

# Запись в Google Sheets
if tickets_to:
    worksheet_to.append_rows(tickets_to)

time.sleep(2)

if tickets_back:
    worksheet_back.append_rows(tickets_back)


today = datetime.today().strftime('%d-%m-%Y')

worksheet_back.append_row(["", "", "", "", "", f"Дата запроса: {today}"])