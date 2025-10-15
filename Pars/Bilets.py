import requests
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import time
import os
import smtplib
from email.mime.text import MIMEText
from dotenv import load_dotenv
load_dotenv()

token = os.environ.get("AVIASALES_TOKEN")
url = 'https://api.travelpayouts.com/aviasales/v3/prices_for_dates'

#email рассылка
SMTP_SERVER = "smtp.yandex.ru"
SMTP_PORT = 465
EMAIL_SENDER = "m.krylov.a@yandex.ru"
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD")
EMAIL_RECEIVERS = ["cattrap.3s@gmail.com"]


THRESHOLD_TO = 26000
THRESHOLD_BACK = 37000

watch_routes_to = {'KZN→KIX', 'KZN→HND', 'KZN→NRT', 'KZN→FUK', 'KZN→NGO', 'KZN→SPK', 'KZN→KMJ', 'KZN→SDJ', 'KZN→TOY'}
watch_routes_back = {'KIX→KZN', 'HND→KZN', 'NRT→KZN', 'FUK→KZN', 'NGO→KZN', 'SPK→KZN', 'KMJ→KZN'}

alerts = []

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
    'KZN': 'Казань', 'NGO': 'Нагоя', 'SPK': 'Саппоро', 'KMJ': 'Кумамото',
    'KIX': 'Осака', 'HND': 'Токио.Ханеда', 'NRT': 'Токио.Нарита', 'FUK': 'Фукуока',
    'SDJ': 'Sendai', 'TOY': 'Toyama'
}

# Отдельные списки
origins_to = ['KZN']
destinations_to = ['KIX', 'NRT', 'FUK', 'HND', 'NGO', 'SPK', 'KMJ']
origins_back = ['KIX', 'NRT', 'HND', 'FUK', 'NGO', 'SPK', 'KMJ']
destinations_back = ['KZN']


tickets_to = []
tickets_back = []


# Туда
for origin in origins_to:
    for destination in destinations_to:
        params_to = {
            'origin': origin,
            'destination': destination,
            'currency': 'rub',
            'sorting': 'price',
            'limit': 50,
            'token': token,
            'one_way': 'true'
        }
        response = requests.get(url, params=params_to)
        data_to = response.json()

        if data_to.get('success') and data_to.get('data'):
            for ticket in data_to['data']:
                if ticket.get('transfers', 2) <= 2:
                    duration_min = ticket.get('duration', 0)
                    hours, minutes = divmod(duration_min, 60)
                    price = ticket.get('price', 999999)
                    
                    # Преобразование даты из API в нужный формат
                    date_from_api = ticket.get('departure_at', '—')
                    if date_from_api != '—':
                        try:
                            parsed_date = datetime.fromisoformat(date_from_api.replace('Z', '+00:00'))
                            formatted_date = parsed_date.strftime('%d.%m.%Y')
                            
                            # Проверка диапазона дат для ТУДА
                            ticket_date = parsed_date.date()
                            range_start = datetime(2026, 3, 19).date()
                            range_end = datetime(2026, 5, 9).date()
                            
                            if range_start <= ticket_date <= range_end:
                                route_key = f"{origin}→{destination}"
                                if route_key in watch_routes_to and price <= THRESHOLD_TO:
                                    alerts.append(
                                        f"🔥 Дешевый билет ТУДА!\n"
                                        f"{airport_city_map[origin]} → {airport_city_map[destination]}\n"
                                        f"Дата: {formatted_date}\n"
                                        f"Цена: {price} руб.\n"
                                        f"Авиакомпания: {ticket.get('airline', '—')}\n"
                                        f"Рейс: {ticket.get('flight_number', '—')}"
                                        f"-----------------------------------------------------\n"
                                    )
                                    
                                tickets_to.append([
                                    formatted_date, price, ticket.get('airline', '—'), ticket.get('flight_number', '—'),
                                    f"{hours}ч {minutes}м",
                                    f"{airport_city_map.get(origin)} → {airport_city_map.get(destination)}"
                                ])
                        except:
                            formatted_date = '—'
        time.sleep(1)


# Обратно
for origin in origins_back:
    for destination in destinations_back:
        params_back = {
            'origin': origin,
            'destination': destination,
            'currency': 'rub',
            'sorting': 'price',
            'limit': 50,
            'token': token,
            'one_way': 'true'
        }
        response = requests.get(url, params=params_back)
        data_back = response.json()

        if data_back.get('success') and data_back.get('data'):
            for ticket in data_back['data']:
                if ticket.get('transfers', 2) <= 2:
                    duration_min = ticket.get('duration', 0)
                    hours, minutes = divmod(duration_min, 60)
                    price = ticket.get('price', 999999)

                    date_from_api = ticket.get('departure_at', '—')
                    if date_from_api != '—':
                        try:
                            parsed_date = datetime.fromisoformat(date_from_api.replace('Z', '+00:00'))
                            formatted_date = parsed_date.strftime('%d.%m.%Y')
                            
                            # Проверка диапазона дат для ОБРАТНО
                            ticket_date = parsed_date.date()
                            range_start = datetime(2026, 3, 25).date()
                            range_end = datetime(2026, 5, 15).date()
                            
                            if range_start <= ticket_date <= range_end:
                                route_key = f"{origin}→{destination}"
                                if route_key in watch_routes_back and price <= THRESHOLD_BACK:
                                    alerts.append(
                                        f"🔥 Дешевый билет ОБРАТНО!\n"
                                        f"{airport_city_map[origin]} → {airport_city_map[destination]}\n"
                                        f"Дата: {formatted_date}\n"
                                        f"Цена: {price} руб.\n"
                                        f"Авиакомпания: {ticket.get('airline', '—')}\n"
                                        f"Рейс: {ticket.get('flight_number', '—')}"
                                        f"-----------------------------------------------------\n"
                                    )
            
                                tickets_back.append([
                                    formatted_date, price, ticket.get('airline', '—'), ticket.get('flight_number', '—'),
                                    f"{hours}ч {minutes}м",
                                    f"{airport_city_map.get(origin)} → {airport_city_map.get(destination)}"
                                ])
                        except:
                            formatted_date = '—'
        time.sleep(1)



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

# --- Отправка email если есть дешевые билеты ---
if alerts:
    message_body = "\n\n".join(alerts)
    msg = MIMEText(message_body, "plain", "utf-8")
    msg['Subject'] = "🚨 Внимание! Внимание! Найдены дешевые авиабилеты! Кто не купит тот гау"
    msg['From'] = EMAIL_SENDER
    msg['To'] = ", ".join(EMAIL_RECEIVERS)

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.send_message(msg)
    except Exception as e:
        print("❌ Ошибка при отправке email:", e)
