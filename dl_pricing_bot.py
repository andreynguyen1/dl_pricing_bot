import telebot
from openpyxl import load_workbook
import requests
from io import BytesIO
from telebot import types

# Инициализация бота
bot = telebot.TeleBot('7123896415:AAHaVQeFj4Pc9_zlxi3DuprdRH9RIO1qoK4')

# ID файла Excel на Google Диске (из его общедоступной ссылки)
file_id = '1fo3Wbb6noDILXL96qkodV4uVInZYZY-p'

# Функция загрузки данных из Excel-файла, расположенного на Google Диске
def load_excel_data_from_google_drive(file_id):
    url = f'https://drive.google.com/uc?export=download&id={file_id}'
    response = requests.get(url)
    workbook = load_workbook(filename=BytesIO(response.content))
    sheet = workbook.active
    data = {}
    for row in sheet.iter_rows(values_only=True):
        sku = row[0]
        if sku == 'SKU':  # Пропускаем заголовок
            continue
        if sku == input_sku:  # Проверяем точное совпадение SKU
            prices = {sheet.cell(row=sku_index, column=i).value: row[i] for i in range(1, len(row) - 1)}
            comment = row[-1]
            data[sku] = {'prices': prices, 'comment': comment}
            break  # Прекращаем поиск после нахождения точного совпадения
    return data

# Загрузка данных из Excel-файла на Google Диске
excel_data = load_excel_data_from_google_drive(file_id)

# Обработка команды /start
@bot.message_handler(commands=['start'])
def send_welcome(message):
    bot.reply_to(message, "Привет! Я бот, который поможет тебе найти значения цен и комментариев по SKU и типу цены.")

# Обработка текстовых сообщений
@bot.message_handler(func=lambda message: True)
def handle_message(message):
    global user_price_type, user_sku
    if user_price_type is None:
        user_price_type = message.text
        bot.reply_to(message, "Теперь введите SKU:")
    elif user_sku is None:
        user_sku = message.text
        excel_data = load_excel_data_from_google_drive(file_id)
        if user_sku in excel_data:
            price = excel_data[user_sku]['price']
            comment = excel_data[user_sku]['comment']
            if comment:
                bot.reply_to(message, f"Цена для SKU {user_sku} ({user_price_type}): {price}. Комментарий: {comment}")
            else:
                bot.reply_to(message, f"Цена для SKU {user_sku} ({user_price_type}): {price}")
        else:
            bot.reply_to(message, f"SKU {user_sku} не найден.")

# Запуск бота
bot.polling()