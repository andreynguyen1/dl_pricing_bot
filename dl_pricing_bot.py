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
        #prices = {sheet.cell(row=row[0], column=i).value: row[i] for i in range(1, len(row) - 1)}
        sku_index = int(sku)
        prices = {sheet.cell(row=sku_index, column=i).value: row[i] for i in range(1, len(row) - 1)}
        comment = row[-1]
        data[sku] = {'prices': prices, 'comment': comment}
    return data

# Загрузка данных из Excel-файла на Google Диске
excel_data = load_excel_data_from_google_drive(file_id)

# Обработка команды /start
@bot.message_handler(commands=['start'])
def send_welcome(message):
    bot.reply_to(message, "Привет! Я бот, который поможет тебе найти значения цен по SKU и типу цены.")

# Обработка текстовых сообщений
@bot.message_handler(func=lambda message: True)
def handle_message(message):
    if message.text.startswith('/'):
        return
    if message.text in excel_data:
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        for price_type in excel_data[message.text]['prices']:
            markup.add(types.KeyboardButton(price_type))
        bot.send_message(message.chat.id, "Выберите тип цены:", reply_markup=markup)
        bot.register_next_step_handler(message, process_price_type_selection)
    else:
        bot.reply_to(message, "SKU не найден.")

# Обработка выбора типа цены
def process_price_type_selection(message):
    price_type = message.text
    sku = message.chat.id  # Предполагаем, что id чата можно использовать как SKU
    if sku in excel_data:
        if price_type in excel_data[sku]['prices']:
            price = excel_data[sku]['prices'][price_type]
            comment = excel_data[sku]['comment']
            if comment:
                bot.send_message(message.chat.id, f"Цена для SKU {sku} ({price_type}): {price}. Комментарий: {comment}")
            else:
                bot.send_message(message.chat.id, f"Цена для SKU {sku} ({price_type}): {price}")
        else:
            bot.send_message(message.chat.id, f"Тип цены {price_type} для SKU {sku} не найден.")
    else:
        bot.send_message(message.chat.id, f"SKU {sku} не найден.")

# Запуск бота
bot.polling()
