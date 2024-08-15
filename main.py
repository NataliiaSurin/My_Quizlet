import openpyxl as op
import random
import telebot
from telebot import types
from urllib.request import urlretrieve

TOKEN = "7379053950:AAEDhfRYutZmdcFxXMOut8ZjbSYmngudYSA"  # Token for Telegram Bot
bot = telebot.TeleBot(TOKEN)
user_counters = {}  # Initiation dictionary for score counting inside bot.callback_query_handler
global total_rows


# Telebot /start command
@bot.message_handler(commands=['start'])
def start(message):
    user_id = message.from_user.id
    user_counters[user_id] = 0
    markup = types.InlineKeyboardMarkup()
    markup.add(types.InlineKeyboardButton("Go to the Cambridge dictionary website", url='https://dictionary.cambridge.org/plus/myWordlists'))
    bot.send_message(message.chat.id, f'Hello, {message.from_user.first_name} ! Upload Exel file with your words list from Cambridge dictionary', reply_markup=markup)


# Exel file uploading handler
@bot.message_handler(content_types=['document'])
def get_file(message):
    markup2 = types.InlineKeyboardMarkup()
    markup2.add(types.InlineKeyboardButton("Start game", callback_data='start_game'))
    bot.reply_to(message, "Your words list uploaded successfully", reply_markup=markup2)
    bot.get_file(message.document.file_id)
    url = str(bot.get_file_url(message.document.file_id))
    file_name = message.document.file_name
    df = op.load_workbook(file_name)
    data = df.active
    print(url, file_name)
    urlretrieve(url, file_name)
    global total_rows
    total_rows = data.max_row
    global words, definitions
    words = []
    definitions = []

    # Creating list with words
    for row_number in range(3, total_rows + 1):
        cell_obj = data.cell(row=row_number, column=1)
        words.append(cell_obj.value)
    print(words)

    # Creating list with definitions
    for row_number in range(3, total_rows + 1):
        cell_obj = data.cell(row=row_number, column=3)
        definitions.append(cell_obj.value)
    print(definitions)


# Choosing 4 random rows for game round
def randomize(total_rows):
    global random_row_1, random_row_2, random_row_3, random_row_4, random_row_main
    random_row_1 = int(random.uniform(3, total_rows+1))
    random_row_2 = int(random.uniform(3, total_rows+1))
    random_row_3 = int(random.uniform(3, total_rows+1))
    random_row_4 = int(random.uniform(3, total_rows+1))
    random_row_main = int(random.choice([random_row_1, random_row_2, random_row_3, random_row_4]))


# Function for creating new game round with 4 answer variants buttons
def send_question(callback):
    randomize(total_rows)
    markup3 = types.InlineKeyboardMarkup()
    markup3.add(types.InlineKeyboardButton(f"{words[random_row_1]}", callback_data=str(random_row_1)))
    markup3.add(types.InlineKeyboardButton(f"{words[random_row_2]}", callback_data=str(random_row_2)))
    markup3.add(types.InlineKeyboardButton(f"{words[random_row_3]}", callback_data=str(random_row_3)))
    markup3.add(types.InlineKeyboardButton(f"{words[random_row_4]}", callback_data=str(random_row_4)))
    bot.send_message(callback.message.chat.id, f'Choose the right word that best matches to this definition: <b>{definitions[random_row_main]}</b>', parse_mode='html', reply_markup=markup3)
    print(random_row_1, random_row_2, random_row_3, random_row_4, random_row_main)


# Analysis pressed buttons, checking for answer correctness, score counting
@bot.callback_query_handler(func=lambda callback: True)
def callback_message(callback):
    user_id = callback.from_user.id
    print(user_counters[user_id])
    if callback.data == 'start_game':
        send_question(callback)
        bot.answer_callback_query(callback.id)
        print(callback.data)
    elif callback.data == str(random_row_main):
        if user_id in user_counters:
            user_counters[user_id] += 1  # Score counting
        print(f'Correct, you score now {user_counters[user_id]}')
        bot.answer_callback_query(callback.id)
        bot.send_message(callback.message.chat.id, f'Correct, you score now {user_counters[user_id]}')
        send_question(callback)
    else:
        print(f'Wrong, you score now {user_counters[user_id]}')
        bot.answer_callback_query(callback.id)
        bot.send_message(callback.message.chat.id, f'Wrong, you score now {user_counters[user_id]}')
        send_question(callback)


bot.polling(non_stop=True)
