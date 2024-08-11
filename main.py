import openpyxl as op
import random
import telebot
from telebot import types

TOKEN = "7379053950:AAEDhfRYutZmdcFxXMOut8ZjbSYmngudYSA"  # Token for Telegram Bot
bot = telebot.TeleBot(TOKEN)

df = op.load_workbook('Courses.xlsx')
data = df.active
total_rows = data.max_row
words = []
definitions = []

# Creating list with words
for row_number in range(3, total_rows+1):
    cell_obj = data.cell(row=row_number, column=1)
    words.append(cell_obj.value)
print(words)

# Creating list with definitions
for row_number in range(3, total_rows+1):
    cell_obj = data.cell(row=row_number, column=3)
    definitions.append(cell_obj.value)
print(definitions)

# Choosing the random row number
# random_row_number=int(random.uniform(3,total_rows+1))
# print("For word *",words[random_row_number],"* definition is *",definitions[random_row_number],"*")


# Choosing 4 random rows for game round
def randomize(total_rows):
    global random_row_1, random_row_2, random_row_3, random_row_4, random_row_main
    random_row_1 = int(random.uniform(3, total_rows+1))
    random_row_2 = int(random.uniform(3, total_rows+1))
    random_row_3 = int(random.uniform(3, total_rows+1))
    random_row_4 = int(random.uniform(3, total_rows+1))
    random_row_main = int(random.choice([random_row_1, random_row_2, random_row_3, random_row_4]))

    return random_row_main, random_row_1, random_row_2, random_row_3, random_row_4


# print("For word 1 *",words[random_row_1],"* definition is *",definitions[random_row_1],"*")
# print("For word 2 *",words[random_row_2],"* definition is *",definitions[random_row_2],"*")
# print("For word 3 *",words[random_row_3],"* definition is *",definitions[random_row_3],"*")
# print("For word 4 *",words[random_row_4],"* definition is *",definitions[random_row_4],"*")


# Telebot /start command
@bot.message_handler(commands=['start'])
def start(message):
    markup = types.InlineKeyboardMarkup()
    markup.add(types.InlineKeyboardButton("Go to Cambridge dictionary website", url='https://dictionary.cambridge.org/plus/myWordlists'))
    bot.send_message(message.chat.id, f'Hello, {message.from_user.first_name} ! Upload Exel file with your words list from Cambrdge dictinoray', reply_markup=markup)


# Exel file uploading handler
@bot.message_handler(content_types=['photo'])
def get_file(message):
    markup2 = types.InlineKeyboardMarkup()
    markup2.add(types.InlineKeyboardButton("Start game", callback_data='start_game'))
    bot.reply_to(message, "Your words list uploaded successfully", reply_markup=markup2)


@bot.callback_query_handler(func=lambda callback: True)
def callback_message(callback):
    if callback.data == 'start_game':
        randomize(total_rows)
        markup3 = types.InlineKeyboardMarkup()
        markup3.add(types.InlineKeyboardButton(f"{definitions[random_row_1]}", callback_data=str(random_row_1)))
        markup3.add(types.InlineKeyboardButton(f"{definitions[random_row_2]}", callback_data=str(random_row_2)))
        markup3.add(types.InlineKeyboardButton(f"{definitions[random_row_3]}", callback_data=str(random_row_3)))
        markup3.add(types.InlineKeyboardButton(f"{definitions[random_row_4]}", callback_data=str(random_row_4)))
        bot.send_message(callback.message.chat.id, f'Choose the right definition for word <b>{words[random_row_main]}</b>', parse_mode='html', reply_markup=markup3)


bot.polling(non_stop=True)
