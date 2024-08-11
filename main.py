import openpyxl as op
import random
import telebot
from telebot import types

TOKEN = "7379053950:AAEDhfRYutZmdcFxXMOut8ZjbSYmngudYSA" #Token for Telegram Bot
bot = telebot.TeleBot(TOKEN)

df = op.load_workbook('Courses.xlsx')
data = df.active
total_rows = data.max_row
words = []
definitions = []

#Creating list with words
for row_number in range(3,total_rows+1):
    cell_obj = data.cell(row=row_number, column=1)
    words.append(cell_obj.value)
print(words)

#Creating list with definitions
for row_number in range(3,total_rows+1):
    cell_obj = data.cell(row=row_number, column=3)
    definitions.append(cell_obj.value)
print(definitions)

#Choosing the random row number
#random_row_number=int(random.uniform(3,total_rows+1))
#print("For word *",words[random_row_number],"* definition is *",definitions[random_row_number],"*")

#Choosing 4 random rows for game round
random_row_1=int(random.uniform(3,total_rows+1))
random_row_2=int(random.uniform(3,total_rows+1))
random_row_3=int(random.uniform(3,total_rows+1))
random_row_4=int(random.uniform(3,total_rows+1))
# print("For word 1 *",words[random_row_1],"* definition is *",definitions[random_row_1],"*")
# print("For word 2 *",words[random_row_2],"* definition is *",definitions[random_row_2],"*")
# print("For word 3 *",words[random_row_3],"* definition is *",definitions[random_row_3],"*")
# print("For word 4 *",words[random_row_4],"* definition is *",definitions[random_row_4],"*")

#Telebot /start command
@bot.message_handler(commands=['start'])
def start(message):
    markup = types.InlineKeyboardMarkup()
    markup.add(types.InlineKeyboardButton("Go to Cambridge dictionary website", url='https://dictionary.cambridge.org/plus/myWordlists'))
    bot.send_message(message.chat.id, f'Hello, {message.from_user.first_name}! Upload Exel file with your words list from Cambrdge dictinoray',reply_markup=markup)

#Exel file uploading handler
@bot.message_handler(content_types=['photo'])
def get_file(message):
    markup2 = types.InlineKeyboardMarkup()
    markup2.add(types.InlineKeyboardButton("Start game",callback_data='start_game'))
    bot.reply_to(message, "Your words list uploaded successfully",reply_markup=markup2)




bot.polling(non_stop=True)




