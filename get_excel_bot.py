import telebot
import datetime as dt
from telebot import types




# create bot with hash code from BotFather
bot = telebot.TeleBot('5587191876:AAFzEuQgRSlcJsMLUxXxHYLKo-1-4SdEACE')

name = ''

# start Bot
@bot.message_handler(commands=['start'])
def start(m, res=False):
    # Create buttons
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    button1 = types.KeyboardButton('Белотелов')
    button2 = types.KeyboardButton('Орлова')
    button3 = types.KeyboardButton('Кулик')
    button4 = types.KeyboardButton('Сброс данных по наименованию отчета')
    markup.add(button1)
    markup.add(button2)
    markup.add(button3)
    markup.add(button4)
    bot.send_message(m.chat.id,
                     'Чей отчет подгружается', reply_markup=markup)


    @bot.message_handler(content_types=['text'])
    def mes_name(message):
        global name
        if message.text != 'Вернуться назад/Отмена':
            name = message.text
        else:
            @bot.message_handler(commands=['Вернуться назад/Отмена'])
            def handle_start(message):
                menu_plumbing = True
                if menu_plumbing:
                    global name
                    name = ''
                    menu_plumbing = False
                    # тут можно уже или как то так
                    handle_start('Другое меню')


    @bot.message_handler(content_types=['document'])
    def handle_file(message):
        try:
            file_info = bot.get_file(message.document.file_id)
            downloaded_file = bot.download_file(file_info.file_path)
            src = 'excel_docs/' + f'{name}-{dt.datetime.date(dt.datetime.now())}';
            with open(src, 'wb') as new_file:
                new_file.write(downloaded_file)
            bot.reply_to(message, "Сохранено")
        except Exception as e:
            bot.reply_to(message, e)

# launch bot
bot.polling(none_stop=True, interval=0)