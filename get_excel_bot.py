import json
import os
import warnings
from logging.handlers import RotatingFileHandler
from dotenv import load_dotenv

import logging

import openpyxl
import telebot
import datetime as dt
from telebot import types

from pars_table import dict_article_count, update_table_count_fbo

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
log_dir = os.path.join(BASE_DIR, 'logs/')
log_file = os.path.join(BASE_DIR, 'logs/bot.log')
console_handler = logging.StreamHandler()
file_handler = RotatingFileHandler(
    log_file,
    maxBytes=100000,
    backupCount=3,
    encoding='utf-8'
)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s, %(levelname)s, %(message)s',
    handlers=(
        file_handler,
        console_handler
    )
)

dotenv_path = os.path.join(os.path.dirname(__file__), '.env')


if os.path.exists(dotenv_path):
    load_dotenv(dotenv_path)
load_dotenv('.env ')
token = os.getenv('TELEGRAM_TOKEN')

bot = telebot.TeleBot(token)

date = '0'


@bot.message_handler(commands=['start'])
def start(message):
    markup_start = types.ReplyKeyboardMarkup(resize_keyboard=True)
    button_get_reports = types.KeyboardButton('Загрузить отчеты в бота')
    global date
    if date == '0':
        date = 'Дата не выбрана'
    button_update_table = types.KeyboardButton(f'Обновить данные в Гугл Таблицах ({date})')
    button_date = types.KeyboardButton(f'Выбрать дату для работы с таблицей')
    button_cancel = types.KeyboardButton('Сброс')
    button_start = types.KeyboardButton('/start')
    markup_start.add(button_get_reports)
    markup_start.add(button_update_table)
    markup_start.add(button_date)
    markup_start.add(button_cancel, button_start)
    bot.send_message(message.chat.id,
                     'Выберите действие', reply_markup=markup_start)
    if message.text == 'Сброс':
        global day
        global name
        global month
        global year

        day, date, name, month, year = '', 'Дата не выбрана', '', '', ''

    @bot.message_handler(content_types=['text'])
    def get_reports(message):
        global date
        global day
        if message.text == 'Загрузить отчеты в бота':
            markup_report_date = types.ReplyKeyboardMarkup(resize_keyboard=True)
            button_present = types.KeyboardButton('Загрузить отчет за текущий день')
            button_last = types.KeyboardButton('Загрузить отчет за прошедший день (итоговый)')
            button_random = types.KeyboardButton('Другая дата отчета')
            markup_report_date.add(button_present)
            markup_report_date.add(button_last)
            markup_report_date.add(button_random)
            button_cancel = types.KeyboardButton('Сброс')
            button_start = types.KeyboardButton('/start')
            markup_report_date.add(button_cancel, button_start)
            bot.reply_to(message,
                         'Выберите нужный период \nили укажите другую дату загрузки',
                         reply_markup=markup_report_date)
            bot.register_next_step_handler(message, get_date)
        if message.text == f'Обновить данные в Гугл Таблицах ({date})':
            if date !='Дата не выбрана':
                cred_file = os.path.join(BASE_DIR, 'credentials.json')
                with open(cred_file, 'r', encoding="utf-8") as f:
                    cred = json.load(f)
                try:
                    for i in cred:
                        if i != 'Савельева':
                            table_id = cred[i].get('table_id')
                            with warnings.catch_warnings(record=True):
                                warnings.simplefilter("always")
                                excel_file = openpyxl.load_workbook(f'excel_docs/{i}-{date}.xlsx')
                            employees_sheet = excel_file['Sheet1']
                            update_table_count_fbo(day, month, year, table_id,
                                                   dict_article_count(employees_sheet))
                            if i == 'Кулик':
                                table_id = cred[i].get('table_id_december')
                            with warnings.catch_warnings(record=True):
                                warnings.simplefilter("always")
                                excel_file = openpyxl.load_workbook(f'excel_docs/{i}-{date}.xlsx')
                            employees_sheet = excel_file['Sheet1']
                            update_table_count_fbo(day, month, year, table_id,
                                                   dict_article_count(employees_sheet))
                        from pars_table import report_article
                        if len(report_article) > 0:
                            d = report_article
                            bot.send_message(message.chat.id,
                                             f'В таблице отсутствуют следующие артикулы по которым имеются данные о продажах '
                                             f'в отчете {i}, внесите данные артикулы в Гугл таблицу и повторите выгрузку отчетов.',
                                             reply_markup=markup_start)
                            for i in d.keys():
                                bot.send_message(message.chat.id, text=i)
                    bot.reply_to(message,
                             'Гугл Таблица с отчетам за выбранный период обновлена.', reply_markup=markup_start)
                except:
                    bot.reply_to(message, 'Ошибка.\n', reply_markup=markup_start)
            else:
                bot.send_message(message.chat.id,
                             'Не задана дата для обновления таблицы.\n'
                             'Выберите дату.',
                             reply_markup=markup_start)
        if message.text == 'Выбрать дату для работы с таблицей':
            markup_date_table = types.ReplyKeyboardMarkup(resize_keyboard=True)
            button_cancel = types.KeyboardButton('Сброс')
            button_start = types.KeyboardButton('/start')
            markup_date_table.add(button_cancel, button_start)
            bot.send_message(message.chat.id, 'Сообщите боту дату которую необходимо обновить в таблице в формате ГГГГ-ММ-ДД\n'
                                              'пример: 2022-01-15')
            bot.register_next_step_handler(message, table_date)
        if message.text == 'Сброс':
            bot.send_message(message.chat.id, 'Подтвердите сброс повторным нажатием на соотвествущий пункт')
            bot.register_next_step_handler(message, cancel)

    def get_date(message):
        markup_report_name = types.ReplyKeyboardMarkup(resize_keyboard=True)
        button_bel = types.KeyboardButton('Белотелов')
        button_orl = types.KeyboardButton('Орлова')
        button_kul = types.KeyboardButton('Кулик')
        markup_report_name.add(button_bel, button_kul, button_orl)
        markup_report_name.add(button_cancel, button_start)
        global day
        global date
        global month
        global year
        if message.text == 'Загрузить отчет за текущий день':
            day = dt.datetime.now().strftime('%d')
            month = dt.datetime.now().strftime("%m")
            year = dt.datetime.now().strftime("%Y")
            date = dt.datetime.date(dt.datetime.now())
            bot.send_message(message.chat.id, 'Выберите наименования отчета для загрузки',
                             reply_markup=markup_report_name)
            bot.register_next_step_handler(message, final_get_report)
        elif message.text == 'Загрузить отчет за прошедший день (итоговый)':
            day = (dt.datetime.now() - dt.timedelta(days=1)).strftime('%d')
            month = (dt.datetime.now()- dt.timedelta(days=1)).strftime("%m")
            year = (dt.datetime.now()- dt.timedelta(days=1)).strftime("%Y")
            date = dt.datetime.date(dt.datetime.now()) - dt.timedelta(days=1)
            bot.send_message(message.chat.id, 'Выберите наименования отчета для загрузки',
                             reply_markup=markup_report_name)
            bot.register_next_step_handler(message, final_get_report)
        elif message.text == 'Другая дата отчета':
            bot.send_message(message.chat.id, 'Сообщите боту дату отчета в формате ГГГГ-ММ-ДД\n'
                                                  'пример: 2022-01-15')
            bot.register_next_step_handler(message, get_date_report)
        if message.text == 'Сброс':
            bot.send_message(message.chat.id, 'Подтвердите сброс повторным нажатием на соотвествущий пункт')
            bot.register_next_step_handler(message, cancel)


    def get_date_report(message):
        markup_report_name = types.ReplyKeyboardMarkup(resize_keyboard=True)
        button_bel = types.KeyboardButton('Белотелов')
        button_orl = types.KeyboardButton('Орлова')
        button_kul = types.KeyboardButton('Кулик')
        markup_report_name.add(button_bel, button_kul, button_orl)
        markup_report_name.add(button_cancel, button_start)
        text = message.text
        if len(text) != 0:
            global date
            date = text
            global day
            day = date[-2:]
            global month
            month = date[5:7]
            global year
            month = date[0:4]
            bot.reply_to(message,f'Введенная дата отчета - {text}')
            bot.send_message(message.chat.id,'Выберите наименования отчета для загрузки',
                         reply_markup=markup_report_name)
            bot.register_next_step_handler(message, final_get_report)
        if message.text == 'Сброс':
            bot.send_message(message.chat.id, 'Подтвердите сброс повторным нажатием на соотвествущий пункт')
            bot.register_next_step_handler(message, cancel)


    def final_get_report(message):
        global name
        name = message.text
        markup_get_name = types.ReplyKeyboardMarkup(resize_keyboard=True)
        button_cancel = types.KeyboardButton('Сброс')
        button_start = types.KeyboardButton('/start')
        markup_get_name.add(button_cancel,button_start)
        bot.reply_to(message,
                     f'Выбрано наименование отчета - {name}. Дата отчета - {date}.\n'
                     f'Теперь загрузите файл отчета',
                     reply_markup=markup_get_name)
        if message.text == 'Сброс':
            bot.send_message(message.chat.id, 'Подтвердите сброс повторным нажатием на соотвествущий пункт')
            bot.register_next_step_handler(message, cancel)
        return bot.register_next_step_handler(message, handle_file)

    def table_date(message):
        global date
        text = message.text
        if len(text) != 0:
            date = text
            global day
            day = date[-2:]
            global month
            month = date[5:7]
            global year
            month = date[0:4]
            markup_start = types.ReplyKeyboardMarkup(resize_keyboard=True)
            button_get_reports = types.KeyboardButton('Загрузить отчеты в бота')
            button_update_table = types.KeyboardButton(f'Обновить данные в Гугл Таблицах ({date})')
            button_date = types.KeyboardButton(f'Выбрать дату для работы с таблицей')
            button_cancel = types.KeyboardButton('Сброс')
            button_start = types.KeyboardButton('/start')
            markup_start.add(button_get_reports)
            markup_start.add(button_update_table)
            markup_start.add(button_date)
            markup_start.add(button_cancel, button_start)
            bot.reply_to(message, f'Введенная дата - {text}', reply_markup=markup_start)
        if message.text == 'Сброс':
            bot.send_message(message.chat.id, 'Подтвердите сброс повторным нажатием на соотвествущий пункт')
            bot.register_next_step_handler(message, cancel)


    def cancel(message):
        if message.text == 'Сброс':
            global day
            global name
            global month
            global year
            global date
            day, date, name, month, year = '', 'Дата не выбрана', '','',''
            markup_start = types.ReplyKeyboardMarkup(resize_keyboard=True)
            button_get_reports = types.KeyboardButton('Загрузить отчеты в бота')
            button_update_table = types.KeyboardButton(f'Обновить данные в Гугл Таблицах ({date})')
            button_date = types.KeyboardButton(f'Выбрать дату для работы с таблицей')
            button_cancel = types.KeyboardButton('Сброс')
            button_start = types.KeyboardButton('/start')
            markup_start.add(button_get_reports)
            markup_start.add(button_update_table)
            markup_start.add(button_date)
            markup_start.add(button_cancel, button_start)
            bot.reply_to(message,
                         'Произведен сброс.\nВыберите период загрузки отчета\n            или\nОбновление отчетов в Гугл Таблице',
                         reply_markup=markup_start)


    @bot.message_handler(content_types=['document'])
    def handle_file(message):
        try:
            file_info = bot.get_file(message.document.file_id)
            downloaded_file = bot.download_file(file_info.file_path)
            global date
            src = 'excel_docs/' + f'{name}-{date}.xlsx';
            with open(src, 'wb') as new_file:
                new_file.write(downloaded_file)
            date = 'Дата не выбрана'
            bot.reply_to(message, " Сохранено, вы можете:\n"
                                  " 1) Загрузить новый отчет\n 2) Обновить данные в Гугл таблице"
                                  "\n 3) Сделать сброс наименований и периода",reply_markup=markup_start)

        except Exception as e:
            bot.reply_to(message, e)




# launch bot
bot.polling(none_stop=True, interval=0)
