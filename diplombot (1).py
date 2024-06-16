# pip install pytelegrambotapi

# pip install beautifulsoup4

# pip install requests

# !pip install selenium
# !apt-get update
# !apt install -y chromium-chromedriver

# !pip install --upgrade selenium
# !apt-get update
# !apt install -y chromium-chromedriver

# pip install xlrd

import telebot
import random
from openpyxl import load_workbook
import time
from telebot import types
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import pandas as pd
import xlrd
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from openpyxl.utils import get_column_letter
from urllib.parse import urljoin
from urllib.request import urlretrieve
import os
from PIL import ImageGrab
from openpyxl import Workbook
import win32com.client as win32
from spire.xls import *
from spire.xls.common import *
import xlwings as xw

bot = telebot.TeleBot('6551232828:AAGHHErdbGQEyFW7fZKdK_9uatLx7sNDj8U')
api_weather = "4dd800010f75e613d40eef88ca54cd40"
city_name = 'London'

class UserState:
    def __init__(self):
        self.institute = None
        self.group = None
        self.course = None

user_states = {}
chat_states = {}

@bot.message_handler(commands=['start'])
def start_message(message):
    bot.send_message(message.chat.id, 'Привет!\n\n\u2005\u2005*Я СтудентоСпаситель - цифровой помощник*😎\u2005\u2005\n\nЯ помогаю студентам и абитуриентам БГИТУ быстро находить нужную информацию. Скорее выбирай раздел в Главном меню – я знаю много интересного про университет!', parse_mode='Markdown')

@bot.message_handler(commands=['menu'])
def menu_message(message):
    markup = types.InlineKeyboardMarkup(row_width=1)
    btn1 = types.InlineKeyboardButton(text="Абитуриент", callback_data ='first')
    btn2 = types.InlineKeyboardButton(text="Студент", callback_data ='second')
    btn3 = types.InlineKeyboardButton(text="Школьник", callback_data ='thirst')
    markup.add(btn1, btn2, btn3)
    bot.send_message(message.chat.id, 'И так, ты у нас...', reply_markup=markup)

@bot.message_handler(commands=['menu'])
def menu_message_vozvrat(message):
    markup = types.InlineKeyboardMarkup(row_width=1)
    btn1 = types.InlineKeyboardButton(text="Абитуриент", callback_data ='first')
    btn2 = types.InlineKeyboardButton(text="Студент", callback_data ='second')
    btn3 = types.InlineKeyboardButton(text="Школьник", callback_data ='thirst')
    markup.add(btn1, btn2, btn3)
    bot.send_message(message.chat.id, '*Вот и снова в меню.*\n\nНавигация здесь проста, выбирай свой путь.', reply_markup=markup, parse_mode='Markdown')
# ОБРАБОТКА АБИТУРИЕНТА И ДАЛЬНЕЙШИЕ КНОПКИ С НИМ
                                   #ШКОЛЬНИК
@bot.callback_query_handler(func=lambda call: call.data == 'thirst')
def abiturient_menu(call):
    markup = types.InlineKeyboardMarkup(row_width=2)
    btn1 = types.InlineKeyboardButton(text="Сайт университета", url='http://www.bgitu.ru/')
    btn2 = types.InlineKeyboardButton(text="Объявления", url='http://bgitu.ru/schoolboys/')
    btn3 = types.InlineKeyboardButton(text="Курсы подготовки к ЕГЭ", url='http://bgitu.ru/schoolboys/kursy-podgotovki-k-ege/')
    btn4 = types.InlineKeyboardButton(text="Код будущего", url="http://bgitu.ru/schoolboys/kod-budushchego/")
    btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='back_to_menu')
    markup.add(btn1, btn2, btn3, btn4, btn6)
    bot.send_message(call.message.chat.id, 'Чтобы найти нужную тебе информацию, выбери интересующий раздел:', reply_markup=markup)
                                   #ШКОЛЬНИК
                                   #АБИТУРИЕНТ
@bot.callback_query_handler(func=lambda call: call.data == 'first')
def abiturient_menu(call):
    markup = types.InlineKeyboardMarkup(row_width=1)
    btn1 = types.InlineKeyboardButton(text="Сайт университета", url='http://www.bgitu.ru/')
    btn2 = types.InlineKeyboardButton(text="Приемная комиссия", callback_data ='comis')
    btn3 = types.InlineKeyboardButton(text="Бакалавриат/Специалитет", callback_data ='baknspec')
    btn4 = types.InlineKeyboardButton(text="Магистратура", url="http://www.bgitu.ru/abitur/magistr/")
    btn5 = types.InlineKeyboardButton(text="Аспирантура", url="http://www.bgitu.ru/abitur/aspirant/")
    btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='back_to_menu')
    markup.add(btn1, btn2, btn3, btn4, btn5, btn6)
    bot.send_message(call.message.chat.id, 'Чтобы найти нужную тебе информацию, выбери интересующий раздел:\n\nТакже мы предлагаем тебе зарегистрироваться Личном кабинете будущего студента БГИТУ, чтобы не пропустить важную информацию о процессе поступления и зачисления, сэкономить свое время при подаче документов.', reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data == 'comis')
def comis_menu(call):
    markup = types.InlineKeyboardMarkup(row_width=1)
    btn1 = types.InlineKeyboardButton(text="Как подать документы онлайн", url='http://www.bgitu.ru/')
    btn2 = types.InlineKeyboardButton(text="Сроки подачи документов", callback_data ='srokpod')
    btn3 = types.InlineKeyboardButton(text="План приема", url="http://www.bgitu.ru/abitur/")
    btn4 = types.InlineKeyboardButton(text="Документы, необходимые для поступления", callback_data ='docmust')
    btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='naz_to_first')
    btn7 = types.InlineKeyboardButton(text="Меню📖", callback_data ='back_to_menu')
    markup.add(btn1, btn2, btn3, btn4, btn6, btn7)
    bot.send_message(call.message.chat.id, 'Часто задаваемые вопросы о поступлении', reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data == 'baknspec')
def baknspec_menu(call):
    markup = types.InlineKeyboardMarkup(row_width=1)
    btn1 = types.InlineKeyboardButton(text="Направления подготовки", url='http://www.bgitu.ru/abitur/napravleniya-podgotovki-/')
    btn2 = types.InlineKeyboardButton(text="Институты", url='http://www.bgitu.ru/universitet/instituty/')
    btn3 = types.InlineKeyboardButton(text="Специалитет", url="http://bgitu.ru/abitur/special/")
    btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='naz_to_first')
    btn7 = types.InlineKeyboardButton(text="Меню📖", callback_data ='back_to_menu')
    markup.add(btn1, btn2, btn3, btn6, btn7)
    bot.send_message(call.message.chat.id, '*Полезные ссылки:*', reply_markup=markup, parse_mode='Markdown')

@bot.callback_query_handler(func=lambda call: call.data == 'srokpod')
def srok_menu(call):
    markup = types.InlineKeyboardMarkup(row_width=1)
    btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='naz_to_comis')
    btn7 = types.InlineKeyboardButton(text="Меню📖", callback_data ='back_to_menu')
    markup.add(btn6, btn7)
    bot.send_message(call.message.chat.id, 'Срок подачи документов ограничен: начинается он для всех абитуриентов 20 июня, а вот срок окончания нужно уточнить\n\n🔴*Исключение: с вступительными испытаниями,'
    'проводимыми Университетом прием документов начинается 18 июня (Очная форма обучения)*\n\n*Сроки окончания приема документов:*\n\n🚩Очная форма обучения\n-Без вступительных испытаний/с результатами ЕГЭ - 25 июля\n'
    '-С вступительными испытаниями, проводимыми Университетом - 12 июля\n\n🚩Заочная форма обучения\n-Без вступительных испытаний/с результатами ЕГЭ/С вступительными испытаниями, проводимыми Университетом - 10 августа\n'
    '-С вступительными испытаниями, проводимыми Университетом (второе высшее) - 15 сентября\n\n🚩Магистратура\n-Срок окончания приема документов - 30 июля', reply_markup=markup, parse_mode='Markdown')

@bot.callback_query_handler(func=lambda call: call.data == 'docmust')
def srok_menu(call):
    markup = types.InlineKeyboardMarkup(row_width=1)
    btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='naz_to_comis')
    btn7 = types.InlineKeyboardButton(text="Меню📖", callback_data ='back_to_menu')
    markup.add(btn6, btn7)
    bot.send_message(call.message.chat.id, 'Поступающие на 1-й курс БГИТУ предоставляют в приёмную комиссию следующие документы:\n\n1. *Заявление* (заполняется в Университете) на имя ректора о приёме документов,'
    ' о допуске к участию в конкурсе (выбранные для обучения направления подготовки / специальности указываются в Заявлении в приоритетном для абитуриента порядке) и о согласии на зачисление при условии прохождения по конкурсу;\n'
    '2. Предъявляют *паспорт* или его копию (страницы 2 и 3 с личными данными);\n'
    '3. *Оригинал документа* установленного образца об образовании или его копию (для абитуриентов с особыми правами приёма без вступительных испытаний, на места в пределах особой (10 %) квоты) сразу необходим оригинал;\n'
    '4. *Копии дипломов победителей или призёров олимпиад школьников*, дающих особые права на приём без вступительных испытаний или на участие в конкурсе с максимально возможной оценкой в 100 баллов по предмету, соответствующему профилю олимпиады;\n'
    '*копии документов, подтверждающих возможное наличие у абитуриентов прав:*\n ➖на приём по результатам вступительных испытаний, проводимых БГИТУ самостоятельно;\n'
    ' ➖на приём на места в пределах особой (10 %) квоты;\n -преимущественное право зачисления;\n ➖копии документов, подтверждающих наличие у абитуриентов индивидуальных достижений, оцениваемых дополнительными баллами;\n'
    '5. *4 фотографии 3х4* -  при подаче оригиналов документов об образовании;\n6.*6 фотографий 3х4*  - для сдающих экзамены, проводимые БГИТУ самостоятельно;\n-*другие документы*, представление которых отвечает интересам абитуриентов.\n'
    'Заявление о приёме и, в частности, указываемые в нём выбранные для обучения направления подготовки / специальности до 26 июля включительно могут быть изменены абитуриентом.\n❗️*Копии документов у нотариуса заверять не требуется.*❗️', reply_markup=markup, parse_mode='Markdown')

                                   #АБИТУРИЕНТ

@bot.callback_query_handler(func=lambda call: call.data == 'naz_to_comis')
def naz_to_comis(call):
    comis_menu(call)

@bot.callback_query_handler(func=lambda call: call.data == 'naz_to_first')
def back_to_first(call):
    abiturient_menu(call)


@bot.callback_query_handler(func=lambda call: call.data == 'back_to_menu')
def back_to_menu(call):
    menu_message_vozvrat(call.message)

# ОБРАБОТКА АБИТУРИЕНТА И ДАЛЬНЕЙШИЕ КНОПКИ С НИМ
                                   #ОБРАБОТКА СТУДЕНТА И ДАЛЬНЕЙШИЕ КНОПКИ С НИМ

@bot.callback_query_handler(func=lambda call: call.data == 'second')
def student_menu(call):
    markup = types.InlineKeyboardMarkup(row_width=2)
    btn1 = types.InlineKeyboardButton(text="Да", callback_data ='verot')
    btn2 = types.InlineKeyboardButton(text="Нет", callback_data ='back_to_menu')
    markup.add(btn1, btn2)
    bot.send_message(call.message.chat.id,'*Ты точно студент?*🤨', reply_markup=markup, parse_mode='Markdown')

@bot.callback_query_handler(func=lambda call: call.data == 'verot')
def perex_to_reg(call):
    markup = types.InlineKeyboardMarkup(row_width=3)
    btn1 = types.InlineKeyboardButton(text="Строительный институт", callback_data ='stroy_inst')
    btn2 = types.InlineKeyboardButton(text="Институт лесного комплекса", callback_data ='les_inst')
    btn3 = types.InlineKeyboardButton(text="Инженерно-экономический институт", callback_data ='inger_inst')
    btn4 = types.InlineKeyboardButton(text="Меню📖", callback_data ='back_to_menu')
    markup.add(btn1, btn2, btn3, btn4)
    bot.send_message(call.message.chat.id, 'Так, студент, выбери свой институт:', reply_markup=markup, parse_mode='Markdown')

@bot.callback_query_handler(func=lambda call: call.data in ['stroy_inst', 'les_inst', 'inger_inst'])
def group_menu(call):
    rus_inst = {
        'stroy_inst': 'Строительный институт',
        'les_inst': 'Институт лесного комплекса, ландшафтной архитектуры, транспорта и экологии',
        'inger_inst': 'Инженерно-экономический институт'
    }
    user_id = call.from_user.id
    user_states[user_id] = UserState()
    user_states[user_id].institute = rus_inst[call.data]
    if user_states[user_id].institute == "Строительный институт":
        markup = types.ForceReply(selective=False)
        markup = types.InlineKeyboardMarkup(row_width=3)
        btn1 = types.InlineKeyboardButton(text="ПГС", callback_data ='pgs_group')
        btn2 = types.InlineKeyboardButton(text="ГСХ", callback_data ='gsx_group')
        btn3 = types.InlineKeyboardButton(text="ТГСВ", callback_data ='tgsv_group')
        btn4 = types.InlineKeyboardButton(text="ПСК", callback_data ='psk_group')
        btn5 = types.InlineKeyboardButton(text="АД", callback_data ='ad_group')
        btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='back_to_inst')
        markup.add(btn1, btn2, btn3, btn4, btn5, btn6)
        bot.send_message(call.message.chat.id,'Выбери свою группу', reply_markup=markup, parse_mode='Markdown')
    if user_states[user_id].institute == "Институт лесного комплекса, ландшафтной архитектуры, транспорта и экологии":
        markup = types.ForceReply(selective=False)
        markup = types.InlineKeyboardMarkup(row_width=3)
        btn1 = types.InlineKeyboardButton(text="ММ", callback_data ='mm_group')
        btn2 = types.InlineKeyboardButton(text="МЛП", callback_data ='mlp_group')
        btn3 = types.InlineKeyboardButton(text="ТД", callback_data ='td_group')
        btn4 = types.InlineKeyboardButton(text="ТБ", callback_data ='tb_group')
        btn5 = types.InlineKeyboardButton(text="ЭРСП", callback_data ='ersp_group')
        btn6 = types.InlineKeyboardButton(text="ЛХ", callback_data ='lx_group')
        btn7 = types.InlineKeyboardButton(text="ЛА", callback_data ='la_group')
        btn8 = types.InlineKeyboardButton(text="САД", callback_data ='sad_group')
        btn9 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='back_to_inst')
        markup.add(btn1, btn2, btn3, btn4, btn5, btn6, btn7, btn8, btn9)
        bot.send_message(call.message.chat.id,'Выбери свою группу', reply_markup=markup, parse_mode='Markdown')
    if user_states[user_id].institute == "Инженерно-экономический институт":
        block1 = 1
        if block1 == 1:
            markup = types.ForceReply(selective=False)
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn1 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='back_to_inst')
            markup.add(btn1)
            bot.send_message(call.message.chat.id,'Ваш расдел сейчас недоступен', reply_markup=markup, parse_mode='Markdown')
        else:
            markup = types.ForceReply(selective=False)
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn1 = types.InlineKeyboardButton(text="ИВТ", callback_data ='ivt_group')
            btn2 = types.InlineKeyboardButton(text="ИСТ", callback_data ='ist_group')
            btn3 = types.InlineKeyboardButton(text="ПрИ", callback_data ='pri_group')
            btn4 = types.InlineKeyboardButton(text="Экон", callback_data ='econ_group')
            btn5 = types.InlineKeyboardButton(text="ПИ", callback_data ='pi_group')
            btn6 = types.InlineKeyboardButton(text="ГМУ", callback_data ='gmu_group')
            btn7 = types.InlineKeyboardButton(text="ЭБ", callback_data ='eb_group')
            btn8 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='back_to_inst')
            markup.add(btn1, btn2, btn3, btn4, btn5, btn6, btn7, btn8)
            bot.send_message(call.message.chat.id,'Выбери свою группу', reply_markup=markup, parse_mode='Markdown')

#####################################################ИНЖ-ЭКОН#######################################################

@bot.callback_query_handler(func=lambda call: call.data in ['ivt_group', 'ist_group', 'pri_group', 'econ_group', 'pi_group', 'gmu_group', 'eb_group'])
def process_group_selection11(call):
    rus_gr = {
        'ivt_group': 'ИВТ',
        'ist_group': 'ИСТ',
        'pri_group': 'ПРИ',
        'econ_group': 'Экон',
        'pi_group': 'ПИ',
        'gmu_group': 'ГМУ',
        'eb_group': 'ЭБ'
    }
    user_id = call.from_user.id
    user_states[user_id].group = rus_gr[call.data]

    markup = types.ForceReply(selective=False)
    markup = types.InlineKeyboardMarkup(row_width=5)
    btn1 = types.InlineKeyboardButton(text="1 курс", callback_data ='first_group2')
    btn2 = types.InlineKeyboardButton(text="2 курс", callback_data ='sec_group2')
    btn3 = types.InlineKeyboardButton(text="3 курс", callback_data ='thr_group2')
    btn4 = types.InlineKeyboardButton(text="4 курс", callback_data ='four_group2')
    btn5 = types.InlineKeyboardButton(text="Магистратура", callback_data ='magiss_group2')
    btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='group_menu')
    markup.add(btn1, btn2, btn3, btn4, btn5, btn6)
    bot.send_message(call.message.chat.id,'Выберите свой курс', reply_markup=markup, parse_mode='Markdown')

@bot.callback_query_handler(func=lambda call: call.data in ['first_group2', 'sec_group2', 'thr_group2', 'four_group2', 'magiss_group2'])
def process_group_selection12(call):
    rus_cr = {
        'first_group2': '1',
        'sec_group2': '2',
        'thr_group2': '3',
        'four_group2': '4',
        'magiss_group2': 'Магистратура'
    }
    user_id = call.from_user.id
    user_states[user_id].course = rus_cr[call.data]
    institute = user_states[user_id].institute
    group = user_states[user_id].group
    course = user_states[user_id].course
    if (course == "4") and (group == "ПГС" or group == "ГСХ" or group == "ТГСВ" or group == "АД"):
        file_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps\raspisonstr.xls"
        # Проверяем, существует ли файл
        if os.path.exists(file_path):
            # Если файл существует, удаляем его
            os.remove(file_path)
        base_url = "https://bgitu.ru/studentu/raspisanie/ochnoe-obuchenie/2024/rv-2/"
        url = "https://bgitu.ru/studentu/raspisanie/ochnoe-obuchenie/2024/rv-2/"  # второй раз указываем url, чтобы получить содержимое ссылки
        # Отправляем GET-запрос и получаем содержимое страницы
        response = requests.get(url)
        html_content = response.content
        # Создаем объект BeautifulSoup для парсинга HTML
        soup = BeautifulSoup(html_content, "html.parser")
        # Находим ссылку на странице, которая содержится в теге <a> с определенным текстом
        link_tag = soup.find("a", string="Расписание на 2 семестр ПГС ГСХ ПСК АД ТГСВ 4 курс.xls")
        # Если найден тег <a>, получаем его URL и объединяем с базовым URL
        if link_tag:
            link_href = link_tag.get("href")
            full_url = urljoin(base_url, link_href)
            # Определяем путь, по которому нужно сохранить файл
            save_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps"
            # Создаем директорию, если ее нет
            os.makedirs(save_path, exist_ok=True)
            # Скачиваем файл по объединенной ссылке и сохраняем в указанную папку
            file_path, headers = urlretrieve(full_url, os.path.join(save_path, "raspisonstr.xls"))
        else:
            markup = types.ForceReply(selective=False)
            markup = types.InlineKeyboardMarkup(row_width=5)
            btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
            btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
            btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
            markup.add(btn1, btn3, btn6)
            bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
            # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
            bot.send_message(call.message.chat.id,'Расписание сейчас не получиться посмотреть', reply_markup=markup, parse_mode='Markdown')
            bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
        save_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps"
        # Путь к файлу для проверки
        file_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps\raspisonstr.xlsx"
        # Проверяем, существует ли файл
        if os.path.exists(file_path):
            # Если файл существует, удаляем его
            os.remove(file_path)
        # Путь к скриншоту для проверки
        screenshot_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\raspison_rangestr.png"
        # Проверяем, существует ли скриншот
        if os.path.exists(screenshot_path):
        # Если скриншот существует, удаляем его
            os.remove(screenshot_path)
        time.sleep(1)
        # Конвертируем .xls файл в .xlsx
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(os.path.join(save_path, "raspisonstr.xls"))
        wb.SaveAs(os.path.join(save_path, "raspisonstr.xlsx"), FileFormat=51)
        wb.Close()
        excel.Quit()
        time.sleep(0.5)
        xlsx_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps\raspisonstr.xlsx"
        client = win32.Dispatch("Excel.Application")
        wb = client.Workbooks.Open(xlsx_path)
        ws = wb.ActiveSheet
        print(1)
        ws.Range("A2:I49").CopyPicture(Format = 2)
        img = ImageGrab.grabclipboard()
        img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
        wb.Close()
        client.Quit()
        markup = types.ForceReply(selective=False)
        markup = types.InlineKeyboardMarkup(row_width=5)
        btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
        btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
        btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
        btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
        markup.add(btn1, btn2, btn3, btn6)
        bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
        # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
        bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
    else:
        file_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps\raspisonstr.xls"
        # Проверяем, существует ли файл
        if os.path.exists(file_path):
            # Если файл существует, удаляем его
            os.remove(file_path)

        base_url = "https://bgitu.ru/studentu/raspisanie/ochnoe-obuchenie/2024/rv-2/"
        url = "https://bgitu.ru/studentu/raspisanie/ochnoe-obuchenie/2024/rv-2/"  # второй раз указываем url, чтобы получить содержимое ссылки
        # Отправляем GET-запрос и получаем содержимое страницы
        response = requests.get(url)
        html_content = response.content
        # Создаем объект BeautifulSoup для парсинга HTML
        soup = BeautifulSoup(html_content, "html.parser")
        # Находим ссылку на странице, которая содержится в теге <a> с определенным текстом
        link_tag = soup.find("a", string="Расписание на 2 семестр ИВТ ИСТ ПИ ПрИ Экон ГМУ 1-3 к. ЭБ 1-4 к..xls")
        # Если найден тег <a>, получаем его URL и объединяем с базовым URL
        if link_tag:
            link_href = link_tag.get("href")
            full_url = urljoin(base_url, link_href)
            # Определяем путь, по которому нужно сохранить файл
            save_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps"
            # Создаем директорию, если ее нет
            os.makedirs(save_path, exist_ok=True)
            # Скачиваем файл по объединенной ссылке и сохраняем в указанную папку
            file_path, headers = urlretrieve(full_url, os.path.join(save_path, "raspisonstr.xls"))
        else:
            bot.send_message(call.message.chat.id, 'Все плохо')
        save_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps"
        # Путь к файлу для проверки
        file_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps\raspisonstr.xlsx"
        # Проверяем, существует ли файл
        if os.path.exists(file_path):
            # Если файл существует, удаляем его
            os.remove(file_path)
        # Путь к скриншоту для проверки
        screenshot_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\raspison_rangestr.png"
        # Проверяем, существует ли скриншот
        if os.path.exists(screenshot_path):
        # Если скриншот существует, удаляем его
            os.remove(screenshot_path)
        time.sleep(1)
        # Конвертируем .xls файл в .xlsx
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(os.path.join(save_path, "raspisonstr.xls"))
        wb.SaveAs(os.path.join(save_path, "raspisonstr.xlsx"), FileFormat=51)
        wb.Close()
        excel.Quit()
        time.sleep(0.5)
        xlsx_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps\raspisonstr.xlsx"
        client = win32.Dispatch("Excel.Application")
        wb = client.Workbooks.Open(xlsx_path)
        ws = wb.ActiveSheet

        if (course == "1") and (group == "ИВТ"):
            print(1)
            ws.Range("A2:F43").CopyPicture(Format = 2)
            img = ImageGrab.grabclipboard()
            img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
            wb.Close()
            client.Quit()
            markup = types.ForceReply(selective=False)
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
            btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
            btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
            btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
            markup.add(btn1, btn2, btn3, btn6)
            bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
            # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
            bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
        if (course == "2") and (group == "ИВТ"):
            print(2)
            ws.Range("G2:H42").CopyPicture(Format = 2)
            img = ImageGrab.grabclipboard()
            img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
            wb.Close()
            client.Quit()
            markup = types.ForceReply(selective=False)
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
            btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
            btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
            btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
            markup.add(btn1, btn2, btn3, btn6)
            bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
            # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
            bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
        if (course == "3") and (group == "ИВТ"):
            print(3)
            ws.Range("I2:J45").CopyPicture(Format = 2)
            img = ImageGrab.grabclipboard()
            img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
            wb.Close()
            client.Quit()
            markup = types.ForceReply(selective=False)
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
            btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
            btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
            btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
            markup.add(btn1, btn2, btn3, btn6)
            bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
            # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
            bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
        if (course == "1" or course == "2" or course == "3" ) and (group == "ИСТ"):
            print(1)
            ws.Range("K2:M45").CopyPicture(Format = 2)
            img = ImageGrab.grabclipboard()
            img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
            wb.Close()
            client.Quit()
            markup = types.ForceReply(selective=False)
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
            btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
            btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
            btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
            markup.add(btn1, btn2, btn3, btn6)
            bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
            # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
            bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
        if (course == "1" or course == "2" or course == "3" ) and (group == "ИСТ"):
            print(1)
            ws.Range("N2:P43").CopyPicture(Format = 2)
            img = ImageGrab.grabclipboard()
            img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
            wb.Close()
            client.Quit()
            markup = types.ForceReply(selective=False)
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
            btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
            btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
            btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
            markup.add(btn1, btn2, btn3, btn6)
            bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
            # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
            bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
        if (course == "1" or course == "2" ) and (group == "ПИ"):
            print(1)
            ws.Range("Q2:R42").CopyPicture(Format = 2)
            img = ImageGrab.grabclipboard()
            img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
            wb.Close()
            client.Quit()
            markup = types.ForceReply(selective=False)
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
            btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
            btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
            btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
            markup.add(btn1, btn2, btn3, btn6)
            bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
            # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
            bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
        if (course == "3" ) and (group == "ПРИ"):
            print(1)
            wb.Close()
            client.Quit()
            markup = types.ForceReply(selective=False)
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
            btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='prepod')
            btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
            btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
            markup.add(btn1, btn2, btn3, btn6)
            bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
            # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
            bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
        if (course == "1" or course == "2" or course == "3" ) and (group == "ГМУ"):
            print(1)
            ws.Range("S2:U44").CopyPicture(Format = 2)
            img = ImageGrab.grabclipboard()
            img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
            wb.Close()
            client.Quit()
            markup = types.ForceReply(selective=False)
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
            btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
            btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
            btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
            markup.add(btn1, btn2, btn3, btn6)
            bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
            # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
            bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')

##################################################ЛЕСНИКИ##############################################################

@bot.callback_query_handler(func=lambda call: call.data in ['mm_group', 'mlp_group', 'td_group', 'tb_group', 'ersp_group', 'lx_group', 'la_group', 'sad_group'])
def process_group_selection3(call):
    rus_gr = {
        'mm_group': 'ММ',
        'mlp_group': 'МЛП',
        'td_group': 'ТД',
        'tb_group': 'ТБ',
        'ersp_group': 'ЭРСП',
        'lx_group': 'ЛХ',
        'la_group': 'ЛА',
        'sad_group': 'САД'
    }
    user_id = call.from_user.id
    user_states[user_id].group = rus_gr[call.data]
    markup = types.ForceReply(selective=False)
    markup = types.InlineKeyboardMarkup(row_width=5)
    btn1 = types.InlineKeyboardButton(text="1 курс", callback_data ='first_group1')
    btn2 = types.InlineKeyboardButton(text="2 курс", callback_data ='sec_group1')
    btn3 = types.InlineKeyboardButton(text="3 курс", callback_data ='thr_group1')
    btn4 = types.InlineKeyboardButton(text="4 курс", callback_data ='four_group1')
    btn5 = types.InlineKeyboardButton(text="Магистратура", callback_data ='magiss_group1')
    btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='group_menu1')
    markup.add(btn1, btn2, btn3, btn4, btn5, btn6)
    bot.send_message(call.message.chat.id,'Выберите свой курс', reply_markup=markup, parse_mode='Markdown')

@bot.callback_query_handler(func=lambda call: call.data in ['first_group1', 'sec_group1', 'thr_group1', 'four_group1', 'magiss_group1'])
def process_group_selection4(call):
    chat_states[call.message.chat.id] = 'group_selection'
    rus_cr = {
        'first_group1': '1',
        'sec_group1': '2',
        'thr_group1': '3',
        'four_group1': '4',
        'magiss_group1': 'Магистратура'
    }
    user_id = call.from_user.id
    user_states[user_id].course = rus_cr[call.data]
    institute = user_states[user_id].institute
    group = user_states[user_id].group
    course = user_states[user_id].course
    if (course == "4") and ((group == "ЛХ") or (group == "ЛА") or (group == "МЛП") or (group == "АД")):
        block = 1
        if block == 1:
            markup = types.ForceReply(selective=False)
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn1 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='back_to_inst')
            markup.add(btn1)
            bot.send_message(call.message.chat.id,'Ваш расдел сейчас недоступен', reply_markup=markup, parse_mode='Markdown')
        else:
            file_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps\raspisonstr.xls"
            # Проверяем, существует ли файл
            if os.path.exists(file_path):
                # Если файл существует, удаляем его
                os.remove(file_path)

            base_url = "https://bgitu.ru/studentu/raspisanie/ochnoe-obuchenie/2024/rv-2/"
            url = "https://bgitu.ru/studentu/raspisanie/ochnoe-obuchenie/2024/rv-2/"  # второй раз указываем url, чтобы получить содержимое ссылки
            # Отправляем GET-запрос и получаем содержимое страницы
            response = requests.get(url)
            html_content = response.content
            # Создаем объект BeautifulSoup для парсинга HTML
            soup = BeautifulSoup(html_content, "html.parser")
            # Находим ссылку на странице, которая содержится в теге <a> с определенным текстом
            link_tag = soup.find("a", string="Расписание на 2 семестр ПГС ГСХ ПСК АД ТГСВ 4 курс.xls")
            # Если найден тег <a>, получаем его URL и объединяем с базовым URL
            if link_tag:
                link_href = link_tag.get("href")
                full_url = urljoin(base_url, link_href)
                # Определяем путь, по которому нужно сохранить файл
                save_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps"
                # Создаем директорию, если ее нет
                os.makedirs(save_path, exist_ok=True)
                # Скачиваем файл по объединенной ссылке и сохраняем в указанную папку
                file_path, headers = urlretrieve(full_url, os.path.join(save_path, "raspisonstr.xls"))
            else:
                bot.send_message(call.message.chat.id, 'Все плохо')
            save_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps"
            # Путь к файлу для проверки
            file_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps\raspisonstr.xlsx"
            # Проверяем, существует ли файл
            if os.path.exists(file_path):
                # Если файл существует, удаляем его
                os.remove(file_path)
            # Путь к скриншоту для проверки
            screenshot_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\raspison_rangestr.png"
            # Проверяем, существует ли скриншот
            if os.path.exists(screenshot_path):
            # Если скриншот существует, удаляем его
                os.remove(screenshot_path)
            time.sleep(1)
            # Конвертируем .xls файл в .xlsx
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(os.path.join(save_path, "raspisonstr.xls"))
            wb.SaveAs(os.path.join(save_path, "raspisonstr.xlsx"), FileFormat=51)
            wb.Close()
            excel.Quit()
            time.sleep(0.5)
            xlsx_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps\raspisonstr.xlsx"
            client = win32.Dispatch("Excel.Application")
            wb = client.Workbooks.Open(xlsx_path)
            ws = wb.ActiveSheet
            print(1)
            ws.Range("A2:I49").CopyPicture(Format = 2)
            img = ImageGrab.grabclipboard()
            img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
            wb.Close()
            client.Quit()
            markup = types.ForceReply(selective=False)
            markup = types.InlineKeyboardMarkup(row_width=5)
            btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
            btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
            btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
            btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
            markup.add(btn1, btn2, btn3, btn6)
            bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
            # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
            bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
    else:
        

        if group == "ЛХ" or group == "ЛА" or group == "САД":
            file_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps\raspisonstr.xls"
            # Проверяем, существует ли файл
            if os.path.exists(file_path):
                # Если файл существует, удаляем его
                os.remove(file_path)

            base_url = "https://bgitu.ru/studentu/raspisanie/ochnoe-obuchenie/2024/rv-2/"
            url = "https://bgitu.ru/studentu/raspisanie/ochnoe-obuchenie/2024/rv-2/"  # второй раз указываем url, чтобы получить содержимое ссылки
            # Отправляем GET-запрос и получаем содержимое страницы
            response = requests.get(url)
            html_content = response.content
            # Создаем объект BeautifulSoup для парсинга HTML
            soup = BeautifulSoup(html_content, "html.parser")
            # Находим ссылку на странице, которая содержится в теге <a> с определенным текстом
            link_tag = soup.find("a", string="Расписание на 2 семестр ЛХ ЛА САД 1-3 курсы.xls")
            # Если найден тег <a>, получаем его URL и объединяем с базовым URL
            if link_tag:
                link_href = link_tag.get("href")
                full_url = urljoin(base_url, link_href)
                # Определяем путь, по которому нужно сохранить файл
                save_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps"
                # Создаем директорию, если ее нет
                os.makedirs(save_path, exist_ok=True)
                # Скачиваем файл по объединенной ссылке и сохраняем в указанную папку
                file_path, headers = urlretrieve(full_url, os.path.join(save_path, "raspisonstr.xls"))
            else:
                bot.send_message(call.message.chat.id, 'Все плохо')
            save_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps"
            # Путь к файлу для проверки
            file_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps\raspisonstr.xlsx"
            # Проверяем, существует ли файл
            if os.path.exists(file_path):
                # Если файл существует, удаляем его
                os.remove(file_path)
            # Путь к скриншоту для проверки
            screenshot_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\raspison_rangestr.png"
            # Проверяем, существует ли скриншот
            if os.path.exists(screenshot_path):
            # Если скриншот существует, удаляем его
                os.remove(screenshot_path)
            time.sleep(1)
            # Конвертируем .xls файл в .xlsx
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(os.path.join(save_path, "raspisonstr.xls"))
            wb.SaveAs(os.path.join(save_path, "raspisonstr.xlsx"), FileFormat=51)
            wb.Close()
            excel.Quit()
            time.sleep(0.5)
            xlsx_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps\raspisonstr.xlsx"
            client = win32.Dispatch("Excel.Application")
            wb = client.Workbooks.Open(xlsx_path)
            ws = wb.ActiveSheet

            if (course == "1") and (group == "ЛХ" or group == "САД" or group == "ЛА"):
                print(1)
                ws.Range("A2:F43").CopyPicture(Format = 2)
                img = ImageGrab.grabclipboard()
                img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
                wb.Close()
                client.Quit()
                markup = types.ForceReply(selective=False)
                markup = types.InlineKeyboardMarkup(row_width=2)
                btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
                btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
                btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
                btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
                markup.add(btn1, btn2, btn3, btn6)
                bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
                # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
                bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
            if (course == "2") and (group == "ЛХ" or group == "САД" or group == "ЛА"):
                print(2)
                ws.Range("G2:I43").CopyPicture(Format = 2)
                img = ImageGrab.grabclipboard()
                img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
                wb.Close()
                client.Quit()
                markup = types.ForceReply(selective=False)
                markup = types.InlineKeyboardMarkup(row_width=2)
                btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
                btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
                btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
                btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
                markup.add(btn1, btn2, btn3, btn6)
                bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
                # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
                bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
            if (course == "3") and (group == "ЛХ" or group == "САД" or group == "ЛА"):
                print(3)
                ws.Range("J2:M43").CopyPicture(Format = 2)
                img = ImageGrab.grabclipboard()
                img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
                wb.Close()
                client.Quit()
                markup = types.ForceReply(selective=False)
                markup = types.InlineKeyboardMarkup(row_width=2)
                btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
                btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
                btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
                btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
                markup.add(btn1, btn2, btn3, btn6)
                bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
                # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
                bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
            if (course == "4") and (group == "САД"):
                block = 1
                if block == 1:
                    wb.Close()
                    client.Quit()
                    markup = types.ForceReply(selective=False)
                    markup = types.InlineKeyboardMarkup(row_width=2)
                    btn1 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='back_to_inst')
                    markup.add(btn1)
                    bot.send_message(call.message.chat.id,'Ваш расдел сейчас недоступен', reply_markup=markup, parse_mode='Markdown')

        else:
            file_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps\raspisonstr.xls"
            # Проверяем, существует ли файл
            if os.path.exists(file_path):
                # Если файл существует, удаляем его
                os.remove(file_path)

            base_url = "https://bgitu.ru/studentu/raspisanie/ochnoe-obuchenie/2024/rv-2/"
            url = "https://bgitu.ru/studentu/raspisanie/ochnoe-obuchenie/2024/rv-2/"  # второй раз указываем url, чтобы получить содержимое ссылки
            # Отправляем GET-запрос и получаем содержимое страницы
            response = requests.get(url)
            html_content = response.content
            # Создаем объект BeautifulSoup для парсинга HTML
            soup = BeautifulSoup(html_content, "html.parser")
            # Находим ссылку на странице, которая содержится в теге <a> с определенным текстом
            link_tag = soup.find("a", string="Расписание на 2 семестр ММ МЛП ТД ТБ ЭРСП 1-3 курсы.xls")
            # Если найден тег <a>, получаем его URL и объединяем с базовым URL
            if link_tag:
                link_href = link_tag.get("href")
                full_url = urljoin(base_url, link_href)
                # Определяем путь, по которому нужно сохранить файл
                save_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps"
                # Создаем директорию, если ее нет
                os.makedirs(save_path, exist_ok=True)
                # Скачиваем файл по объединенной ссылке и сохраняем в указанную папку
                file_path, headers = urlretrieve(full_url, os.path.join(save_path, "raspisonstr.xls"))
            else:
                bot.send_message(call.message.chat.id, 'Все плохо')
            save_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps"
            # Путь к файлу для проверки
            file_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps\raspisonstr.xlsx"
            # Проверяем, существует ли файл
            if os.path.exists(file_path):
                # Если файл существует, удаляем его
                os.remove(file_path)
            # Путь к скриншоту для проверки
            screenshot_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\raspison_rangestr.png"
            # Проверяем, существует ли скриншот
            if os.path.exists(screenshot_path):
            # Если скриншот существует, удаляем его
                os.remove(screenshot_path)
            time.sleep(1)
            # Конвертируем .xls файл в .xlsx
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(os.path.join(save_path, "raspisonstr.xls"))
            wb.SaveAs(os.path.join(save_path, "raspisonstr.xlsx"), FileFormat=51)
            wb.Close()
            excel.Quit()
            time.sleep(0.5)
            xlsx_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps\raspisonstr.xlsx"
            client = win32.Dispatch("Excel.Application")
            wb = client.Workbooks.Open(xlsx_path)
            ws = wb.ActiveSheet

            if ((course == "2" and group == "ММ") or (group == "ТБ" and (course =="1" or course == "2"))):
                ws.Range("A2:F45").CopyPicture(Format = 2)
                img = ImageGrab.grabclipboard()
                img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
                wb.Close()
                client.Quit()
                markup = types.ForceReply(selective=False)
                markup = types.InlineKeyboardMarkup(row_width=2)
                btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
                btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
                btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
                btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
                markup.add(btn1, btn2, btn3, btn6)
                bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
                # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
                bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
            if (course == "2" or course == "1" or course == "3") and group == "МЛП":
                ws.Range("G2:I45").CopyPicture(Format = 2)
                img = ImageGrab.grabclipboard()
                img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
                wb.Close()
                client.Quit()
                markup = types.ForceReply(selective=False)
                markup = types.InlineKeyboardMarkup(row_width=2)
                btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
                btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
                btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
                btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
                markup.add(btn1, btn2, btn3, btn6)
                bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
                # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
                bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
            if (course == "2" or course == "1" or course == "3") and group == "ТД":
                ws.Range("J2:L45").CopyPicture(Format = 2)
                img = ImageGrab.grabclipboard()
                img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
                wb.Close()
                client.Quit()
                markup = types.ForceReply(selective=False)
                markup = types.InlineKeyboardMarkup(row_width=2)
                btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
                btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
                btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
                btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
                markup.add(btn1, btn2, btn3, btn6)
                bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
                # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
                bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
            if group == "ЭРСП":
                print(1)
                ws.Range("M2:M45").CopyPicture(Format = 2)
                img = ImageGrab.grabclipboard()
                img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
                wb.Close()
                client.Quit()
                markup = types.ForceReply(selective=False)
                markup = types.InlineKeyboardMarkup(row_width=2)
                btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
                btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
                btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
                btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
                markup.add(btn1, btn2, btn3, btn6)
                bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
                # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
                bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
            if (course == "1" or course == "2" or course == "3" ) and (group == "ТГСВ"):
                ws.Range("N2:P43").CopyPicture(Format = 2)
                img = ImageGrab.grabclipboard()
                img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
                wb.Close()
                client.Quit()
                markup = types.ForceReply(selective=False)
                markup = types.InlineKeyboardMarkup(row_width=2)
                btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
                btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
                btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
                btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
                markup.add(btn1, btn2, btn3, btn6)
                bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
                # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
                bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
            if (course == "1" or course == "2" ) and (group == "ПСК"):
                ws.Range("Q2:R42").CopyPicture(Format = 2)
                img = ImageGrab.grabclipboard()
                img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
                wb.Close()
                client.Quit()
                markup = types.ForceReply(selective=False)
                markup = types.InlineKeyboardMarkup(row_width=2)
                btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
                btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
                btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
                btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
                markup.add(btn1, btn2, btn3, btn6)
                bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
                # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
                bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
            if (course == "3" ) and (group == "ПСК"):
                wb.Close()
                client.Quit()
                markup = types.ForceReply(selective=False)
                markup = types.InlineKeyboardMarkup(row_width=2)
                btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
                btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='prepod')
                btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
                btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
                markup.add(btn1, btn2, btn3, btn6)
                bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
                # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
                bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
            if (course == "1" or course == "2" or course == "3" ) and (group == "АД"):
                ws.Range("S2:U44").CopyPicture(Format = 2)
                img = ImageGrab.grabclipboard()
                img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
                wb.Close()
                client.Quit()
                markup = types.ForceReply(selective=False)
                markup = types.InlineKeyboardMarkup(row_width=2)
                btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
                btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
                btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
                btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
                markup.add(btn1, btn2, btn3, btn6)
                bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
                # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
                bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')

###################################СТРОИТЕЛИ##############################################################################
@bot.callback_query_handler(func=lambda call: call.data in ['pgs_group', 'gsx_group', 'tgsv_group', 'psk_group', 'ad_group'])
def process_group_selection1(call):
    rus_gr = {
        'pgs_group': 'ПГС',
        'gsx_group': 'ГСХ',
        'tgsv_group': 'ТГСВ',
        'psk_group': 'ПСК',
        'ad_group': 'АД'
    }
    user_id = call.from_user.id
    user_states[user_id].group = rus_gr[call.data]
    markup = types.ForceReply(selective=False)
    markup = types.InlineKeyboardMarkup(row_width=5)
    btn1 = types.InlineKeyboardButton(text="1 курс", callback_data ='first_group')
    btn2 = types.InlineKeyboardButton(text="2 курс", callback_data ='sec_group')
    btn3 = types.InlineKeyboardButton(text="3 курс", callback_data ='thr_group')
    btn4 = types.InlineKeyboardButton(text="4 курс", callback_data ='four_group')
    btn5 = types.InlineKeyboardButton(text="Магистратура", callback_data ='magiss_group')
    btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='group_menu')
    markup.add(btn1, btn2, btn3, btn4, btn5, btn6)
    bot.send_message(call.message.chat.id,'Выберите свой курс', reply_markup=markup, parse_mode='Markdown')

@bot.callback_query_handler(func=lambda call: call.data in ['first_group', 'sec_group', 'thr_group', 'four_group', 'magiss_group'])
def process_group_selection2(call):
    chat_states[call.message.chat.id] = 'group_selection'
    rus_cr = {
        'first_group': '1',
        'sec_group': '2',
        'thr_group': '3',
        'four_group': '4',
        'magiss_group': 'Магистратура'
    }
    user_id = call.from_user.id
    user_states[user_id].course = rus_cr[call.data]
    institute = user_states[user_id].institute
    group = user_states[user_id].group
    course = user_states[user_id].course
    if (course == "4") and (group == "ПГС" or group == "ГСХ" or group == "ТГСВ" or group == "АД"):
        file_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps\raspisonstr.xls"
        # Проверяем, существует ли файл
        if os.path.exists(file_path):
            # Если файл существует, удаляем его
            os.remove(file_path)
        base_url = "https://bgitu.ru/studentu/raspisanie/ochnoe-obuchenie/2024/rv-2/"
        url = "https://bgitu.ru/studentu/raspisanie/ochnoe-obuchenie/2024/rv-2/"  # второй раз указываем url, чтобы получить содержимое ссылки
        # Отправляем GET-запрос и получаем содержимое страницы
        response = requests.get(url)
        html_content = response.content
        # Создаем объект BeautifulSoup для парсинга HTML
        soup = BeautifulSoup(html_content, "html.parser")
        # Находим ссылку на странице, которая содержится в теге <a> с определенным текстом
        link_tag = soup.find("a", string="Расписание на 2 семестр ПГС ГСХ ПСК АД ТГСВ 4 курс.xls")
        # Если найден тег <a>, получаем его URL и объединяем с базовым URL
        if link_tag:
            link_href = link_tag.get("href")
            full_url = urljoin(base_url, link_href)
            # Определяем путь, по которому нужно сохранить файл
            save_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps"
            # Создаем директорию, если ее нет
            os.makedirs(save_path, exist_ok=True)
            # Скачиваем файл по объединенной ссылке и сохраняем в указанную папку
            file_path, headers = urlretrieve(full_url, os.path.join(save_path, "raspisonstr.xls"))
        else:
            bot.send_message(call.message.chat.id, 'Все плохо')
        save_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps"
        # Путь к файлу для проверки
        file_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps\raspisonstr.xlsx"
        # Проверяем, существует ли файл
        if os.path.exists(file_path):
            # Если файл существует, удаляем его
            os.remove(file_path)
        # Путь к скриншоту для проверки
        screenshot_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\raspison_rangestr.png"
        # Проверяем, существует ли скриншот
        if os.path.exists(screenshot_path):
        # Если скриншот существует, удаляем его
            os.remove(screenshot_path)
        time.sleep(1)
        # Конвертируем .xls файл в .xlsx
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(os.path.join(save_path, "raspisonstr.xls"))
        wb.SaveAs(os.path.join(save_path, "raspisonstr.xlsx"), FileFormat=51)
        wb.Close()
        excel.Quit()
        time.sleep(0.5)
        xlsx_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps\raspisonstr.xlsx"
        client = win32.Dispatch("Excel.Application")
        wb = client.Workbooks.Open(xlsx_path)
        ws = wb.ActiveSheet
        print(1)
        ws.Range("A2:I49").CopyPicture(Format = 2)
        img = ImageGrab.grabclipboard()
        img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
        wb.Close()
        client.Quit()
        markup = types.ForceReply(selective=False)
        markup = types.InlineKeyboardMarkup(row_width=5)
        btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
        btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
        btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
        btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
        markup.add(btn1, btn2, btn3, btn6)
        bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
        # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
        bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
    else:
        file_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps\raspisonstr.xls"
        # Проверяем, существует ли файл
        if os.path.exists(file_path):
            # Если файл существует, удаляем его
            os.remove(file_path)

        base_url = "https://bgitu.ru/studentu/raspisanie/ochnoe-obuchenie/2024/rv-2/"
        url = "https://bgitu.ru/studentu/raspisanie/ochnoe-obuchenie/2024/rv-2/"  # второй раз указываем url, чтобы получить содержимое ссылки
        # Отправляем GET-запрос и получаем содержимое страницы
        response = requests.get(url)
        html_content = response.content
        # Создаем объект BeautifulSoup для парсинга HTML
        soup = BeautifulSoup(html_content, "html.parser")
        # Находим ссылку на странице, которая содержится в теге <a> с определенным текстом
        link_tag = soup.find("a", string="Расписание на 2 семестр ПГС ГСХ ПСК АД ТГСВ 1-3 курсы.xls")
        # Если найден тег <a>, получаем его URL и объединяем с базовым URL
        if link_tag:
            link_href = link_tag.get("href")
            full_url = urljoin(base_url, link_href)
            # Определяем путь, по которому нужно сохранить файл
            save_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps"
            # Создаем директорию, если ее нет
            os.makedirs(save_path, exist_ok=True)
            # Скачиваем файл по объединенной ссылке и сохраняем в указанную папку
            file_path, headers = urlretrieve(full_url, os.path.join(save_path, "raspisonstr.xls"))
        else:
            bot.send_message(call.message.chat.id, 'Все плохо')
        save_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps"
        # Путь к файлу для проверки
        file_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps\raspisonstr.xlsx"
        # Проверяем, существует ли файл
        if os.path.exists(file_path):
            # Если файл существует, удаляем его
            os.remove(file_path)
        # Путь к скриншоту для проверки
        screenshot_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\raspison_rangestr.png"
        # Проверяем, существует ли скриншот
        if os.path.exists(screenshot_path):
        # Если скриншот существует, удаляем его
            os.remove(screenshot_path)
        time.sleep(1)
        # Конвертируем .xls файл в .xlsx
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(os.path.join(save_path, "raspisonstr.xls"))
        wb.SaveAs(os.path.join(save_path, "raspisonstr.xlsx"), FileFormat=51)
        wb.Close()
        excel.Quit()
        time.sleep(0.5)
        xlsx_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps\raspisonstr.xlsx"
        client = win32.Dispatch("Excel.Application")
        wb = client.Workbooks.Open(xlsx_path)
        ws = wb.ActiveSheet
        if (course == "1") and (group == "ПГС"):
            ws.Range("A2:F43").CopyPicture(Format = 2)
            img = ImageGrab.grabclipboard()
            img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
            wb.Close()
            client.Quit()
            markup = types.ForceReply(selective=False)
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
            btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
            btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
            btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
            markup.add(btn1, btn2, btn3, btn6)
            bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
            # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
            bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
        if (course == "2") and (group == "ПГС"):
            ws.Range("G2:H42").CopyPicture(Format = 2)
            img = ImageGrab.grabclipboard()
            img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
            wb.Close()
            client.Quit()
            markup = types.ForceReply(selective=False)
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
            btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
            btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
            btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
            markup.add(btn1, btn2, btn3, btn6)
            bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
            # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
            bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
        if (course == "3") and (group == "ПГС"):
            ws.Range("I2:J45").CopyPicture(Format = 2)
            img = ImageGrab.grabclipboard()
            img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
            wb.Close()
            client.Quit()
            markup = types.ForceReply(selective=False)
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
            btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
            btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
            btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
            markup.add(btn1, btn2, btn3, btn6)
            bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
            # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
            bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
        if (course == "1" or course == "2" or course == "3" ) and (group == "ГСХ"):
            ws.Range("K2:M45").CopyPicture(Format = 2)
            img = ImageGrab.grabclipboard()
            img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
            wb.Close()
            client.Quit()
            markup = types.ForceReply(selective=False)
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
            btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
            btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
            btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
            markup.add(btn1, btn2, btn3, btn6)
            bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
            # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
            bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
        if (course == "1" or course == "2" or course == "3" ) and (group == "ТГСВ"):
            ws.Range("N2:P43").CopyPicture(Format = 2)
            img = ImageGrab.grabclipboard()
            img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
            wb.Close()
            client.Quit()
            markup = types.ForceReply(selective=False)
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
            btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
            btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
            btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
            markup.add(btn1, btn2, btn3, btn6)
            bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
            # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
            bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
        if (course == "1" or course == "2" ) and (group == "ПСК"):
            ws.Range("Q2:R42").CopyPicture(Format = 2)
            img = ImageGrab.grabclipboard()
            img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
            wb.Close()
            client.Quit()
            markup = types.ForceReply(selective=False)
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
            btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
            btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
            btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
            markup.add(btn1, btn2, btn3, btn6)
            bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
            # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
            bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
        if (course == "3" ) and (group == "ПСК"):
            wb.Close()
            client.Quit()
            markup = types.ForceReply(selective=False)
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
            btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='prepod')
            btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
            btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
            markup.add(btn1, btn2, btn3, btn6)
            bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
            # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
            bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')
        if (course == "1" or course == "2" or course == "3" ) and (group == "АД"):
            ws.Range("S2:U44").CopyPicture(Format = 2)
            img = ImageGrab.grabclipboard()
            img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrffstr.jpg", "jpeg")
            wb.Close()
            client.Quit()
            markup = types.ForceReply(selective=False)
            markup = types.InlineKeyboardMarkup(row_width=2)
            btn1 = types.InlineKeyboardButton(text="Список преподавателей", callback_data ='back_to_inst')
            btn2 = types.InlineKeyboardButton(text="Расписание", callback_data ='raspis13st')
            btn3 = types.InlineKeyboardButton(text="График приема задолжностей", callback_data ='back_to_inst')
            btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
            markup.add(btn1, btn2, btn3, btn6)
            bot.send_message(call.message.chat.id, f'И так, вы: {institute} \nГруппа: {group} \nКурс: {course}', parse_mode='Markdown')
            # bot.send_message(call.message.chat.id,'*Институт: ' + institute + '*\n*Группа: ' + group + '*\nКурс: ' + course + '*', parse_mode='Markdown')
            bot.send_message(call.message.chat.id,'Выберите интересующий вас раздел', reply_markup=markup, parse_mode='Markdown')


@bot.callback_query_handler(func=lambda call: call.data == 'raspis13stF')
def sunmon_menu(call):

    markup = types.InlineKeyboardMarkup(row_width=5)
    btn1 = types.InlineKeyboardButton(text="Пн", callback_data='mondayst13')
    btn2 = types.InlineKeyboardButton(text="Вт", callback_data='tuesdayst13')
    btn3 = types.InlineKeyboardButton(text="Ср", callback_data="wednesdayst13")
    btn4 = types.InlineKeyboardButton(text="Чт", callback_data="thursdayst13")
    btn5 = types.InlineKeyboardButton(text="Пт", callback_data="fridayst13")
    btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='naz_to_first')
    btn7 = types.InlineKeyboardButton(text="Меню📖", callback_data ='back_to_menu')
    markup.add(btn1, btn2, btn3, btn4, btn5, btn6, btn7)
    bot.send_message(call.message.chat.id, 'Выбирай день недели\nУчти!Порядок такой: ЛХ, САД, ЛА', reply_markup=markup, parse_mode='Markdown')

@bot.callback_query_handler(func=lambda call: call.data == 'raspis13st')
def mond13st(call):
    markup = types.InlineKeyboardMarkup(row_width=1)
    btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='bkgru')
    markup.add(btn6)
    # Отправляем скриншот пользователю
    bot.send_photo(call.message.chat.id, open('E:\\otkat\\NvidiaOTK\\Grand Theft Auto  San Andreas\\pybot\\scr\\scrffstr.jpg', 'rb'))
    bot.send_message(call.message.chat.id, 'Пожалуйста)', reply_markup=markup, parse_mode='Markdown')

@bot.callback_query_handler(func=lambda call: call.data == 'bkgru')
def bkgru(call):
    perex_to_reg(call)

@bot.callback_query_handler(func=lambda call: call.data == 'naz_to_comis')
def naz_to_comis(call):
    comis_menu(call)

@bot.callback_query_handler(func=lambda call: call.data == 'back_to_inst')
def back_to_inst(call):
    perex_to_reg(call)

                                   #ОБРАБОТКА СТУДЕНТА И ДАЛЬНЕЙШИЕ КНОПКИ С НИМ

@bot.callback_query_handler(func=lambda call: call.data == 'prepod')
def back_to_inst(call):
    markup = types.InlineKeyboardMarkup(row_width=1)
    btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='naz_to_comis')
    markup.add(btn6)
    bot.send_message(call.message.chat.id, 'Будет дополняться, сейчас тут пусто', reply_markup=markup, parse_mode='Markdown')

@bot.callback_query_handler(func=lambda call: call.data == 'prepod')
def back_to_inst(call):
    markup = types.InlineKeyboardMarkup(row_width=1)
    btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='naz_to_comis')
    markup.add(btn6)
    bot.send_message(call.message.chat.id, 'Будет дополняться, сейчас тут пусто', reply_markup=markup, parse_mode='Markdown')

@bot.callback_query_handler(func=lambda call: call.data == 'prepod1')
def back_to_inst(call):
    markup = types.InlineKeyboardMarkup(row_width=1)
    btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
    markup.add(btn6)
    bot.send_message(call.message.chat.id, 'Будет дополняться, сейчас тут пусто', reply_markup=markup, parse_mode='Markdown')

@bot.callback_query_handler(func=lambda call: call.data == 'prepod2')
def back_to_inst(call):
    markup = types.InlineKeyboardMarkup(row_width=1)
    btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection2')
    markup.add(btn6)
    bot.send_message(call.message.chat.id, 'Будет дополняться, сейчас тут пусто', reply_markup=markup, parse_mode='Markdown')

@bot.callback_query_handler(func=lambda call: call.data == 'zadol')
def back_to_inst(call):
    markup = types.InlineKeyboardMarkup(row_width=1)
    btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='naz_to_comis')
    markup.add(btn6)
    bot.send_message(call.message.chat.id, 'Будет дополняться, сейчас тут пусто', reply_markup=markup, parse_mode='Markdown')

@bot.callback_query_handler(func=lambda call: call.data == 'zadol1')
def back_to_inst(call):
    markup = types.InlineKeyboardMarkup(row_width=1)
    btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection1')
    markup.add(btn6)
    bot.send_message(call.message.chat.id, 'Будет дополняться, сейчас тут пусто', reply_markup=markup, parse_mode='Markdown')

@bot.callback_query_handler(func=lambda call: call.data == 'zadol2')
def back_to_inst(call):
    markup = types.InlineKeyboardMarkup(row_width=1)
    btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='process_group_selection2')
    markup.add(btn6)
    bot.send_message(call.message.chat.id, 'Будет дополняться, сейчас тут пусто', reply_markup=markup, parse_mode='Markdown')


@bot.message_handler(content_types=['text'])
def send_text(message):
    if message.text == "Привет":
      bot.send_message(message.chat.id, 'Привет')
      bot.register_next_step_handler(message, group_menu())


bot.polling()











####################################ПРЕПОДОВАТЕЛИ##############################

@bot.callback_query_handler(func=lambda call: call.data == 'prepod')
def back_to_inst(call):
    markup = types.InlineKeyboardMarkup(row_width=1)
    btn6 = types.InlineKeyboardButton(text="Назад↩️", callback_data ='naz_to_comis')
    btn7 = types.InlineKeyboardButton(text="Меню📖", callback_data ='back_to_menu')
    markup.add(btn6, btn7)
    bot.send_message(call.message.chat.id, 'Будет дополняться, сейчас тут пусто', reply_markup=markup, parse_mode='Markdown')



# Загрузка расписания

    # base_url = "https://bgitu.ru/studentu/raspisanie/ochnoe-obuchenie/2024/rv-2/"
    # url = "https://bgitu.ru/studentu/raspisanie/ochnoe-obuchenie/2024/rv-2/"  # второй раз указываем url, чтобы получить содержимое ссылки
    # # Отправляем GET-запрос и получаем содержимое страницы
    # response = requests.get(url)
    # html_content = response.content
    # # Создаем объект BeautifulSoup для парсинга HTML
    # soup = BeautifulSoup(html_content, "html.parser")
    # # Находим ссылку на странице, которая содержится в теге <a> с определенным текстом
    # link_tag = soup.find("a", string="Расписание на 2 семестр ЛХ ЛА САД 1-3 курсы.xls")
    # # Если найден тег <a>, получаем его URL и объединяем с базовым URL
    # if link_tag:
    #     link_href = link_tag.get("href")
    #     full_url = urljoin(base_url, link_href)
    #     # Определяем путь, по которому нужно сохранить файл
    #     save_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps"
    #     # Создаем директорию, если ее нет
    #     os.makedirs(save_path, exist_ok=True)
    #     # Скачиваем файл по объединенной ссылке и сохраняем в указанную папку
    #     file_path, headers = urlretrieve(full_url, os.path.join(save_path, "raspison.xls"))
    # else:
    #     bot.send_message(call.message.chat.id, 'Все плохо')
    # # Путь к файлу для проверки
    # file_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps\raspison.xlsx"
    # # Проверяем, существует ли файл
    # if os.path.exists(file_path):
    #     # Если файл существует, удаляем его
    #     os.remove(file_path)
    # # Путь к скриншоту для проверки
    # screenshot_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\raspison_range.png"
    # # Проверяем, существует ли скриншот
    # if os.path.exists(screenshot_path):
    # # Если скриншот существует, удаляем его
    #     os.remove(screenshot_path)
    # time.sleep(1)
    # # Конвертируем .xls файл в .xlsx
    # excel = win32.gencache.EnsureDispatch('Excel.Application')
    # wb = excel.Workbooks.Open(os.path.join(save_path, "raspison.xls"))
    # wb.SaveAs(os.path.join(save_path, "raspison.xlsx"), FileFormat=51)
    # wb.Close()
    # excel.Quit()
    # time.sleep(0.5)
    # xlsx_path = r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\raps\raspison.xlsx"
    # client = win32.Dispatch("Excel.Application")
    # wb = client.Workbooks.Open(xlsx_path)
    # ws = wb.ActiveSheet
    # ws.Range("A2:F11").CopyPicture(Format = 2)
    # img = ImageGrab.grabclipboard()
    # img.save(r"E:\otkat\NvidiaOTK\Grand Theft Auto  San Andreas\pybot\scr\scrff.jpg", "jpeg")
    # wb.Close()
    # client.Quit()
    # # Отправляем скриншот пользователю
    # bot.send_photo(call.message.chat.id, open('E:\\otkat\\NvidiaOTK\\Grand Theft Auto  San Andreas\\pybot\\scr\\scrff.jpg', 'rb'))