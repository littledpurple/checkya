#!/usr/bin/env python3

######
# Скрипт для проверки отправления/получения почты с Zimbra
# Отправляет письма с локального на удаленный ящик и с удаленного на локальный 
# Далее проверяет, пришли ли письма на удалённый и локальный ящик
# Генерирует и отправляет отчёт при провале отправки/получения gbcmvf
# Либо при аргументе --debug
#####

import smtplib, imaplib, ssl, telebot, emoji, argparse, configparser
from time import sleep
from datetime import datetime, timedelta
from os import path

''' Парсинг аргумента для дебага '''
parser = argparse.ArgumentParser(
                    prog='checkYa',
                    description="Скрипт для проверки исходящей/входящей связи для Zimbra\n"\
                                "Отправляет письма с локального и удаленного почтового ящиков\n"\
                                "Затем проверяет, пришли ли они\n",
                    formatter_class=argparse.RawTextHelpFormatter)

parser.add_argument('--debug', action='store_true', help='Отладочная информация в Телеграм')
parser.set_defaults(debug=False)
debug = parser.parse_args().debug

''' Парсинг конфигурации '''
try:
    config = configparser.ConfigParser()
    config.read(path.join(path.dirname(__file__), 'config.ini'))
    ''' Основные настройки '''
    s_port = int(config["main"]["s_port"])
    r_port = int(config["main"]["r_port"])
    delete = bool(config["main"]["delete"])
    delay = int(config["main"]["delay"])
    ''' Локальный почтовый сервер '''
    s_server = config["local"]["smtp"]
    s_r_server = config["local"]["imap"]
    s_email = config["local"]["email"]
    s_password = config["local"]["password"]
    ''' Удалённый почтовый сервер '''
    r_server = config["remote"]["imap"]
    r_s_server = config["remote"]["smtp"]
    r_email = config["remote"]["email"]
    r_password = config["remote"]["password"]
    ''' Настройки Telegram '''
    tg_token = config["telegram"]["token"]
    tg_chat = config["telegram"]["chat"]
    tg_admin = config["telegram"]["admin"]

except Exception as ex:
    print(f"Не удалось прочитать конфигурационный файл\n"
          f"{ex}")
    exit()

''' Нужные функции '''
def tg_send(msg:str, tg_token:str=tg_token, tg_id:str=tg_admin):
    '''
    Функция-обработчик отправки сообщения в Телеграм
    Также добавляет emoji
    msg - сообщение
    tg_token - токен бота
    tg_id - id чата, куда шлём сообщение 
    '''
    try:
        bot = telebot.TeleBot(tg_token)
        bot.send_message(tg_id, emoji.emojize(msg), parse_mode='HTML')
    except Exception as ex:
        print(f'Ошибка отправки сообщения в Telegram\n'
              f'{msg}\n'
              f'{ex}')

def gen_msg(r_email:str, s_email:str, subject:str=f"Тест отправки с {s_email} на {r_email}", text:str=f"Тест отправки с {s_email} на {r_email}"):
    '''
    Генерация email'а со всеми заголовками
    Прежде всего для правильной работы Яндекс почты
    Без них шлёт почту без From
    r_email - email получателя
    s_email - email отправителя
    subject - тема письма
    text - тест письма
    '''
    msg = 'From: {}\nTo: {}\nSubject: {}\n\n{}'.format(s_email,
                                                       r_email, 
                                                       subject, 
                                                       text)
    return msg.encode('utf-8')

def gen_report(sr:bool=False, ss:bool=False, rr:bool=False, rs:bool=False):
    '''
    Генерация отчёта о работе стрипта
    sr - отправка с Зимбры на Яндекс
    ss - отправка с Яндекс на Зимбру
    rr - получение на Яндекс
    rs - получение на Зимбре
    '''

    return f":page_facing_up: <b>Отчёт об отправке/получении писем</b>\n"\
           f"{gen_str(sr, 'Zimbra')} => {gen_str(rr, 'Yandex')}\n"\
           f"{gen_str(ss, 'Yandex')} => {gen_str(rs, 'Zimbra')}"
    
def gen_str(status:bool=False, text:str=""):
    '''
    Генерация строчки по каждому из направлений отправки/получения
    status - статус отправления/получения
    text - текст проверки
    '''
    if status:
        return f":check_mark_button: {text}"
    else:
        return f":cross_mark: {text}"

def em_send(s_server:str, s_port:int, s_mail:str, s_pass:str, r_mail:str, msg:str=f"Тест отправки с {s_email} на {r_email}"):
    '''
    Отправка письма
    s_server - smtp сервер для отправки
    s_port - порт smtp сервера для отправки
    s_mail - email для отправки
    s_pass - пароль от email'а для отправки
    r_mail - email получателя
    msg - сообщения
    '''

    try:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(s_server, s_port, context=context) as server:
            server.login(s_mail, s_pass)
            server.sendmail(s_mail, r_mail, gen_msg(r_mail, s_mail))
    except Exception as ex:
        if debug:
            print(ex)
            #tg_send(ex)
        return False
    else:
        if debug:
            tg_send(f":check_mark_button: Письмо с {s_mail} на {r_mail} <b>успешно отправлено</b>")
        return True

def em_read(imap_srv:str, imap_mail:str, imap_pass:str, s_email:str):
    '''
    Проверка получения письма
    imap_srv - imap сервер
    imap_mail - email для проверки входящих
    imap_pass - пароль от email'а для проверкаи входящих
    s_email - email, с которого должно приходить сообщение для проверки
    '''
    try:
        imap_server = imaplib.IMAP4_SSL(imap_srv)
        imap_server.login(imap_mail, imap_pass)
        imap_server.select("INBOX")
        result, data = imap_server.search(None, f'(UNSEEN FROM "{s_email}")')
        uids = [int(s) for s in data[0].split()]
        uids_l = len(uids)
        if uids_l > 0:
            for uid in uids:
                result, data = imap_server.fetch(str(uid), '(RFC822)')
                if delete: imap_server.store(str(uid), '+FLAGS', '\\Deleted') # Флаг для удаления письма
        if delete: imap_server.expunge()
        imap_server.close()
        imap_server.logout()
    except Exception as ex:
        if debug:
            tg_send(ex)
        return False
    else:
        if uids_l > 0:
            if debug:
                tg_send(f":check_mark_button: Письмо с {s_email} на {imap_mail} <b>успешно получено</b> ({uids_l} шт.)")
            return True
        else:
            if debug:
                tg_send(f":cross_mark: Письмо с {s_email} на {imap_mail} <b>не получено</b>!")
            return False

''' Рабочая часть '''        
# Отправка писем в обе стороны
sr = em_send(s_server, s_port, s_email, s_password, r_email)
ss = em_send(r_s_server, s_port, r_email, r_password, s_email)

sleep(delay) # выжидаем некоторое время, пока сервера обработают входящие

# Проверка получения писем
rr = em_read(r_server, r_email, r_password, s_email)
rs = em_read(s_r_server, s_email, s_password, r_email)

# Отправка отчёта для дебага администратору
if debug:
    tg_send(gen_report(sr, ss, rr, rs))
# Отправка проваленных отчётов в общий чат
if not(sr and ss and rr and rs):
    tg_send(msg=gen_report(sr, ss, rr, rs),tg_id=tg_chat)
