#!/usr/bin/python3
import os
import sys
import time
import smtplib
import threading
from requests import Session
from requests.auth import HTTPBasicAuth
from multiprocessing import Process
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.utils import formatdate

from asterisk.agi import *
from zeep import Client as zClient
from zeep.transports import Transport


agi = AGI()
main_dir = os.path.dirname(__file__)

class Client():
    """Список клиентов (сотрудников) и работа с ними"""
    _ITILIUM_URL = 'http://ITILLIUM WEB SERVICE'
    _ITILIUM_USER = 'USER'
    _ITILIUM_PASSWORD = 'PASSWORD'

    _list = []

    def __init__(self, name, mail, phone, ctype):
        self.name = name
        self.mail = mail
        self.phone = phone
        self.ctype = ctype
        Client._list.append(self)

    @staticmethod
    def get_list():
        return Client._list

    @staticmethod
    def get_count():
        return len(Client.get_list())

    @staticmethod
    def initialize():
        """Первичное заполнение списка клиентов"""
        session = Session()
        session.auth = HTTPBasicAuth(Client._ITILIUM_USER, Client._ITILIUM_PASSWORD)
        try:
            itilium_soap = zClient(
                Client._ITILIUM_URL, 
                transport=Transport(
                    session=session,
                    timeout=5,
                    ),
                )
            soap_result = itilium_soap.service.GetIDs()
        except Exception:
            return
        Client._parse_clients(soap_result.ClientArray.Client)
        Client._parse_clients(soap_result.InitiatorArray.Initiator)

    @staticmethod
    def format_phone(phone_str):
        """Приведение номера телефона к единому стилю"""
        result_phone = None
        tmp_phone = []
        phone_str = phone_str.strip()
        # Убираем из номера +7 или 8 для упрощения
        if phone_str.startswith('+7'):
            phone_str = phone_str[2:]
        elif phone_str.startswith('8') or phone_str.startswith('7'):
            phone_str = phone_str[1:]
        # Отрежем все незначащие символы
        for symbol in phone_str:
            if not symbol.isdigit():
                continue
            tmp_phone.append(symbol) # Конкатенировать строку сразу затратно по ресурсам
        if len(tmp_phone) == 10: # Городской (не внутренний) номер
            tmp_phone.insert(0, '+7')
        elif len(tmp_phone) == 6: # Городской номер без кода города
            tmp_phone.insert(0, '+74742') # Предполагаем липецкий код по-умолчанию
        if tmp_phone:
            result_phone = "".join(tmp_phone)
        return result_phone

    @staticmethod    
    def search_client(phone): 
        """Поиск клиента по телефону"""
        phone = Client.format_phone(phone)
        if not phone:
            return
        for cli in Client.get_list():
            for ph in cli.phone:
                if ph == phone:
                    return cli

    @staticmethod
    def _parse_clients(cli_list):
        def parse_cli_phone(phone_str):
            if not phone_str:
                return
            result_list = []
            phone_str = phone_str.split(',')        
            for phone in phone_str:
                phone = Client.format_phone(phone)
                if not phone:
                    continue
                result_list.append(phone)
            return result_list

        for cli in cli_list:
            try:
                if not cli.ClientEmail:
                    continue
                cli_phone = parse_cli_phone(cli.ClientPhone)
                if not cli_phone:
                    continue
                Client(
                    name=cli.Name,
                    ctype=cli.Client_Name,
                    phone=cli_phone,
                    mail=cli.ClientEmail,
                )
            except Exception:
                agi.verbose("client parse failed " + str(cli.Name))
                continue


def exit_procedures():
    """Системный выход"""
    agi.verbose("bye!")
    agi.hangup()
    sys.exit()


def play_menu(sound, avail_choices):
    """Проиграть звуковое меню
    
    sound - название звукового файла для проигрывания, без расширения
    avail_choices - список с доступными цифрами для выбора в меню, кроме 0, напр. [1, 2, 3]
    """
    MAX_CYCLE_COUNT = 5 # Максимальное кол-во циклов проигрывания меню перед завершением звонка

    menu_choose = 0
    cycle_count = 0    
    while not menu_choose and cycle_count < MAX_CYCLE_COUNT:
        cycle_count = cycle_count + 1
        try:
            menu_choose = agi.get_option(
                    filename=sound, 
                    escape_digits=[0] + avail_choices,
                    timeout=3000,
                )
            menu_choose = int(menu_choose)
        except Exception:
            agi.verbose("menu failed")
        agi.verbose(f"choosed menu: {menu_choose}")
    if not menu_choose: # При неверном выборе в меню или таймауте - завершаем звонок
        exit_procedures()
    return menu_choose


def get_client_type(pnumber):
    """Определение типа клиента
    0 - нет связи с Итилиум
    1 - неизвестный клиент
    2 - Орг-ция
    3 - Орг-ция2
    4 - Орг-ция3
    """
    client_type = 0
    client_type_name = ''
    client = None
    if Client.get_count(): 
        client_type = 1
        client = Client.search_client(pnumber)
        if client:
            if client.ctype == 'Орг-ция2':
                client_type = 3
                client_type_name = 'Орг-ция2'
            elif client.ctype == 'Орг-ция3':
                client_type = 4
                client_type_name = 'Орг-ция3'
            else:
                client_type = 2
        agi.verbose(f"client type {client_type}")
    return client_type, client_type_name, client


def structure_menu(client_type, menu_choices=[]):
    """Заданная структура меню (его содержимое)

    client_type - тип клиента    
    menu_choices - пункты меню, выбранные пользователем (список по-порядку)

    Выходные значения:    
    struct['sound'] - файл озвучки меню (без расширения)
    struct['avail_choices'] - возможные варианты выбора, списком, напр. [1, 2, 3]
    struct['avail_choices_name'] - справочник названий пунктов меню (вариантов выбора)
    Пустой ответ - предыдущее меню было последним в иерархии, пользхователь выбрал конечный пункт
    """
    # Определение меню
    # Число - пункт меню, которые необходимо нажать пользователю, чтобы его выбрать
    # пункт - 'sound - путь к файлу озвучки меню
    if client_type == 3:
        # Меню Орг-ция2 (Торговая сеть, Магазины)
        menu = {
            (): { # Верхний уровень меню                
                'sound': main_dir + "/sound/upper_menu", # Аудио - верхнее меню Орг-ция2
                1: 'Проблема / вопрос с программой',
                2: 'Проблемы с компьютером, монитором, принтером, ИБП',
                3: 'Прочее',
            },
            (1,): { # Подуровень пункта 1 'Проблема / вопрос с программой'                
                'sound': main_dir + "/sound/second_menu_org2", # Аудио - подменю 'Проблема / вопрос с программой' Орг-ция2
                1: 'Проблема / вопрос с ЗУП, Бухгалтерия',
                2: 'Проблема / вопрос с Управление Торговлей',
                3: 'Проблема / вопрос с Розница',
                4: 'Проблема / вопрос по другой программе',
            },
        }
    elif client_type == 4:
        # Меню Орг-ция3 ()
        menu = {
            (): { # Верхний уровень меню                
                'sound': main_dir + "/sound/upper_menu", # Аудио - верхнее меню Орг-ция3
                1: 'Проблема / вопрос с программой',
                2: 'Проблемы с компьютером, монитором, принтером, ИБП',
                3: 'Прочее',
            },
            (1,): { # Подуровень пункта 1 'Проблема / вопрос с программой'                
                'sound': main_dir + "/sound/second_menu_org3", # Аудио - подменю 'Проблема / вопрос с программой' Орг-ция3
                1: 'Проблема / вопрос с ЗУП, Бухгалтерия',
                2: 'Проблема / вопрос по другой программе',
            },
        }
    else:
        # Меню для Орг-ция + неизвестные номера        
        menu = {
            (): { # Верхний уровень меню                
                'sound': main_dir + "/sound/upper_menu", # Аудио - верхнее меню Орг-ция
                1: 'Проблема / вопрос с программой',
                2: 'Проблемы с компьютером, монитором, принтером, ИБП',
                3: 'Прочее',
            },
            (1,): { # Подуровень пункта 1 'Проблема / вопрос с программой'                
                'sound': main_dir + "/sound/second_menu_org", # Аудио - подменю 'Проблема / вопрос с программой' Орг-ция
                1: 'Проблемы с почтой',
                2: 'Проблема / вопрос с ЕРП',
                3: 'Проблема / вопрос с СЭД',
                4: 'Проблема / вопрос с ЗУП, Бухгалтерия',
                5: 'Проблема / вопрос с ...',
                6: 'Проблема / вопрос по другой программе',
            },
        }

    struct = {}
    try:
        menu_item = menu[tuple(menu_choices)]
    except KeyError:
        return
    for key,val in menu_item.items():
        if key == 'sound':
            struct['sound'] = val            
        elif isinstance(key, int) and key > 0 and key < 10:
            try:
                struct['avail_choices']
            except KeyError:
                struct['avail_choices'] = []
                struct['avail_choices_name'] = {}
            struct['avail_choices'].append(key)
            struct['avail_choices_name'][key] = val            

    if struct: # Проверка на наличие обязательных полей
        try:
            if not struct['avail_choices'] or not struct['avail_choices_name'] or not struct['sound']:
                raise ValueError
        except (KeyError, ValueError):
            agi.verbose("!!! Wrong menu configuration !!!")
            return

    return struct


def check_filemessage(filename):
    """Проверка файла на доступность для отправки (записано ли обращение целиком)"""
    GLOBAL_TIMEOUT = 10 * 60 # Через это количество секунд выполнение процедуры прекращается

    # Проверка файла на то, что он дозаписан и готов к отправке
    timer_start = time.time()
    before_size = None
    while True:
        if time.time() > timer_start + GLOBAL_TIMEOUT:
            return False #  Неудача          
        time.sleep(10)
        if not os.path.isfile(filename):
            continue
        if before_size is None:
            before_size = os.stat(filename).st_size
            continue
        last_size = os.stat(filename).st_size
        if before_size < last_size:
            before_size = last_size
            continue # Файл еще пишется
        break
    return True


def send_filemessage(filename, message_sender, message_theme, message_body): 
    """Отправить записываемый файл сообщения в итилиум заявкой
    Данная функция работает уже после завершения звонка

    """
    MESSAGE_TO = 'help@org.ru'
    MESSAGE_DEFAULT_SENDER = 'no-reply@org.ru'
    SMTP_HOST = '127.0.0.1'
    SMTP_LOGIN = ''
    SMTP_PASSWORD = ''

    # Обработка и считывание вложения
    file_ready = check_filemessage(filename)
    if file_ready:
        attach_header = 'Content-Disposition', 'attachment; filename="%s"' % os.path.basename(filename)
        attachment = MIMEBase('application', "octet-stream")
        try:
            with open(filename, "rb") as fh:
                data = fh.read()
            attachment.set_payload(data)
            encoders.encode_base64(attachment)
            attachment.add_header(*attach_header)
        except IOError:
            file_ready = False
    # Создание сообщения
    msg = MIMEMultipart()
    msg["From"] = message_sender
    msg["To"] = MESSAGE_TO
    msg["Subject"] = Header(message_theme, 'utf-8')
    msg["Date"] = formatdate(localtime=True)
    # Добавление вложения
    if file_ready:
        msg.attach(attachment)
    else:            
        message_body += '\r\n' + 'Не удалось прикрепить аудиофайл обращения! Уточните у пользователя причину обращения самостоятельно' + \
            '\r\n\r\n' + filename
        filename = None
    if message_body:
        msg.attach(MIMEText(message_body, 'plain', 'utf-8'))

    # Отправка
    server = smtplib.SMTP(SMTP_HOST)
    #server.login(SMTP_LOGIN, SMTP_PASSWORD)
    # try:
    #     server.sendmail(message_sender, MESSAGE_TO, msg.as_string())
    # except smtplib.SMTPSenderRefused:
    #     server.sendmail(MESSAGE_DEFAULT_SENDER, MESSAGE_TO, msg.as_string())
    server.sendmail(MESSAGE_DEFAULT_SENDER, MESSAGE_TO, msg.as_string()) # Всегда логинимся под default_sender, т.к. в теле письма e-mail обрабатывается корректно
    server.quit()      
      
    if file_ready:
        os.remove(filename)
    sys.exit()


def record_and_send_filemessage(pnumber, message_sender, message_theme, message_body):
    """Запись обращения и отправка в итилиум
    Важно: Если пользователь повесит трубку на этом этапе - дальнейший код в текущем процессе не будет выполнен, стоит это учитывать

    pnumber - номер звонящего
    message_sender - от чьего имени (адреса) отправлять письмо (заявку)
    message_theme - тема письма
    message_body - тело письма
    """
    uniq_id = agi.get_variable("UNIQUEID")
    agi.stream_file( main_dir + "/sound/record_notify") # Аудио - уведомление перед записью обращения
    filename = f"{main_dir}/records/{pnumber}-{uniq_id}"
    fileext = 'wav'
    full_filename = f"{filename}.{fileext}"
    agi.verbose(full_filename)
    # Перед началом записи - запускаем второй процесс, который закончит работу и отправит сообщение после окончания звонка
    Process(target=send_filemessage, args=(full_filename,message_sender,message_theme,message_body)).start()
    # Запись
    agi.record_file(
        filename=filename,
        format=fileext,
        beep="beep", 
        timeout=-1,
        silence=5,
    )
    return full_filename


def main():
    agi.verbose(f"python agi started from {main_dir}")
    pnumber = agi.get_variable("CALLERID(num)")
    agi.verbose(f"call from {pnumber}")
    pnumber = Client.format_phone(pnumber) # Форматируем номер в единый формат
    agi.answer()
    threading.Thread(target=Client.initialize).start() # Запуск получения данных из итилиума (асинхронно)
    time.sleep(2) # Без этого первые секунды аудио неслышны

    agi.stream_file(main_dir + "/sound/welcome") # Аудио - приветствие    

    # Определение типа клиента
    client_type, client_type_name, client = get_client_type(pnumber)

    # Различные ограничения    
    if client_type == 1: 
        if pnumber.startswith('+79'): # Если неизвестный клиент с неизвестным мобильным номером, не разрешаем
            agi.stream_file(main_dir + "/sound/num_not_found") # Аудио - неизвестный номер
            exit_procedures()

    # Голосовое меню
    menu_choices = []
    choice_name = ''    
    while True:
        struct = structure_menu(client_type, menu_choices)
        if not struct:            
            break        
        choice = play_menu(struct['sound'], struct['avail_choices'])
        if choice:
            menu_choices.append(choice)
            choice_name = struct['avail_choices_name'][choice]
    agi.verbose(f"{menu_choices} {choice_name}")

    # Обработка конечного выбора в голосовом меню
    # сейчас все пункты меню кончаются записью обращения в итилиум, так что безусловно выполняем ее
    message_sender = 'no-reply@org.ru' if not client else client.mail
    message_theme = '[' + (client_type_name + ' ' + choice_name).strip() + ']' + ' Новая заявка по телефону'
    message_body = 'Поступила новая заявка по телефону '
    if client:
        message_body += f'от {client.name} ({pnumber}, {client.mail})'  
    else:
        message_body += f'от неизвестного абонента ({pnumber})' + '\r\n' + \
                        'Требуется внести данный номер телефона в справочник итилиум (почта, телефон и подразделение)!'  
    message_body += '\r\n' + 'Тема обращения: ' + choice_name
    record_and_send_filemessage(pnumber, message_sender, message_theme, message_body)

    # !!! На данном этапе звонок может быть уже завершен, дальнейший код может быть не выполнен
    exit_procedures()


if __name__ == "__main__":
    main()