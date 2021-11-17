from inserts import *
import psycopg2
import win32print
from itertools import chain
import datetime as dt
from configparser import ConfigParser
from PyQt5.QtCore import QDate as Date
import logging
import docxtpl
from PyQt5.QtWidgets import QMessageBox as mes
from inserts import db_keys
import os
si = ["тн", "т", "кг", "м2", "м", "м/п", "мм", "м3", "л", "мм", "шт"]
count_days = (31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
statues = ["работает", "отпуск", "уволен"]
types_docs = {"1": "/ОТ.docx", "2": "/ПТМ.docx", "3": "/ЭБ.docx"}
types_card = {"1": "/ОТ_уд.docx", "2": "/ПТМ_уд.docx", "3": "/ЭБ_уд.docx"}
ERR = -1

GET_UI = "Не удалось найти файл дизайна: "
GET_INI = "Не удалось получить данные из config: "
GET_FILE = "Не удалось найти файл: "
GET_PAGE = "Нет странице в файле: "
GET_DB = "Не удалось получить данные из Базы данных"
GET_NEXT_ID = "Ошибка при работе с БД"
LOAD_UI = "Не удалось загрузить файл дизайна"
CREATE_FOLDER = "Не удалось создать папку: "
CONNECT_DB = "Нет подключения к базе данных"
UPDATE_DB = "Не удалось обновить данные в базе данных"
KILL_DB = "Не удалось удалить данные из базы данных"
ADD_DB = "Не удалось добавить данные в базу данных"
FULL_ALL = "Заполните все поля"
ADD_CONTRACT = "Добавьте договор"
ADD_PEOPLE = "Добавьте сотрудников"
OPEN_WEB = "Не удалось открыть ссылку "
BIG_INDEX = "Ошибка: big index"
PERMISSION_ERR = "Проблема с правом доступа. " \
                 " 1. - вручную скопируйте файл в папку со счетами," \
                 " 2 - Снова нажмите Выбрать файл и укажите файл чека в новом месте"

CREATE_DB = "База данных успешно создана"
CREATE_DOCS = "Не удалось создать файлы с документацией"
CREATE_ACT = "Не удалось создать исполнительную"
CREATE_JOURNAL = "Не удалось создать журнал"
WRONG_DATE = "Дата начала старше даты окончания договора"
CHECK_COVID = "Не правильно заполнены данные по вакцинации"

KILLD_NOTE = "Запись удалена"
KILL_NOTE = "Вы действительно хотите удалить запись: "
ER_KILL_NOTE = "Не удалось удалить данные"
CHANGED_NOTE = "Запись изменена"
CHANGE_NOTE = "Вы действительно хотите изменить запись на "
ADDED_FILE = "Файл добавлен"
ADDED_NOTE = "Запись добавлена"
YES = "да"
NO = "нет"
PLACE_VAC = "Укажите место прививки"
OLD_DOC = "Сертификат устарел. С даты получения прошло более, чем {0} дней."
ERR_VAC_MANY = "Между прививками прошло {0} дней. Это менее {1} дней"
ERR_VAC_MACH = "Между прививками прошло {0} дней. Это менее {1} дней"
NUMBER_DOC = "Укажите номер сертификата"

CUSTOMER = "Заказчик"
CONTRACTOR = "Подрядчик"
SPUTNIK = "2 дозы"
SP_LITE = "1 доза"
COVID = "болел"
ZERO = "01.01.2000"
NOT = "(нет)"

#TB
ALL = "*"
WORKERS = "workers"
ITRS = "itrs"
AUTO = "auto"
DRIVERS = "drivers"
BOSSES = "bosses"
COMPANY = "company"
CONTRACTS = "contracts"
TOOLS = "tools"
TENDERS = "tenders"
BILLS = "bills"
FINANCE = "finance"
MATERIALS = "materials"
LINKS = "links"

ASR_FILE = "/asr.docx"
JOURNAL_FILE = "/Журнал.docx"
PDF = ".pdf"
UI = ".ui"
XLSX = ".xlsx"
DOCX = ".docx"

TO_PDF = "to_pdf"
TO_PAPER = "to_paper"

dictionary = {"Производитель работ": {"gent": "производителя работ", "datv": "производителю работ"},
              "Технический директор": {"gent": "технического директора", "datv": "техническому директору"}}
path_log = os.getcwd() + "/log_file.log"
logging.basicConfig(filename=path_log, level=logging.INFO)
month = ["январь", "февраль", "март", "апрель", "май", "июнь",
         "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"]


def msg_er(widgets, text):
    mes.question(widgets, "Сообщение", text, mes.Ok)
    logging.info(text + " " + str(dt.datetime.now()))
    return -1


def msg_info(widgets, text, btn=mes.Ok):
    mes.question(widgets, "Сообщение", text, btn)
    return True


def msg_q(widgets, text):
    return mes.question(widgets, "Сообщение", text, mes.Ok | mes.Cancel)


class DataBase:
    def __init__(self, parent=None):
        self.parent = parent
        self.conn = None
        self.cursor = None
        self.conf = parent.conf
        try:
            self.ip = self.conf.config.get('database', 'ip')
            self.name_db = self.conf.config.get('database', 'name_db')
            self.user_db = self.conf.config.get('database', 'user_db')
            self.password_db = self.conf.config.get('database', 'password_db')
            self.port = self.conf.config.get('database', 'port')
        except:
            return

    def connect_to_db(self):
        try:
            self.conn = psycopg2.connect(dbname=self.name_db,
                                         user=self.user_db,
                                         password=self.password_db,
                                         host=self.ip,
                                         port=self.port)
            if not self.conn:
                return msg_er(self.parent, CONNECT_DB)
            self.cursor = self.conn.cursor()
            if not self.cursor:
                return msg_er(self.parent, CONNECT_DB)
        except:
            return msg_er(self.parent, CONNECT_DB)

    def get_data(self, fields, table):
        try:
            row = get_from_db(fields, table)
            self.connect_to_db()
            self.execute(row)
            rows = self.cursor.fetchall()
            man = []
            for row in rows:
                tmp = []
                for item in row:
                    tmp.append(item)
                man.append(tmp)
            return man
        except:
            return msg_er(self.parent, GET_DB)

    def execute(self, text):
        self.connect_to_db()
        self.cursor.execute(text)
        try:
            pass
        except:
            return msg_er(self.parent, GET_DB)
        return True

    def get_next_id(self, table):
        try:
            self.connect_to_db()
            rows = self.get_data("id", table)
            if not rows:
                return 1
            return int(max(rows)[0]) + 1
        except:
            return msg_er(self.parent, GET_NEXT_ID)

    def my_commit(self, data):
        self.connect_to_db()
        self.cursor.execute(data)
        self.conn.commit()
        if data:
            try:

                return True
            except:
                return msg_er(self.parent, ADD_DB)

    def init_list(self, widget, fields, table, people=False):
        rows = self.get_data(fields, table)
        man = []
        for row in rows:
            tmp = []
            for item in row:
                tmp.append(item)
            man.append(tmp)
        man.sort(key=lambda x: x[-1])
        if not man and man != []:
            return False
        if not people:
            for row in man:
                widget.addItems([row[-1] + ". " + row[0]])
            return man
        else:
            for row in man:
                widget.addItems([str(row[-1]) + ". " + short_name(row[:3])])
            return man

    def my_update(self, data, table):
        try:
            self.connect_to_db()
            self.cursor.execute(my_update(data, table))
            self.conn.commit()
        except:
            return msg_er(self.parent, UPDATE_DB)

    def kill_value(self, my_id, table):
        try:
            self.connect_to_db()
            self.execute("DELETE FROM {0} WHERE id = '{1}'".format(table, my_id))
            self.conn.commit()
        except:
            return msg_er(self.parent, KILL_DB)

    def new_note(self, date, name, number):
        self.my_commit(add_to_db((date, name, number), "notes"))

    def create_db(self):
        create = "CREATE TABLE "
        drop = "DROP TABLE "
        for key in db_keys.keys():
            fields = db_keys[key][1:-1].split(", ")
            res_str = "("
            for word in fields:
                res_str += word + " text, "
            row = create + key + " " + res_str[:-2] + ")"
            try:
                self.cursor.execute(row)
                self.conn.commit()
            except Exception as ex:
                try:
                    row = drop + key
                    self.connect_to_db()
                    self.cursor.execute(row)
                    self.conn.commit()
                    row = create + key + " " + res_str[:-2] + ")"
                    self.cursor.execute(row)
                    self.conn.commit()
                except:
                    msg_info(self, "Ошибка создание Базы данных " + key)
                    logging.debug(key + " is NOT created. " + str(ex))
                continue
            logging.debug(key + " is created")


class Ini:
    def __init__(self, parent=None):
        self.parent = parent
        self.path_conf = os.getcwd() + "/config.ini"
        self.path_ui = os.getcwd() + "/ui_files/"
        self.config = ConfigParser()
        self.config.read(self.path_conf, encoding="windows-1251")
        self.path_folder = self.get_from_ini("path", "path")
        self.sections = {"conf": "config", "path": "path", "ui": "ui_files"}

    def get_path(self, my_type):
        self.path_folder + str(self.config.get(self.sections["path"], my_type))
        try:
            return self.path_folder + str(self.config.get(self.sections["path"], my_type))
        except:
            msg_er(self.parent, GET_INI)
            return ERR

    def get_config(self, my_type):
        try:
            return self.config.get(self.sections["conf"], my_type)
        except:
            msg_er(self.parent, GET_INI)
            return ERR

    def get_from_ini(self, my_type, part):
        try:
            return self.config.get(part, my_type)
        except:
            msg_er(self.parent, GET_INI)
            return ERR

    def get_ui(self, ui_file):
        return self.path_ui + ui_file + UI

    def get_next_number(self):
        try:
            number_note = self.config.get(self.sections["conf"], 'number')
            return int(number_note)
        except:
            msg_er(self.parent, GET_INI)
            return ERR

    def set_next_number(self, n):
        try:
            number_note = self.config.get(self.sections["conf"], 'number')
            next_number = n
            self.config.set(self.sections["conf"], 'number', str(next_number))
            with open(self.path_conf, 'w') as configfile:
                self.config.write(configfile)
            return int(number_note)
        except:
            msg_er(self.parent, GET_INI)
            return ERR

    def set_val(self, section, field, val):
        try:
            self.config.set(section, field, str(val))
            with open(self.path_conf, 'w') as configfile:
                self.config.write(configfile)
            return True
        except:
            msg_er(self.parent, GET_INI)
            return ERR


def short_name(data):
    if not data:
        return ""
    return data[0] + " " + data[1][0] + "." + data[2][0] + "."


def full_name(data):
    return " ".join(data[:3])


def time_delta(date_1, date_2):
    if date_1 == "now":
        a = dt.datetime(dt.datetime.now().year, dt.datetime.now().month, dt.datetime.now().day)
    else:
        a = dt.datetime(int(date_1[6:]), int(date_1[3:5]), int(date_1[:2]))
    b = dt.datetime(int(date_2[6:]), int(date_2[3:5]), int(date_2[:2]))
    return (a - b).days


def from_str(date):
    symbols = [".", ",", "-", "/", "_"]
    for item in symbols:
        tmp = date.split(item)
        if len(tmp) == 3:
            if len(tmp[0]) == 4:
                return Date(int(tmp[0]), int(tmp[1]), int(tmp[2]))
            if len(tmp[2]) == 4:
                return Date(int(tmp[2]), int(tmp[1]), int(tmp[0]))

zero = from_str(ZERO)


def yong_date(young, old):
    if int(young[6:10]) > int(old[6:10]):
        return True
    if int(young[6:10]) < int(old[6:10]):
        return False

    if int(young[3:5]) > int(old[3:5]):
        return True
    if int(young[3:5]) < int(old[3:5]):
        return False

    if int(young[:2]) > int(old[:2]):
        return True
    if int(young[:2]) < int(old[:2]):
        return False

    else:
        return False


def get_val(ui):
    return "".join(ui.currentText().split(". ")[1:])


def get_index(cb, text):
    for ind in range(cb.count()):
        cb.setCurrentIndex(ind)
        if text in cb.currentText():
            return ind
    return -1


def set_fix_size(wnd):
    wnd.setFixedSize(wnd.geometry().width(), wnd.geometry().height())


def print_to(file, type_=TO_PAPER):
    conf = Ini()
    printer = conf.get_config(type_)
    win32print.SetDefaultPrinter(printer)
    try:
        os.startfile(file, "print")
    except:
        pass


# _________ Безопасные функции ___________


def save_open(file):
    try:
        os.startfile(file)
    except:
        return False


def save_replace(target, sourсe):
    try:
        os.replace(sourсe, target)
    except:
        return False


def save_open_docxtpl(path):
    try:
        doc = docxtpl.DocxTemplate(path)
        return doc
    except:
        return False


def save_save_doc(path, doc):
    try:
        doc.save(path)
        return True
    except:
        return False

def add_unic_items(wgt, table, key):
    rows = DataBase().get_data(key, table)
    added = list()
    for item in rows:
        if not item[0] in added:
            added.append(item)
            for item in wgt:
                item.addItem(item[0])