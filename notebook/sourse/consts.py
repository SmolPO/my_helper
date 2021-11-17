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
# _____ Коды ошибок _____
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

#_______ База данных___________
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

ASR_FILE = "/asr.docx"
JOURNAL_FILE = "/Журнал.docx"
INSTRUCTION = "/Инструкция.docx"
TRAVEL_1 = "/Путевой_1.xlsx"
TRAVEL_2 = "/Путевой_2.xlsx"
ATTORNEY = "/Доверенность.xlsx"
INVOICE = "/Накладная.xlsx"
BLANK = "/Бланк.docx"
DAYLY_M = "/Суточные.xlsx"
MARK = "/Бирки.xlsx"
"/Наряды.jpg"
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

