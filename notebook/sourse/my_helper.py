from PyQt5.QtWidgets import QMainWindow, QApplication, QCheckBox, QAction, QInputDialog,  QFileDialog
import sys
import requests
import zipfile
import shutil
import win32print
from new_boss import NewBoss
from new_itr import NewITR
from new_worker import NewWorker
from nw_company import NewCompany
from new_auto import NewAuto
from new_driver import NewDriver
from new_bill import NewBill
from covid19 import NewCovid
from table import NewTable
from pdf_module import PDFModule
from new_contract import NewContact
from new_material import NewMaterial
from new_tool import NewTool
from new_TB import NewTB
from new_tender import NewTender
from acts import Acts
from my_email import *
from pass_week import WeekPass
from pass_unlock import UnlockPass
from pass_month import MonthPass
from pass_get import GetPass
from pass_auto import AutoPass
from pass_drive import DrivePass
from pass_tools import ToolsPass
from pass_dt import DTPass
from settings import Settings
from database import *
from consts import *
from my_tools import Notepad
from music import Web
from get_money import GetMoney
"""
1. ТБ (собирать удостоверения на одном листе, инструктаж) - 4 часа
2. Исполнительная (создание, заполнение) - 10 часов
+ 3. Отправка данных сотрудника (кнопка добавления +, тупо открыть менеджер и выбрать доки +) - 1 час
4. Отчеты по материалам (в пандас в эксель) - 1 час
- 5. Коммерческие (Форма + в договоре кнопку + и заполнить по форме) - 1 час
+ 6. Бланки на суточные (сделать кнопку меню и вынести все кнопки туда) - 30 мин +
+ 7. Путевые листы (в меню) - 30 мин +
8. Расширить работу в выходные на N чел (придумать как, скорее вбок добавлять) - 1 час
9. Дизаин формочек (разобрать как дизайнить) - 1 час
+ 10. Проверка работы (тестирование) - 2 часа
- 11. Исходящие в другие заводы (добавление папки, ведение коммерческих, добавление доков) - 1 час
- 12. Папка на тендер (сформировать) - 1 час +
13. Вакциная (решить с границами как задать) - 2 часа
+ 14. порядок в папке +
+ 15. сортировка списков +
+ 16. Магический список 
+ 17. Организация босса +
+ 18. Автоматическое формирование списка по объекту.
+ 19. Дизаин окон (2)
+ 20. Добавить выбор протокола
+ 21. Дизайн формы Сотрудника
+ 22. Создание ТБ
+ 23. Печать ТБ
24. Безопасное открытие и сохранение
25. Тестирование
Итого 26 часов, реально 20 часов - 2 полных дня.
Сделать
 """

class Instruct(QDialog):
    def __init__(self, parent):
        super(Instruct, self).__init__()
        path = parent.conf.get_ui("instruction")
        self.parent = parent
        uic.loadUi(path, self)
        self.b_print.clicked.connect(self.my_print)

    def my_print(self):
        if not save_open(os.getcwd() + INSTRUCTION):
            msg_er(self, GET_FILE + os.getcwd() + INSTRUCTION)
            return False


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.conf = Ini(self)
        path = self.conf.get_ui("main_menu")
        uic.loadUi(path, self)
        self.my_help = False
        self.db = DataBase(self)
        if self.db.connect_to_db() == ERR:
            return
        check = self.conf.get_config("is_created_db")
        if check == ERR:
            return
        if check == "NO":
            if self.db.create_db() == ERR:
                return
            if self.conf.set_val("config", "is_created_db", YES) == ERR:
                return
        self.check_start()
        self.b_pass_week.clicked.connect(self.start_wnd)
        self.b_pass_month.clicked.connect(self.start_wnd)
        self.b_pass_drive.clicked.connect(self.start_wnd)
        self.b_pass_unlock.clicked.connect(self.start_wnd)
        self.b_pass_issue.clicked.connect(self.start_wnd)
        self.b_tb.clicked.connect(self.start_wnd)
        self.b_new_person.clicked.connect(self.start_wnd)
        self.b_new_build.clicked.connect(self.start_wnd)
        self.b_new_boss.clicked.connect(self.start_wnd)
        self.b_new_itr.clicked.connect(self.start_wnd)
        self.b_new_company.clicked.connect(self.start_wnd)
        self.b_new_auto.clicked.connect(self.start_wnd)
        self.b_new_driver.clicked.connect(self.start_wnd)
        self.b_new_material.clicked.connect(self.start_wnd)
        self.b_new_bill.clicked.connect(self.start_wnd)
        self.b_notepad.clicked.connect(self.start_wnd)
        self.b_music.clicked.connect(self.start_wnd)
        self.b_get_money.clicked.connect(self.start_wnd)
        self.b_act.clicked.connect(self.start_wnd)
        self.b_pdf_check.clicked.connect(self.start_wnd)
        self.b_journal.clicked.connect(self.start_wnd)
        self.b_tabel.clicked.connect(self.start_wnd)
        self.b_new_tool.clicked.connect(self.start_wnd)
        self.b_tools.clicked.connect(self.start_wnd)
        self.b_tender.clicked.connect(self.start_wnd)

        exitAction = QAction('Настройки', self)
        exitAction.setStatusTip('Настройки')
        sqlAction = QAction('Запрос в базу данных', self)
        sqlAction.setStatusTip('SQL запрос')
        instrAction = QAction('Инструкция', self)
        instrAction.setStatusTip('Инструкция')
        helpAction = QAction('Включить подсказки', self)
        helpAction.setStatusTip('Вкл/выкл подсказки')
        helpAction.triggered.connect(self.helpers)
        instrAction.triggered.connect(self.instruct)
        sqlAction.triggered.connect(self.sql_mes)
        exitAction.triggered.connect(self.ev_settings)

        menu = self.menu
        fileMenu = menu.addMenu("Настройки")
        fileMenu.addAction(instrAction)
        fileMenu.addAction(sqlAction)
        fileMenu.addAction(exitAction)
        self.b_empty.clicked.connect(self.start_file)
        self.b_days.clicked.connect(self.start_file)
        self.b_birki.clicked.connect(self.start_file)
        self.b_travel.clicked.connect(self.travel)
        self.b_dt.clicked.connect(self.get_dt)
        self.b_some.clicked.connect(self.start_file)
        self.b_tb.setEnabled(True)
        self.b_act.setEnabled(False)
        self.b_plan.setEnabled(False)
        self.b_attorney.clicked.connect(self.start_file)
        self.b_invoice.clicked.connect(self.start_file)

        self.get_param_from_widget = None
        self.company = self.conf.get_config("company")
        self.customer = self.conf.get_config("customer")
        if self.company == ERR:
            self.company = "<Подрядчик>"
        if self.customer == ERR:
            self.company = "<Заказчик>"
        self.company_ = ""
        self.customer_ = ""
        self.check_company()
        self.data_to_db = None
        self.get_weather()
        self.city = self.conf.get_config("city")
        if self.city == ERR:
            self.city = "<город>"

    def check_start(self):
        if self.conf.get_config("status") == "start":
            self.start_func()
            self.instruct()
            self.conf.set_val("config", "status", "worked")
            self.main_path = os.getcwd()
        else:
            self.main_path = self.conf.get_from_ini("path", "path")

    def get_dt(self):
        wnd = DTPass(self)
        wnd.exec_()
        pass

    def travel(self):
        count, ok = QInputDialog.getInt(self, "Кол-во копий", "Копий")
        if ok:
            for item in range(count):
                file = self.conf.get_path("patterns") + TRAVEL_1
                if not print_to(file):
                    return msg_info(self, GET_FILE + file)
            ok = msg_info(self, "Переверните распечатанную стопку и "
                                "вставьте в принтер повторно для печати оборотной стороны")
            for item in range(count):
                file = self.conf.get_path("patterns") + TRAVEL_2
                if not print_to(file):
                    return msg_info(self, GET_FILE + file)

    def helpers(self):
        self.my_help = not self.my_help

    def instruct(self):
        wnd = Instruct(self)
        wnd.exec_()

    def start_func(self):
        ok = msg_info(self, "Добро пожаловать! Кратко расскажу как работать с данной программой. Начнем!")
        # ok = msg_info(self, "Выберите где разместить главную папку, с которой будете работать")
        # self.main_path = QFileDialog.getExistingDirectory(self, "Open Directory", os.getcwd(), QFileDialog.ShowDirsOnly | QFileDialog.DontResolveSymlinks)
        # if self.main_path:
        #    shutil.copytree("/my_helper", self.main_path + "/Документация")
        # ok = msg_info(self, "Папка проекта создана!")

    def ev_settings(self):
        wnd = Settings(self)
        set_fix_size(wnd)
        wnd.exec_()

    def start_wnd(self):
        name = self.sender().text()
        dct = {"Продление на месяц": MonthPass,
               "Пропуск на выходные": WeekPass,
               "Разблокировка пропуска": UnlockPass,
               "Продление на машину": AutoPass,
               "Выдать пропуск": GetPass,
               "Разовый пропуск": DrivePass,
               "Ввоз инстр-ов": ToolsPass,

               "Материал": NewMaterial,
               "Инструмент": NewTool,
               "Договор": NewContact,
               "Сотрудник": NewWorker,
               "Автомобиль": NewAuto,
               "Организация": NewCompany,
               "Водитель": NewDriver,
               "Прораб": NewITR,
               "Чек": NewBill,
               "Тендер": NewTender,
               "Босс": NewBoss,
               "Распечатать ТБ": NewTB,
               "Журнал-ковид": NewCovid,
               "Табель": NewTable,

               "Блокнот": Notepad,
               "Сайты": Web,
               "Сканер": PDFModule,
               "Заявка на деньги": GetMoney,
               "Исполнительная": Acts}
        wnd = dct.get(name)(self)
        check = self.is_have_some(name)
        if check == ERR or not check:
            return
        set_fix_size(wnd)
        wnd.exec_()
        if self.sender().text() == "Организация":
           self.check_company()

    def check_company(self):
        company = self.db.get_data("*", COMPANY)
        if company == ERR:
            return
        for item in company:
            if item[-2] == CONTRACTOR:
                self.company_ = item
            if item[-2] == CUSTOMER:
                self.customer_ = item

    def start_file(self):
        files = {"Доверенность": ATTORNEY,
                 "Накладная": INVOICE,
                 "Бланк": BLANK,
                 "Суточные": DAYLY_M,
                 "Бирки на инстр.": MARK,
                 "Бланки на нар-ы": "/Наряды.jpg"}
        name = self.sender().text()
        folder = self.conf.get_path("pat_patterns")
        path = folder + files[name]
        if name in ["Доверенность", "Накладная", "Бланки на нар-ы"]:
            if files[name][1:] in os.listdir(folder):
                count, ok = QInputDialog.getInt(self, name, "Кол-во копий")
                if ok:
                    for ind in range(count):
                        print_to(path, TO_PAPER)
            else:
                msg_info(self, GET_FILE + path + files[name])
                return
        else:
            if not save_open(path):
                msg_info(self, GET_FILE + path)
            return

    def get_new_data(self, data):
        self.data_to_db = data

    def get_weather(self):
        s_city = self.conf.get_from_ini("city", "weather")
        city_id = 0
        appid = self.conf.get_from_ini("appid", "weather")
        try:
            res = requests.get(self.conf.get_from_ini("site_find", "weather"),
                               params={'q': s_city, 'type': 'like', 'units': 'metric', 'APPID': appid})
            data = res.json()
            city_id = data['list'][0]['id']
        except Exception as e:
            self.l_weather.setText("погода")
            self.l_temp.setText("температура")
        try:
            res = requests.get(self.conf.get_from_ini("site_weather", "weather"),
                               params={'id': city_id, 'units': 'metric', 'lang': 'ru', 'APPID': appid})
            data = res.json()
            self.l_weather.setText(data['weather'][0]['description'])
            self.l_temp.setText(str(round(data['main']['temp_max'])) + " C")
        except Exception as e:
            self.l_weather.setText("погода")
            self.l_temp.setText("температура")

    def is_have_some(self, name):
        table = ""
        text = ""
        if name == "Табель":
            wnd = NewTable(self)
            wnd.create_table()
            return False
        if name == "Журнал-ковид":
            table = WORKERS
            text = "Добавьте сначала хотя бы одного рабочего"
            data = self.db.get_data(ALL, table)
            if data == ERR:
                return ERR
            if not data:
                msg_info(self, text)
                return False
            wnd = NewCovid(self)
            wnd.create_covid()
            return False
        elif name == "Сотрудники":
            table = CONTRACTS
            text = "Добавьте сначала хотя бы один договор"
        elif name == "Договор":
            data = self.db.get_data("status", COMPANY)
            if not [CUSTOMER] in data:
                msg_info(self, "Введите сначала данные Заказчика")
                return False
            table = COMPANY
            text = "Добавьте сначала хотя бы одну организацию"
        elif name == "Босс":
            table = COMPANY
            text = "Добавьте сначала хотя бы одну организацию"
        else:
            return True
        data = self.db.get_data(ALL, table)
        if data == ERR:
            return ERR
        if not data:
            msg_info(self, text)
            return False
        return True

    def sql_mes(self):
        text, ok = QInputDialog.getText(self, "SQL запрос", "Введите запрос")
        if ok:
            if text:
                self.db.execute(text)
                try:
                    pass
                except:
                    msg_info(self, "Запрос не удался")
                    return
                file = open("text.txt", "w")
                try:
                    file.write(str(self.db.cursor.fetchall()))
                except:
                    pass


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = MainWindow()
    ex.setFixedSize(ex.geometry().width(), ex.geometry().height())
    ex.show()
    sys.exit(app.exec())

