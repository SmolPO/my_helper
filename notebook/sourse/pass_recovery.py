from PyQt5 import uic
from PyQt5.QtWidgets import QDialog, QMessageBox
from PyQt5.QtCore import Qt
from inserts import get_from_db
import datetime as dt
import os
import docx
import docxtpl
from configparser import ConfigParser
#  сделать мессаджбоксы на Сохранить
count_days = (31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
designer_file = '../designer_ui/pass_drive.ui'


class DrivePass(QDialog):
    def __init__(self, parent):
        super(DrivePass, self).__init__()
        uic.loadUi(designer_file, self)
        # my_pass
        self.parent = parent
        self.b_ok.clicked.connect(self.ev_OK)
        self.b_cancel.clicked.connect(self.ev_cancel)
        self.b_save.clicked.connect(self.save_pattern)
        self.b_kill.clicked.connect(self.kill_pattern)
        self.b_open.clicked.connect(self.my_open_file)
        self.count_people = 0

        self.cb_all.stateChanged.connect(self.set_enabled_workers)

        self.d_note.setDate(dt.datetime.now().date())
        self.number.setValue(self.parent.get_next_number())

        self.list_ui = (self.worker_1, self.worker_2, self.worker_3, self.worker_4, self.worker_5,
                        self.worker_6, self.worker_7, self.worker_8, self.worker_9, self.worker_10)
        self.list_month = ["январь", "февраль", "март", "апрель",
                           "май", "июнь", "июль", "август", "сентябрь",
                           "октябрь", "ноябрь", "декабрь"]
        self.data = {"customer": "", "company": "", "start_date": "", "end_date": "",
                     "post": "", "number": "", "data": ""}
        self.init_workers()
        self.init_cb_month()
        self.main_file += "/pass_drive.docx"

    # инициализация
    def init_cb_month(self):
        for elem in self.list_month:
            self.cb_month.addItem(elem)

    def init_workers(self):
        rows = self.parent.db.get_data("family, name, surname, post, passport, "
                                       "passport_got, birthday, adr,  live_adr", "workers")
        for item in self.list_ui:
            item.addItem("(нет)")
            item.activated[str].connect(self.new_worker)
            item.setEnabled(False)
        self.list_ui[0].setEnabled(True)
        for name in rows:
            family = name[0] + " " + ".".join([name[1][0], name[2][0]]) + "."
            for item in self.list_ui:
                item.addItem(family)

    # флаг на выбор всех
    def set_enabled_workers(self, state):
        for elem in self.list_ui:
            elem.setEnabled(state != Qt.Checked)

    # получить данные
    # для заполнения текста
    def get_data(self):
        next_month = self.list_month.index(self.cb_month.currentText()) + 1
        print("month: ", self.cb_month.currentText(), " index:", next_month)
        # если конец года: увеличить год и месяц в 1
        if next_month == 13:
            next_day = "09"
            next_month = "01"  # MessageBox для ввода первого дня
            next_year = str(dt.datetime.now().year + 1)
        else:
            next_day = "01"
            next_month = str(next_month)
            if int(next_month) < 10:
                next_month = "0" + next_month
            next_year = str(dt.datetime.now().year)

        if int(next_year) / 4 == 0:
            end_next_month = str(count_days[12])
        else:
            end_next_month = str(count_days[int(next_month)])

        self.data["customer"] = self.parent.customer
        self.data["company"] = self.parent.company
        self.data["start_date"] = ".".join((next_day, next_month, next_year))
        self.data["end_date"] = ".".join((end_next_month, next_month, next_year))

    def get_worker(self, family):
        rows = self.from_db("family, name, surname, post, passport, "
                            "passport_got, birthday, adr,  live_adr", "workers")
        if family == "all":
            return rows
        for row in rows:
            if family[:-5] == row[0]:  # на форме фамилия в виде Фамилия И.О.
                return row

    # обработчики кнопок
    def ev_ok(self):
        self.data["number"] = "Исх. № " + self.number.text()
        self.data["data"] = "от. " + self.d_note.text()

        self.get_data()
        doc = docxtpl.DocxTemplate(main_file)
        doc.render(self.data)

        doc.save(print_file)
        doc = docx.Document(print_file)
        # Заполнить таблицу
        workers = []
        if self.cb_all.isChecked():
            workers = self.get_worker("all")
        else:
            for elem in self.list_ui:
                if elem.currentText() != "(нет)":
                    workers.append(self.get_worker(elem.currentText()))
        i = 1
        for people in workers:
            doc.tables[1].add_row()
            doc.tables[1].rows[i].cells[0].text = str(i)
            doc.tables[1].rows[i].cells[1].text = " ".join(people[0:2])
            doc.tables[1].rows[i].cells[2].text = people[3]
            doc.tables[1].rows[i].cells[3].text = people[6]
            doc.tables[1].rows[i].cells[4].text = " ".join(people[4:6])
            doc.tables[1].rows[i].cells[5].text = people[7]
            doc.tables[1].rows[i].cells[6].text = people[8]
            i += 1
        doc.save(print_file)
        self.close()
        os.startfile(print_file)

    def new_worker(self):
        flag = True
        for item in self.list_ui:
            if item.currentText() != "(нет)":
                item.setEnabled(True)
            else:
                item.setEnabled(flag)
                flag = False
        pass

    def ev_cancel(self):
        self.close()

    def save_pattern(self):
        return True

    def my_open_file(self):
        print("open file")
        pass

    def kill_pattern(self):
        pass





