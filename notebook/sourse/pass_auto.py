from PyQt5 import uic
from PyQt5.QtCore import Qt
import os
import docxtpl
from pass_template import TempPass
from database import *


class AutoPass(TempPass):
    def __init__(self, parent):
        self.status_ = True
        self.conf = Ini(self)
        ui_file = self.conf.get_ui("pass_auto")
        if not ui_file or ui_file == ERR:
            self.status_ = False
            return
        super(AutoPass, self).__init__(ui_file, parent, "auto")
        uic.loadUi(ui_file, self)
        # my_pass
        self.b_open.clicked.connect(self.my_open_file)
        self.b_clean.clicked.connect(self.clean_data)
        self.add_auto.clicked.connect(self.next_auto)
        self.cb_drivers.activated[str].connect(self.driver_changed)
        self.cb_activ.stateChanged.connect(self.manual_set)

        self.data = {"number": "", "date": "", "start_date": "", "end_date": "",
                     "auto": list(), "gov_numbers": list(), "people": list(list())}

        if self.init_auto() == ERR:
            self.status_ = False
            return
        if self.init_drivers() == ERR:
            self.status_ = False
            return
        self.list_ui = list([self.driver_1, self.driver_2, self.driver_3, self.driver_4,
                             self.driver_5, self.driver_6, self.driver_7])
        self.count = 0
        path_1 = self.conf.get_path("pat_notes")
        path_2 =  self.conf.get_path("notes_docs")
        self.main_file = path_1 + "/pass_auto.docx"
        self.print_file = path_2

    # инициализация
    def init_drivers(self):
        drivers = self.parent.db.get_data("family, name", self.table)
        if drivers == ERR:
            return ERR
        for item in self.list_ui:
            item.addItem(NOT)
        for row in drivers:
            for item in self.list_ui:
                item.addItem(" ".join((row[0], row[1][0] + ".")))
                item.activated[str].connect(self.new_driver)

    def init_auto(self):
        auto = self.parent.db.get_data("model, gov_number", "auto")
        if auto == ERR:
            return ERR
        auto.append([NOT])
        for row in auto:
            self.cb_auto.addItem(row[0])

    # для заполнения текста
    def get_data(self):
        rows = self.parent.db.get_data("*", "auto")
        if rows == ERR or not rows:
            return ERR
        for row in rows:
            if self.cb_auto.currentText() in row:
                self.data["auto"].append(" ".join(row[:2]))
                self.data["gov_number"].append(row[2])
        rows = self.parent.db.get_data("family, name, surname, birthday, passport", "drivers")
        if rows == ERR or not rows:
            return ERR
        for row in rows:
            for item in self.list_ui:
                if item.currentText()[:-3] in row:
                    self.data["people"][self.count].append(" ".join(row))

        self.data["number"] = "Исх. № " + self.number.text()
        self.data["date"] = "от. " + self.d_note.text()
        if "" in self.data:
            return False
        return True

    # обработчики кнопок
    def ev_ok(self):
        data = self.get_data()
        if data == ERR or not data:
            return
        try:
            path = self.main_file
            doc = docxtpl.DocxTemplate(path)
            doc.render(self.data)
            path = self.print_file
            doc.save(path)
            os.startfile(path)
        except:
            return msg_er(self, GET_FILE + path)

    def _set_enabled(self, status):
        self.d_note.setEnabled(status)
        self.number.setEnabled(status)
        self.cb_month.setEnabled(status)
        self.d_from.setEnabled(status)
        self.d_to.setEnabled(status)

    def next_auto(self):
        self.get_date()
        self.my_setEnabled(False)
        self.list_auto.append(self.cb_auto.currentText())

    def kill_last_auto(self):
        self.date["auto"].pop()
        self.data["people"].pop()
        self.list_auto.pop()
        if len(self.date["auto"]) == 1:
            self._set_enabled(True)

    def clean_data(self):
        self.data = {"number": "", "date": "", "start_date": "", "end_date": "",
                     "auto": list(), "gov_numbers": list(), "people": list(list())}
        self.list_auto.clear()

    def ev_cancel(self):
        self.close()

    def my_open_file(self):
        pass

    def manual_set(self, state):
        if state == Qt.Checked:
            self.cb_mounth.setEnabled(False)
            self.d_from.setEnabled(True)
            self.d_to.setEnabled(True)
        else:
            self.cb_mounth.setEnabled(True)
            self.d_from.setEnabled(False)
            self.d_to.setEnabled(False)

    def get_dates(self):
        if not self.cb_chouse.isChecked():
            month = self.list_month.index(self.cb_month.currentText()) + 1
            if month == 13:
                new_year_day = self.conf.get_config("new_year")
                if new_year_day == ERR or not new_year_day:
                    return ERR
                day, month, year = new_year_day, "01", str(dt.datetime.now().year + 1)  # работаем с 9 января
            else:
                day, month, year = "01", str(month), str(dt.datetime.now().year)
                if int(month) < 10:
                    month = "0" + month
            end_month = str(count_days[12]) if int(year) / 4 == 0 else str(count_days[int(month)])
            self.data["start_date"] = ".".join((day, month, year))
            self.data["end_date"] = ".".join((end_month, month, year))
        else:
            self.data["start_date"] = self.d_from.text()
            self.data["end_date"] = self.d_to.text()

    def check_input(self):
        return True
