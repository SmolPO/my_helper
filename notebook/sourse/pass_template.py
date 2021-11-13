from PyQt5 import uic
from PyQt5.QtWidgets import QDialog
from PyQt5.QtCore import Qt
import os
import docxtpl
from database import *


class TempPass(QDialog):
    def __init__(self, ui_file, parent, table):
        self.status_ = True
        self.conf = Ini(self)
        super(TempPass, self).__init__()
        self.parent = parent
        if not self.check_start(ui_file):
            return
        self.table = table
        # my_pass
        self.b_ok.clicked.connect(self.ev_ok)
        self.b_cancel.clicked.connect(self.close)
        self.b_open.clicked.connect(self.my_open_file)

        self.d_note.setDate(dt.datetime.now().date())
        self.number.setValue(self.conf.get_next_number())

        self.list_month = ["январь", "февраль", "март", "апрель",
                           "май", "июнь", "июль", "август", "сентябрь",
                           "октябрь", "ноябрь", "декабрь"]
        self.data = dict()
        self.main_file = self.conf.get_path("path_pat_notes")
        self.print_folder = self.conf.get_path("path_notes_docs")

        fields = "family, name, surname, post, passport, passport_got, birthday, adr,  live_adr"
        self.people_all = self.parent.db.get_data(fields, "workers") + self.parent.db.get_data(fields, "itrs")
        if self.people_all == ERR:
            return
        self.people_mark = list()
        fields = "family, name, surname, post, passport, passport_got, " \
                 "birthday, adr,  live_adr, d_vac_1, d_vac_2, place, vac_doc, vac_type, id"
        rows = [self.parent.db.get_data(fields, "workers"), self.parent.db.get_data(fields, "itrs")]
        if ERR in rows:
            return
        self.all_people = rows[0] + rows[1]

    def check_start(self, ui_file):
        self.status_ = True
        self.path_ = ui_file
        try:
            uic.loadUi(ui_file, self)
            return True
        except:
            self.status_ = False
            return msg_er(self, GET_UI + self.path_)

    # флаг на выбор всех
    def set_dates(self, state):
        self.d_to.setEnabled(state == Qt.Checked)
        self.d_from.setEnabled(state == Qt.Checked)
        self.cb_month.setEnabled(state != Qt.Checked)

    # обработчики кнопок
    def ev_ok(self):
        if not self._ev_ok():
            return False
        if self._get_data() == ERR:
            return
        self.data["number"] = "Исх. № " + self.number.text()
        self.data["date"] = "от. " + self.d_note.text()
        self.data["customer"] = self.parent.customer
        self.data["company"] = self.parent.company
        if not self.check_input():
            return False
        print_file = self.print_folder + "/" + self.number.text() + "_" + self.d_note.text() + ".docx"

        path = self.main_file
        try:
            doc = docxtpl.DocxTemplate(path)
            doc.render(self.data)
            path = print_file
            doc.save(path)
        except:
            return msg_er(self, GET_FILE + path)
        if self._create_data(path) == ERR:
            return
        if self.conf.set_next_number(self.conf.get_next_number()) == ERR:
            return
        os.startfile(path)
        self.close()

    def get_cb_list(self, cb):
        return []

    def clean_cb(self, cb):
        cb.clear()
        pass

    def new_worker(self):
        flag = True
        workers = self.parent.db.get_data("*", "workers")
        itrs = self.parent.db.get_data("*", "itrs")
        peoples_ = workers + itrs
        peoples = []
        for man in peoples_:
            if man[-2] == "уволен":
                continue
            else:
                peoples.append(man)
        i = 0
        for item in self.list_ui:
            if item.currentText() == NOT:
                break
            for el in peoples:
                if el[0] in item.currentText().split(". ")[1]:
                    if el[1][0] in item.currentText().split(". ")[1][-4] and \
                            el[2][0] in item.currentText().split(". ")[1][-2]:
                        peoples.pop(i)
                i += 1
        for item in peoples:
            for el in self.list_ui[:i]:
                el.addItem(str(item[-1]) + ". " + short_name(item[:3]))

        for item in self.list_ui:
            if item.currentText() != NOT:
                item.setEnabled(True)
            else:
                item.setEnabled(flag)
                flag = False
        pass

    def my_open_file(self):
        try:
            os.startfile(self.print_file)
        except:
            return msg_er(self, GET_FILE + self.print_file)

    def save_pattern(self):
        pass

    def kill_pattern(self):
        pass

    def get_contract(self, name):
        # получить номер договора по короткому имени
        rows = self.parent.db.get_data("number, date, object, place, type_work, name", "contracts")
        if rows == ERR:
            return ERR
        for row in rows:
            if name in row:
                self.data["contract"] = " от ".join(row[:2])
                self.data["object_name"] = row[2]
                self.data["part"] = row[3]
                self.data["type_work"] = row[4]

    def get_worker(self, family):
        rows = self.parent.db.get_data("family, name, surname, post, passport, "
                                       "passport_got, birthday, adr,  live_adr, id", "workers")
        if rows == ERR:
            return ERR
        if family == "all":
            return rows
        for row in rows:
            if family[:-5] == row[0]:  # на форме фамилия в виде Фамилия И.О.
                return row

    def get_worker_week(self, family):
        # получить номер договора по короткому имени
        rows = self.parent.db.get_data("family, name, surname, post, passport, passport_got, "
                                           "birthday, adr,  live_adr", "workers")
        if rows == ERR:
            return ERR
        for row in rows:
            if family[:-5] == row[0]:
                return row
            return row

    def get_week_days(self):
        if self.cb_other.isChecked():
            self.data["week_day"] = "в выходные дни с " + self.d_from.text() + " до " + self.d_to.text()
        else:
            if len(self.get_days()) > 1:
                self.data["week_day"] = "в выходные дни с " + str(self.get_days()[0]) + " до " + str(self.get_days()[1])
            else:
                self.data["week_day"] = "в выходной день " + str(self.get_days()[0])

    def other_days(self, state):
        if state == Qt.Checked:
            self.cb_sun.setEnabled(False)
            self.cb_sub.setEnabled(False)
            self.d_from.setEnabled(True)
            self.d_to.setEnabled(True)
        else:
            self.cb_sun.setEnabled(True)
            self.cb_sub.setEnabled(True)
            self.d_from.setEnabled(False)
            self.d_to.setEnabled(False)

    def week_days(self, state):
        if state == Qt.Checked:
            self.cb_other.setEnabled(False)
            self.d_from.setEnabled(False)
            self.d_to.setEnabled(False)
        elif not self.cb_sun.isChecked() and not self.cb_sub.isChecked():
            self.cb_other.setEnabled(True)
        self.set_date()

    def set_date(self):
        data = []
        now_weekday = dt.datetime.now().weekday()
        if self.cb_sun.isChecked() and not self.cb_sub.isChecked():
            sub_day = dt.datetime.now() + dt.timedelta(6 - now_weekday)
            date = ".".join((str(sub_day.day), str(sub_day.month), str(sub_day.year)))
            self.d_from.setDate(from_str(date))
        if self.cb_sub.isChecked() and not self.cb_sun.isChecked():
            sub_day = dt.datetime.now() + dt.timedelta(5 - now_weekday)
            date = ".".join((str(sub_day.day), str(sub_day.month), str(sub_day.year)))
            self.d_from.setDate(from_str(date))
        if self.cb_sub.isChecked() and self.cb_sun.isChecked():
            sub_day = dt.datetime.now() + dt.timedelta(5 - now_weekday)
            date_sub = ".".join((str(sub_day.day), str(sub_day.month), str(sub_day.year)))
            sun_day = dt.datetime.now() + dt.timedelta(6 - now_weekday)
            date_sun = ".".join((str(sun_day.day), str(sun_day.month), str(sun_day.year)))
            self.d_from.setDate(from_str(date_sub))
            self.d_to.setDate(from_str(date_sun))
        return data

    def get_end_month(self):
        next_month = dt.datetime.now().month
        next_day = count_days[next_month]
        next_month = str(next_month)
        if int(next_month) < 10:
            next_month = "0" + next_month
        next_year = str(dt.datetime.now().year)

        if int(next_year) / 4 == 0:
            end_next_month = str(count_days[12])
        else:
            end_next_month = str(count_days[int(next_month)])
        return (end_next_month, next_month, next_year)

    def check_row(self, row):
        my_id = int(row.split(".")[0])
        family = row.split(".")[1][1:-2]
        s_name = row.split(".")[1][-1]
        s_sur = row.split(".")[2]
        for item in self.all_people:
            if str(my_id) == str(item[-1]):
                if family == item[0]:
                    if s_name == item[1][0]:
                        if s_sur == item[2][0]:
                            return item