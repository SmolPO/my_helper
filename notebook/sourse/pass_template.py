from PyQt5 import uic
from PyQt5.QtWidgets import QDialog
from PyQt5.QtCore import Qt
import os
import docxtpl
from database import *


class TempPass(QDialog):
    def __init__(self, ui_file, parent):
        self.status_ = True
        self.conf = Ini(self)
        self.db = DataBase(self)
        super(TempPass, self).__init__()
        self.parent = parent
        if not self.check_start(ui_file):
            return
        self.db = DataBase(self)
        # my_pass
        self.b_ok.clicked.connect(self.ev_ok)
        self.b_cancel.clicked.connect(self.close)
        try:
            self.b_open.clicked.connect(self.my_open_file)
        except:
            pass
        self.d_note.setDate(dt.datetime.now().date())
        self.number.setValue(self.conf.get_next_number())

        self.list_month = ["январь", "февраль", "март", "апрель",
                           "май", "июнь", "июль", "август", "сентябрь",
                           "октябрь", "ноябрь", "декабрь"]
        self.data = dict()
        self.main_file = self.conf.get_path("pat_notes")
        self.print_folder = self.conf.get_path("notes_docs")

        fields = "family, name, surname, post, passport, passport_got, birthday, adr,  live_adr"
        self.people_all = self.parent.db.get_data(fields, "workers") + self.parent.db.get_data(fields, "itrs")
        if self.people_all == ERR:
            return
        self.people_mark = list()
        fields = "family, name, surname, post, passport, passport_got, " \
                 "birthday, adr,  live_adr, d_vac_1, d_vac_2, place, vac_doc, vac_type, status, id"
        rows = [self.parent.db.get_data(fields, "workers"), self.parent.db.get_data(fields, "itrs")]
        if ERR in rows:
            return
        self.all_people = rows[0] + rows[1]
        self.vac = False

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
        try:
            print_file = self.print_folder + "/" + self.number.text() + "_" + self.d_note.text() + ".docx"
            path = self.main_file
            doc = docxtpl.DocxTemplate(path)
            doc.render(self.data)
            path = print_file
            doc.save(path)
            if self.vac:
                os.startfile(path)
        except:
            return msg_er(self, GET_FILE + path)
        if self._create_data(path) == ERR:
            return
        if self.conf.set_next_number(self.conf.get_next_number() + 1) == ERR:
            return
        os.startfile(path)
        self.close()

    def get_cb_list(self, cb):
        return []

    def clean_cb(self, cb):
        cb.clear()
        pass

    def compare_family(self, family_short, family_full):
        if family_short == family_full:
            return True
        pass

    def add_item(self, people, ind):
        for item in self.list_ui:
            if self.list_ui.index(item) == ind:
                continue
            man = [item.itemText(i) for i in range(item.count())]
            if not people == NOT:
                man.append(people)
            man.sort()
            i = man.index(people)
            item.insertItem(i, people)
            print([item.itemText(i) for i in range(item.count())])

    def kill_item(self, people, ind):
        for item in self.list_ui:
            if self.list_ui.index(item) == ind:
                continue
            for i in range(item.count()):
                if item.itemText(i) == people:
                    item.removeItem(i)
                    break

    def new_worker(self, some, man=None, indx=None):
        if man:
            val = man
            ind = indx
        else:
            val = self.sender().currentText()
            ind = self.list_ui.index(self.sender())
        flag = True
        if val == NOT:
            if not self.list_cb[ind] == NOT:
                self.add_item(self.list_cb[ind], ind)
                self.list_cb[ind] = NOT
        else:
            self.kill_item(val, ind)
            if self.list_cb[ind] != NOT:
                self.add_item(self.list_cb[ind], ind)
            self.list_cb[ind] = val

        for item in self.list_ui:
            if item.currentText() != NOT:
                item.setEnabled(True)
            else:
                item.setEnabled(flag)
                flag = False

    def get_contract(self, name):
        # получить номер договора по короткому имени
        rows = self.parent.db.get_data("number, date, datv, name", CONTRACTS)
        if rows == ERR:
            return ERR
        for row in rows:
            if name in row:
                self.data["contract"] = " от ".join(row[:2])
                self.data["work"] = row[2]

    def get_worker(self, family):
        rows = self.parent.db.get_data("family, name, surname, post, passport, "
                                       "passport_got, birthday, adr,  live_adr, id", WORKERS) + \
               self.parent.db.get_data("family, name, surname, post, passport, "
                                       "passport_got, birthday, adr,  live_adr, id", ITRS)
        if rows == ERR:
            return ERR
        if family == "all":
            return rows
        for row in rows:
            if family[:-5] in row:  # на форме фамилия в виде Фамилия И.О.
                return row

    def get_worker_week(self, family):
        # получить номер договора по короткому имени
        rows = self.parent.db.get_data("family, name, surname, post, passport, passport_got, " 
                                       "birthday, adr,  live_adr", WORKERS)
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
            self.cb_to.setEnabled(True)
        else:
            self.cb_sun.setEnabled(True)
            self.cb_sub.setEnabled(True)
            self.d_from.setEnabled(False)
            self.d_to.setEnabled(False)
            self.cb_to.setEnabled(False)

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
        family = row.split(".")[0][:-2]
        s_name = row.split(".")[0][-1]
        s_sur = row.split(".")[1]
        for item in self.all_people:
            if family == item[0]:
                if s_name == item[1][0]:
                    if s_sur == item[2][0]:
                        return item