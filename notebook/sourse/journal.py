from PyQt5 import uic
from PyQt5.QtWidgets import QDialog
import docxtpl
from database import *


class Journal(QDialog):
    def __init__(self, parent):
        super(Journal, self).__init__()
        self.conf = Ini(self)
        self.ui_file = self.conf.get_ui("journal")
        if self.ui_file == ERR:
            self.status_ = False
            return
        if not self.check_start():
            return
        self.parent = parent
        self.path = parent.path + self.conf.get_path("docs") + JOURNAL_FILE
        self.b_print.clicked.connect(self.ev_print)
        self.data = dict()
        if self.init_bosses() == ERR:
            return
        self.name.append("По ")

    def check_start(self):
        self.status_ = True
        self.path_ = self.conf.get_ui("journal")
        if self.path_ == ERR:
            return
        try:
            uic.loadUi(self.ui_file, self)
            return True
        except:
            msg_info(self, "Не удалось открыть форму " + self.ui_file)
            self.status_ = False
            return False

    def init_bosses(self):
        if self.parent.parent.db.init_list(self.boss_1, "*", "itrs", people=True) == ERR:
            self.status_ = False
            return False
        if self.parent.parent.db.init_list(self.boss_2, "*", "itrs", people=True) == ERR:
            self.status_ = False
            return False
        if self.parent.parent.db.init_list(self.boss_3, "*", "bosses", people=True) == ERR:
            self.status_ = False
            return False
        if self.parent.parent.db.init_list(self.boss_4, "*", "bosses", people=True) == ERR:
            self.status_ = False
            return False
        pass

    def get_data(self):
        data = dict()
        data["name"] = self.name.toPlainText()
        data["numb"] = str(self.number.value())
        data["boss_1"] = "".join(self.boss_1.currentText().split(". ")[1:])
        data["boss_2"] = "".join(self.boss_2.currentText().split(". ")[1:])
        data["boss_3"] = "".join(self.boss_3.currentText().split(". ")[1:])
        data["boss_4"] = "".join(self.boss_4.currentText().split(". ")[1:])
        data["post_1"] = self.get_post(self.boss_1.currentText().split(".")[0], "itrs")
        data["post_2"] = self.get_post(self.boss_2.currentText().split(".")[0], "itrs")
        data["post_3"] = self.get_post(self.boss_3.currentText().split(".")[0], "bosses")
        data["post_4"] = self.get_post(self.boss_4.currentText().split(".")[0], "bosses")
        data["date"] = self.date.text()
        data["year"] = str(dt.datetime.now().year)
        data["company"] = self.parent.parent.company
        data["customer"] = self.parent.parent.customer
        return data

    def get_post(self, my_id, table):
        rows = self.parent.parent.db.get_data("*", table)
        if rows == ERR:
            return ERR
        for item in rows:
            if str(item[-1]) == my_id:
                return item[3]
        return "."

    def ev_print(self):
        self.data = self.get_data()
        if not self.check_input():
            return False
        path = self.conf.get_path("path") + self.conf.get_path("pat_patterns") + "/Журнал.docx"
        try:
            doc = docxtpl.DocxTemplate(path)
        except:
            msg_er(self, GET_FILE + path)
            return False
        doc.render(self.data)
        try:
            doc.save(self.parent.path + JOURNAL_FILE)
            msg_info(self, "Журнал создан")
            if not save_open(path):
                msg_er(self, GET_FILE + path)
                return False
            self.close()
        except:
            msg_er(self, GET_FILE + self.parent.path + JOURNAL_FILE)
            return False
        self.close()

    def check_input(self):
        if self.number.value() == 0:
            msg_info(self, "Укажите номер журнала")
            return False
        if self.date.text() == "01.01.2000":
            msg_info(self, "Укажите дату начала работ")
            return False
        if self.boss_1.currentText() == "(нет)":
            msg_info(self, "Укажите первого босса")
            return False
        if self.boss_2.currentText() == "(нет)":
            msg_info(self, "Укажите второго босса")
            return False
        if self.boss_3.currentText() == "(нет)":
            msg_info(self, "Укажите третьего босса")
            return False
        if self.boss_4.currentText() == "(нет)":
            msg_info(self, "Укажите четвертого босса")
            return False
        if self.name.toPlainText() == "" or len(self.name.toPlainText()) < 3:
            msg_info(self, "Укажите название журнала")
            return False
        return True