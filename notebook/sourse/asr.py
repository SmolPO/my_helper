from PyQt5 import uic
from PyQt5.QtWidgets import QDialog
from PyQt5.QtWidgets import QMessageBox as mes
from database import *
import docxtpl
import os
import datetime as dt
import pymorphy2
si = ["тн", "т", "кг", "м2", "м", "м/п", "мм", "м3", "л", "мм", "шт"]


class Asr(QDialog):
    def __init__(self, parent):
        super(Asr, self).__init__()
        self.conf = Ini(self)
        self.ui_file = self.conf.get_ui("asr")
        if self.check_start() == ERR:
            return
        self.parent = parent
        self.contract = parent.contract
        self.year.setCurrentIndex(dt.datetime.now().year-2021)
        self.b_print.clicked.connect(self.ev_print)
        self.b_change.clicked.connect(self.ev_change)
        self.b_kill.clicked.connect(self.ev_kill)
        self.b_close.clicked.connect(self.ev_close)
        self.cb_select.activated[str].connect(self.ev_select)
        self.numbers.textChanged.connect(self.change_count)
        if self.init_bosses() == ERR:
            self.status_ = False
            return
        if self.init_SI() == ERR:
            self.status_ = False
            return
        self.ind = 0
        path_1 = self.conf.get_path("path")
        path_2 = self.conf.get_path("path_pat_patterns")
        if path_1 == ERR or path_2 == ERR:
            self.status_ = False
            return
        self.path = path_1 + path_2
        self.my_id = 0

    def check_start(self):
        self.status_ = True
        try:
            uic.loadUi(self.ui_file, self)
            return True
        except:
            self.status_ = False
            return msg_er(self, GET_UI)

    def change_count(self):
        self.count.setValue(len(self.numbers.toPlainText().split(",")))

    def init_bosses(self):
        if self.parent.parent.db.init_list(self.boss_1, "id, family, name, surname", "itrs", people=True) == ERR:
            return msg_er(self, GET_DB)
        if self.parent.parent.db.init_list(self.boss_2, "id, family, name, surname", "itrs", people=True) == ERR:
            return msg_er(self, GET_DB)
        if self.parent.parent.db.init_list(self.boss_3, "id, family, name, surname", "bosses", people=True) == ERR:
            return msg_er(self, GET_DB)
        if self.parent.parent.db.init_list(self.boss_4, "id, family, name, surname", "bosses", people=True) == ERR:
            return msg_er(self, GET_DB)
        pass

    def init_SI(self):
        for item in si:
            self.cb_SI.addItem(item)

    def create_data(self, day_start, day_end, number):
        if len(day_end) == 1:
            day_end = "0" + day_end
        if len(day_start) == 1:
            day_start = "0" + day_start
        morph = pymorphy2.MorphAnalyzer()
        data = dict()
        bosses = [self.boss_1, self.boss_2, self.boss_3, self.boss_4]
        tables = ["itrs", "itrs", "bosses", "bosses"]
        for ind in range(1, 5):
            boss = bosses[ind-1].currentText()
            post = self.get_post(boss.split(".")[0], tables[ind-1])
            if post == ERR:
                return ERR
            family = boss[boss.index(".")+2:]
            val = post + "__" + family + "_"*(100-len(post + "__" + family))
            data["boss_" + str(ind)] = val
        val = "_______" if self.sb_value.value() == 0 else "__" + str(self.sb_value.value()) + "__"
        data["work"] = self.work.toPlainText() + val + self.cb_SI.currentText()
        materials = self.material.currentText() if self.material.currentText() != "(нет)" else "_"*40
        data["material"] = materials
        data["next_work"] = self.next_work.toPlainText()
        data["day_end"] = day_end
        data["day_start"] = day_start
        data["month"] = morph.parse(self.month.currentText())[0].inflect({'gent'})[0].capitalize().lower()
        data["year"] = self.year.currentText()
        data["number"] = number
        company = self.conf.get_config("company")
        if company == ERR:
            return ERR
        data["company"] = company
        return data

    def get_data(self):
        data = list()
        data.append(self.boss_1.currentText() + self.get_post(self.boss_1.currentText().split(".")[0], "itrs"))
        data.append(self.boss_2.currentText() + self.get_post(self.boss_2.currentText().split(".")[0], "itrs"))
        data.append(self.boss_3.currentText() + self.get_post(self.boss_3.currentText().split(".")[0], "bosses"))
        data.append(self.boss_4.currentText() + self.get_post(self.boss_4.currentText().split(".")[0], "bosses"))
        data.append(self.work.toPlainText() + "_" * 10 + self.cb_SI.currentText())
        data.append(self.material.currentText())
        data.append(self.next_work.text())
        return data

    def get_post(self, my_id, table):
        rows = self.parent.parent.db.get_data("*", table)
        for item in rows:
            if str(item[-1]) == my_id:
                return item[3]
        return "."

    def ev_print(self):
        if not self.check_input():
            return False
        ind = 1
        if self.days_start.toPlainText() == "":
            day_end = "______"
            day_start = "______"
        else:
            day_end = self.days_end.toPlainText().split(",")[ind]
            day_start = self.days_start.toPlainText().split(",")[ind]
        if self.numbers.toPlainText() == "":
            number = "______"
        else:
            number = self.numbers.toPlainText().split(",")[ind]
        self.data = self.create_data(day_start, day_end, number)
        if self.data == ERR:
            return
        path = self.path + ASR_FILE
        try:
            doc = docxtpl.DocxTemplate(path)
        except:
            return msg_er(self, GET_FILE + path)

        doc.render(self.data)
        path_1 = self.conf.get_path("path")
        path_2 = self.conf.get_path("path_contracts")
        path = path_1 + path_2 + "/1030/102021" + "/1.docx"
        try:
            doc.save(path)
            os.startfile(path)
        except:
            return msg_er(self, GET_FILE + path)

    def save_pattern(self):
        data = list()
        data.append(self.data["work"])
        data.append(self.data["value"])
        data.append(self.data["si"])
        data.append(self.data["material"])
        data.append(self.data["day_start"])
        data.append(self.data["day_end"])
        data.append(self.data["month"])
        data.append(self.data["year"])
        data.append(self.data["boss_1"])
        data.append(self.data["boss_2"])
        data.append(self.data["boss_3"])
        data.append(self.data["boss_4"])
        return data

    def ev_select(self):
        rows = self.parent.parent.db.get_data("*", "asrs")
        if rows == ERR:
            return
        for row in rows:
            if self.cb_select.currentText().split(". ")[0] == row[-1]:
                print(row)
                self.work.clean()
                self.next_work.clean()
                self.work.append(row[0])
                self.material.setText(row[1])
                self.next_work.append(row[2])
                self.cb_SI.setCurrentIndex(si.index(row[3]))
                self.month.setCurrentIndex(row[4])
                self.days.clean()
                self.days.append(row[5])
                self.my_id = row[-1]
            pass

    def ev_change(self):
        data = self.get_data()
        if data == ERR:
            return
        answer = msg_q(self, "Вы действительно хотите изменить запись на " + str(data) + "?")
        if answer == mes.Ok:
            data.append(str(self.my_id))
            if self.parent.db.my_update(data, self.table) == ERR:
                return
            answer = msg_q(self, CHANGED_NOTE)
            self.close()

    def ev_kill(self):
        data = self.get_data()
        if data == ERR:
            return
        answer = msg_q(self, "Вы действительно хотите удалить запись " + str(data) + "?")
        if answer == mes.Ok:
            if self.parent.db.kill_value(self.my_id, self.table) == ERR:
                return
            msg_info(self, KILLD_NOTE)
            self.close()

    def ev_close(self):
        self.close()

    def check_input(self):
        if len(self.work.toPlainText()) < 3:
            msg_info(self, "Укажите вид работ")
            return False
        elif len(self.next_work.toPlainText()) < 3:
            msg_info(self, "Укажите следующие работы")
            return False
        elif self.cb_SI.currentText() == "(нет)":
            msg_info(self, "Укажите единицы измерения")
            return False
        elif self.month.currentText() == "(нет)":
            msg_info(self, "Укажите месяц")
            return False
        elif self.boss_1.currentText() == "(нет)":
            msg_info(self, "Укажите первого босса")
            return False
        elif self.boss_2.currentText() == "(нет)":
            msg_info(self, "Укажите второго босса")
            return False
        elif self.boss_3.currentText() == "(нет)":
            msg_info(self, "Укажите третьего босса")
            return False
        elif self.boss_4.currentText() == "(нет)":
            msg_info(self, "Укажите четвертого босса")
            return False
        elif self.count.value() == 0:
            msg_info(self, "Укажите количество актов")
            return False
        elif len(self.days_start.toPlainText().split(",")) != \
                len(self.days_end.toPlainText().split(",")):
            msg_info(self, "Количесттво дней начала не совпадает с количеством дней окончания")
            return False

        elif self.numbers.toPlainText().split(",") != "" and \
             len(self.days_start.toPlainText().split(",")) != \
             len(self.numbers.toPlainText().split(",")):
            msg_info(self, "Не совпадает кол-во дней и кол-во номеров актов. Проверьте. Дней начала: " +
                         str(len(self.days_start.toPlainText().split(","))) + ", Номеров: " +
                         str(len(self.numbers.toPlainText().split(","))))
            return False
        elif self.days_start.toPlainText().split(",") != "" and \
             len(self.days_start.toPlainText().split(",")) != self.count.value():
            msg_info(self, "Количество бланков не совпадает с количеством дат")
            return False
        return True