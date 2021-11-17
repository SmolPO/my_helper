from PyQt5.QtWidgets import QMessageBox
import docxtpl
import inserts as ins
from my_email import *
from database import *


class GetMoney(QDialog):
    def __init__(self, parent):
        self.status_ = True
        self.conf = Ini(self)
        ui_file = self.conf.get_ui("get_money")
        if not ui_file or ui_file == ERR:
            self.status_ = False
            return
        super(GetMoney, self).__init__()
        uic.loadUi(ui_file, self)
        self.parent = parent
        self.table = "finance"
        self.b_ok.clicked.connect(self.ev_ok)
        self.b_cancel.clicked.connect(self.ev_cancel)
        self.b_kill.clicked.connect(self.ev_kill)
        self.b_change.clicked.connect(self.ev_change)

        self.cb_recipient.activated[str].connect(self.change_note)
        self.cb_select.activated[str].connect(self.ev_select)
        self.cb_manual_set.setEnabled(False)
        self.cb_day.stateChanged.connect(self.day_money)

        self.note.textChanged.connect(self.change_note)
        # self.sb_value.valueChanged[int].connect(self.change_note)
        self.sb_days.valueChanged[int].connect(self.change_note)
        self.sb_emploeeyrs.valueChanged[int].connect(self.change_note)
        self.sb_some_value.valueChanged[int].connect(self.change_note)
        self.sb_cost.valueChanged[int].connect(self.change_note)
        self.cb_some.stateChanged.connect(self.change_note)
        self.cb_day.stateChanged.connect(self.change_note)
        self.date.setDate(dt.datetime.now().date())

        self.rows_from_db = self.parent.db.get_data("*", self.table)
        rows = self.parent.db.get_data("id, date", self.table)
        notes = []
        for row in rows:
            tmp = []
            for item in row:
                tmp.append(item)
            notes.append(tmp)
        notes.sort(key=lambda x: x[0])
        for row in notes:
            self.cb_select.addItems([". ".join((row[0], row[1]))])
        if self.parent.db.init_list(self.cb_recipient, "*", "itrs", people=True) == ERR:
            return
        if self.parent.db.init_list(self.cb_customer, "*", "itrs", people=True) == ERR:
            return
        self.next_id = self.parent.db.get_next_id(self.table)
        if self.next_id == ERR:
            return
        self.current_id = self.next_id
        self.my_id.setValue(self.next_id)
        self.data = {"date": "", "post": "", "family": "", "text": ""}
        self.change_note()

        self.main_file = self.conf.get_path("pat_notes") + \
                         self.conf.get_from_ini("get_money", "patterns")
        self.print_folder = self.conf.get_path("bills") + \
                            "/" + str(dt.datetime.now().year) + "/" + \
                            "/Заявки на деньги"

    def ev_ok(self):
        if not self.check_input():
            return False
        data = self.get_data()
        if not data:
            return
        # morph = pymorphy2.MorphAnalyzer()
        if self.parent.db.my_commit(ins.add_to_db(data, self.table)) == ERR:
            return msg_er(self, ADD_DB)
        rows = self.parent.db.get_data("post, family, name, surname, id", "itrs")
        if rows == ERR:
            return
        for row in rows:
            if self.cb_customer.currentText().split(".")[0] == str(row[-1]):
                self.data["post_g"] = dictionary[row[0]]["gent"]
                self.data["post_i"] = row[0]
                people = self.cb_customer.currentText()
                family = people.split(". ")[1][:-5]
                # fam = morph.parse(family)[0].inflect({'gent'})[0].capitalize()
                self.data["family_g"] = family + " " + people[-4:]
                self.data["text"] = self.note_result.toPlainText()
                self.data["date"] = self.date.text()
                self.data["family_i"] = "".join(people.split(". ")[1:])
        print_file = self.print_folder + "/" + str(dt.datetime.now().date()) + "_" + \
                                               str(self.sb_value.value()) + ".docx"
        if not self.data:
            return False
        try:
            doc = docxtpl.DocxTemplate(self.main_file)
        except:
            msg_er(self, GET_FILE + self.main_file)
            return
        doc.render(self.data)
        try:
            doc.save(print_file)
        except:
            msg_er(self, GET_FILE + print_file)
            return
        self.close()
        if not save_open(print_file):
            msg_er(self, GET_FILE + print_file)
            return
        mes.question(self, "Сообщение", "Запись добавлена", mes.Ok)
        answer = mes.question(self, "Сообщение", "Отправить бухгалтеру?", mes.Ok | mes.Cancel)
        if answer == mes.Ok:
            wnd = SendPost(self.parent.db, print_file)
            if not wnd.status_:
                return
            wnd.exec_()
        self.close()

    def ev_cancel(self):
        self.close()

    def ev_select(self, text):
        if text == NOT:
            self.clean_data()
            self.but_status("add")
            return
        else:
            self.but_status("change")
        for row in self.rows_from_db:
            if text.split(",")[0] == row[0]:
                self.set_data(row)

    def change_note(self, state=None):
        self.sb_some_value.setEnabled(True) if self.cb_some.isChecked() else self.sb_some_value.setEnabled(False)
        self.day_money(self.cb_day.isChecked())
        itr = ""
        # morph = pymorphy2.MorphAnalyzer()
        people = self.parent.db.get_data("post, family, name, surname, id", "itrs")
        if people == ERR:
            return ERR
        for boss in people:
            if self.cb_recipient.currentText().split(".")[0] == str(boss[-1]):
                itr = boss
                break
        text = list()
        if not self.note.toPlainText():
            self.sb_value.setValue(self.sb_days.value() * self.sb_emploeeyrs.value() * self.sb_cost.value() +
                                   self.sb_some_value.value())
        text.append("Прошу Вас выслать ")
        text.append(str(self.sb_value.value()))
        text.append("р. на банковскую карту ")
        for item in itr[:4]:
            post = item
            if post in dictionary.keys():
                word = dictionary[post]["gent"]
            else:
                word = post
            #   word = morph.parse(item)[0].inflect({'gent'})[0].capitalize()
            text.append(word)
        text.append(" для:\n")
        if self.cb_day.isChecked():
            cost = self.sb_days.value() * self.sb_emploeeyrs.value() * self.sb_cost.value()
            text.append("- {0}р. суточные из расчета {1} дней/я {2} чел. {3}р. ставка\n".format(cost,
                                                                                           self.sb_days.value(),
                                                                                           self.sb_emploeeyrs.value(),
                                                                                           self.sb_cost.value()))
        if self.cb_some.isChecked():
            text.append("- {0}р. на производственные нужды\n".format(self.sb_some_value.value()))
        if self.note.toPlainText():
            text.append(" - " + self.note.toPlainText())
        self.note_result.setText(" ".join(text))

    def day_money(self, state):
        list_ui = [self.sb_days, self.sb_emploeeyrs, self.sb_cost, self.l_1, self.l_2, self.l_3]
        tmp = [item.setEnabled(state) for item in list_ui]

    def ev_change(self):
        answer = mes.question(self, "Изменение записи", "Вы действительно хотите изменить запись на " +
                              str(self.get_data()) + "?", mes.Ok | mes.Cancel)
        if answer == mes.Ok:
            data = self.get_data()
            data[-1] = str(self.current_id)
            if self.parent.db.my_update(data, self.table) == ERR:
                return msg_er(self, UPDATE_DB)
            answer = mes.question(self, "Сообщение", "Запись изменена", mes.Ok)
            if answer == mes.Ok:
                self.close()

    def ev_kill(self):
        answer = mes.question(self, "Удаление записи", "Вы действительно хотите удалить запись " +
                              str(self.get_data()) + "?", mes.Ok | mes.Cancel)
        if answer == mes.Ok:
            if self.parent.db.kill_value(self.current_id, self.table) == ERR:
                return msg_er(self, KILLD_NOTE)
            answer = mes.question(self, "Сообщение", "Запись удалена", mes.Ok)
            if answer == mes.Ok:
                self.close()

    def set_data(self, data):
        self.my_id.setValue(data[0])
        self.date.setDate(Date.fromString(data[1], "dd.mm.yyyy"))
        self.sb_value.Value(int(data[2]))
        i = range(len(self.rows_from_db))
        for row in self.rows_from_db:
            self.cb_recipient.setCurrentIndex(next(i)) if data[3] in row else next(i)
        self.note.clear()
        self.note.append(data[4])

    def get_data(self):
        cost = self.sb_days.value() * self.sb_emploeeyrs.value() * self.sb_cost.value()
        if cost + self.sb_some_value.value() > self.sb_value.value():
            QMessageBox.question(self, "Внимание",
                                 "Сумма в итоге меньше, чем сумма пунктов. Вы хотите за своих оплачивать?))",
                                 QMessageBox.Ok | QMessageBox.Cancel)
            return
        if not self.note.toPlainText():
            if cost + self.sb_some_value.value() != self.sb_value.value():
                QMessageBox.question(self, "Внимание", "Не сходится сумма, напишите куда именно вы потратите разницу",
                                     QMessageBox.Ok)
                return
        data = list()
        data.append(self.date.text())
        data.append(str(self.sb_value.value()))
        data.append(self.cb_customer.currentText().split(".")[1] + "." +
                    self.cb_customer.currentText().split(".")[2] + ".")

        if self.cb_day.isChecked():
            data.append("суточные {0} чел {1} дней {2}р. ставка".format(self.sb_days.value(),
                                                                        self.sb_emploeeyrs.value(),
                                                                        self.sb_cost.value()))
        if self.cb_some.isChecked():
            if len(data) == 2:
                data.append("производственные нужды " + self.sb_some_value.value())
            else:
                data[3] = data[3] + "| производственные нужды " + str(self.sb_some_value.value())
        if self.note.toPlainText():
            if len(data) == 2:
                data.append(self.note.toPlainText())
            else:
                data[3] = data[3] + "| " + self.note.toPlainText()
        data.append(str(self.my_id.value()))
        return data

    def check_input(self):
        if self.sb_value.value() == 0:
            mes.question(self, "Сообщение", "Укажите общую сумму", mes.Cancel)
            return False
        if self.cb_recipient.currentText() == "(нет)":
            mes.question(self, "Сообщение", "Укажите получателя перевода", mes.Cancel)
            return False
        if self.cb_day.isChecked():
            cost = self.sb_days.value() * self.sb_emploeeyrs.value() * self.sb_cost.value()
            if cost == 0:
                mes.question(self, "Сообщение", "Укажите значения в суточных или уберите галочку", mes.Cancel)
                return False
        if self.cb_some.isChecked():
            if self.sb_some_value.value() == 0:
                mes.question(self, "Сообщение", "Укажите значения в производственных нуждах или уберите галочку",
                             mes.Cancel)
                return False
        return True

    def clean_data(self):
        self.sb_recipient.setCurrentIndex(0)
        self.sb_value.setValue(0)
        self.note.clear()
        self.cb_day.setCheacked(False)
        self.sb_day.setValue(0)
        self.sb_emploeers.setValue(0)

    def but_status(self, status):
        if status == "add":
            self.b_ok.setEnabled(True)
            self.b_change.setEnabled(False)
            self.b_kill.setEnabled(False)
        if status == "change":
            self.b_ok.setEnabled(False)
            self.b_change.setEnabled(True)
            self.b_kill.setEnabled(True)
