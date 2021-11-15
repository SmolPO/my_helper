from PyQt5.QtCore import Qt
from pass_template import TempPass
from database import *


class DrivePass(TempPass):
    def __init__(self, parent):
        self.status_ = True
        self.conf = Ini(self)
        ui_file = self.conf.get_ui("pass_driver")
        if not ui_file or ui_file == ERR:
            self.status_ = False
            return
        super(DrivePass, self).__init__(ui_file, parent, "drivers")
        if not self.status_:
            return
        self.cb_contracts.activated[str].connect(self.contract_changed)
        self.cb_auto.activated[str].connect(self.auto_changed)
        self.cb_drivers.activated[str].connect(self.driver_changed)

        self.cb_activ.stateChanged.connect(self.manual_set)
        self.d_arrive.dateChanged.connect(self.date_changed)
        self.input_text.textChanged.connect(self.change_note)
        self.input_text.textChanged.connect(self.change_note)

        self.data = {"number": "", "date": "", "text_note": "", "auto": "",
                     "gov_number": "", "track_number": " ", "passport": "", "adr": ""}
        self.work = ""
        self.data_arrive = ""
        self.contract = ""
        self.my_text = ""
        self.contracts = [[]]
        if self.init_auto() == ERR:
            self.status_ = False
            return
        if self.init_drivers() == ERR:
            self.status_ = False
            return
        if self.init_contracts() == ERR:
            self.status_ = False
            return
        self.d_arrive.setDate(dt.datetime.now().date())
        self.main_file += "/pass_drive.docx"

    # инициализация
    def init_drivers(self):
        self.cb_drivers.addItem(NOT)
        people = self.parent.db.get_data("family, name", self.table)
        if people == ERR or not people:
            return ERR
        for row in people:
            self.cb_drivers.addItem(" ".join((row[0], row[1][0] + ".")))

    def init_auto(self):
        self.cb_auto.addItem(NOT)
        auto = self.parent.db.get_data("gov_number", "auto")
        if auto == ERR or not auto:
            return ERR
        for row in auto:
            self.cb_auto.addItem(row[0])

    def init_contracts(self):
        self.cb_contracts.addItem(NOT)
        contracts = self.parent.db.get_data("name", "contracts")
        if contracts == ERR or not contracts:
            return ERR
        for row in contracts:
            self.cb_contracts.addItem(row[0])

    # для заполнения текста
    def _get_data(self):
        self.data["text_note"] = self.note.toPlainText()
        self.data["contract"] = self.cb_contracts.currentText()
        self.data["d_arrive"] = self.d_arrive.text()

    # обработчики кнопок
    def _ev_ok(self):
        return True

    def date_changed(self):
        self.data_arrive = self.d_arrive.text()
        self.change_note()

    def contract_changed(self):
        rows = self.parent.db.get_data("*", "contracts")
        if rows == ERR or not rows:
            return ERR
        for row in rows:
            if self.cb_contracts.currentText() == row[0]:
                self.work = " ".join(row[4:7])
                self.contract = " от ".join(row[2:4])
            if self.cb_contracts.currentText() == NOT:
                self.work = ""
                self.contract = ""
        self.change_note()

    def change_note(self):
        text = list()
        text.append("Прошу Вашего разрешения оформить разовый электронный пропуск на ")
        text.append(self.data_arrive)
        text.append("с 08-00 до 17-00 для проезда на территорию")
        text.append(self.parent.customer)
        text.append(" и ")
        text.append(self.input_text.toPlainText())
        text.append(self.work)
        text.append(self.contract)
        self.note.setText(" ".join(text))

    def manual_set(self, state):
        if state == Qt.Checked:
            self.note.setEnabled(True)
            self.input_text.setEnabled(False)
            self.d_arrive.setEnabled(False)
            self.cb_contracts.setEnabled(False)
        else:
            self.note.setEnabled(False)
            self.input_text.setEnabled(True)
            self.d_arrive.setEnabled(True)
            self.cb_contracts.setEnabled(True)

    def auto_changed(self):
        rows = self.parent.db.get_data("*", "auto")
        if rows == ERR or not rows:
            return ERR
        for row in rows:
            if self.cb_auto.currentText() == row[0]:
                self.data["auto"] = " ".join(row[1:3])
                self.data["gov_number"] = row[0]
                self.data["track"] = " " if row[-2] == NOT else "п/п " + row[-2]

    def driver_changed(self):
        people = self.parent.db.get_data("*", self.table)
        if people == ERR or not people:
            return ERR
        for row in people:
            if self.cb_drivers.currentText()[:-3] == row[0]:
                self.data["passport"] = " ".join(row[:-1])
                self.data["adr"] = row[-2]

    def _create_data(self, doc):
        pass

    def check_input(self):
        for key in self.data.keys():
            if self.data[key] == NOT or self.data[key] == "":
                msg_info(self, FULL_ALL)
                return False
        return True