from PyQt5.QtCore import Qt
from pass_template import TempPass
from database import *


class DTPass(TempPass):
    def __init__(self, parent):
        self.status_ = True
        self.conf = Ini(self)
        ui_file = self.conf.get_ui("pass_dt")
        self.table = DRIVERS
        super(DTPass, self).__init__(ui_file, parent)
        self.cb_contracts.activated[str].connect(self.change_note)
        self.d_to.setDate(dt.datetime.now().date())
        self.d_from.setDate(dt.datetime.now().date())
        self.data = {"number": "", "date": "", "d_from": "",  "d_to": ""}
        self.contract = ""
        self.my_text = ""
        self.contracts = []
        self.init_contracts()
        self.main_file += "/Ввоз_ДТ.docx"
        self.change_note()

    # инициализация
    def init_contracts(self):
        self.cb_contracts.addItem(NOT)
        contracts = self.parent.db.get_data("number, id", CONTRACTS)
        if contracts == ERR or not contracts:
            return ERR
        for row in contracts:
            self.cb_contracts.addItem(row[0])

    # для заполнения текста
    def _get_data(self):
        if NOT in self.cb_contracts.currentText():
            msg_info(self, "Выберите договор")
            return
        rows = self.parent.db.get_data("number, date, datv, place", CONTRACTS)
        for row in rows:
            if self.cb_contracts.currentText() in row:
                self.data["work"] = row[-2]
                self.data["place"] = row[-1]
                self.data["contract"] = row[0] + " от " + row[1]
                self.data["d_to"] = self.d_to.text()
                self.data["d_from"] = self.d_from.text()

    # обработчики кнопок
    def _ev_ok(self):
        return True

    def change_note(self):
        pass

    def _create_data(self, doc):
        pass

    def check_input(self):
        for key in self.data.keys():
            if self.data[key] == NOT or self.data[key] == "":
                msg_info(self, FULL_ALL)
                return False
        return True