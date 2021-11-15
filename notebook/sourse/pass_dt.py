from PyQt5.QtCore import Qt
from pass_template import TempPass
from database import *


class DTPass(TempPass):
    def __init__(self, parent):
        self.status_ = True
        self.conf = Ini(self)
        ui_file = self.conf.get_ui("pass_dt")
        super(DTPass, self).__init__(ui_file, parent, "drivers")
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
        contracts = self.parent.db.get_data("number, id", "contracts")
        if contracts == ERR or not contracts:
            return ERR
        for row in contracts:
            self.cb_contracts.addItem(row[0])

    # для заполнения текста
    def _get_data(self):
        if NOT in self.cb_contracts.currentText():
            msg_info(self, "Выберите договор")
            return
        rows = self.parent.db.get_data("number, date, object, type_work, place", "contracts")
        contract = ""
        place = ""
        work = ""
        for row in rows:
            if self.cb_contracts.currentText() in row:
                contract = row
                work = row[-2]
                place = row[-1]
                break
        if "цех" in contract[-1].lower():
            place = contract[-1].replace("цех", "цеха")
        if "ремонт" in contract[-2].lower():
            work = contract[-1].replace("ремонт", "по ремонту ")
            work += 'на объекте: "' + contract[-2] + '" '
        self.data["work"] = work + " " + place
        self.data["contract"] = contract[0] + " от " + contract[1]
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