from PyQt5.QtCore import QRegExp as QRE
from PyQt5.QtGui import QRegExpValidator as QREVal
from new_template import TempForm
from database import *


class NewBoss(TempForm):
    def __init__(self, parent):
        self.status_ = True
        self.conf = Ini(self)
        self.db = DataBase(self)
        ui_file = self.conf.get_ui("new_boss")
        self.rows_from_db = self.db.get_data(ALL, BOSSES)
        super(NewBoss, self).__init__(ui_file, parent)
        if not self.status_:
            return
        if self.parent.db.init_list(self.cb_select, ALL, BOSSES, people=True) == ERR:
            self.status_ = False
            return
        self.parent.db.init_list(self.cb_company, "company, id", COMPANY, people=False )
        self.list_ui = list([self.family, self.name, self.surname, self.post,
                             self.email, self.phone, self.cb_sex, self.cb_company])

    def init_mask(self):
        symbols = QREVal(QRE("[а-яА-Я]{30}"))

        for item in self.list_ui[0:4]:
            item.setValidator(symbols)
        self.phone.setValidator(QREVal(QRE("[0-9]{11}")))
        self.email.QREVal(QRE("[a-zA-Z ._@ 0-9]{30}"))

    def _set_data(self, data):
        self.cb_sex.setCurrentIndex(0) if data[6] == "М" else self.cb_sex.setCurrentIndex(1)
        self.cb_company.setCurrentIndex(get_index((self.cb_company, data[-2])))

    def _select(self, data):
        return True