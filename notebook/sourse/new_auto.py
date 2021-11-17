from PyQt5.QtCore import QRegExp as QRE
from PyQt5.QtGui import QRegExpValidator as QREVal
from PyQt5.QtCore import Qt
from new_template import TempForm
from database import *
fields = ["model", "brand", "gov_number", "track_number", "id"]


class NewAuto(TempForm):
    def __init__(self, parent=None):
        self.conf = Ini(self)
        self.db = DataBase(self)
        ui_file = self.conf.get_ui("new_auto")
        self.rows_from_db = self.db.get_data(ALL, AUTO)
        super(NewAuto, self).__init__(ui_file, parent)
        if not self.status_:
            return
        self.is_track.stateChanged.connect(self.have_track)
        self.init_list()
        self.track_number.setEnabled(False)
        self.list_ui = [self.gov_number, self.brand, self.model,  self.track_number]
        self.track_number.setText(NO)

    def init_list(self):
        rows = self.parent.db.get_data("id, gov_number, model", AUTO)
        if rows == ERR:
            return
        self.cb_select.addItem(NOT)
        for row in rows:
            self.cb_select.addItems([row[0] + ". " + row[1] + " " + row[2]])

    def init_mask(self):
        symbols = QREVal(QRE("[а-яА-Я 0-9]{9}"))
        for item in self.list_ui[:4]:
            item.setValidator(symbols)

    def _set_data(self, data):
        if not self.track_number.text() or self.track_number.text() == NO:
            self.is_track.setChecked(False)
        else:
            self.is_track.setChecked(True)
        return True

    def have_track(self, state):
        if state == Qt.Checked:
            self.track_number.setEnabled(True)
        else:
            self.track_number.setEnabled(False)
            self.track_number.setText(NO)

    def _clean_data(self):
        self.is_track.setChecked(False)
        return True

    def _select(self, text):
        return True