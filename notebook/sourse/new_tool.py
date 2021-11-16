from PyQt5.QtCore import QRegExp as QRE
from PyQt5.QtGui import QRegExpValidator as QREVal
from new_template import TempForm
from database import *
from itertools import chain


class NewTool(TempForm):
    def __init__(self, parent=None):
        self.conf = Ini(self)
        self.db = DataBase(self)
        ui_file = self.conf.get_ui("new_tool")
        super(NewTool, self).__init__(ui_file, parent, TOOLS)
        self.date.setDate(dt.datetime.now().date())
        self.init_select()
        self.list_ui = [self.name, self.code, self.count_, self.date]

        self.tools = True
        self.cb_names.activated[str].connect(self.change_name)
        self.init_names()

    def init_names(self):
        rows = self.db.get_data("name", TOOLS)
        add_list = set(chain(*rows))
        list(map(self.cb_names.addItem, add_list))

    def init_select(self):
        rows = self.db.get_data("id, name, date", TOOLS)
        names = list(map(" | ".join, [*rows]))
        list(map(self.cb_select.addItem, names))
        for row in rows:
            self.cb_select.addItem(" | ".join(row))

    def _select(self, text):
        return True

    def _set_data(self, data):
        return True

    def _get_data(self, data):
        return True

    def change_name(self, name):
        if name == NOT:
            return
        self.name.setText(name)
