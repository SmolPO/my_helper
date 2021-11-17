import webbrowser
from new_template import TempForm
from database import *
from PyQt5 import uic


class Web(TempForm):
    def __init__(self, parent=None):
        self.status_ = True
        self.conf = Ini(self)
        self.db = DataBase(self)
        self.table = LINKS
        ui_file = self.conf.get_ui("web")
        self.rows_from_db = self.db.get_data(ALL, LINKS)
        super(Web, self).__init__(ui_file, parent)
        self.b_start.clicked.connect(self.ev_start)
        self.db.init_list(self.cb_select, "name, id", LINKS)
        self.rows_from_db = self.parent.db.get_data(ALL, LINKS)
        if self.rows_from_db == ERR:
            return
        self.list_ui = [self.name, self.add_link]

    def ev_start(self):
        rows = self.parent.db.get_data("name, link", LINKS)
        if rows == ERR:
            return msg_er(self, GET_DB)
        if not rows:
            msg_er(self, "Выберите или добавьте новую ссылку")
            return
        for row in rows:
            if self.cb_select.currentText() in row:
                try:
                    webbrowser.open(row[1])
                except:
                    return msg_er(self, GET_DB)
        self.close()

    def _select(self, text):
        return True

