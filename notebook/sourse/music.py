import webbrowser
from new_template import TempForm
from database import *


class Web(TempForm):
    def __init__(self, parent=None):
        self.status_ = True
        self.conf = Ini(self)
        ui_file = self.conf.get_ui("web")
        super(Web, self).__init__(ui_file, parent, "links")
        if not self.status_:
            return
        self.b_start.clicked.connect(self.ev_start)
        self.parent.db.init_list(self.cb_select, "name, id", self.table)
        self.rows_from_db = self.parent.db.get_data("*", self.table)
        if self.rows_from_db == ERR:
            return
        self.list_ui = [self.name, self.add_link]

    def ev_start(self):
        rows = self.parent.db.get_data("name, link", self.table)
        if rows == ERR:
            return msg_er(self, GET_DB)
        for row in rows:
            if self.cb_select.currentText() in row:
                try:
                    webbrowser.open(row[1])
                except:
                    return msg_er(self, GET_DB)
        self.close()

    def _select(self, text):
        return True

