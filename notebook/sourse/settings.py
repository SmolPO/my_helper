from PyQt5 import uic
from PyQt5.QtWidgets import QDialog, QFileDialog
from database import *
import win32print
from PyQt5.QtWidgets import QMessageBox as mes


class Settings(QDialog):
    def __init__(self, parent=None):
        super(Settings, self).__init__()
        self.conf = Ini(self)
        self.ui_file = self.conf.get_ui("settings")
        if not self.check_start():
            return
        self.parent = parent
        # my_pass
        self.b_ok.clicked.connect(self.ev_OK)
        self.b_cancel.clicked.connect(self.ev_cancel)
        self.b_explorer.clicked.connect(self.ev_explorer)
        self.init_data()

    def check_start(self):
        self.status_ = True
        try:
            uic.loadUi(self.ui_file, self)
            return True
        except:
            msg_er(self, GET_UI + self.ui_file)
            self.status_ = False
            return False

    def ev_OK(self):
        data = self.get_data()
        self.save_data(data)
        self.close()
        pass

    def ev_cancel(self):
        self.close()

    def ev_explorer(self):
        self.main_path = QFileDialog.getExistingDirectory(self, "Open Directory", self.parent.main_path,
                                                          QFileDialog.ShowDirsOnly | QFileDialog.DontResolveSymlinks)

        if not self.filename:
            return
        self.path.setText(self.self.main_path)

    def get_data(self):
        data = dict()
        data["path"] = self.path.text()
        data["ip"] = self.ip.text()
        data["dname"] = self.name_db.text()
        data["duser"] = self.user_db.text()
        data["password"] = self.password_db.text()
        data["new_year"] = str(self.new_year.value())
        data["to_paper"] = self.to_paper.text()
        data["to_pdf"] = self.to_pdf.text()
        return data

    def init_data(self):
        config = ConfigParser()
        try:
            config.read('config.ini')
            self.path.setText(config.get('path', 'path'))
            self.ip.setText(config.get('database', 'ip'))
            self.name_db.setText(config.get('database', 'name_db'))
            self.user_db.setText(config.get('database', 'user_db'))
            self.password_db.setText(config.get('database', 'password_db'))
            self.new_year.setValue(int(config.get('config', 'new_year')))
        except:
            return  msg_er(self, GET_INI)

    def save_data(self, data):
        config = ConfigParser()
        try:
            config.read('config.ini')
            config.set('path', 'path', data["path"])
            config.set('database', 'ip', data["ip"])
            config.set('database', 'name_db', data["dname"])
            config.set('database', 'user_db', data["duser"])
            config.set('database', 'password_db', data["password"])
            config.set('config', 'new_year', data["new_year"])
            config.set('config', 'to_paper', data["to_paper"])
            config.set('config', 'to_pdf', data["to_pdf"])
            with open('config.ini', 'w') as configfile:
                config.write(configfile)
        except:
            return msg_er(self, GET_INI)

