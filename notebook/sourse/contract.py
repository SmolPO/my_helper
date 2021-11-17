from PyQt5 import uic
from PyQt5.QtWidgets import QDialog
from database import *


class Contract(QDialog):
    def __init__(self, parent):
        super(Contract, self).__init__()
        self.conf = Ini(self)
        self.ui_file = self.conf.get_ui("contract")
        uic.loadUi(self.ui_file, self)
        self.parent = parent

        self.b_contract.clicked.connect(self.ev_open)
        self.b_def.clicked.connect(self.ev_open)
        self.b_ekr.clicked.connect(self.ev_open)
        self.b_exit.clicked.connect(self.close)
        self.path = self.conf.get_path("docs")

    def ev_open(self):
        if self.sender().text() + PDF in os.listdir(self.path):
            path = self.path + "/" + self.sender().text() + PDF
            if not save_open(path):
                msg_er(self, GET_FILE + path)
                return
        else:
            msg_info(self, "Нет файла документации: " + self.sender().text())
            return
