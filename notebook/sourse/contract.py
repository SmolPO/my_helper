from PyQt5 import uic
import os
from PyQt5.QtWidgets import QDialog
from database import *


class Contract(QDialog):
    def __init__(self, parent):
        super(Contract, self).__init__()
        self.conf = Ini(self)
        self.ui_file = self.conf.get_ui("contract")
        if not self.check_start():
            return
        self.parent = parent

        self.b_contract.clicked.connect(self.ev_open)
        self.b_def.clicked.connect(self.ev_open)
        self.b_ekr.clicked.connect(self.ev_open)
        self.b_exit.clicked.connect(self.close)
        self.path = self.conf.get_path("docs")

    def check_start(self):
        self.status_ = True
        try:
            uic.loadUi(self.ui_file, self)
            return True
        except:
            msg_er(self, GET_UI)
            self.status_ = False
            return False

    def ev_open(self):
        if self.sender().text() + ".pdf" in os.listdir(self.path):
            path = self.path + "/" + self.sender().text() + ".pdf"
            try:
                os.startfile(path)
            except:
                msg_er(self, GET_FILE + path)
                return
        else:
            msg_info(self, "Нет файла документации: " + self.sender().text())
            return
