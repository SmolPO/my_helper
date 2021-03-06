from PyQt5 import uic
from PyQt5.QtWidgets import *
from PyQt5.QtWidgets import QDialog
import os
from asr import Asr
from report import CreateReport
from database import *


class Acts(QDialog):
    def __init__(self, parent):
        super(Acts, self).__init__()
        self.conf = Ini(self)
        self.db = DataBase(self)
        self.ui_file = self.conf.get_ui("acts")
        self.parent = parent
        uic.loadUi(self.ui_file, self)
        self.b_save.clicked.connect(self.ev_save)
        self.b_asr.clicked.connect(self.ev_start)
        self.b_add.clicked.connect(self.ev_add)
        self.b_latter.clicked.connect(self.ev_latter)
        self.b_month.clicked.connect(self.ev_month)
        self.b_create.clicked.connect(self.ev_xlsx)
        self.b_exit.clicked.connect(self.ev_exit)
        self.cb_select.activated[str].connect(self.ev_select)
        self.contract = ""
        if self.init_contracts() == ERR:
            self.status_ = False
            return

    def init_contracts(self):
        rows = self.db.get_data("id, number", CONTRACTS)
        if rows == ERR:
            return ERR
        for item in rows:
            self.cb_select.addItem(". ".join(item))

    def ev_start(self):
        if self.cb_select.currentText() == NOT:
            msg_info(self, "Сначала выберите договор")
            return False
        menu = {self.b_asr.text(): Asr}
        name = self.sender().text()
        self.path = self.parent.main_path + "/Договор/" + "".join(self.cb_select.currentText().split(". ")[1:])
        wnd = menu[name](self)
        set_fix_size(wnd)
        wnd.exec_()
        return

    def ev_save(self):
        pass

    def ev_add(self):
        if self.cb_select.currentText() == NOT:
            return msg_er(self, "Укажите сначала номер договора")
        self.filename, tmp = QFileDialog.getOpenFileName(self,
                                                         "Выбрать файл",
                                                         self.conf.get_path("scan"),
                                                         "*.*(*.*)")
        if not self.filename:
            return
        tmp = self.filename.split(".")[-1]
        path_save = self.path + "/" + "".join(self.cb_select.currentText().split(". ")[1:]) + self.conf.get_path("others")
        name, ok = QInputDialog.getText(self, "Введите имя файла", "Имя (без расширения)")
        if ok:
            if not name:
                return
            path_save = path_save + "/" + name + "." + tmp
            if not save_replace(self.filename, path_save):
                msg_info(self, GET_FILE + save_replace + " или " + path_save)
                return False
            msg_info(self, "Сообщение", ADDED_FILE)

    def ev_latter(self):
        path_from = self.conf.get_path("pat_patterns") + "/Бланк.docx"
        path_to = self.conf.get_path("path") + "/Исходящие/Письма/Письмо.docx"
        if not save_replace(path_from, path_to):
            mes.question(self, "Сообщение", "Файл " + path_from + " не найден", mes.Ok)
            return False
        if not save_open(path_to):
            msg_info(self, GET_FILE + path_to)

    def ev_month(self):
        if not save_open(self.path):
            return msg_info(self, GET_FILE + self.path)
        pass

    def ev_xlsx(self):
        wnd = CreateReport(self)
        if not wnd.status_:
            return
        wnd.exec_()
        if not save_open(self.path):
            return msg_info(self, GET_FILE + self.path)
        pass

    def ev_exit(self):
        self.close()

    def ev_select(self):
        if self.cb_select.currentText() != NOT:
            self.contract = self.cb_select.currentText().split(".")[1]
        pass
