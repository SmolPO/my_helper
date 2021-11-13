import os
from PyQt5.QtWidgets import QDialog
from PyQt5 import uic
from database import *
from PyQt5.QtWidgets import QMessageBox as mes


class Notepad(QDialog):
    def __init__(self, parent=None):
        super(Notepad, self).__init__(parent)
        self.conf = Ini(self)
        self.ui_file = self.conf.get_ui("notepad")
        if not self.check_start():
            return
        self.b_close.clicked.connect(self.ev_close)
        self.file_note = os.getcwd() + "/notepad.txt"
        text = self.get_text()
        if text == ERR:
            return
        self.ui_text_area.appendPlainText(text)
        self.show()

    def check_start(self):
        self.status_ = True
        try:
            uic.loadUi(self.ui_file, self)
            return True
        except:
            msg_er(self, GET_UI + self.ui_file)
            self.status_ = False
            return False

    def get_text(self):
        try:
            file = open(self.file_note, 'r')
        except:
            msg_er(self, GET_FILE + self.file_note)
            return ERR
        text = file.read()
        file.close()
        return text

    def set_text(self):
        text_ui = self.ui_text_area.toPlainText()
        try:
            file = open(self.file_note, 'w')
        except:
            msg_er(self, GET_FILE + self.file_note)
            return ERR
        file.write(text_ui + '\n')
        file.close()

    def ev_close(self):
        if self.ui_check_box.isChecked():
            if self.set_text() == ERR:
                return
        self.close()

    def calc_tap_segments(self):
        d = int(input("Введите диаметр голой трубы"))
        b = int(input("Введите толщину изоляции"))
        neck = int(input("Введите размер шейки"))
        back = int(input("Введите размер затылка"))
        L = int(input("Введите длину окружности по металлу"))
        count = int(input("Введите число сегментов"))
        my_count = count
        if back < 630:
            my_count = 5
        elif back < 1230:
            my_count = 7
        elif back < 2050:
            my_count = 9
        elif back < 3050:
            my_count = 11
        else:
            print("Слишком огромный отвод, сочувствую....")

        h = (back - neck) / (2 * my_count - 2)
        min_neck = neck / my_count
        max_back = min_neck + h
        radius = L / 6.28
        f = open("tmp.txt")
        f.write("высота h = " + str(h))
        f.write("шейка рыбки = " + str(min_neck))
        f.write("затылок рыбки = " + str(max_back))
        f.write("Радиус окружности = " + str(radius))
        f.close()
        os.startfile("tmp.txt", "print")
        os.remove("tmp.txt")
        return

