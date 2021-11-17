from PyQt5 import uic
from PyQt5.QtWidgets import QDialog, QFileDialog, QInputDialog
from PyQt5.QtGui import QRegExpValidator as QREVal
from PyQt5.QtCore import QRegExp as QRE
import inserts as ins
from datetime import datetime as dt
import PyPDF2
import os
import docxtpl
from database import *


class NewTender(QDialog):
    def __init__(self, parent):
        super(NewTender, self).__init__()
        self.status_ = True
        self.conf = Ini(self)
        self.table = TENDERS
        ui_files = self.conf.get_ui("new_tender")
        uic.loadUi(ui_files, self)
        self.parent = parent
        self.b_ok.clicked.connect(self.ev_ok)
        self.b_cancel.clicked.connect(self.close)
        self.b_kill.clicked.connect(self.ev_kill)
        self.b_change.clicked.connect(self.ev_change)
        self.b_kp_1.clicked.connect(self.kp_1_)
        self.b_kp_2.clicked.connect(self.kp_2_)
        self.b_add.clicked.connect(self.add_docs)
        self.b_open.clicked.connect(self.open_folder)
        self.list_ui = [self.name, self.customer, self.object, self.work, self.part,
                        self.kp_1, self.out_1, self.d_out_1,
                        self.kp_2, self.out_2, self.d_out_2,
                        self.nomn, self.datv]
        self.cb_select.activated[str].connect(self.ev_select)
        self.but_status("add")
        self.rows_from_db = self.parent.db.get_data(ALL, TENDERS)
        self.init_tender()
        self.init_masks()
        if not self.rows_from_db:
            self.current_id = 1
        else:
            self.current_id = int(self.rows_from_db[-1][-1]) + 1

    def init_tender(self):
        for row in self.rows_from_db:
            self.cb_select.addItem(row[-1] + ". " + row[0])

    def init_masks(self):
        self.kp_1.setValidator(QREVal(QRE("[0-9,]{11}")))
        self.kp_2.setValidator(QREVal(QRE("[0-9,]{10}")))

    def check_input(self):
        data = self.get_data()
        if data == ERR:
            return ERR
        empty = {"", ZERO, NOT}
        if empty.intersection(set(data[:5])):
            msg_info(self, FULL_ALL)
            return False
        return data

    def ev_ok(self):
        data = self.check_input()
        if not data or data == ERR:
            return
        if self.parent.db.my_commit(ins.add_to_db(data, TENDERS)) == ERR:
            msg_info(self, "Запись не добавлена")
            return
        self.check_have_folder()
        answer = msg_q(self, ADDED_NOTE)
        if answer == mes.Ok:
            self.close()

    def check_have_folder(self):
        folder = self.conf.get_path("outdoor")
        companies = os.listdir(folder)
        if not self.customer.text() in companies:
            os.mkdir(folder + "/" + self.customer.text())

    def ev_select(self, text):
        if text == NOT:
            self.clean_data()
            self.but_status("add")
            if not self.rows_from_db:
                self.current_id = 1
            else:
                self.current_id = int(self.rows_from_db[-1][-1]) + 1
            return
        else:
            self.but_status("change")
        for row in self.rows_from_db:
            if str(text.split(". ")[0]) == str(row[-1]):
                self.set_data(row)

    def set_data(self, data):
        k = 0
        for item in self.list_ui:
            if "QLineEdit" in str(type(item)):
                item.setText(data[k])
            if "QTextEdit" in str(type(item)):
                item.clear()
                item.append(data[k])
            if "QDateEdit" in str(type(item)):
                item.setDate(from_str(data[k]))
            k = k + 1
        self.current_id = data[-1]

    def get_data(self):
        data = list()
        val = ""
        for item in self.list_ui:
            if "QLineEdit" in str(type(item)):
                val = item.text()
            elif "QTextEdit" in str(type(item)):
                val = item.toPlainText()
            elif "QDateEdit" in str(type(item)):
                val = item.text()
            data.append(val)
        if not data:
            return False
        data.append(str(self.current_id))
        return data

    def clean_data(self):
        for item in self.list_ui:
            if "QLineEdit" in str(type(item)):
                item.setText("")
            if "QTextEdit" in str(type(item)):
                item.clear()
            if "QDateEdit" in str(type(item)):
                item.setDate(zero)

    def ev_change(self):
        data = self.get_data()
        if not data:
            return
        answer = msg_q(self, CHANGE_NOTE + str(data) + "?")
        if answer == mes.Ok:
            data[-1] = str(self.current_id)
            if self.parent.db.my_update(data, self.table) == ERR:
                return
            answer = msg_q(self, CHANGED_NOTE)
            if answer == mes.Ok:
                self.close()

    def ev_kill(self):
        data = self.get_data()
        if not data:
            return
        answer = msg_q(self, KILL_NOTE + str(data) + "?")
        if answer == mes.Ok:
            if self.parent.db.kill_value(self.current_id, TENDERS) == ERR:
                return
            answer = msg_q(self, KILLD_NOTE)
            if answer == mes.Ok:
                self.close()

    def but_status(self, status):
        if status == "add":
            self.b_ok.setEnabled(True)
            self.b_change.setEnabled(False)
            self.b_kill.setEnabled(False)
        if status == "change":
            self.b_ok.setEnabled(False)
            self.b_change.setEnabled(True)
            self.b_kill.setEnabled(True)

    # ____________ Коммерческие____________
    def kp_1_(self):
        data = dict()
        data["number"] = "Исх. " + str(self.conf.get_next_number())
        data["date"] = "от " + str(dt.datetime.now().date())
        data["out"] = self.out_1.text()
        data["d_out"] = self.d_out_1.text()
        data["price"] = self.kp_1.text()
        data["work"] = self.datv.toPlainText()
        for key in data.keys():
            if not data[key] or data[key] == ZERO:
                msg_info(self, FULL_ALL)
                return
        path = self.conf.get_path("pat_notes") + "/КП.docx"
        doc = docxtpl.DocxTemplate(path)
        doc.render(data)
        path = self.conf.get_path("outdoor") + "/" + self.customer.text().replace('"', "") + "/КП_1" + DOCX
        doc.save(path)
        os.startfile(path)

    def kp_2_(self):
        data = dict()
        data["number"] = "Исх. " + str(self.conf.get_next_number())
        data["date"] = "от " + str(dt.datetime.now().date())
        data["out"] = self.out_2.text()
        data["d_out"] = self.d_out_2.text()
        data["price"] = self.kp_2.text()
        data["sale"] = " с учетом максимальной возможной скидки "
        data["work"] = self.datv.toPlainText()
        path = self.conf.get_path("pat_notes") + "/КП.docx"
        doc = docxtpl.DocxTemplate(path)
        doc.render(data)
        path = self.conf.get_path("outdoor") + "/" + self.customer.text().replace('"', "") + "/КП_2" + DOCX
        doc.save(path)
        os.startfile(path)

    def add_docs(self):
        answer = msg_q(self, "Отсканировать файлы?")
        if answer == mes.Ok:
            answer = msg_q(self, "Отсканируйте файлы и нажмите Ок. "
                                    "Затем программа сама их объединит и сохранит куда выберите.")
            if answer == mes.Ok:
                folder = self.conf.get_path("scan")
                files = os.listdir(folder)
                if not files:
                    msg_info(self, "Файлы не найдены. Отсканируйте в PDF и повторите операцию")
                    return
                folder = self.conf.get_path("outdoor") + "/" + self.customer.text().replace('"', "")
                dirlist = QFileDialog.getExistingDirectory(self, "Выбрать папку назначения", folder)
                if dirlist:
                    text, ok = QInputDialog.getText(self, "Название", "Название объединенного документа")
                    if not text or not ok:
                        return
                    path = dirlist + "/" + text + PDF
                else:
                    return
                pdf_merger = PyPDF2.PdfFileMerger()
                for doc in files:
                    if PDF in doc:
                        pdf_merger.append(str(folder + "/" + doc))
                pdf_merger.write(path)
                msg_info(self, "Файл успешно объединен и сохранен")
                return
        else:
            folder = self.conf.get_path("outdoor") + "/" + self.customer.text().replace('"', "")
            sourse, tmp = QFileDialog.getOpenFileName(self, "Выбрать файл", folder, "*(*.*)")
            if sourse:
                target = QFileDialog.getExistingDirectory(self, "Выбрать папку назначения", folder)
                if target:
                    name, ok = QInputDialog.getText(self, "Введите имя", "Введите имя файла")
                    if name:
                        end = sourse.split(".")[-1]
                        os.replace(sourse, target + "/" + name + "." + end)

    def open_folder(self):
        path = self.conf.get_path("outdoor") + "/" + self.customer.text().replace('"', "")
        try:
            os.startfile(path)
        except:
            ok = msg_info(self, GET_FILE + path)


def set_cb_text(combobox, data, rows):
    i = iter(range(1000))
    for item in rows:
        if item[2] == data:
            combobox.setCurrentIndex(next(i) + 1)
            return
        next(i)
