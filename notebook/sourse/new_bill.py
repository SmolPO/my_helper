from PyQt5.QtWidgets import *
from datetime import datetime as dt
import os
import openpyxl
from database import *
from new_template import TempForm


class NewBill(TempForm):
    def __init__(self, parent):
        self.status_ = True
        self.conf = Ini(self)
        ui_file = self.conf.get_ui("new_bill")
        if not ui_file or ui_file == ERR:
            self.status_ = False
            return
        super(NewBill, self).__init__(ui_file, parent, "bills")
        if not self.status_:
            return
        self.b_bill.clicked.connect(self.ev_bill)
        self.b_report.clicked.connect(self.report)
        if self.parent.db.init_list(self.cb_buyer, "*", "itrs", people=True) == ERR:
            self.status_ = False
            return
        self.filename = ""
        self.list_ui = [self.date, self.sb_value, self.cb_buyer, self.file_path, self.note]
        self.date.setDate(dt.datetime.now().date())
        if self.init_operations() == ERR:
            self.status_ = False
            return
        self.bill = True

    def init_mask(self):
        return

    def init_operations(self):
        rows = self.parent.db.get_data("*", self.table)
        rows.sort(key=lambda x: x[-1])
        if rows == ERR:
            self.status_ = False
            return ERR
        self.cb_select.addItem(NOT)
        for row in rows:
            self.cb_select.addItems([". ".join((row[-1], row[0]))])

    def ev_bill(self):
        self.filename, tmp = QFileDialog.getOpenFileName(self, "Выбрать файл", self.conf.get_path("scan"),
                                                         "PDF Files(*.pdf)")
        if self.filename:
            self.file_path.setText(self.filename.split("/")[-1])

    def _select(self, text):
        return True

    def create_note(self, value, date, people):
        path_ = self.conf.get_path("bills")
        path = path_ + "/" + str(dt.now().year) + \
                                    "/" + str(dt.now().month) + \
                                    "/" + str(dt.now().month) + str(dt.now().year) + ".xlsx"
        try:
            wb = openpyxl.load_workbook(path)
        except:
            return msg_er(self, GET_FILE + path)
        try:
            sheet = wb['bills']
        except:
            return msg_er(self, GET_PAGE + 'bills')
        row = sheet['F2'].value
        sheet['A' + str(row + 3)].value = int(row) + 1
        sheet['B' + str(row + 3)].value = date
        sheet['C' + str(row + 3)].value = value
        sheet['D' + str(row + 3)].value = people
        sheet['F2'].value = int(row) + 1
        try:
            wb.save(path)
            os.startfile(path)
        except:
            return msg_er(self, GET_FILE)

    def report(self):
        cur_month, ok = QInputDialog.getItem(self.parent, "Выберите месяц", "Месяц", month, dt.datetime.now().month)
        if not ok:
            return
        month_ = str(month.index(cur_month) + 1)
        year_ = str(dt.datetime.now().year)
        path_ = "/" + str(dt.datetime.now().year) + "/Чеки/" + month_ + "/" + month_ + year_ + XLSX
        path = self.parent.main_path + "/Бухгалтерия" + path_
        try:
            os.startfile(path)
        except:
            msg_info(self, "За данный месяц отчет не найден")

