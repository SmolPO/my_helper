from PyQt5.QtWidgets import QMessageBox as mes
from PyQt5.QtWidgets import QDialog, QCheckBox, QLabel
from PyQt5 import uic
import docx
import docxtpl
from database import *
types_docs = {"1": "/ot_doc.docx", "2": "/ptm_doc.docx", "3": "/eb_doc.docx"}
types_card = {"1": "/ot_card.docx", "2": "/ptm_card.docx", "3": "/es_card.docx"}


class NewTB(QDialog):
    def __init__(self, parent=None):
        self.status_ = True
        self.conf = Ini(self)
        ui_file = self.conf.get_ui("new_TB")
        super(NewTB, self).__init__()
        uic.loadUi(ui_file, self)
        self.parent = parent
        self.table = "workers"
        self.rows_from_db = self.parent.db.get_data("*", self.table)
        if self.rows_from_db == ERR:
            self.status_ = False
            return
        self.count = self.parent.count_people_tb
        self.b_ok.clicked.connect(self.ev_ok)
        self.b_cancel.clicked.connect(self.ev_cancel)
        self.path = dict()
        self.list_ui = [self.worker_1, self.worker_2, self.worker_3, self.worker_4, self.worker_5,
                        self.worker_6, self.worker_7, self.worker_8, self.worker_9, self.worker_10,
                        self.worker_11, self.worker_12, self.worker_13, self.worker_14, self.worker_15,
                        self.worker_16, self.worker_17, self.worker_18, self.worker_19, self.worker_20,
                        self.worker_21, self.worker_22, self.worker_23, self.worker_24, self.worker_25,
                        self.worker_26, self.worker_27, self.worker_28, self.worker_29, self.worker_30]
        self.list_cb = [NOT for i in range(len(self.list_ui))]
        rows = [self.parent.db.get_data(ALL, WORKERS), self.parent.db.get_data(ALL, ITRS)]
        if ERR in rows:
            return
        self.all_people = rows[0] + rows[1]
        self.path["main_folder"] = self.conf.get_path("pat_tb")
        self.path["print_folder"] = self.conf.get_path("tb")
        self.list_ui = list()
        self.init_list()

    def init_workers(self):
        for item in self.list_ui:
            item.addItem(NOT)
            item.activated[str].connect(self.new_worker)
            item.setEnabled(False)
        self.list_ui[0].setEnabled(True)
        self.all_people.sort()
        workers = list()
        for item in self.all_people:
            if "работает" in item[-2] or "в отпуске" in item[-2]:
                workers.append(item)
        workers.sort(key=lambda x: x[0])
        for people in workers:
            for item in self.list_ui:
                item.addItem(short_name(people))

    def get_list_people(self):
        list_people = list()
        for elem in self.list_ui:
            family = elem.currentText()
            if family != NOT:
                item = self.check_row(family)
                if item:
                    list_people.append(item)
        list_people.sort(key=lambda x: x[0])
        return list_people

    def check_row(self, row):
        family = row.split(".")[0][:-2]
        s_name = row.split(".")[0][-1]
        s_sur = row.split(".")[1]
        for item in self.all_people:
            if family == item[0]:
                if s_name == item[1][0]:
                    if s_sur == item[2][0]:
                        return item

    def ev_ok(self):
        people = self.get_list_people()
        if self.create_protocols(people) == ERR:
            return
        if self.create_cards(people) == ERR:
            return
        return True

    def ev_cancel(self):
        self.close()

    def create_protocols(self, people):
        printed_docs = list()
        worker = list()
        fields = [0, 1, 2, 4, 20, 21, 22, 24]
        _key = 20
        for man in people:
            if not man[_key] in printed_docs:
                for field in fields:
                    worker.append(man[field])
                for key in types_docs.keys():
                    if self.print_doc(worker, key) == ERR:
                        return ERR
                printed_docs.append(man[_key])
        return
    
    def create_cards(self, people):
        data = dict()
        for worker in people:
            data["family"] = " ".join(worker[:3])
            data["post"] = worker[3]
            data["number_doc"] = worker[20]
            data["number"] = worker[21]
            data["date"] = worker[22]
            for name in types_card.keys():
                try:
                    path = self.path["main_folder"] + types_card[name]
                    doc = docxtpl.DocxTemplate(path)
                    doc.render(self.data)
                    path = self.path["print_folder"] + types_card[name]
                    doc.save(path)
                    os.startfile(path, "print")
                except:
                    return msg_er(self, GET_FILE + path)
                self.close()

    def print_doc(self, workers, number_type):
        data = dict()
        data["date"] = workers[7] + "/" + str(number_type)  # номер столбца с датой
        data["number_doc"] = workers[6]  # номер протокол
        self.create_table(workers, number_type)
        try:
            path = self.path["main_folder"] + types_docs[number_type]
            doc = docxtpl.DocxTemplate(path)
            doc.render(self.data)
            path = self.path["print_folder"] + types_docs[number_type]
            doc.save(path)
            os.startfile(self.print_file)
        except:
            return msg_er(self, GET_FILE + path)

    def create_table(self, data, number_type):
        g = iter(range(len(data)))
        try:
            path = self.path["main_folder"] + types_docs[number_type]
            doc = docx.Document(path)
            for item in data:
                i = next(g)
                doc.tables[0].add_row()
                doc.tables[0].rows[i].cells[0].text = str(i)
                doc.tables[0].rows[i].cells[1].text = " ".join(item[:3]) # ФИО
                doc.tables[0].rows[i].cells[2].text = item[3]   # профессия
                doc.tables[0].rows[i].cells[3].text = "Сдал №" + item[5]
            path = self.path["print_folder"] + types_docs[number_type]
            doc.save(path)
        except:
            return msg_er(self, GET_FILE + path)


class CountPeople(QDialog):
    def __init__(self, parent=None):
        self.status_ = True
        self.conf = Ini(self)
        ui_file = self.conf.get_ui("count")
        super(CountPeople, self).__init__(parent)
        uic.loadUi(ui_file, self)

        self.parent = parent
        self.table = "workers"
        self.b_ok.clicked.connect(self.ev_ok)
        self.b_cancel.clicked.connect(self.ev_cancel)
        count_people = len(self.parent.db.get_data("*", self.table))
        if count_people == ERR:
            return
        self.count.setMaximum(count_people)

    def ev_ok(self):
        self.parent.count_people_tb = self.count.value()
        self.close()
        return

    def ev_cancel(self):
        self.parent.count_people_tb = -1
        self.close()



