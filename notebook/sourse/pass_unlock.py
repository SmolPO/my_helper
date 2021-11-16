import docx
import os
# import pymorphy2
import docxtpl
from docx.oxml.text.paragraph import CT_P
from docx.text.paragraph import Paragraph
from pass_template import TempPass
from database import *


class UnlockPass(TempPass):
    def __init__(self, parent):
        self.status_ = True
        self.conf = Ini(self)
        ui_file = self.conf.get_ui("pass_unlock")
        super(UnlockPass, self).__init__(ui_file, parent, "workers")
        self.parent = parent
        # my_pass
        self.d_from.setDate(dt.datetime.now().date())
        self.d_to.setDate(from_str(".".join([str(count_days[dt.datetime.now().month - 1]),
                                             str(dt.datetime.now().month),
                                             str(dt.datetime.now().year)])))
        self.rows_from_db = self.parent.db.get_data("*", self.table)
        if self.rows_from_db == ERR:
            self.status_ = False
            return
        self.init_workers()
        self.data = {"number": "", "data": "", "customer": "", "company": "", "start_date": "", "end_date": "",
                     "post": "", "family": "", "name": "", "surname": "", "adr": ""}
        self.vac_path = self.parent.conf.get_path("pat_notes")
        self.main_file += "/Разблокировка.docx"
        self.count_days = 14
        self.vac = True

    def init_workers(self):
        for people in self.all_people:
            if "работает" in people[-2]:
                self.cb_worker.addItem(short_name(people))

    def _get_data(self):
        family = self.cb_worker.currentText()
        # morph = pymorphy2.MorphAnalyzer()
        people = self.check_row(family)  # на форме фамилия в виде Фамилия И.
        post = people[3]
        if post in dictionary.keys():
            post = dictionary[post]['datv']
        else:
            post = people[3]
        #    try:
        #         post = morph.parse(people[3])[0].inflect({'datv'})[0]
        #    except:
        #        post = people[3]
        self.data["family"] = people[0]
        self.data["name"] = people[1]
        self.data["surname"] = people[2]
        # self.data["family"] = morph.parse(people[0])[0].inflect({'datv'})[0].capitalize()
        # self.data["name"] = morph.parse(people[1])[0].inflect({'datv'})[0].capitalize()
        # self.data["surname"] = morph.parse(people[2])[0].inflect({'datv'})[0].capitalize()
        self.data["post"] = post
        self.data["adr"] = people[8]
        self.data["start_date"] = self.d_from.text()
        self.data["end_date"] = self.d_to.text()

    def check_input(self):
        return True

    def _ev_ok(self):
        return True

    def all_days(self, state):
        self.sb_days.setEnabled(not state)
        self.sb_days.setValue(14)
        self.count_days = 14
        pass

    def _create_data(self, _):
        family = self.cb_worker.currentText()
        if self.create_vac(family) == ERR:
            return ERR
        pass

    def create_vac(self, family):
        note = ["Настоящим письмом информируем Вас о прохождение вакцинации от Covid-19 сотрудником ООО «Вертикаль»"]
        people = self.check_row(family)
        data_vac = people[-7:-2]
        next_id = self.conf.set_next_number(int(self.number.value()) + 2)
        if SPUTNIK in data_vac:
            self.vac_path += "/Вакцинация_3.docx"
        if SP_LITE in data_vac:
            self.vac_path += "/Вакцинация_2.docx"
        if COVID in data_vac:
            self.vac_path += "/Вакцинация_4.docx"
        try:
            doc = docxtpl.DocxTemplate(self.vac_path)
        except:
            return ERR
        data = dict()
        data["number"] = "Исх. №" + str(next_id)
        data["date"] = "от " + self.d_note.text()
        data["family"] = " ".join(people[:3])
        data["post"] = people[3]
        if SPUTNIK in data_vac:
            data["d_vac_1"] = data_vac[0]
            data["d_vac_2"] = data_vac[1]
            data["place"] = data_vac[2]
        elif SP_LITE in data_vac:
            data["date_vac"] = data_vac[0]
            data["place"] = data_vac[2]

        elif COVID in data_vac:
            data["title"] = note
            data["date_doc"] = data_vac[0]
            data["num_doc"] = data_vac[2]

        path = self.conf.get_path("notes_docs") + "/" + str(int(next_id) + 1) + "_" + self.d_note.text() + ".docx"
        doc.render(data)
        try:
            doc.save(path)
            os.startfile(path)
        except:
            return msg_er(self, GET_FILE + path)
