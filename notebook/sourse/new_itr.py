from PyQt5.QtCore import QRegExp as QRE
from PyQt5.QtGui import QRegExpValidator as QREVal
from database import *
from new_template import TempForm


class NewITR(TempForm):
    def __init__(self, parent):
        self.status_ = True
        self.conf = Ini(self)
        self.db = DataBase(self)
        ui_file = self.conf.get_ui("new_itr")
        self.rows_from_db = self.db.get_data(ALL, ITRS)
        super(NewITR, self).__init__(ui_file, parent)
        # my_pass
        self.init_mask()
        self.cb_vac.activated[str].connect(self.change_vac)
        self.rows_from_db = self.parent.db.init_list(self.cb_select, ALL, ITRS, people=True)
        if self.rows_from_db == ERR:
            self.status_ = False
            return
        if self.parent.db.init_list(self.cb_auto, ALL, AUTO) == ERR:
            self.status_ = False
            return
        self.list_ui = [self.family, self.name, self.surname, self.post,
                        self.passport, self.passport_got,
                        self.adr, self.live_adr, self.cb_auto,
                        self.inn, self.snils,
                        self.n_td, self.d_td,
                        self.n_OT_p, self.d_OT, self.n_OT_c,
                        self.n_PTM_p, self.d_PTM, self.n_PTM_c,
                        self.n_ES_p, self.n_ES_g, self.n_ES_c, self.d_ES,
                        self.n_H_p, self.d_H, self.n_H_g, self.n_H_c,
                        self.promsave,
                        self.n_ST_p, self.n_ST_c, self.d_ST,
                        self.bday,
                        self.d_vac_1, self.d_vac_2, self.place, self.vac_doc, self.cb_vac, self.status]
        self.my_mem = ""
        self.vac = True

    def _select(self, text):
        return True

    def init_mask(self):
        list_valid = [[self.family,  self.name, self.surname, self.post],
                      [self.n_OT_p, self.n_OT_c, self.n_PTM_p, self.n_PTM_c, self.n_ES_p,  self.n_ES_c,
                       self.n_H_p, self.n_H_c, self.n_ST_p, self.n_ST_c]]
        for item in list_valid[0]:
            item.setValidator(QREVal(QRE("[а-яА-Я ]{30}")))
        for item in list_valid[1]:
            item.setValidator(QREVal(QRE("[А-Яа-я /- 0-9]{10}")))

        self.passport.setValidator(QREVal(QRE("[0-9]{10}")))
        self.inn.setValidator(QREVal(QRE("[0-9]{9}")))
        self.snils.setValidator(QREVal(QRE("[0-9]{11}")))
        self.n_td.setValidator(QREVal(QRE("[0-9]{3}")))
        self.n_ES_g.setValidator(QREVal(QRE("[0-9]{3}")))
        self.n_H_g.setValidator(QREVal(QRE("[0-9]{3}")))