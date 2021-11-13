from PyQt5.QtCore import QRegExp as QRE
from PyQt5.QtGui import QRegExpValidator as QREVal
from database import *
from new_template import TempForm
msgs = {"mes": "Сообщение", "atn": "Внимание"}


class NewWorker(TempForm):
    def __init__(self, parent=None):
        self.status_ = True
        self.conf = Ini(self)
        ui_file = self.conf.get_ui("new_worker")
        if not ui_file or ui_file == ERR:
            self.status_ = False
            return
        super(NewWorker, self).__init__(ui_file, parent, "workers")
        if not self.status_:
            return
        # my_pass
        self.cb_auto.stateChanged.connect(self.ev_auto)
        self.cb_vac.activated[str].connect(self.change_vac)
        self.init_mask()
        self.list_ui = [self.family, self.name, self.surname, self.bday, self.post,
                        self.passport, self.passport_post,
                        self.adr, self.live_adr, self.phone, self.inn, self.snils,
                        self.n_td, self.d_td, self.n_hght, self.n_group_h,
                        self.d_height, self.n_study, self.n_study_card, self.d_study,
                        self.n_prot, self.n_card, self.d_prot, self.cb_contract,
                        self.d_vac_1, self.d_vac_2, self.place,
                        self.vac_doc, self.cb_vac, self.status]
        self.rows_from_db = self.parent.db.init_list(self.cb_select, "*", self.table, people=True)
        if self.rows_from_db == ERR:
            self.status_ = False
            return
        if self.parent.db.init_list(self.cb_contract, "name, id", "contracts") == ERR:
            self.status_ = False
            return
        self.auto_numbers = ()
        self.my_mem = ""
        self.vac = True

    def init_mask(self):
        symbols = QREVal(QRE("[а-яА-Я ]{30}"))
        number_prot = QREVal(QRE("[А-Яа-я _/- 0-9]{10}"))

        self.family.setValidator(symbols)
        self.name.setValidator(symbols)
        self.surname.setValidator(symbols)
        self.post.setValidator(symbols)
        self.phone.setValidator(QREVal(QRE("[0-9]{11}")))
        self.passport.setValidator(QREVal(QRE("[0-9]{10}")))
        self.inn.setValidator(QREVal(QRE("[0-9]{8}")))
        self.snils.setValidator(QREVal(QRE("[0-9]{8}")))
        self.n_td.setValidator(QREVal(QRE("[0-9]{2}")))
        self.n_hght.setValidator(number_prot)
        self.n_study.setValidator(number_prot)
        self.n_study_card.setValidator(number_prot)
        self.n_prot.setValidator(number_prot)
        self.n_card.setValidator(number_prot)

    def _select(self, text):
        return True

    def get_next_id(self, data):
        g = iter(range(len(self.rows_from_db) + 1))
        for item in self.rows_from_db:
            next(g)
            if data[-1] == item[-1]:
                self.cb_contract.setCurrentIndex(next(g))
                break

    def ev_auto(self, state):
        if state == 2:
            number = list()
            card = list()
            i = 0
            for worker in self.rows_from_db:
                print((i, worker))
                number.append(int(worker[20]))
                card.append(int(worker[21]))
            delta = 1 if dt.datetime.now().weekday() >= 1 else 3
            date = dt.datetime.now().date() - dt.timedelta(delta)
            if not number:
                number.append(0)
                card.append(0)
            self.auto_numbers = max(number), max(card), str(date)
            self.n_prot.setText(str(max(number) + 1))
            self.n_card.setText(str(max(card) + 1))
            self.d_prot.setDate(from_str(str(date)))

            self.n_prot.setEnabled(False)
            self.n_card.setEnabled(False)
            self.d_prot.setEnabled(False)
        else:
            self.n_prot.setEnabled(True)
            self.n_card.setEnabled(True)
            self.d_prot.setEnabled(True)