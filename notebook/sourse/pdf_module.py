from PyQt5 import uic
from PyQt5.QtWidgets import QMessageBox as mes
from PyQt5.QtWidgets import QDialog, QFileDialog
from PyQt5.QtWidgets import QInputDialog
import PyPDF2
import pytesseract
from datetime import datetime as dt
from database import *
from my_email import *

pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'
sub = "Ковидный журнал " + "ООО 'Вертикаль'"
to_email = "kalent_ivan@mail.ru"
body_text = "Доброе утро!"
folder = "B:/my_helper/Сканы"


class PDFModule(QDialog):
    def __init__(self, parent):
        super(PDFModule, self).__init__()
        self.conf = Ini(self)
        self.ui_file_1 = self.conf.get_ui("pdf_module")
        if self.ui_file_1 == ERR:
            self.status_ = False
            return
        self.ui_file_2 = self.conf.get_ui("email")
        if self.ui_file_2 == ERR:
            self.status_ = False
            return
        self.parent = parent
        if not self.check_start():
            return
        self.b_covid.clicked.connect(self.ev_covid)
        self.b_note.clicked.connect(self.ev_note)
        self.b_other.clicked.connect(self.ev_open)

    def check_start(self):
        self.status_ = True
        self.path_ = self.ui_file_1
        try:
            uic.loadUi(self.ui_file_1, self)
            return True
        except:
            msg_er(self, GET_UI + self.ui_file_1)
            self.status_ = False
            return False

    def ev_covid(self):
        path_file = self.check_input_c19()
        if not path_file or path_file == ERR:
            return False
        path = self.conf.get_path("path_covid")
        path_to = path + "/" + str(dt.now().date()) + PDF
        answer = msg_info(self, "Отправить ковидный журнал на почту " + to_email)
        if answer == mes.Ok:
            try:
                os.replace(path_file, path_to)
            except:
                msg_info(self, GET_FILE + path_file)
            SendPost(self).send(sub, body_text, to_email, path_to)
        pass

    def check_input_c19(self):
        folder = self.get_folder()
        if len(os.listdir(folder)) > 1:
            file_name = QFileDialog.getOpenFileName(self, "Выбрать файл",
                                                         folder, "PDF Files(*.pdf)")
            if not file_name:
                return
            return folder + "/" + file_name
        elif not os.listdir(folder):
            msg_info(self, "В папке {0} нет отсканированных файлов. "
                           "Отсканируйте файл и повторите попытку".format(folder))
            return False
        else:
            return folder + "/" + os.listdir(folder)[0]

    def check_input_n(self):
        folder = self.get_folder()
        try:
            files = os.listdir(folder)
        except:
            msg_info(self, GET_FILE + folder)
            return
        files.sort()
        if not files:
            msg_info(self, "В папке {0} нет отсканированных файлов. "
                           "Отсканируйте файл и повторите попытку".format(folder))
            return False
        for file in files:
            if not ".pdf" in file:
                msg_info(self, "В папке {0} нет отсканированных файлов в формате .pdf, "
                               "отсканируйте файл в формате .pdf и повторите попытку".format(folder))
                return False
        return True

    def ev_note(self):
        folder = self.get_folder()
        try:
            files = os.listdir(folder)
            if not files:
                msg_info(self, "Файлы не найдены. Отсканируйте в PDF и программа сама их объединит по порядку")
                return
        except:
            mes.question(self, "Сообщение", GET_FILE + folder, mes.Cancel)
            return False
        files.sort()
        if not self.check_input_n():
            return
        text, ok = QInputDialog.getInt(self, 'Сообщение', 'Укажите номер исходящего')
        if not text:
            return
        pdf_merger = PyPDF2.PdfFileMerger()
        paths = [self.conf.get_path("path"), self.conf.get_path("path_notes_pdf")]
        if ERR in paths:
            return
        path_to = "".join(paths) + "/" + str(text) + "_" + str(dt.datetime.now().date()) + PDF
        for doc in files:
            if PDF in doc:
                pdf_merger.append(str(folder + "/" + doc))
        pdf_merger.write(path_to)
        wnd = SendPost(self.parent.db, path_to)
        if not wnd.status_:
            mes.question(self, "Сообщение", "Не найден файл дизайна " + wnd._path, mes.Cancel)
            return
        wnd.setFixedSize(wnd.geometry().width(), wnd.geometry().height())
        wnd.exec_()
        for doc in files:
            os.remove(str(folder + "/" + doc))

    def ev_open(self):
        folder = self.get_folder()
        files = os.listdir(folder)
        if not files:
            msg_info(self, "Файлы не найдены. Отсканируйте в PDF и программа сама их объединит по порядку")
            return
        text, ok = QInputDialog.getText(self, "Название", "Название документа")
        if not text:
            return
        if not ok:
            return
        dirlist = QFileDialog.getExistingDirectory(self, "Выбрать папку", folder)
        if ok:
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

    def get_folder(self):
        return self.conf.get_path("path_scan")

"""def check_file(parent):
    files = os.listdir(folder)
    files.sort()
    path_covid = "B:/my_helper/Служебки/covid"
    count_files = len(files)
    for file in files:
        print(file)
        text = str((pytesseract.image_to_string(Image.open(folder + "/" + file), lang='rus')))
        print(text)
        if "Температура" in text:
            if count_files == 1:
                Это утренний ковид журнал. 
                Переместить в папку Ковид, преобразовать в PDF и отправить по почте. Адрес записан в настройках.
                date_ind = text.find(".2021")
                date = text[date_ind:date_ind + 10]
                print(date)
                path_pdf = path_covid + "/" + date + "_" + get_config("company") + ".pdf"
                convert_2pdf(file, path_pdf)
                answer = mes.question(parent, "Сообщение", "Отправить ковидны журнал на почту " + to_email, mes.Ok | mes.Cansel)
                if answer == mes.Ok:
                    send_post(sub, body_text, to_email, path_pdf)
                pass
        elif "выдать" in text:
            create_note(parent, file, files)
        elif "разблокировать" in text:
            create_note(parent, file, files)
        elif "продлить " in text:
            create_note(parent, file, files)
        elif "ИТОГ" in text:
            date = next_bill(text, "B:\my_helper\scan\scan.jpg")
            return date
        elif "выходные" or "выходной" in text:
            create_note(parent, file, files)


def convert_2pdf(path_from, path_to):
    image = Image.open(path_from)
    pdf_bytes = img2pdf.convert(image.filename)
    file = open(path_to, "wb")
    file.write(pdf_bytes)
    image.close()
    file.close()

def create_note(parent, path, files):
    pdf_merger = PyPDF2.PdfFileMerger()
    output_file = "B:/my_helper/Служебки/Архив/PDF" + "1.pdf"
    for doc in files.sort():
        pdf_merger.append(str(doc))
    pdf_merger.write(output_file)
    answer = mes.question(parent, "Сообщение", "Отправаить ковидны журнал на почту " + to_email, mes.Ok | mes.Cansel)
    if answer == mes.Ok:
        send_post(sub, body_text, to_email, output_file)
    for doc in files:
        os.remove(doc)

def next_bill(text, file):
    img = cv2.imread(file)
    detector = cv2.QRCodeDetector()
    data, bbox, tmp = detector.detectAndDecode(img)
    return data


def go_img2pdf(file, to_folder, new_name):
    a4_page_size = [img2pdf.in_to_pt(8.3), img2pdf.in_to_pt(11.7)]
    layout_function = img2pdf.get_layout_fun(a4_page_size)
    pdf = img2pdf.convert(file, layout_fun=layout_function)
    with open(to_folder + "/" + new_name, 'wb') as f:
        f.write(pdf)

"""