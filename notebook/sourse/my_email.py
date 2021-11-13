from PyQt5 import uic
from PyQt5.QtWidgets import QDialog
import os
import smtplib
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from database import *


class SendPost(QDialog):
    def __init__(self, db, path):
        self.status_ = True
        self.conf = Ini(self)
        super(SendPost, self).__init__()
        uic.loadUi(self.conf.get_ui("email"), self)
        self.db = db
        self.b_ok.clicked.connect(self.ev_send)
        self.b_cancel.clicked.connect(self.ev_cancel)
        self.path = path
        self.init_emails()
        self.sub = ""
        self.to_email = ""
        self.body_text = ""

    def ev_send(self):
        self.my_sub = self.topic.text()
        self.to_email = self.email.currentText()
        self.body_text = self.note.toPlainText()
        if not self.check_input():
            return False
        answer = msg_q(self, "Отправить письмо на почту ? " + self.to_email)
        if answer == mes.Ok:
            self.send(self.my_sub, self.body_text, self.to_email, self.path)
            self.close()

    def ev_cancel(self):
        self.close()

    def init_emails(self):
        rows = self.db.get_data("email", "bosses")
        if rows == ERR:
            return False
        for item in rows:
            self.email.addItem(item[0])

    def check_input(self):
        return True

    def send(self, subject, body_text, to_email, file_to_attach):
        data = [self.conf.get_from_ini("smtp", "post"),
                self.conf.get_from_ini("my_email", "post"),
                self.conf.get_from_ini("password", "post")]
        if ERR in data:
            return False
        recipients = to_email
        sender = data[1]
        subject = subject
        text = body_text

        filepath = file_to_attach
        basename = os.path.basename(filepath)
        filesize = os.path.getsize(filepath)

        msg = MIMEMultipart('alternative')
        msg['Subject'] = subject
        msg['From'] = data[1]
        msg['To'] = recipients

        part_text = MIMEText(text, 'plain')
        part_file = MIMEBase('application', 'octet-stream; name="{}"'.format(basename))
        part_file.set_payload(open(filepath, "rb").read())
        part_file.add_header('Content-Description', basename)
        part_file.add_header('Content-Disposition', 'attachment; filename="{}"; size={}'.format(basename, filesize))
        encoders.encode_base64(part_file)

        msg.attach(part_text)
        msg.attach(part_file)

        mail = smtplib.SMTP_SSL(data[0], 465)
        mail.login(data[1], data[2])
        mail.sendmail(sender, recipients, msg.as_string())
        mail.quit()
        return
