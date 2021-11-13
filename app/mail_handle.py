import imap_tools as imap
import configparser
import smtplib
import datetime
import datefinder

from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.utils import formatdate

time_last_check = datetime.datetime.today()
config = configparser.ConfigParser()
config.read('../config.ini')
my_imap = config.get("post", "my_imap")
login = config.get("post", "login")
password = config.get("post", "password")
inbox = config.get("post", "inbox")
host = config.get("post", "smtp")

def connect_to_email():
    mailbox = imap.MailBox(my_imap)
    mailbox.login(login, password, initial_folder=inbox)

def break_connect():
    mailbox = imap.MailBox(my_imap)
    mailbox.logout()

def send_post(to_email, sub, text):
    message = "\r\n".join((
        "From: %s" % login,
        "To: %s" % to_email,
        "Subject: %s" % sub,
        "", text))
    server = smtplib.SMTP(host)
    server.sendmail(login, [to_email], message)
    server.quit()
    pass

def check_new_letters():
    result, data = mailbox.search(None, "ALL")
    return result, data

def get_new_post():
    new_message = []
    for msg in mailbox.fetch():
        if msg.is_new_post(msg):
            new_message.append(msg)
        else:
            break
    time_last_check = datetime.datetime.today()
    return new_message

def is_new_post(msg):
    data = [x for x in next(datefinder.find_dates(msg.data_str))]
    post_data = datetime.datetime(*data[0:5])
    now_data = datetime.datetime.today()
    delte_1 = now_data - time_last_check
    delte_2 = now_data - post_data
    return delte_1.seconds - delte_2.seconds < 0

def send_post_with_file(to_email, sub, text, file):
    msg = MIMEMultipart()
    msg["From"] = login
    msg["To"] = to_email
    msg["Subject"] = sub
    msg["Date"] = formatdate(localtime=True)

    if text:
        msg.attach(MIMEText(text))

    attachment = MIMEBase('application', "octet-stream")

    try:
        with open(file, "rb") as fh:
            data = fh.read()

        attachment.set_payload(data)
        encoders.encode_base64(attachment)
        msg.attach(attachment)
    except IOError:
        msg = "Error opening attachment file %s" % file
        print(msg)
        return False
    server = smtplib.SMTP(host)
    server.sendmail(login, to_email, msg.as_string())
    server.quit()
    return True

