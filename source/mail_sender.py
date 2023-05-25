import csv
import os
import smtplib
import ssl
from configparser import ConfigParser
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


CURRENT_DIRECTORY = os.path.dirname(os.path.realpath(__file__)).replace("\\", "/")
MAIL_ATTACHMENTS_DIRECTORY = f"{CURRENT_DIRECTORY}/mail_attachments"
MAIL_LIST_FILE_NAME = "mail_list.csv"
MAIL_LIST_FILE_PATH = f"{CURRENT_DIRECTORY}/{MAIL_LIST_FILE_NAME}"
CONFIG_FILE_NAME = "mail_sender.ini"
CONFIG_FILE_PATH = f"{CURRENT_DIRECTORY}/{CONFIG_FILE_NAME}"

# 0 - ID
# 1 - Ученик
# 2 - email
# 3 - File name
# 4 - присъствия от 2 събития и 2 обучения
# 5 - Крайни срокове общо
# 6 - брой срещи
INDEX_STUDENT_NAME = 1
INDEX_STUDENT_EMAIL = 2
INDEX_STUDENT_FILE_NAME = 3


def get_config():
    config = ConfigParser(allow_no_value=True)
    config.read(CONFIG_FILE_PATH)
    return config


def create_message(sender_email: str, receiver_email: str, receiver_name: str, attachment_file_path: str):
    message = MIMEMultipart("alternative")
    message["Subject"] = "ABLE Mentor | Колко точки имаш? 🚀"
    message["From"] = sender_email
    message["To"] = receiver_email
    body_text = """\
Здравей, ✌️

Както споделихме в началото на програмата, всеки екип трупа точки за активно участие. 🎲 Измерваме го чрез три категории: присъствие на събития (общо 4), спазване на крайни срокове (общо 3) и брой срещи с ментора (без максимум). 🐝

Като прикачен файл ще откриеш твоите точки за всяка категория към този момент. 🎉 И помни, никога не е късно да наваксаш! 😏

Поздрави,
Екипът на ABLE Mentor
"""

    # Attach body
    message.attach(MIMEText(body_text, "plain"))

    # Add attachment
    file_name = f"{receiver_name}.pdf"
    payload = MIMEBase("application", "octet-stream", Name=file_name)  # attachment as application/octet-stream
    with open(attachment_file_path, "rb") as binary_file:
        payload.set_payload(binary_file.read())

    # Encode file in ASCII characters to send by email
    encoders.encode_base64(payload)

    # Add header as key/value pair to attachment part
    payload.add_header('Content-Decomposition', 'attachment', filename=file_name)
    message.attach(payload)
    return message


def send_mails():
    smtp_server = "smtp.gmail.com"
    port = 465  # For SSL
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
        config = get_config()
        sender_email = config["General"]["sender_email"]
        password = config["General"]["password"]
        server.login(sender_email, password)

        with open(MAIL_LIST_FILE_PATH, encoding="utf-8", mode="r") as fstream:
            reader = csv.reader(fstream, delimiter=',', quotechar='"')
            for idx, row in enumerate(reader):
                if idx == 0:
                    continue  # skip first row

                receiver_email = row[INDEX_STUDENT_EMAIL]
                receiver_name = row[INDEX_STUDENT_NAME]
                attachment_file_name = row[INDEX_STUDENT_FILE_NAME]
                attachment_file_path = f"{MAIL_ATTACHMENTS_DIRECTORY}/{attachment_file_name}.pdf"
                message = create_message(sender_email, receiver_email, receiver_name, attachment_file_path)

                print(f"Sending email to '{receiver_email}'. Attached file: '{receiver_name}.pdf'")
                server.sendmail(sender_email, receiver_email, message.as_string())


if __name__ == "__main__":
    send_mails()
