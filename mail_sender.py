import csv
import logging
import os
import smtplib
import ssl
import sys
import xml.etree.ElementTree as ET
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# For debugging purposes you can test on a local server.
# First you have to start an smtp server and then run the script with a '--debug' argument.
# 1. $ python -m aiosmtpd -n (for python 3.12+)
# 1. $ python -m smtpd -c DebuggingServer -n localhost:8025 (for python 3.11-)
# 2. $ python mail_sender.py --debug

CURRENT_DIRECTORY = os.path.dirname(os.path.realpath(__file__)).replace("\\", "/")
CONFIG_FILE_NAME = "mail_sender.xml"
CONFIG_FILE_PATH = f"{CURRENT_DIRECTORY}/{CONFIG_FILE_NAME}"

# These are the names of the XML tags
SENDER_EMAIL = "sender_email"
PASSWORD = "password"
SUBJECT = "subject"
BODY = "body"
ATTACHMENTS_FOLDER_NAME = "attachments_folder_name"
CSV_FILE_NAME = "file_name"
CSV_RECEIVER_NAME_INDEX = "receiver_name_index"
CSV_RECEIVER_EMAIL_INDEX = "receiver_email_index"
CSV_ATTACHMENT_FILE_INDEX = "attachment_file_index"

logging.basicConfig(filename=f"{CURRENT_DIRECTORY}/mail_sender.log", filemode="w", level=logging.DEBUG)


def in_debug_mode():
    return len(sys.argv) > 1 and sys.argv[1] == "--debug"


def log(message: str):
    print(message)
    logging.info(message)


def log_error(message: str):
    print(message)
    logging.error(message)


def get_config():
    config = dict()
    root_node = ET.parse(CONFIG_FILE_PATH)

    # Login
    login_node = root_node.find("login")
    config[SENDER_EMAIL] = login_node.find(SENDER_EMAIL).text
    config[PASSWORD] = login_node.find(PASSWORD).text

    # Message
    message_node = root_node.find("message")
    config[SUBJECT] = message_node.find(SUBJECT).text
    config[BODY] = message_node.find(BODY).text.strip()
    config[ATTACHMENTS_FOLDER_NAME] = message_node.find(ATTACHMENTS_FOLDER_NAME).text

    # CSV
    csv_node = root_node.find("csv")
    config[CSV_FILE_NAME] = csv_node.find(CSV_FILE_NAME).text
    config[CSV_RECEIVER_NAME_INDEX] = int(csv_node.find(CSV_RECEIVER_NAME_INDEX).text)
    config[CSV_RECEIVER_EMAIL_INDEX] = int(csv_node.find(CSV_RECEIVER_EMAIL_INDEX).text)
    config[CSV_ATTACHMENT_FILE_INDEX] = int(csv_node.find(CSV_ATTACHMENT_FILE_INDEX).text)
    return config


def get_csv_file_path(csv_file_name: str):
    csv_file_path = f"{CURRENT_DIRECTORY}/{csv_file_name}"
    return csv_file_path


def get_attachments_folder_path(attachments_folder_name: str):
    attachments_folder_path = f"{CURRENT_DIRECTORY}/{attachments_folder_name}"
    return attachments_folder_path


def create_message(
        sender_email: str,
        receiver_email: str,
        subject: str,
        body: str,
        attachment_file_path: str):
    message = MIMEMultipart("alternative")
    message["Subject"] = subject
    message["From"] = sender_email
    message["To"] = receiver_email

    # Attach body
    message.attach(MIMEText(body, "plain"))

    # Attach file
    if attachment_file_path is not None:
        file_name = os.path.basename(attachment_file_path)
        payload = MIMEBase("application", "octet-stream", Name=file_name)
        with open(attachment_file_path, "rb") as binary_file:
            payload.set_payload(binary_file.read())

        encoders.encode_base64(payload)

        payload.add_header('Content-Decomposition', 'attachment', filename=file_name)
        message.attach(payload)

    return message


def send_mails():
    try:
        smtp_server = "localhost" if in_debug_mode() else "smtp.gmail.com"
        port = 8025 if in_debug_mode() else 587  # For starttls
        server = smtplib.SMTP(smtp_server, port)
        server.ehlo()
        if not in_debug_mode():
            context = ssl.create_default_context()
            server.starttls(context=context)  # Secure the connection
            server.ehlo()

        config = get_config()
        sender_email = config[SENDER_EMAIL]
        if not in_debug_mode():
            password = config[PASSWORD]
            server.login(sender_email, password)

        csv_file_path = get_csv_file_path(config[CSV_FILE_NAME])
        with open(csv_file_path, encoding="utf-8", mode="r") as fstream:
            reader = csv.reader(fstream, delimiter=',', quotechar='"')
            for idx, row in enumerate(reader):
                if idx == 0:
                    continue  # skip first row
                subject = config[SUBJECT]
                body = config[BODY]
                receiver_email = row[config[CSV_RECEIVER_EMAIL_INDEX]]
                attachment_file_name = None
                attachment_file_path = None
                if config[CSV_ATTACHMENT_FILE_INDEX] > -1 and row[config[CSV_ATTACHMENT_FILE_INDEX]]:
                    attachment_file_name = row[config[CSV_ATTACHMENT_FILE_INDEX]]
                    attachment_file_path = f"{get_attachments_folder_path(config[ATTACHMENTS_FOLDER_NAME])}/{attachment_file_name}"
                    if not os.path.exists(attachment_file_path):
                        attachment_file_name = None
                        attachment_file_path = None
                message = create_message(sender_email, receiver_email, subject, body, attachment_file_path)

                log(f"Sending email to '{receiver_email}' | Attached file: '{attachment_file_name}'")
                server.sendmail(sender_email, receiver_email, message.as_string())
    except Exception as ex:
        log_error(ex)
    finally:
        server.quit()


if __name__ == "__main__":
    send_mails()
