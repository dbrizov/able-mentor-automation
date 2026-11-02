import csv
import logging
import mimetypes
import os
import smtplib
import ssl
import sys
import xml.etree.ElementTree as ET
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# For a dry run, the emails are sent to a local server, which prints them instead of sending them.
# First you have to start an smtp server and then run the script with a '--dry-run' argument.
# 1. $ .venv\Scripts\python.exe -m aiosmtpd -n
# 2. $ python mail_sender.py --dry-run

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


def is_dry_run():
    return len(sys.argv) > 1 and sys.argv[1] == "--dry-run"


def log(message: str):
    print(message)
    logging.info(message)


def log_error(message: str):
    print(message)
    logging.error(message)


def get_text(parent_node, tag: str, required: bool = True) -> str:
    # An empty tag has no text (None), so it's read as ""
    text = (parent_node.find(tag).text or "").strip()
    if required and not text:
        raise ValueError(f"<{tag}> is empty in {CONFIG_FILE_NAME}")

    return text


def get_config():
    config = dict()
    root_node = ET.parse(CONFIG_FILE_PATH)

    # Login (not needed in a dry run)
    login_node = root_node.find("login")
    config[SENDER_EMAIL] = get_text(login_node, SENDER_EMAIL, required=not is_dry_run())
    config[PASSWORD] = get_text(login_node, PASSWORD, required=not is_dry_run())

    # Message
    message_node = root_node.find("message")
    config[SUBJECT] = get_text(message_node, SUBJECT)
    config[BODY] = get_text(message_node, BODY)
    config[ATTACHMENTS_FOLDER_NAME] = get_text(message_node, ATTACHMENTS_FOLDER_NAME)

    # CSV
    csv_node = root_node.find("csv")
    config[CSV_FILE_NAME] = get_text(csv_node, CSV_FILE_NAME)
    config[CSV_RECEIVER_NAME_INDEX] = int(get_text(csv_node, CSV_RECEIVER_NAME_INDEX))
    config[CSV_RECEIVER_EMAIL_INDEX] = int(get_text(csv_node, CSV_RECEIVER_EMAIL_INDEX))
    config[CSV_ATTACHMENT_FILE_INDEX] = int(get_text(csv_node, CSV_ATTACHMENT_FILE_INDEX))
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
        attachment_file_path: str | None):
    message = MIMEMultipart("mixed")
    message["Subject"] = subject
    message["From"] = sender_email
    message["To"] = receiver_email

    # Attach body
    message.attach(MIMEText(body, "plain", "utf-8"))

    # Attach file
    if attachment_file_path is not None:
        file_name = os.path.basename(attachment_file_path)
        mime_type = mimetypes.guess_type(file_name)[0] or "application/octet-stream"
        main_type, sub_type = mime_type.split("/", 1)
        payload = MIMEBase(main_type, sub_type)
        with open(attachment_file_path, "rb") as binary_file:
            payload.set_payload(binary_file.read())

        encoders.encode_base64(payload)

        payload.set_param("name", file_name, charset="utf-8")
        payload.add_header("Content-Disposition", "attachment", filename=("utf-8", "", file_name))
        message.attach(payload)

    return message


def send_mails():
    try:
        smtp_server = "localhost" if is_dry_run() else "smtp.gmail.com"
        port = 8025 if is_dry_run() else 587  # For starttls
        server = smtplib.SMTP(smtp_server, port)
        server.ehlo()
        if not is_dry_run():
            context = ssl.create_default_context()
            server.starttls(context=context)  # Secure the connection
            server.ehlo()

        config = get_config()
        sender_email = config[SENDER_EMAIL]
        if not is_dry_run():
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
                        log_error(f"NOT SENT to '{receiver_email}': missing attachment '{attachment_file_path}'")
                        continue

                message = create_message(sender_email, receiver_email, subject, body, attachment_file_path)

                log(f"Sending email to '{receiver_email}' | Attached file: '{attachment_file_name}'")
                server.sendmail(sender_email, receiver_email, message.as_string())
    except Exception as ex:
        log_error(ex)
    finally:
        server.quit()


if __name__ == "__main__":
    send_mails()
