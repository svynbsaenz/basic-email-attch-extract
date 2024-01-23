import datetime
import imaplib
import email
from email.header import decode_header
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import os
import time
import getpass  
import argparse
import re

def get_email_credentials():
    # Input Gmail credentials
    email_address = input("Enter your Gmail address: ")
    # password = getpass.getpass("Enter your Gmail password: ")
    app_password = getpass.getpass("Enter your Gmail App Password: ")
    # return email_address, password
    return email_address, app_password

def connect_to_gmail(email_address, password):
    # Connect to Gmail
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(email_address, password)
    return mail

def get_senders_list(sender_list_file):
    # Read the list of senders from the file
    with open(sender_list_file, "r") as file:
        senders = [line.strip() for line in file if line.strip()]
    return senders

def extract_email_add(email_sender):
    pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'

    matches = re.findall(pattern, email_sender)
    return matches[0]

def process_emails(mail, senders, save_directory):
    # Select the inbox
    mail.select("inbox")

    today = datetime.datetime.today()
    date_since = today.strftime('%d-%b-%Y')

    # Search for all unread emails
    result, data = mail.search(None, f'UNSEEN SINCE "{date_since}"')
    email_ids = data[0].split()

    for email_id in email_ids:
        result, message_data = mail.fetch(email_id, "(RFC822)")
        raw_email = message_data[0][1]
        msg = email.message_from_bytes(raw_email)

        sender = extract_email_add(msg.get("From"))
        subject = decode_header(msg.get("Subject"))[0][0]

        if sender in senders:
            print("New email received..")
            print(f"Processing email from {sender}: {subject}")
            process_attachments(msg, save_directory)
            send_confirmation_email(mail, sender)

def process_attachments(msg, save_directory):
    # Extract attachments and save to the specified folder
    date_str = time.strftime("%Y%m%d_%H%M%S")
    folder_name = os.path.join(save_directory, f"{date_str}")
    os.makedirs(folder_name)

    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        if part.get("Content-Disposition") is None:
            print(f'The email sent on {time.strftime("%Y%m%d")} at {time.strftime("%H%M%S")} has no attachments..')
            continue

        filename = part.get_filename()
        if filename:
            filepath = os.path.join(folder_name, filename)
            with open(filepath, "wb") as f:
                f.write(part.get_payload(decode=True))
            print(f"Attachment saved: {filename}")

def send_confirmation_email(mail, sender):
    # Send confirmation email
    subject = "Confirmation: Email Received"
    body = "Thank you for your email. We will get back to you as soon as possible."

    confirmation_email = MIMEMultipart()
    confirmation_email["From"] = email_address
    confirmation_email["To"] = sender
    confirmation_email["Subject"] = subject
    confirmation_email.attach(MIMEText(body, "plain"))

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(email_address, app_password)
        server.sendmail(email_address, sender, confirmation_email.as_string())
        print(f"Confirmation email sent to {sender}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Automate email extraction and confirmation")
    parser.add_argument("sender_list_file", help="Path to the text file containing the list of senders")
    parser.add_argument("save_directory", help="Path to the directory for saving extracted attachments")
    args = parser.parse_args()

    email_address, app_password = get_email_credentials()
    print("Logging in..")
    mail = connect_to_gmail(email_address, app_password)
    print("Loging successful. Getting senders..")
    senders = get_senders_list(args.sender_list_file)
    print("Senders are noted..")
    print("Waiting for new emails..")

    while True:
        try:
            process_emails(mail, senders, args.save_directory)
        except Exception as e:
            print(f"An error occurred: {e}")


        time.sleep(120)
