from docx import Document
from docx2pdf import convert
import smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
from dotenv import load_dotenv
import pandas as pd
from os.path import exists
import time

load_dotenv()

EMAIL = os.getenv('EMAIL')
PASSWORD = os.getenv('PASSWORD')

def send_mail(sender, reciever, subject, body, password):
    message = MIMEMultipart()
    message["From"] = sender
    message["To"] = reciever
    message["Subject"] = subject

    message.attach(MIMEText(body, "plain"))

    filename = "Volunteering_Hours.pdf"

    with open(filename, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    encoders.encode_base64(part)
    
    part.add_header("Content-Disposition", f"attachment; filename= {filename}")

    message.attach(part)
    text = message.as_string()

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender, password)
        server.sendmail(sender, reciever, text)


document = Document("HOPE-Volunteering-Letter.docx")

def edit_doc(first_name, last_name, hours):
    paragraph6 = document.paragraphs[6]
    paragraph6.text = paragraph6.text.replace('<<First Name>>', first_name)
    paragraph6.text = paragraph6.text.replace('<<Last Name>>', last_name)
    paragraph6.text = paragraph6.text.replace('<<Hours>>', str(hours))
    document.save("Volunteering_Hours.docx")
    convert("Volunteering_Hours.docx")

def check_file_existence(text):
    while True:
        user_input = input(f"{text}")
        if exists(user_input):
            return user_input
        else:
            print("Invalid file path")


def determine_correct_pandas_conversion(file):
    extension = file.split('.')[1]
    if extension == 'xlsx':
        return pd.read_excel(file, sheet_name=0)
    elif extension == 'csv':
        return pd.read_csv(file)
    elif extension == 'tsv':
        return pd.read_csv(file, delimiter='\t')
    else:
        print(f'\nError: Invalid extension {extension}')


def main():
    print("Welcome to the Program")
    main_file = 'test.xlsx'
    if not exists('test.xlsx'):
        main_file = check_file_existence('\nEnter the main spreadsheet: ')

    subject = ''
    body = ''

    df = determine_correct_pandas_conversion(main_file)

    number_of_entry = len(df.index)
    
    for index, row in df.iterrows():
        #edit_doc(row['First Name'], row['Last Name'], row['Total Hours'])
        #send_mail(EMAIL, row['Email'], subject, body, password=PASSWORD)
        time.sleep(1)
        print(f" Progress: {(int(index/number_of_entry * 100))}%\r", end='')
    print("Complete       ")

if __name__ == "__main__":
    main()