import smtplib
from email.message import EmailMessage


SENDER_EMAIL = "name.email@gmail.com"
APP_PASSWORD = "password"

def send_mail_with_excel(recipient_email, subject, content, excel_file):
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = SENDER_EMAIL
        msg['To'] = recipient_email
        msg.set_content(content)

        with open(excel_file, 'rb') as f:
            file_data = f.read()
        msg.add_attachment(file_data, maintype="application", subtype="xls", filename=excel_file)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(SENDER_EMAIL, APP_PASSWORD)
            smtp.send_message(msg)
        