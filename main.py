import smtplib
from email.mime.text import MIMEText

def send_email(sender_email, sender_password, recipient_email, subject, html_message):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587

    message = MIMEText(html_message, 'html')
    message['Subject'] = subject
    message['From'] = sender_email
    message['To'] = recipient_email

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, message.as_string())
        print("Email berhasil dikirim ke:", recipient_email)
    except Exception as e:
        print("Terjadi kesalahan saat mengirim email ke:", recipient_email)
        print("Kesalahan:", str(e))
    finally:
        server.quit()

def read_email_database(filename):
    with open(filename, 'r') as file:
        emails = file.read().splitlines()
    return emails


sender_email = 'diesnataliesskomda@gmail.com'
sender_password = 'suksessukses'
subject = 'Contoh Email HTML'

database_filename = 'database_email.txt'
html_file = 'email.html'

recipient_emails = read_email_database(database_filename)

with open(html_file, 'r') as file:
    html_content = file.read()

for recipient_email in recipient_emails:
    send_email(sender_email, sender_password, recipient_email, subject, html_content)
