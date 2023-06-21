import smtplib, os # Import modul smtplib (untuk SMTP protocol) dan os (untuk akses sistem operasi)
import pandas as pd # Import pandas (untuk manipulasi data)
from dotenv import load_dotenv # Import fungsi load_dotenv dari modul dotenv (untuk mengambil data dari file .env)
from email.mime.text import MIMEText # Import MIMEText (untuk membuat email dengan format MIME)
from string import Template # Import Template dari string (untuk substitusi string)
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Mengambil data dari file .env.development dan memasukkannya ke variabel lingkungan
env_path = os.path.join(os.path.dirname(__file__), '.env.development') # Ubah nama file ke .env.production waktu udah fix
load_dotenv(env_path)

def send_email(sender_email, sender_password, recipient_email, recipient_company, subject, html_message_template, attachment_path):
    smtp_server = os.getenv('smtp_server')  # Mengambil server SMTP dari variabel lingkungan
    smtp_port = int(os.getenv('smtp_port'))  # Mengambil port SMTP dari variabel lingkungan

     # Membuat pesan email dengan format HTML dan melakukan substitusi nama perusahaan
    html_message = Template(html_message_template).safe_substitute(nama_pt=recipient_company)
    message = MIMEText(html_message, 'html')

    email = MIMEMultipart() #Buat Attachment
    message['Subject'] = subject
    message['From'] = sender_email
    message['To'] = recipient_email

    with open(attachment_path, 'rb') as attachment:
        mime_attachment = MIMEBase('application', 'octet-stream')
        mime_attachment.set_payload(attachment.read())
        encoders.encode_base64(mime_attachment)
        mime_attachment.add_header('Content-Disposition', f'attachment; filename={attachment_path}')

    email.attach(mime_attachment)
    email.attach(message)

    # Mencoba mengirim email
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)  # Membuka koneksi ke server SMTP
        server.starttls()  # Mulai sesi TLS
        server.login(sender_email, sender_password)  # Login ke server SMTP
        server.sendmail(sender_email, recipient_email, message.as_string())  # Mengirim email
        print("Email berhasil dikirim ke:", recipient_email)
    except Exception as e:  # Jika terjadi kesalahan, tampilkan pesan kesalahan
        print("Terjadi kesalahan saat mengirim email ke:", recipient_email)
        print("Kesalahan:", str(e))
    finally:
        if 'server' in locals(): # Pastikan koneksi telah dibuka sebelum mencoba menutupnya
            server.quit()

# Fungsi untuk membaca database email dari file Excel
def read_email_database(filename):
    dataframe = pd.read_excel(filename)  # Membaca file Excel dan memasukkannya ke DataFrame
    data = list(dataframe.itertuples(index=False, name=None))  # Mengubah DataFrame menjadi list dari tuple
    return data  # Mengembalikan data

# Mengambil data pengirim dan password dari variabel lingkungan
sender_email = os.getenv('sender_email')
sender_password = os.getenv('sender_password')

subject = 'Contoh Email HTML' #Subjek email

# Nama file untuk database email dan template email tambahan attachment
database_filename = 'database_email.xlsx'
html_file = 'email.html'
attachment_path = 'path/ke/file_attachment.pdf'

# Membaca excel database email
recipient_emails = read_email_database(database_filename)

# Membaca template email
html_message_template = open(html_file, 'r').read()

# Mengirim email ke setiap penerima di database
# File Excel nya memiliki tiga kolom untuk alamat email, nama perusahaan, dan status.
for recipient_email, recipient_company, status in recipient_emails:
    send_email(sender_email, sender_password, recipient_email, recipient_company, subject, html_message_template, attachment_path)
