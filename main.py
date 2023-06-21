import smtplib, os # Import modul smtplib (untuk SMTP protocol) dan os (untuk akses sistem operasi)
import pandas as pd # Import pandas (untuk manipulasi data)
from dotenv import load_dotenv # Import fungsi load_dotenv dari modul dotenv (untuk mengambil data dari file .env)
from email.mime.text import MIMEText # Import MIMEText (untuk membuat email dengan format MIME)
from string import Template # Import Template dari string (untuk substitusi string)
from email.mime.multipart import MIMEMultipart # Import MIMEMultipart (untuk membuat email multi-part)
from email.mime.base import MIMEBase # Import MIMEBase (untuk membuat attachment)
from email import encoders # Import encoders (untuk encoding attachment)
from tqdm import tqdm # Import tqdm (untuk progress bar)

# Mengambil data dari file .env.development dan memasukkannya ke variabel lingkungan
env_path = os.path.join(os.path.dirname(__file__), '.env.development') # Ubah nama file ke .env.production waktu udah fix
load_dotenv(env_path) # Load file .env

def send_email(sender_name, sender_email, sender_password, recipient_email, recipient_company, subject, html_message_template, attachment_path):
    smtp_server = os.getenv('smtp_server')  # Mengambil server SMTP dari variabel lingkungan
    smtp_port = int(os.getenv('smtp_port'))  # Mengambil port SMTP dari variabel lingkungan

    # Membuat pesan email dengan format HTML dan melakukan substitusi nama perusahaan
    html_message = Template(html_message_template).safe_substitute(nama_pt=recipient_company)

    message = MIMEMultipart() #Buat email multi-part (supaya bisa mengirim attachment)
    message['Subject'] = subject
    message['From'] = sender_name
    message['To'] = recipient_email

    # Tambahkan body emailnya
    message.attach(MIMEText(html_message, 'html'))

    # Baca file PDF di mode binary
    with open(attachment_path, 'rb') as attachment:
        mime_attachment = MIMEBase('application', 'octet-stream')
        mime_attachment.set_payload(attachment.read())

        # Encode file PDF ke base64 agar bisa dikirim lewat email
        encoders.encode_base64(mime_attachment)

        # Tambahkan header agar file PDF bisa diunduh
        proposal_filename = os.path.basename(attachment_path)
        mime_attachment.add_header('Content-Disposition', f'attachment; filename={proposal_filename}')

    # Tambahkan attachment ke email
    message.attach(mime_attachment)

    # Mencoba mengirim email
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)  # Membuka koneksi ke server SMTP
        server.starttls()  # Mulai sesi TLS
        server.login(sender_email, sender_password)  # Login ke server SMTP
        server.sendmail(sender_email, recipient_email, message.as_string())  # Mengirim email
        print("\nEmail berhasil dikirim ke:", recipient_email)
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

# Input kustom oleh user. Tapi kalau user tekan enter, pakai nilai default
sender_name = input("Ketik nama pengirim atau tekan enter untuk default [OSIS SMK Telkom Sidoarjo]: ") or 'OSIS SMK Telkom Sidoarjo'
subject = input("Ketik subjek email untuk dikirm atau tekan enter untuk default [Contoh Email HTML]: ") or 'Contoh Email HTML'
html_file = input("Ketik nama file template html atau tekan enter untuk default [email.html]: ") or 'email.html'
attachment_path = input("Ketik nama file attachment atau tekan enter untuk default [Proposal.pdf]: ") or 'Proposal.pdf'
database_filename = input("Ketik nama file Excel atau tekan enter untuk default [database_email.xlsx]: ") or 'database_email.xlsx'

# Membaca excel database email
recipient_emails = read_email_database(database_filename)

# Membaca template email
with open(html_file, 'r') as f:
    html_message_template = f.read()

# Konfirmasi ke user sebelum mengirim email, yakin kirim atau tidak
execute_confirmation = input('\nApakah Anda yakin ingin mengirim email ke semua penerima? (y/n) ')
if execute_confirmation == 'y':
    os.system('cls' if os.name == 'nt' else 'clear')  # Membersihkan layar
    print('Email pengirim       : ' + sender_email)
    print('Nama pengirim        : ' + sender_name)
    print('Subjek               : ' + subject)
    print('Mengirim email ke    : ' + str(len(recipient_emails)) + ' alamat email\n')

# Mengirim email ke setiap penerima di database
# File Excel nya memiliki tiga kolom untuk alamat email, nama perusahaan, dan status.
    for recipient_email, recipient_company, status in tqdm(recipient_emails, desc='Progress Pengiriman', unit='email'): # TQDM untuk progress bar
        send_email(sender_name, sender_email, sender_password, recipient_email, recipient_company, subject, html_message_template, attachment_path)
else:
    print('Pengiriman email dibatalkan')