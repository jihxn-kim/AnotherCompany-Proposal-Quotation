import os
import smtplib
from ntpath import basename

from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.header import Header
from email.utils import formataddr

from config import GMAIL


def send_gmail(my_id, my_app_pass, send_to, subject, content, filenames):
    msg = MIMEMultipart()

    try:
        for file in filenames:
            with open(file, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
                encoders.encode_base64(part)

                # Base64로 파일 이름 인코딩
                encoded_filename = Header(basename(file), 'utf-8').encode()
                part.add_header('Content-Disposition', 'attachment', filename=encoded_filename)
                msg.attach(part)
    except Exception as e:
        print(e)

    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(my_id, my_app_pass)
    msg.attach(MIMEText(content, 'plain', 'utf-8'))
    msg['Subject'] = Header(subject, 'utf-8')
    msg['From'] = formataddr((str(Header('Sender Name', 'utf-8')), my_id))
    msg['To'] = ', '.join(send_to)

    s.sendmail(my_id, send_to, msg.as_string())
    s.quit()

def set_config():
    my_id = GMAIL['id']
    my_app_pass = GMAIL['app_pass']
    send_to = ['hun@anothercompany.co.kr', 'gkqjemwlgns@naver.com']
    subject = '이것은 제목입니다.'
    content = '이것은 메일의 내용입니다.'

    attach_path = os.path.dirname(os.path.realpath(__file__)) + '/attachments/'
    filenames = [os.path.join(attach_path, f) for f in os.listdir(attach_path)]

    print(my_id, send_to, subject, content, attach_path, filenames)
    return (my_id, my_app_pass, send_to, subject, content, attach_path, filenames)

if __name__ == "__main__":
    (my_id, my_app_pass, send_to, subject, content, attach_path, filenames) = set_config()
    send_gmail(my_id, my_app_pass, send_to, subject, content, filenames)