import os
import smtplib
from ntpath import basename
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.header import Header
from email.utils import formataddr

from mail.config import GMAIL

class Gmail:
    def __init__(self, filenames=None, grade_num=None, school_name=None, first_grade=None, first_program_1=None):
        self.my_id = GMAIL['id']
        self.my_app_pass = GMAIL['app_pass']
        self.send_to = ['gkqjemwlgns@naver.com']
        self.bcc = ['hun@anothercompany.co.kr']  # 숨김참조
        self.content = '이것은 메일의 내용입니다.'
        self.attach_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'attachments')
        self.filenames = filenames if filenames else []

        if grade_num == "1개 학년":
            self.subject = f"{school_name} {first_grade} {first_program_1} 계획안&견적서 송부 [어나더컴퍼니]"
        else:
            self.subject = f"{school_name} 시뮬레이션진로캠프 계획안&견적서 송부 [어나더컴퍼니]"

    def send_gmail(self):
        msg = MIMEMultipart()

        try:
            for file in self.filenames:
                with open(file, 'rb') as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                    encoders.encode_base64(part)

                    # Base64로 파일 이름 인코딩
                    encoded_filename = Header(basename(file), 'utf-8').encode()
                    part.add_header('Content-Disposition', 'attachment', filename=encoded_filename)
                    msg.attach(part)
        except Exception as e:
            print(f"파일 첨부 오류: {e}")

        s = smtplib.SMTP('smtp.gmail.com', 587)
        s.starttls()
        s.login(self.my_id, self.my_app_pass)
        msg.attach(MIMEText(self.content, 'plain', 'utf-8'))
        msg['Subject'] = Header(self.subject, 'utf-8')
        msg['From'] = formataddr((str(Header('Sender Name', 'utf-8')), self.my_id))
        msg['To'] = ', '.join(self.send_to)

        # 숨김참조(BCC) 추가
        msg['Bcc'] = ', '.join(self.bcc)

        s.sendmail(self.my_id, self.send_to + self.bcc, msg.as_string())
        s.quit()