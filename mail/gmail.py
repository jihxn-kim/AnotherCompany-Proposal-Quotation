import os
import smtplib
import textwrap
from ntpath import basename
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.header import Header
from email.utils import formataddr

from mail.config import GMAIL

class Gmail:
    def __init__(self, filenames=None, grade_num=None, school_name=None,
                 email=None, teacher=None, class_date=None, class_num=None,
                 first_program_1=None, first_grade=None, first_class=None,
                 second_program_1=None, second_grade=None, second_class=None,
                 last_program_1=None, last_grade=None, last_class=None,
                 id=None, app_pass=None, manager=None, bcc=None):
        self.my_id = id
        self.my_app_pass = app_pass
        self.send_to = [email]
        self.bcc = bcc # 숨김참조
        self.attach_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'attachments')
        self.filenames = filenames if filenames else []

        # 제목
        if grade_num == "1개 학년":
            self.subject = f"[어나더컴퍼니] {school_name} {first_grade} {first_program_1} 계획안 및 견적서 송부드립니다."
        else:
            self.subject = f"[어나더컴퍼니] {school_name} 시뮬레이션진로캠프 계획안 및 견적서 송부드립니다."

        # 프로그램 종류 치환
        if (first_program_1 == "수상한스튜디오"):
            program_1 = "비전욕망설계 캠프"
        elif (first_program_1 == "어나더랜드"):
            program_1 = "창업가정신 캠프"
        elif (first_program_1 == "비밀상담소"):
            program_1 = "고교학점제 캠프"

        if (second_program_1 == "수상한스튜디오"):
            program_2 = "비전욕망설계 캠프"
        elif (second_program_1 == "어나더랜드"):
            program_2 = "창업가정신 캠프"
        elif (second_program_1 == "비밀상담소"):
            program_2 = "고교학점제 캠프"

        if (last_program_1 == "수상한스튜디오"):
            program_3 = "비전욕망설계 캠프"
        elif (last_program_1 == "어나더랜드"):
            program_3 = "창업가정신 캠프"
        elif (last_program_1 == "비밀상담소"):
            program_3 = "고교학점제 캠프"

        # 내용 - 대상 & 내용
        if (grade_num == "1개 학년"):
            target = f"{first_grade} ({first_class}반)"
            content = f"&nbsp;&nbsp;[{first_program_1} / {program_1} / {class_num}]"
        elif (grade_num == "2개 학년"):
            target = f"{first_grade} ({first_class}반), {second_grade} ({second_class}반)"
            content = f"&nbsp;&nbsp;{first_grade} : [{first_program_1} / {program_1} / {class_num}]<br/>&nbsp;&nbsp;{second_grade} : [{second_program_1} / {program_2} / {class_num}]"
        elif (grade_num == "3개 학년"):
            target = f"{first_grade} ({first_class}반), {second_grade} ({second_class}반), {last_grade} ({last_class}반)"
            content = f"&nbsp;&nbsp;{first_grade} : [{first_program_1} / {program_1} / {class_num}]<br/>&nbsp;&nbsp;{second_grade} : [{second_program_1} / {program_2} / {class_num}]<br/>&nbsp;&nbsp;{last_grade} : [{last_program_1} / {program_3} / {class_num}]"

        self.content = textwrap.dedent(f"""\
            <html>
            <body>
                <p style="font-size: 14px; font-family: Gulim, Arial, sans-serif;">
                    안녕하세요, {teacher} 부장님~<br/>
                    어나더컴퍼니 {manager} 매니저입니다.
                    <br/><br/>

                    이렇게 인사드리게되어 반갑습니다 :)<br/>
                    저희 프로그램들은 학생들이 스토리의 주인공이 되어 다양한 미션(활동)을 수행하며 자신의 진로를 찾아가는 시뮬레이션 형태의 수업입니다.<br/>
                    <strong>전문성우, 영화감독, 디자이너, 교육전문가가 콜라보레이션해서 만든 콘텐츠</strong>입니다. (특허청 상표출원 제40 2020-004000005호)
                    <br/><br/>

                    진행될 프로그램의 내용을 요약하면 아래와 같습니다.
                    <br/><br/>

                    <strong>* {school_name} 내용 요약</strong>
                    <br/>
                        - 날짜 : {class_date}<br/>
                        - 대상 : {target}<br/>
                        - 내용<br/>
                        {content}
                    <br/><br/>

                    프로그램 운영에 궁금한 부분이나 협의해야 하는 부분이 있으시다면 메일로 답장해주세요~<br/>
                    친절히 답변해드리도록 하겠습니다.<br/>
                    믿고 의뢰해주신만큼 좋은 프로그램과 강사진으로 보답하겠습니다..!
                    <br/><br/>
                    감사합니다.
                </p>
            </body>
            </html>
        """)
        
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
        msg.attach(MIMEText(self.content, 'html', 'utf-8'))
        msg['Subject'] = Header(self.subject, 'utf-8')
        msg['From'] = formataddr((str(Header('어나더컴퍼니', 'utf-8')), self.my_id))
        msg['To'] = ', '.join(self.send_to)

        # 숨김참조(BCC) 추가
        msg['Bcc'] = ', '.join(self.bcc)

        s.sendmail(self.my_id, self.send_to + self.bcc, msg.as_string())
        s.quit()