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

class Gmail:
    def __init__(self, filenames=None, grade_num=None, school_name=None, email_set=None, year=None,
                 email=None, teacher=None, class_date=None,
                 first_program_1=None, first_grade=None, first_class=None, class_num_1=None,
                 second_program_1=None, second_grade=None, second_class=None, class_num_2=None,
                 last_program_1=None, last_grade=None, last_class=None, class_num_3=None,
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
            program_1 = "비전욕망설계 프로그램"
        elif (first_program_1 == "어나더랜드"):
            program_1 = "창업가정신 프로그램"
        elif (first_program_1 == "어나더비밀상담소"):
            program_1 = "고교학점제 프로그램"
        elif (first_program_1 == "취업조작단"):
            program_1 = "진로역량 발산 프로그램"
        elif (first_program_1 == "코드5"):
            program_1 = "직무-직업 매칭 프로그램"
        elif (first_program_1 == "AI오피스"):
            program_1 = "인공지능 프로그램"

        if (second_program_1 == "수상한스튜디오"):
            program_2 = "비전욕망설계 프로그램"
        elif (second_program_1 == "어나더랜드"):
            program_2 = "창업가정신 프로그램"
        elif (second_program_1 == "어나더비밀상담소"):
            program_2 = "고교학점제 프로그램"
        elif (second_program_1 == "취업조작단"):
            program_2 = "진로역량 발산 프로그램"
        elif (second_program_1 == "코드5"):
            program_2 = "직무-직업 매칭 프로그램"
        elif (second_program_1 == "AI오피스"):
            program_2 = "인공지능 프로그램"

        if (last_program_1 == "수상한스튜디오"):
            program_3 = "비전욕망설계 프로그램"
        elif (last_program_1 == "어나더랜드"):
            program_3 = "창업가정신 프로그램"
        elif (last_program_1 == "어나더비밀상담소"):
            program_3 = "고교학점제 프로그램"
        elif (last_program_1 == "취업조작단"):
            program_3 = "진로역량 발산 프로그램"
        elif (last_program_1 == "코드5"):
            program_3 = "직무-직업 매칭 프로그램"
        elif (last_program_1 == "AI오피스"):
            program_3 = "인공지능 프로그램"

        # 내용 - 대상 & 내용
        if (grade_num == "1개 학년"):
            target = f"{first_grade}({first_class}반)"
            content = f"&nbsp;&nbsp;[{first_program_1} / {program_1} / {class_num_1}]"
        elif (grade_num == "2개 학년"):
            target = f"{first_grade}({first_class}반) / {second_grade}({second_class}반)"
            content = f"&nbsp;&nbsp;[{first_program_1} / {program_1} / {class_num_1}]<br/>&nbsp;&nbsp;[{second_program_1} / {program_2} / {class_num_2}]"
        elif (grade_num == "3개 학년"):
            target = f"{first_grade}({first_class}반) / {second_grade}({second_class}반) / {last_grade}({last_class}반)"
            content = f"&nbsp;&nbsp;[{first_program_1} / {program_1} / {class_num_1}]<br/>&nbsp;&nbsp;[{second_program_1} / {program_2} / {class_num_2}]<br/>&nbsp;&nbsp;[{last_program_1} / {program_3} / {class_num_3}]"

        if email_set == "신규":
            self.content = textwrap.dedent(f"""\
                <html>
                <body>
                    <p style="font-size: 14px; font-family: Gulim, Arial, sans-serif;">
                        안녕하세요, {teacher} 부장님<br/>
                        어나더컴퍼니 {manager}매니저입니다.
                        <br/><br/>

                        이번에 {school_name} 학생들과 첫 진로수업을 함께할 수 있게 되어 매우 기쁘게 생각합니다.
                        <br/><br/>

                        어나더컴퍼니는 학생들이 스토리의 주인공이 되어 다양한 미션(활동)을 수행하며 자신의 진로를 찾아가는 시뮬레이션 형태의 수업입니다.<br/>
                        <strong>전문성우, 영화감독, 디자이너, 교육전문가가 콜라보레이션해서 만든 콘텐츠</strong>입니다. (특허청 상표출원 제40 2020-004000005호)
                        <br/><br/>

                        진행 될 프로그램의 내용을 요약하면 아래와 같습니다.
                        <br/><br/>

                        <strong>* {school_name} 내용 요약</strong>
                        <br/>
                        - 날짜 : {class_date}<br/>
                        - 대상 : {target}<br/>
                        - 내용<br/>
                            {content}<br/>
                        - 메일 첨부 파일 : 계획안, 견적서, 사업자등록증 & 통장사본
                        <br/><br/>

                        부장님께서 저희를 선택해 주신 만큼,  첫 시작을 좋은 인연으로 이어갈 수 있도록  꼼꼼하게 준비하겠습니다.<br/>
                        수업 진행과 관련하여 더 필요한 사항이 있거나 요청사항이 있으시면 언제든 연락주세요.
                        <br/><br/>

                        감사합니다.
                        <br/><br/>

                        <strong>{manager}</strong> 교육운영팀 / 매니저<br/>
                        어나더컴퍼니 (AnotheR company)
                        <br/><br/>

                        <strong>Kakao.</strong> <a href="http://pf.kakao.com/_qxlKeb" target="_blank">http://pf.kakao.com/_qxlKeb</a><br/>
                        <strong>Tel.</strong> 02-6953-1718<br/>
                        <strong>Site.</strong> <a href="http://www.anothercompany.co.kr" target="_blank">www.anothercompany.co.kr</a>
                    </p>
                </body>
                </html>
            """)
        elif email_set == "전학년":
            self.content = textwrap.dedent(f"""\
                <html>
                <body>
                    <p style="font-size: 14px; font-family: Gulim, Arial, sans-serif;">
                        안녕하세요, {teacher} 부장님<br/>
                        어나더컴퍼니 {manager}매니저입니다.
                        <br/><br/>

                        이번 {school_name} 전 학년 진로수업을 어나더컴퍼니에 맡겨주셔서 진심으로 감사드립니다.<br/>
                        전학년 수업을 진행하는 만큼 더욱 신뢰에 보답할 수 있도록 꼼꼼하게 준비하겠습니다.
                        <br/><br/>

                        진행 될 프로그램에 대한 내용은 아래와 같습니다.
                        <br/><br/> 

                        <strong>* {school_name} 내용 요약</strong>
                        <br/>
                        - 날짜 : {class_date}<br/>
                        - 대상 : {target}<br/>
                        - 내용<br/>
                            {content}<br/>
                        - 메일 첨부 파일 : 계획안, 견적서, 사업자등록증 & 통장사본
                        <br/><br/>

                        운영 일정이 다가오면 사전 체크를 위해 연락드리겠습니다.<br/>
                        학교에서 큰 결정을 내려주신 만큼, 학생들에게 최고의 경험을 할 수 있도록 최선을 다해 준비하겠습니다.
                        <br/><br/>

                        궁금하신 사항이나 문의사항이 있으시면 언제든 편하게 연락주세요.
                        <br/><br/>

                        감사합니다.
                        <br/><br/>

                        <strong>{manager}</strong> 교육운영팀 / 매니저<br/>
                        어나더컴퍼니 (AnotheR company)
                        <br/><br/>

                        <strong>Kakao.</strong> <a href="http://pf.kakao.com/_qxlKeb" target="_blank">http://pf.kakao.com/_qxlKeb</a><br/>
                        <strong>Tel.</strong> 02-6953-1718<br/>
                        <strong>Site.</strong> <a href="http://www.anothercompany.co.kr" target="_blank">www.anothercompany.co.kr</a>
                    </p>
                </body>
                </html>
            """)
        elif email_set == "재구매":
            self.content = textwrap.dedent(f"""\
                <html>
                <body>
                    <p style="font-size: 14px; font-family: Gulim, Arial, sans-serif;">
                        안녕하세요, {teacher} 부장님<br/>
                        어나더컴퍼니 {manager}매니저입니다.
                        <br/><br/>

                        매년 저희를 믿고 함께해주셔서 감사합니다.<br/>
                        부장님의 꾸준한 신뢰에 깊이 감사드리며, {year}년에도 변함없이 최고의 진로수업으로 보답하겠습니다!
                        <br/><br/>

                        진행 될 프로그램에 대한 내용은 아래와 같습니다.
                        <br/><br/> 

                        <strong>* {school_name} 내용 요약</strong>
                        <br/>
                        - 날짜 : {class_date}<br/>
                        - 대상 : {target}<br/>
                        - 내용<br/>
                            {content}<br/>
                        - 메일 첨부 파일 : 계획안, 견적서, 사업자등록증 & 통장사본
                        <br/><br/>

                        운영 일정이 다가오면 사전 체크를 위해 연락드리겠습니다.<br/>
                        올해도 믿고 맡겨주신 만큼 학생들에게 최고의 경험이 될 수 있도록 더욱 꼼꼼하게 준비하겠습니다!<br/>
                        수업 진행과 관련하여 더 필요한 사항이 있거나 요청사항이 있으시면 언제든 연락주세요.
                        <br/><br/>

                        감사합니다.
                        <br/><br/>

                        <strong>{manager}</strong> 교육운영팀 / 매니저<br/>
                        어나더컴퍼니 (AnotheR company)
                        <br/><br/>

                        <strong>Kakao.</strong> <a href="http://pf.kakao.com/_qxlKeb" target="_blank">http://pf.kakao.com/_qxlKeb</a><br/>
                        <strong>Tel.</strong> 02-6953-1718<br/>
                        <strong>Site.</strong> <a href="http://www.anothercompany.co.kr" target="_blank">www.anothercompany.co.kr</a>
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