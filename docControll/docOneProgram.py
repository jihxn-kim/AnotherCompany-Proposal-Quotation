from pyhwpx import Hwp
import os
import time

# first_grade: 학년
# first_class: 반
# class_num: 차시
# directory_path: 양식 파일 경로
# save_path: 파일 저장 경로
# class_date: 날짜
# first_program_1: 프로그램 정보 (오전)
# first_program_2: 프로그램 정보 (오후)
# school_name: 학교명

class Doc:
    def __init__(self, first_grade, first_class, class_num, directory_path, save_path, class_date,
                 first_program_1, first_program_2,
                 school_name):
        self.first_grade = first_grade
        self.first_class = first_class
        self.class_num = class_num
        self.directory_path = directory_path
        self.save_path = save_path
        self.class_date = class_date
        self.first_program_1 = first_program_1
        self.first_program_2 = first_program_2
        self.school_name = school_name

    @staticmethod
    def to_raw(file_path):
        return fr"{file_path}"

    def makeDoc(self):
        full_directory_path = self.directory_path + "/oneProgram/"

        programPage = ["intro", "title", self.first_program_1 + "_소개", self.first_program_1 + "_교구"]

        if self.class_num != "6차시":
            programPage.append(self.first_program_1 + "_" + self.class_num)
        elif self.class_num == "6차시":
            programPage.append(self.first_program_1 + "_6차시_" + self.first_program_2)

        doc_list = []
        for i in programPage:
            program_file_path = full_directory_path + i + ".hwp"
            raw_file_path = self.to_raw(os.path.normpath(program_file_path))
            doc_list.append(raw_file_path)

        hwp = Hwp(new=True)

        curr_time = time.strftime("%y%m%d").strip()
        curr_year = time.strftime("%Y") + "년"
        curr_month = time.strftime("%m") + "월"
        curr_day = time.strftime("%d") + "일"

        # 문서 병합
        for i in range(0, len(doc_list)):
            if (i == 0):
                hwp.insert_file(doc_list[i], keep_section=True)
                hwp.MoveDocEnd()
                hwp.find_replace_all("(날짜)", f"{curr_year} {curr_month} {curr_day}")
            elif (i == 1):
                hwp.insert_file(doc_list[i], keep_section=True)
                hwp.MoveDocEnd()
            elif (i >= 3):
                hwp.insert_file(doc_list[i], move_doc_end=True)
            else:
                hwp.insert_file(doc_list[i], keep_section=False)
                hwp.MoveDocEnd()

        hwp.MoveDocBegin()
        hwp.DeletePage()

        # 텍스트 대체
        hwp.find_replace_all("(학교명)", self.school_name)
        hwp.find_replace_all("(날짜)", self.class_date)
        hwp.find_replace_all("(학년)", self.first_grade)
        hwp.find_replace_all("(학급)", self.first_class + "학급")
        hwp.find_replace_all("(차시)", self.class_num)
        hwp.find_replace_all("(프로그램명)", self.first_program_1)

        if (self.first_program_1 == "수상한스튜디오"):
            program = "비전욕망설계 프로그램"
        elif (self.first_program_1 == "어나더랜드"):
            program = "창업가정신 프로그램"
        elif (self.first_program_1 == "어나더비밀상담소"):
            program = "고교학점제 프로그램"
        elif (self.first_program_1 == "취업조작단"):
            program = "진로역량 발산 프로그램"
        elif (self.first_program_1 == "코드5"):
            program = "직무-직업 매칭 프로그램"
        elif (self.first_program_1 == "AI오피스"):
            program = "인공지능 프로그램"

        hwp.find_replace_all("(프로그램 종류)", program)

        curr_time = time.strftime("%y%m%d").strip()
        curr_year = time.strftime("%Y") + "년"
        curr_month = time.strftime("%m") + "월"

        save_dir = os.path.join(self.save_path, curr_year, curr_month, self.school_name)
        os.makedirs(save_dir, exist_ok=True)

        file_name = f"{curr_time}_{self.school_name}_{self.first_grade}_{self.first_program_1}_계획안_어나더컴퍼니.hwp"
        file_path = os.path.join(save_dir, file_name)

        # 한글파일 저장
        hwp.save_as(file_path)
        # hwp.save_pdf_as_image()
        hwp.quit()

        absolute_path = os.path.abspath(file_path)
        return absolute_path
    

