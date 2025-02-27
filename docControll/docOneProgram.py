from pyhwpx import Hwp
import os
import time

# first_grade: 학년
# first_class: 반
# class_num: 차시
# directory_path: 양식 파일 경로
# first_program_1: 프로그램 정보 (오전)
# first_program_2: 프로그램 정보 (오후)
# school_name: 학교명
# price: 반당 금액

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

        programPage = ["title", self.first_program_1 + "_소개", self.first_program_1 + "_교구"]

        if self.class_num == "3차시":
            programPage.append(self.first_program_1 + "_3차시")
        elif self.class_num == "4차시":
            programPage.append(self.first_program_1 + "_4차시")
        elif self.class_num == "6차시":
            programPage.append(self.first_program_1 + "_6차시_" + self.first_program_2)

        doc_list = []
        for i in programPage:
            program_file_path = full_directory_path + i + ".hwp"
            raw_file_path = self.to_raw(os.path.normpath(program_file_path))
            doc_list.append(raw_file_path)

        hwp = Hwp(new=True)

        # 문서 병합
        for i in range(0, len(doc_list)):
            if (i >= 2):
                hwp.insert_file(doc_list[i], move_doc_end=True)
            else:
                hwp.insert_file(doc_list[i], keep_section=False)
                hwp.MoveDocEnd()

        # 텍스트 대체
        hwp.find_replace_all("(학교명)", self.school_name)
        hwp.find_replace_all("(날짜)", self.class_date)
        hwp.find_replace_all("(학년)", self.first_grade)
        hwp.find_replace_all("(학급)", self.first_class + "학급")
        hwp.find_replace_all("(차시)", self.class_num)
        hwp.find_replace_all("(프로그램명)", self.first_program_1)

        if (self.first_program_1 == "수상한스튜디오"):
            program = "비전욕망설계 캠프"
        elif (self.first_program_1 == "어나더랜드"):
            program = "창업가정신 캠프"
        elif (self.first_program_1 == "비밀상담소"):
            program = "고교학점제 캠프"

        hwp.find_replace_all("(프로그램 종류)", program)

        curr_time = time.strftime("%y%m%d").strip()
        file_name = self.save_path + "/" + curr_time + "_" + self.school_name + "_" + self.first_grade + "_" + self.first_program_1 + "_계획안_어나더컴퍼니.hwp"

        # 한글파일 저장
        hwp.save_as(file_name)
        # hwp.save_pdf_as_image()
        hwp.quit()

        absolute_path = os.path.join(os.getcwd(), file_name)
        return absolute_path
    

