from pyhwpx import Hwp
import os
import time

# directory_path: 양식 파일 경로
# save_path: 파일 저장 경로
# class_date: 날짜
# first_program_1: 첫 번째 프로그램 정보 (오전)
# first_program_2: 첫 번째 프로그램 정보 (오후)
# first_grade: 첫 번째 학년
# first_class: 첫 번째 반
# class_num_1: 첫 번째 차시
# second_program_1: 두 번째 프로그램 정보 (오전)
# second_program_2: 두 번째 프로그램 정보 (오후)
# second_grade: 두 번째 학년
# second_class: 두 번째 반
# class_num_2: 두 번째 차시
# school_name: 학교명

class Category:
    def __init__(self, directory_path, save_path, class_date,
                 first_program_1, first_program_2, first_grade, first_class, class_num_1,
                 second_program_1, second_program_2, second_grade, second_class, class_num_2,
                 school_name):
        self.directory_path = directory_path
        self.save_path = save_path
        self.class_date = class_date

        self.first_program_1 = first_program_1
        self.first_program_2 = first_program_2
        self.first_grade = first_grade
        self.first_class = first_class
        self.class_num_1 = class_num_1

        self.second_program_1 = second_program_1
        self.second_program_2 = second_program_2
        self.second_grade = second_grade
        self.second_class = second_class
        self.class_num_2 = class_num_2

        self.school_name = school_name

    @staticmethod
    def to_raw(file_path):
        return fr"{file_path}"

    def makeDoc(self):
        full_directory_path = self.directory_path + "/twoProgram/"

        programPage = [
            "title2", 
            self.first_program_1 + "_소개", self.second_program_1 + "_소개", "성과",
            self.first_program_1 + "_성과", self.second_program_1 + "_성과",
            ]

        programPage.append(self.first_program_1 + "_교구")

        if self.class_num_1 == "6차시":
            programPage.append(self.first_program_1 + "_6차시_" + self.first_program_2)
        else:
            programPage.append(self.first_program_1 + "_" + self.class_num_1)

        programPage.append(self.second_program_1 + "_교구")

        if self.class_num_2 == "6차시":
            programPage.append(self.second_program_1 + "_6차시_" + self.second_program_2)
        else:
            programPage.append(self.second_program_1 + "_" + self.class_num_2)

        doc_list = []
        for i in programPage:
            program_file_path = full_directory_path + i + ".hwp"
            raw_file_path = self.to_raw(os.path.normpath(program_file_path))
            doc_list.append(raw_file_path)

        hwp = Hwp(new=True)

        # 문서 병합
        for i in range(0, len(doc_list)):
            if (i <= 2):
                hwp.insert_file(doc_list[i], keep_section=False)
                hwp.MoveDocEnd()
            elif (i == 3):
                hwp.insert_file(doc_list[i], move_doc_end=True)
            elif (i <= 5):
                hwp.insert_file(doc_list[i], keep_section=False)
                hwp.MoveDocEnd()
            elif (i == 6):
                hwp.insert_file(doc_list[i], move_doc_end=True)
                hwp.find_replace_all("(학년)", self.first_grade)
            elif (i == 8):
                hwp.insert_file(doc_list[i], move_doc_end=True)
                hwp.find_replace_all("(학년)", self.second_grade)
            else:
                hwp.insert_file(doc_list[i], move_doc_end=True)

        # 텍스트 대체
        hwp.find_replace_all("(학교명)", self.school_name)
        hwp.find_replace_all("(날짜)", self.class_date)
        hwp.find_replace_all("(학년1)", self.first_grade)
        hwp.find_replace_all("(학급1)", self.first_class + "학급")
        hwp.find_replace_all("(학년2)", self.second_grade)
        hwp.find_replace_all("(학급2)", self.second_class + "학급")
        hwp.find_replace_all("(프로그램명1)", self.first_program_1)
        hwp.find_replace_all("(프로그램명2)", self.second_program_1)

        if (self.first_program_1 == "수상한스튜디오"):
            program_1 = "비전욕망설계 프로그램"
        elif (self.first_program_1 == "어나더랜드"):
            program_1 = "창업가정신 프로그램"
        elif (self.first_program_1 == "어나더비밀상담소"):
            program_1 = "고교학점제 프로그램"
        elif (self.first_program_1 == "취업조작단"):
            program_1 = "진로역량 발산 프로그램"
        elif (self.first_program_1 == "코드5"):
            program_1 = "직무-직업 매칭 프로그램"
        elif (self.first_program_1 == "AI오피스"):
            program_1 = "인공지능 프로그램"

        if (self.second_program_1 == "수상한스튜디오"):
            program_2 = "비전욕망설계 프로그램"
        elif (self.second_program_1 == "어나더랜드"):
            program_2 = "창업가정신 프로그램"
        elif (self.second_program_1 == "어나더비밀상담소"):
            program_2 = "고교학점제 프로그램"
        elif (self.second_program_1 == "취업조작단"):
            program_2 = "진로역량 발산 프로그램"
        elif (self.second_program_1 == "코드5"):
            program_2 = "직무-직업 매칭 프로그램"
        elif (self.second_program_1 == "AI오피스"):
            program_2 = "인공지능 프로그램"

        hwp.find_replace_all("(프로그램 종류1)", program_1)
        hwp.find_replace_all("(프로그램 종류2)", program_2)

        curr_time = time.strftime("%y%m%d").strip()
        curr_year = time.strftime("%Y") + "년"
        curr_month = time.strftime("%m") + "월"

        save_dir = os.path.join(self.save_path, curr_year, curr_month, self.school_name)
        os.makedirs(save_dir, exist_ok=True)

        file_name = f"{curr_time}_{self.school_name}_시뮬레이션진로캠프_계획안_어나더컴퍼니.hwp"
        file_path = os.path.join(save_dir, file_name)

        # 한글파일 저장
        hwp.save_as(file_path)
        # hwp.save_pdf_as_image()
        hwp.quit()

        absolute_path = os.path.abspath(file_path)
        return absolute_path
    

