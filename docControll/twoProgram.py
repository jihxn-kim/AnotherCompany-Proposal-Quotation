from pyhwpx import Hwp
import os
import time

# class_num: 차시
# directory_path: 양식 파일 경로
# first_program_1: 첫 번째 프로그램 정보 (오전)
# first_program_2: 첫 번째 프로그램 정보 (오후)
# first_grade: 첫 번째 학년
# first_class: 첫 번째 반
# second_program_1: 두 번째 프로그램 정보 (오전)
# second_program_2: 두 번째 프로그램 정보 (오후)
# second_grade: 두 번째 학년
# second_class: 두 번째 반
# school_name: 학교명
# price: 반당 금액
class Category:
    def __init__(self, class_num, directory_path, class_date,
                 first_program_1, first_program_2, first_grade, first_class,
                 second_program_1, second_program_2, second_grade, second_class,
                 school_name, price):
        self.class_num = class_num
        self.directory_path = directory_path
        self.class_date = class_date

        self.first_program_1 = first_program_1
        self.first_program_2 = first_program_2
        self.first_grade = first_grade
        self.first_class = first_class

        self.second_program_1 = second_program_1
        self.second_program_2 = second_program_2
        self.second_grade = second_grade
        self.second_class = second_class

        self.school_name = school_name
        self.price = price

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

        if self.class_num == "3차시":
            programPage.append(self.first_program_1 + "_교구")
            programPage.append(self.first_program_1 + "_3차시")
            programPage.append(self.second_program_1 + "_교구")
            programPage.append(self.second_program_1 + "_3차시")
        elif self.class_num == "4차시":
            programPage.append(self.first_program_1 + "_교구")
            programPage.append(self.first_program_1 + "_4차시")
            programPage.append(self.second_program_1 + "_교구")
            programPage.append(self.second_program_1 + "_4차시")
        elif self.class_num == "5차시":
            programPage.append(self.first_program_1 + "_교구")
            programPage.append(self.first_program_1 + "_4차시")
            programPage.append(self.first_program_2 + "_5차시")
            programPage.append(self.second_program_1 + "_교구")
            programPage.append(self.second_program_1 + "_4차시")
            programPage.append(self.second_program_2 + "_5차시")
        elif self.class_num == "6차시":
            programPage.append(self.first_program_1 + "_교구")
            programPage.append(self.first_program_1 + "_4차시")
            programPage.append(self.first_program_2 + "_6차시")
            programPage.append(self.second_program_1 + "_교구")
            programPage.append(self.second_program_1 + "_4차시")
            programPage.append(self.second_program_2 + "_6차시")

        doc_list = []
        for i in programPage:
            program_file_path = full_directory_path + i + ".hwp"
            raw_file_path = self.to_raw(os.path.normpath(program_file_path))
            doc_list.append(raw_file_path)

        hwp = Hwp(new=True)

        # 문서 병합
        for i in range(0, len(doc_list)):
            if (i == 0):
                hwp.insert_file(doc_list[i])
            elif (i == 1):
                hwp.MoveDocEnd()
                hwp.insert_file(doc_list[i], keep_section=False)
                hwp.find_replace_all("(학년)", self.first_grade)
            elif (i == 2):
                hwp.MoveDocEnd()
                hwp.insert_file(doc_list[i], keep_section=False)
                hwp.find_replace_all("(학년)", self.second_grade)
            elif (i == 3):
                hwp.MoveDocEnd()
                hwp.insert_file(doc_list[i], move_doc_end=True)
            elif (i == 4):
                hwp.MoveDocEnd()
                hwp.insert_file(doc_list[i], keep_section=False)
                hwp.find_replace_all("(학년)", self.first_grade)
            elif (i == 5):
                hwp.MoveDocEnd()
                hwp.insert_file(doc_list[i], keep_section=False)
                hwp.find_replace_all("(학년)", self.second_grade)
            elif (i == 6):
                hwp.MoveDocEnd()
                hwp.insert_file(doc_list[i], move_doc_end=True)
            elif (i == 7):
                hwp.MoveDocEnd()
                hwp.insert_file(doc_list[i], move_doc_end=True)
                hwp.find_replace_all("(학년)", self.first_grade)
                hwp.find_replace_all("(학급)", self.first_class + "학급")
            elif (i == 8):
                hwp.MoveDocEnd()
                hwp.insert_file(doc_list[i], move_doc_end=True)
            else:
                hwp.MoveDocEnd()
                hwp.insert_file(doc_list[i], move_doc_end=True)
                hwp.find_replace_all("(학년)", self.second_grade)
                hwp.find_replace_all("(학급)", self.second_class + "학급")

        hwp.MoveDocBegin()
        hwp.DeletePage()

        # 텍스트 대체
        hwp.find_replace_all("(학교명)", self.school_name)
        hwp.find_replace_all("(날짜)", self.class_date)
        hwp.find_replace_all("(학년1)", self.first_grade)
        hwp.find_replace_all("(학급1)", self.first_class + "학급")
        hwp.find_replace_all("(학년2)", self.second_grade)
        hwp.find_replace_all("(학급2)", self.second_class + "학급")
        hwp.find_replace_all("(차시)", self.class_num)
        hwp.find_replace_all("(프로그램명1)", self.first_program_1)
        hwp.find_replace_all("(프로그램명2)", self.second_program_1)

        if (self.first_program_1 == "수상한스튜디오"):
            program_1 = "비전욕망설계 캠프"
        elif (self.first_program_1 == "어나더랜드"):
            program_1 = "창업가정신 캠프"
        elif (self.first_program_1 == "비밀상담소"):
            program_1 = "고교학점제 캠프"

        if (self.second_program_1 == "수상한스튜디오"):
            program_2 = "비전욕망설계 캠프"
        elif (self.second_program_1 == "어나더랜드"):
            program_2 = "창업가정신 캠프"
        elif (self.second_program_1 == "비밀상담소"):
            program_2 = "고교학점제 캠프"

        hwp.find_replace_all("(프로그램 종류1)", program_1)
        hwp.find_replace_all("(프로그램 종류2)", program_2)

        curr_time = time.strftime("%y%m%d").strip()
        file_name = curr_time + "_" + self.school_name + "_시뮬레이션진로캠프_계획안_어나더컴퍼니.hwp"

        # 한글파일 저장
        hwp.save_as(file_name)
        # hwp.save_pdf_as_image()
        hwp.quit()
    

