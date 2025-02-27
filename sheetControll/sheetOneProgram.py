from openpyxl import load_workbook
import time
import os

class Sheet:
    def __init__(self, first_grade, first_class, class_num, directory_path, save_path, class_date,
                 first_program_1, first_program_2,
                 school_name, price):
        self.first_grade = first_grade
        self.first_class = first_class
        self.class_num = class_num
        self.directory_path = directory_path
        self.save_path = save_path
        self.class_date = class_date
        self.first_program_1 = first_program_1
        self.first_program_2 = first_program_2
        self.school_name = school_name
        self.price = price

    def makeSheet(self):
        sampleXlsx = self.directory_path + "/견적서/oneProgram.xlsx"
        
        wb = load_workbook(sampleXlsx)
        ws = wb.active

        ## 학교명
        ws["H3"] = self.school_name 

        ## 사업명
        if self.first_program_1 == "수상한스튜디오":
            program_name = "시뮬레이션 비전욕망설계캠프 [수상한 스튜디오]"
        elif self.first_program_1 == "어나더랜드":
            program_name = "시뮬레이션 기업가정신 프로그램 [어나더랜드]"
        elif self.first_program_1 == "취업조작단":
            program_name = "시뮬레이션 진로역량발산캠프 [취업조작단]"
        elif self.first_program_1 == "비밀상담소":
            program_name = "시뮬레이션 고교학점제캠프 [비밀상담소]"
        elif self.first_program_1 == "코드5":
            program_name = "직무-직업매칭 빅게임 [CODE 5]"
        elif self.first_program_1 == "AI오피스":
            pass

        ws["H4"] = program_name

        ## 견적일자
        curr_time = time.strftime("%Y-%m-%d").strip()
        ws["H7"] = curr_time

        ## 강사비
        # 단가 (50,000 고정)

        # 수량
        ws["D10"] = self.first_class

        # 회
        ws["F10"] = self.class_num[:1]

        # 합계
        ws["G10"] = 50000 * int(self.class_num[:1]) * int(self.first_class)

        # 비고
        ws["H10"] = self.class_num


        ## 재료비
        # 단가
        material_cost = int(self.price) - 50000 * int(self.class_num[:1])
        ws["C11"] = material_cost

        # 수량
        ws["D11"] = self.first_class

        # 합계
        ws["G11"] = material_cost * int(self.first_class)

        # 비고
        if self.first_program_1 == "수상한스튜디오":
            material = "수상한스튜디오 툴킷(비전욕망카드 1set, 숫자카드), 교재(A4, 4P, 컬러)"
        elif self.first_program_1 == "어나더랜드":
            material = "어나더랜드 툴킷(손목밴드, 어나더게임 카드 외), 교재(A4, 8P, 컬러)"
        elif self.first_program_1 == "취업조작단":
            material = "취업조작단 툴킷(강점스캐닝 교구(스티커 1set)), 교재(A4, 양면)"
        elif self.first_program_1 == "비밀상담소":
            material = "고교학점제 툴킷(고교학점제 카드), 교재 (A4, 8P, 컬러)"
        elif self.first_program_1 == "코드5":
            material = "CODE 5 툴킷(직무카드 330장, 플레이매트, 점수카드 등), 교재 (A3, 컬러, 2매)"
        elif self.first_program_1 == "AI오피스":
            pass

        ws["H11"] = material

        ## 전체 합계
        ws["G12"] = material_cost * int(self.first_class) + 50000 * int(self.class_num[:1]) * int(self.first_class)

        ## 파일명
        file_name = self.save_path + "/" + time.strftime("%y%m%d").strip() + "_" + self.school_name + "_" + self.first_grade + "_" + self.first_program_1 + "_견적서_어나더컴퍼니.xlsx"
        wb.save(file_name)

        absolute_path = os.path.join(os.getcwd(), file_name)
        return absolute_path


        