import tkinter.ttk as ttk
import json
import os
import sys
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from tkinter import *
from tkinter import filedialog
from docControll import oneProgram, twoProgram, threeProgram
from mail.gmail import Gmail

#####################################################################################################

root = Tk()
root.title("계획안 제작_어나더컴퍼니")

config_file_path_1 = os.path.join(os.path.expanduser('~'), '.plan_sheet_path.json')

try:
    with open(config_file_path_1, 'r') as f:
        pass
except FileNotFoundError:
    with open(config_file_path_1, 'w') as f:
        json.dump({'file_path': "폴더 경로를 입력해주세요"}, f)

def show_frame_1(event):
    selection_1 = cmb_grade.get()
    if selection_1 == "1개 학년":
        sub_frame_school_info2.pack(side="top", padx=5, pady=5, fill="both")
        sub_frame_school_info3.pack_forget()
        sub_frame_school_info4.pack_forget()
    elif selection_1 == "2개 학년":
        sub_frame_school_info2.pack(side="top", padx=5, pady=5, fill="both")
        sub_frame_school_info3.pack(side="top", padx=5, pady=5, fill="both")
        sub_frame_school_info4.pack_forget()
    elif selection_1 == "3개 학년":
        sub_frame_school_info2.pack(side="top", padx=5, pady=5, fill="both")
        sub_frame_school_info3.pack(side="top", padx=5, pady=5, fill="both")
        sub_frame_school_info4.pack(side="top", padx=5, pady=5, fill="both")

def show_frame_2(event):
    selection_2 = cmb_class.get()
    if selection_2 == "3차시" or selection_2 == "4차시":
        lbl_program_1_1.pack_forget()
        cmb_program_1_1.pack_forget()
        lbl_program_2_1.pack_forget()
        cmb_program_2_1.pack_forget()
        lbl_program_3_1.pack_forget()
        cmb_program_3_1.pack_forget()
    elif selection_2 == "5차시" or selection_2 == "6차시":
        lbl_program_1_1.pack(side="left", padx=5, pady=5)
        cmb_program_1_1.pack(side="left", padx=5, pady=5)
        lbl_program_2_1.pack(side="left", padx=5, pady=5)
        cmb_program_2_1.pack(side="left", padx=5, pady=5)
        lbl_program_3_1.pack(side="left", padx=5, pady=5)
        cmb_program_3_1.pack(side="left", padx=5, pady=5)

def save_config_1(file_path):
    # 설정을 JSON 파일에 저장
    with open(config_file_path_1, 'w') as f:
        json.dump({'file_path': file_path}, f)

def load_config_1():
    # 설정 파일에서 파일 경로 불러오기
    try:
        with open(config_file_path_1, 'r') as f:
            config = json.load(f)
            test = config.get('file_path')
            return config.get('file_path')
    except FileNotFoundError:
        return None
    
def browse_dest_path_1():
    initial_dir = load_config_1()
    folder_selected = filedialog.askdirectory(initialdir=initial_dir)
    if folder_selected == '': # 사용자가 취소를 누를 때
        return
    txt_dest_path_1.delete(0, END)
    txt_dest_path_1.insert(0, folder_selected)

    save_config_1(folder_selected)

def to_raw(file_path):
    return fr"{file_path}"

def start():
    # 학년/차시 정보
    grade_num = cmb_grade.get() # 학년 수
    class_num = cmb_class.get() # 차시
    directory_path = txt_dest_path_1.get() # 양식 파일 경로

    # 수업 날짜 정보
    year = txt_year.get("1.0", END).strip()
    month = cmb_month.get()
    day = cmb_day.get()

    class_date = year + "년 " + month + "월 " + day + "일"

    # 학교명/반당 금액
    school_name = txt_school_name.get("1.0", END).strip()
    price = txt_cost.get("1.0", END).strip()
    formatted_price_1 = "{:,}".format(int(price)) # 세자리마다 쉼표 처리

    # 프로그램 정보
    first_program_1 = cmb_program_1.get()
    first_program_2 = cmb_program_1_1.get()
    first_grade = cmb_grade_1.get()
    first_class = cmb_ban_1.get()

    second_program_1 = cmb_program_2.get()
    second_program_2 = cmb_program_2_1.get()
    second_grade = cmb_grade_2.get()
    second_class = cmb_ban_2.get()

    last_program_1 = cmb_program_3.get()
    last_program_2 = cmb_program_3_1.get()
    last_grade = cmb_grade_3.get()
    last_class = cmb_ban_3.get()

    # 이메일 정보
    email_check = email_var.get()
    email = email_entry.get()

    if grade_num == "1개 학년":
        doc = oneProgram.Category(first_grade, first_class, class_num, directory_path, class_date,
                        first_program_1, first_program_2,
                        school_name, price)
        file_path = doc.makeDoc()
    elif grade_num == "2개 학년":
        doc = twoProgram.Category(class_num, directory_path, class_date,
                         first_program_1, first_program_2, first_grade, first_class,
                         second_program_1, second_program_2, second_grade, second_class,
                         school_name, price)
        file_path = doc.makeDoc()
    elif grade_num == "3개 학년":
        doc = threeProgram.Category(class_num, directory_path, class_date,
                         first_program_1, first_program_2, first_grade, first_class,
                         second_program_1, second_program_2, second_grade, second_class,
                         last_program_1, last_program_2, last_grade, last_class,
                         school_name, price)
        file_path = doc.makeDoc()

    file_list = []
    file_list.append(file_path)

    if email_check:
        gmail = Gmail(filenames=file_list, grade_num=grade_num, school_name=school_name, first_grade=first_grade, first_program_1=first_program_1)
        gmail.send_gmail()

########################################################################################################

# 옵션 프레임
frame_option = Frame(root)
frame_option.pack(padx=5, pady=5, ipady=5, fill="both")

frame_option_sub_1 = LabelFrame(frame_option, text="학년/차시")
frame_option_sub_1.pack(side="top", padx=5, fill="both")

frame_option_sub_2 = LabelFrame(frame_option, text="수업 날짜")
frame_option_sub_2.pack(side="bottom", padx=5, fill="both")

# 1. 제작 옵션
# 학년 레이블
lbl_grade_1 = Label(frame_option_sub_1, text="학년 수", width=8)
lbl_grade_1.pack(side="left", padx=5, pady=5)

# 학년 콤보
opt_grade = ["1개 학년", "2개 학년", "3개 학년"]
cmb_grade = ttk.Combobox(frame_option_sub_1, state="readonly", values=opt_grade, width=10)
cmb_grade.pack(side="left", padx=5, pady=5)
cmb_grade.bind("<<ComboboxSelected>>", show_frame_1)

# 차시 레이블
lbl_class = Label(frame_option_sub_1, text="차시", width=8)
lbl_class.pack(side="left", padx=5, pady=5)

# 차시 콤보
opt_class = ["3차시", "4차시", "5차시", "6차시"]
cmb_class = ttk.Combobox(frame_option_sub_1, state="readonly", values=opt_class, width=10)
cmb_class.pack(side="left", padx=5, pady=5)
cmb_class.bind("<<ComboboxSelected>>", show_frame_2)

# 수업 날짜 - 년도 텍스트박스
txt_year = Text(frame_option_sub_2, width=8, height=1)
txt_year.pack(side="left", padx=10, pady=5)

# 수업 날짜 - 년도 레이블
lbl_year = Label(frame_option_sub_2, text="년", width=4)
lbl_year.pack(side="left", padx=5, pady=5)

# 수업 날짜 - 월 콤보박스
opt_month = [f"{i:02d}" for i in range(1, 13)]
cmb_month = ttk.Combobox(frame_option_sub_2, state="readonly", values=opt_month, width=4)
cmb_month.pack(side="left", padx=5, pady=5)

# 수업 날짜 - 월 레이블
lbl_month = Label(frame_option_sub_2, text="월", width=3)
lbl_month.pack(side="left", pady=5)

# 수업 날짜 - 일 콤보박스
opt_day = [f"{i:02d}" for i in range(1, 32)]
cmb_day = ttk.Combobox(frame_option_sub_2, state="readonly", values=opt_day, width=6)
cmb_day.pack(side="left", padx=5, pady=5)

# 수업 날짜 - 일 레이블
lbl_day = Label(frame_option_sub_2, text="일", width=3)
lbl_day.pack(side="left", pady=5)

# 학교 정보 프레임
frame_school_info = LabelFrame(root, text="학교 정보")
frame_school_info.pack(padx=5, pady=5, fill="both")

# 학교 정보 서브 프레임 1
sub_frame_school_info1 = Frame(frame_school_info)
sub_frame_school_info1.pack(side="top", padx=5, fill="both")

# 1. 학교명 옵션
# 학교명 레이블
lbl_school_name = Label(sub_frame_school_info1, text="학교명", width=5)
lbl_school_name.pack(side="left", padx=5, pady=5)

# 학교명 텍스트박스
txt_school_name = Text(sub_frame_school_info1, width=73, height=1)
txt_school_name.pack(side="left", padx=5, pady=5)

# 반당 금액 레이블
lbl_cost = Label(sub_frame_school_info1, text="반당 금액", width=15)
lbl_cost.pack(side="left", padx=5, pady=5)

# 반당 금액 텍스트박스
txt_cost = Text(sub_frame_school_info1, width=22, height=1)
txt_cost.pack(side="left", padx=5, pady=5)

# 학교 정보 서브 프레임 2
sub_frame_school_info2 = Frame(frame_school_info)

# 학년 레이블
lbl_grade_1 = Label(sub_frame_school_info2, text="학년", width=5)
lbl_grade_1.pack(side="left", padx=5, pady=5)

# 학년 콤보박스
opt_grade = [str(i) + "학년" for i in range(1, 4)]
cmb_grade_1 = ttk.Combobox(sub_frame_school_info2, state="readonly", values=opt_grade, width=8)
cmb_grade_1.pack(side="left", padx=5, pady=5)

# 반 레이블
lbl_ban_1 = Label(sub_frame_school_info2, text="반", width=5)
lbl_ban_1.pack(side="left", padx=5, pady=5)

# 반 콤보박스
opt_ban = [i for i in range(1, 21)]

cmb_ban_1 = ttk.Combobox(sub_frame_school_info2, state="readonly", values=opt_ban, width=8)
cmb_ban_1.pack(side="left", padx=5, pady=5)

# 1~4차시 프로그램 옵션
lbl_program_1 = Label(sub_frame_school_info2, text="1~4차시 프로그램", width=15)
lbl_program_1.pack(side="left", padx=5, pady=5)

# 프로그램 콤보박스
opt_program = ["수상한스튜디오", "어나더랜드", "취업조작단", "비밀상담소", "코드5"]
cmb_program_1 = ttk.Combobox(sub_frame_school_info2, state="readonly", values=opt_program, width=20, height=5)
cmb_program_1.pack(side="left", padx=5, pady=5)

# 5~6차시 프로그램 옵션
lbl_program_1_1 = Label(sub_frame_school_info2, text="5~6차시 프로그램", width=15)

# 프로그램 콤보박스
opt_program_1 = ["코드5", "DISC", "취업조작단1~6차시"]
cmb_program_1_1 = ttk.Combobox(sub_frame_school_info2, state="readonly", values=opt_program_1, width=20, height=1)

# 학교 정보 서브 프레임 3
sub_frame_school_info3 = Frame(frame_school_info)

# 학년 레이블
lbl_grade_2 = Label(sub_frame_school_info3, text="학년", width=5)
lbl_grade_2.pack(side="left", padx=5, pady=5)

# 학년 콤보박스
cmb_grade_2 = ttk.Combobox(sub_frame_school_info3, state="readonly", values=opt_grade, width=8)
cmb_grade_2.pack(side="left", padx=5, pady=5)

# 반 레이블
lbl_ban_2 = Label(sub_frame_school_info3, text="반", width=5)
lbl_ban_2.pack(side="left", padx=5, pady=5)

# 반 콤보박스
cmb_ban_2 = ttk.Combobox(sub_frame_school_info3, state="readonly", values=opt_ban, width=8)
cmb_ban_2.pack(side="left", padx=5, pady=5)

# 1~4차시 프로그램 옵션
lbl_program_2 = Label(sub_frame_school_info3, text="1~4차시 프로그램", width=15)
lbl_program_2.pack(side="left", padx=5, pady=5)

# 프로그램 콤보박스
cmb_program_2 = ttk.Combobox(sub_frame_school_info3, state="readonly", values=opt_program, width=20, height=5)
cmb_program_2.pack(side="left", padx=5, pady=5)

# 5~6차시 프로그램 옵션
lbl_program_2_1 = Label(sub_frame_school_info3, text="5~6차시 프로그램", width=15)

# 프로그램 콤보박스
cmb_program_2_1 = ttk.Combobox(sub_frame_school_info3, state="readonly", values=opt_program_1, width=20, height=1)

# 학교 정보 서브 프레임 4
sub_frame_school_info4 = Frame(frame_school_info)

# 학년 레이블
lbl_grade_3 = Label(sub_frame_school_info4, text="학년", width=5)
lbl_grade_3.pack(side="left", padx=5, pady=5)

# 학년 콤보박스
cmb_grade_3 = ttk.Combobox(sub_frame_school_info4, state="readonly", values=opt_grade, width=8)
cmb_grade_3.pack(side="left", padx=5, pady=5)

# 반 레이블
lbl_ban_3 = Label(sub_frame_school_info4, text="반", width=5)
lbl_ban_3.pack(side="left", padx=5, pady=5)

# 반 콤보박스
cmb_ban_3 = ttk.Combobox(sub_frame_school_info4, state="readonly", values=opt_ban, width=8)
cmb_ban_3.pack(side="left", padx=5, pady=5)

# 1~4차시 프로그램 옵션
lbl_program_3 = Label(sub_frame_school_info4, text="1~4차시 프로그램", width=15)
lbl_program_3.pack(side="left", padx=5, pady=5)

# 프로그램 콤보박스
cmb_program_3 = ttk.Combobox(sub_frame_school_info4, state="readonly", values=opt_program, width=20, height=5)
cmb_program_3.pack(side="left", padx=5, pady=5)

# 5~6차시 프로그램 옵션
lbl_program_3_1 = Label(sub_frame_school_info4, text="5~6차시 프로그램", width=15)

# 프로그램 콤보박스
cmb_program_3_1 = ttk.Combobox(sub_frame_school_info4, state="readonly", values=opt_program_1, width=20, height=1)

# 양식 폴더경로 프레임
path_frame_1 = LabelFrame(root, text="양식 폴더경로")
path_frame_1.pack(fill="x", padx=5, pady=5, ipady=5)

txt_dest_path_1 = Entry(path_frame_1)
txt_dest_path_1.insert(0, load_config_1())
txt_dest_path_1.pack(side="left", fill="x", expand=True, padx=5, pady=5, ipady=4) # 높이 변경

btn_dest_path_1 = Button(path_frame_1, text="찾아보기", width=10, command=browse_dest_path_1)
btn_dest_path_1.pack(side="right", padx=5, pady=5)

# 이메일 전송 여부
frame_email = LabelFrame(root, text="이메일 전송")
frame_email.pack(fill="x", padx=5, pady=5, ipady=5)

email_var = IntVar()
email_var.set(1)

email_checkBox = Checkbutton(frame_email, text="이메일 보내기", variable=email_var)
email_checkBox.pack(side="left", padx=5, pady=5)

email_label = Label(frame_email, text="메일 입력칸")
email_label.pack(side="left", padx=5, pady=5)

email_entry = Entry(frame_email, width=40)
email_entry.pack(side="left", padx=5, pady=5, ipady=4)

# 실행 프레임
frame_run = Frame(root)
frame_run.pack(fill="x", padx=5, pady=5)

btn_close = Button(frame_run, padx=5, pady=5, text="닫기", width=12, command=root.quit)
btn_close.pack(side="right", padx=5, pady=5)

btn_start = Button(frame_run, padx=5, pady=5, text="시작", width=12, command=start)
btn_start.pack(side="right", padx=5, pady=5)

root.resizable(False, False)
root.mainloop()