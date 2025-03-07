import tkinter.ttk as ttk
import json
import os
import sys
if getattr(sys, 'frozen', False):  # PyInstaller로 실행된 경우
    BASE_DIR = sys._MEIPASS
else:  # 일반적인 Python 실행
    BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

sys.path.append(BASE_DIR)

sys.path.append(os.path.join(BASE_DIR, "..", "docControll"))
sys.path.append(os.path.join(BASE_DIR, "..", "mail"))
sys.path.append(os.path.join(BASE_DIR, "..", "sheetControll"))
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from docControll import docOneProgram, docThreeProgram, docTwoProgram
from sheetControll import sheetOneProgram, sheetTwoProgram, sheetThreeProgram
from mail.gmail import Gmail

#####################################################################################################

def format_cost(event):
    value = event.widget.get("1.0", "end-1c").strip().replace(",", "")
    if value.isdigit():
        formatted_value = f"{int(value):,}"
        event.widget.delete("1.0", "end")
        event.widget.insert("1.0", formatted_value)

root = Tk()
root.title("계획안 제작_어나더컴퍼니")

config_file_path_1 = os.path.join(os.path.expanduser('~'), '.plan_sheet_path.json')
config_file_path_2 = os.path.join(os.path.expanduser('~'), '.plan_sheet_path_2.json')
config_file_path_3 = os.path.join(os.path.expanduser('~'), '.plan_sheet_path_3.json')

try:
    with open(config_file_path_1, 'r') as f:
        pass
except FileNotFoundError:
    with open(config_file_path_1, 'w') as f:
        json.dump({'file_path': "폴더 경로를 입력해주세요"}, f)

try:
    with open(config_file_path_2, 'r') as f:
        pass
except FileNotFoundError:
    with open(config_file_path_2, 'w') as f:
        json.dump({'file_path': "폴더 경로를 입력해주세요"}, f)

try:
    with open(config_file_path_3, 'r') as f:
        pass
except FileNotFoundError:
    with open(config_file_path_3, 'w') as f:
        pass

def show_frame_1(event):
    selection = cmb_grade.get()
    if selection == "1개 학년":
        sub_frame_school_info2.pack(side="top", padx=5, pady=5, fill="both")
        sub_frame_school_info3.pack_forget()
        sub_frame_school_info4.pack_forget()
    elif selection == "2개 학년":
        sub_frame_school_info2.pack(side="top", padx=5, pady=5, fill="both")
        sub_frame_school_info3.pack(side="top", padx=5, pady=5, fill="both")
        sub_frame_school_info4.pack_forget()
    elif selection == "3개 학년":
        sub_frame_school_info2.pack(side="top", padx=5, pady=5, fill="both")
        sub_frame_school_info3.pack(side="top", padx=5, pady=5, fill="both")
        sub_frame_school_info4.pack(side="top", padx=5, pady=5, fill="both")

def show_frame_2(event):
    selection = cmb_class_1.get()
    if selection == "3차시" or selection == "4차시":
        lbl_program_1_1.pack_forget()
        cmb_program_1_1.pack_forget()
    elif selection == "6차시":
        lbl_program_1_1.pack(side="left", padx=5, pady=5)
        cmb_program_1_1.pack(side="left", padx=5, pady=5)

def show_frame_3(event):
    selection = cmb_class_2.get()
    if selection == "3차시" or selection == "4차시":
        lbl_program_2_1.pack_forget()
        cmb_program_2_1.pack_forget()
    elif selection == "6차시":
        lbl_program_2_1.pack(side="left", padx=5, pady=5)
        cmb_program_2_1.pack(side="left", padx=5, pady=5)

def show_frame_4(event):
    selection = cmb_class_3.get()
    if selection == "3차시" or selection == "4차시":
        lbl_program_3_1.pack_forget()
        cmb_program_3_1.pack_forget()
    elif selection == "6차시":
        lbl_program_3_1.pack(side="left", padx=5, pady=5)
        cmb_program_3_1.pack(side="left", padx=5, pady=5)

def add_manager(name, email, password, cmb_manager):
    if name == "" or email == "" or password == "":
        messagebox.showwarning("입력 오류", "모든 필수값을 입력해주세요!")
        return

    user_info = {
        "name": name,
        "id": email,
        "app_pass": password
    }

    with open(config_file_path_3, 'r') as f:
        try:
            existing_data = json.load(f)
            if not isinstance(existing_data, list):
                existing_data = [existing_data]
        except json.JSONDecodeError:
            existing_data = []
    
    existing_data.append(user_info)

    with open(config_file_path_3, 'w') as f:
        json.dump(existing_data, f, indent=4)

    update_manager_list(cmb_manager)
    
    messagebox.showinfo("추가 완료", "새로운 매니저가 등록되었습니다!")

    # 확인용으로 파일 내용 출력
    with open(config_file_path_3, 'r') as f:
        loaded_data = json.load(f)
        print(loaded_data)

def delete_manager(name, cmb_manager):
    try:
        with open(config_file_path_3, "r") as f:
            existing_data = json.load(f)

        updated_data = [manager for manager in existing_data if manager["name"] != name]

        with open(config_file_path_3, "w") as f:
            json.dump(updated_data, f, indent=4)
        
        messagebox.showinfo("삭제 성공", "매니저를 삭제했습니다")

        update_manager_list(cmb_manager)

        with open(config_file_path_3, 'r') as f:
            loaded_data = json.load(f)
            print(loaded_data)

    except Exception:
        messagebox.showerror("삭제 실패", "매니저 삭제에 실패했습니다!")

def get_manager_list():
    try:
        with open(config_file_path_3, "r") as f:
            existing_data = json.load(f)

        manager_names = [manager["name"] for manager in existing_data]

        return manager_names
    except:
        messagebox.showerror("오류", "매니저 명단을 불러오는데 실패했습니다!")

def update_manager_list(dropdown):
    manager_names = get_manager_list()
    dropdown['values'] = manager_names

def get_manager_info(name):
    try:
        with open(config_file_path_3, "r") as f:
            existing_data = json.load(f)

        for manager in existing_data:
            if manager["name"] == name:
                return manager["id"], manager["app_pass"]

        return None, None  # 해당 이름을 찾지 못한 경우

    except FileNotFoundError:
        print("파일을 찾을 수 없습니다.")
        return None, 
    except json.JSONDecodeError:
        print("JSON 파일을 읽는 데 문제가 발생했습니다.")
        return None, None

def get_emails_except(name):
    try:
        with open(config_file_path_3, "r") as f:
            data = json.load(f)

        emails = [obj["id"] for obj in data if obj["name"] != name]
        return emails

    except Exception as e:
        print(f"오류 발생: {e}")
        return []

def save_config_1(file_path):
    # 설정을 JSON 파일에 저장
    with open(config_file_path_1, 'w') as f:
        json.dump({'file_path': file_path}, f)

def save_config_2(file_path):
    # 설정을 JSON 파일에 저장
    with open(config_file_path_2, 'w') as f:
        json.dump({'file_path': file_path}, f)

def load_config_1():
    # 설정 파일에서 파일 경로 불러오기
    try:
        with open(config_file_path_1, 'r') as f:
            config = json.load(f)
            return config.get('file_path')
    except FileNotFoundError:
        return None
    
def load_config_2():
    # 설정 파일에서 파일 경로 불러오기
    try:
        with open(config_file_path_2, 'r') as f:
            config = json.load(f)
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

def browse_dest_path_2():
    initial_dir = load_config_2()
    folder_selected = filedialog.askdirectory(initialdir=initial_dir)
    if folder_selected == '': # 사용자가 취소를 누를 때
        return
    txt_dest_path_2.delete(0, END)
    txt_dest_path_2.insert(0, folder_selected)

    save_config_2(folder_selected)

def to_raw(file_path):
    return fr"{file_path}"

def checkProgramValid(program_1, program_2, class_num):
    if class_num == "3차시" and program_1 == "AI오피스":
        messagebox.showerror("실패", f"{program_1} {class_num}는 선택할 수 없습니다.")
        return False
    elif class_num == "6차시":
        if program_1 == "코드5":
            messagebox.showerror("실패", "코드5 는 3차시, 4차시만 가능합니다.")
            return False
        elif program_1 != "취업조작단" and program_2 == "취업조작단1~6차시":
            messagebox.showerror("실패", f"취업조작단1~6차시는 취업조작단으로만 구성해야 합니다.")
            return False

    return True

def checkBlankValid(*args):
    if any(arg == "" for arg in args):
        messagebox.showerror("실패", "모든 필수항목을 입력해주세요.")
        return False
    return True

def start():
    # 학년 정보
    grade_num = cmb_grade.get() # 학년 수
    directory_path = txt_dest_path_1.get() # 양식 파일 경로
    save_path = txt_dest_path_2.get() # 파일 저장 경로

    # 수업 날짜 정보
    year = txt_year.get("1.0", END).strip()
    month = cmb_month.get()
    day = cmb_day.get()

    class_date = year + "년 " + month + "월 " + day + "일"

    # 학교명
    school_name = txt_school_name.get("1.0", END).strip()

    if checkBlankValid(grade_num, directory_path, save_path, year, month, day, school_name) == False:
        return

    # 프로그램 정보
    first_program_1 = cmb_program_1.get() # 1 ~ 4 차시 프로그램
    first_program_2 = cmb_program_1_1.get() # 5 ~ 6 차시 프로그램
    first_grade = cmb_grade_1.get() # 학년
    first_class = cmb_ban_1.get() # 반
    class_num_1 = cmb_class_1.get() # 차시
    price_1 = txt_price_1.get("1.0", END).strip().replace(",", "") # 반 당 금액

    second_program_1 = cmb_program_2.get() # 1 ~ 4 차시 프로그램
    second_program_2 = cmb_program_2_1.get() # 5 ~ 6 차시 프로그램
    second_grade = cmb_grade_2.get() # 학년
    second_class = cmb_ban_2.get() # 반
    class_num_2 = cmb_class_2.get() # 차시
    price_2 = txt_price_2.get("1.0", END).strip().replace(",", "") # 반 당 금액

    last_program_1 = cmb_program_3.get() # 1 ~ 4 차시 프로그램
    last_program_2 = cmb_program_3_1.get() # 5 ~ 6 차시 프로그램
    last_grade = cmb_grade_3.get() # 학년
    last_class = cmb_ban_3.get() # 반
    class_num_3 = cmb_class_3.get() # 차시
    price_3 = txt_price_3.get("1.0", END).strip().replace(",", "") # 반 당 금액

    # 이메일 정보
    email_check = email_var.get()
    email = email_entry.get()
    teacher = teacher_entry.get()
    manager = cmb_manager.get()
    email_set = cmb_email_set.get()

    if email_check:
        if checkBlankValid(email, teacher, manager, email_set) == False:
            return

    if grade_num == "1개 학년":
        if class_num_1 == "6차시":
            if checkBlankValid(first_grade, first_class, class_num_1, first_program_1, first_program_2, price_1) == False:
                return
        else:
            if checkBlankValid(first_grade, first_class, class_num_1, first_program_1, price_1) == False:
                return

        if checkProgramValid(first_program_1, first_program_2, class_num_1) == False:
            return
        
        doc = docOneProgram.Doc(first_grade, first_class, class_num_1, directory_path, save_path, class_date,
                        first_program_1, first_program_2,
                        school_name)
        doc_file_path = doc.makeDoc()

        sheet = sheetOneProgram.Sheet(first_grade, first_class, class_num_1, directory_path, save_path,
                        first_program_1, first_program_2,
                        school_name, price_1)
        
        sheet_file_path = sheet.makeSheet()
        
    elif grade_num == "2개 학년":
        if class_num_1 == "6차시":
            if checkBlankValid(first_grade, first_class, class_num_1, first_program_1, first_program_2, price_1) == False:
                return
        else:
            if checkBlankValid(first_grade, first_class, class_num_1, first_program_1, price_1) == False:
                return

        if class_num_2 == "6차시":
            if checkBlankValid(second_grade, second_class, class_num_2, second_program_1, second_program_2, price_2) == False:
                return
        else:
            if checkBlankValid(second_grade, second_class, class_num_2, second_program_1, price_2) == False:
                return

        if checkProgramValid(first_program_1, first_program_2, class_num_1) == False:
            return
        if checkProgramValid(second_program_1, second_program_2, class_num_2) == False:
            return

        doc = docTwoProgram.Category(directory_path, save_path, class_date,
                         first_program_1, first_program_2, first_grade, first_class, class_num_1,
                         second_program_1, second_program_2, second_grade, second_class, class_num_2,
                         school_name)
        doc_file_path = doc.makeDoc()

        sheet = sheetTwoProgram.Sheet(directory_path, save_path,
                        first_program_1, first_program_2, first_grade, first_class, class_num_1, price_1,
                        second_program_1, second_program_2, second_grade, second_class, class_num_2, price_2,
                        school_name)
        
        sheet_file_path = sheet.makeSheet()

    elif grade_num == "3개 학년":
        if class_num_1 == "6차시":
            if checkBlankValid(first_grade, first_class, class_num_1, first_program_1, first_program_2, price_1) == False:
                return
        else:
            if checkBlankValid(first_grade, first_class, class_num_1, first_program_1, price_1) == False:
                return

        if class_num_2 == "6차시":
            if checkBlankValid(second_grade, second_class, class_num_2, second_program_1, second_program_2, price_2) == False:
                return
        else:
            if checkBlankValid(second_grade, second_class, class_num_2, second_program_1, price_2) == False:
                return
        
        if class_num_3 == "6차시":
            if checkBlankValid(last_grade, last_class, class_num_3, last_program_1, last_program_2, price_3) == False:
                return
        else:
            if checkBlankValid(last_grade, last_class, class_num_3, last_program_1, price_3) == False:
                return

        if checkProgramValid(first_program_1, first_program_2, class_num_1) == False:
            return
        if checkProgramValid(second_program_1, second_program_2, class_num_2) == False:
            return
        if checkProgramValid(last_program_1, last_program_2, class_num_3) == False:
            return
        
        doc = docThreeProgram.Category(directory_path, save_path, class_date,
                         first_program_1, first_program_2, first_grade, first_class, class_num_1,
                         second_program_1, second_program_2, second_grade, second_class, class_num_2,
                         last_program_1, last_program_2, last_grade, last_class, class_num_3,
                         school_name)
        doc_file_path = doc.makeDoc()

        sheet = sheetThreeProgram.Sheet(directory_path, save_path,
                        first_program_1, first_program_2, first_grade, first_class, class_num_1, price_1,
                        second_program_1, second_program_2, second_grade, second_class, class_num_2, price_2,
                        last_program_1, last_program_2, last_grade, last_class, class_num_3, price_3,
                        school_name)
        
        sheet_file_path = sheet.makeSheet()

    file_list = []
    file_list.append(doc_file_path)
    file_list.append(sheet_file_path)
    file_list.append(os.path.join(directory_path, "사업자등록증&통장사본/24년_사업자등록증&통장사본_어나더컴퍼니.pdf"))

    if email_check:
        id, app_pass = get_manager_info(manager)
        bcc = get_emails_except(manager)

        gmail = Gmail(filenames=file_list, grade_num=grade_num, school_name=school_name, email_set=email_set, year=year,
                      email=email, teacher=teacher, class_date=class_date,
                      first_program_1=first_program_1, first_grade=first_grade, first_class=first_class, class_num_1=class_num_1,
                      second_program_1=second_program_1, second_grade=second_grade, second_class=second_class, class_num_2=class_num_2,
                      last_program_1=last_program_1, last_grade=last_grade, last_class=last_class, class_num_3=class_num_3,
                      id=id, app_pass=app_pass, manager=manager, bcc=bcc)
        gmail.send_gmail()

    messagebox.showinfo("성공", "작업이 완료되었습니다!")

########################################################################################################

# 옵션 프레임
frame_option = Frame(root)
frame_option.pack(padx=5, pady=5, ipady=5, fill="both")

frame_option_sub_1 = LabelFrame(frame_option, text="학년")
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
txt_school_name = Text(sub_frame_school_info1, width=40, height=1)
txt_school_name.pack(side="left", padx=5, pady=5)



# 학교 정보 서브 프레임 2
sub_frame_school_info2 = Frame(frame_school_info)

# 학년 레이블
lbl_grade_1 = Label(sub_frame_school_info2, text="학년", width=5)
lbl_grade_1.pack(side="left", padx=5, pady=5)

# 학년 콤보박스
opt_grade = [str(i) + "학년" for i in range(1, 4)]
cmb_grade_1 = ttk.Combobox(sub_frame_school_info2, state="readonly", values=opt_grade, width=5)
cmb_grade_1.pack(side="left", padx=5, pady=5)

# 반 레이블
lbl_ban_1 = Label(sub_frame_school_info2, text="반", width=2)
lbl_ban_1.pack(side="left", padx=5, pady=5)

# 반 콤보박스
opt_ban = [i for i in range(1, 21)]

cmb_ban_1 = ttk.Combobox(sub_frame_school_info2, state="readonly", values=opt_ban, width=3)
cmb_ban_1.pack(side="left", padx=5, pady=5)

# 차시 레이블
lbl_class_1 = Label(sub_frame_school_info2, text="차시", width=3)
lbl_class_1.pack(side="left", padx=5, pady=5)

# 차시 콤보박스
opt_class = ["3차시", "4차시", "6차시"]

cmb_class_1 = ttk.Combobox(sub_frame_school_info2, state="readonly", values=opt_class, width=5)
cmb_class_1.pack(side="left", padx=5, pady=5)
cmb_class_1.bind("<<ComboboxSelected>>", show_frame_2)

# 반 당 금액 레이블
lbl_price_1 = Label(sub_frame_school_info2, text="반 당 금액", width=7)
lbl_price_1.pack(side="left", padx=5, pady=5)

# 반 당 금액 텍스트박스
txt_price_1 = Text(sub_frame_school_info2, width=10, height=1)
txt_price_1.pack(side="left", padx=5, pady=5)
txt_price_1.bind("<KeyRelease>", format_cost)

# 1~4차시 프로그램 옵션
lbl_program_1 = Label(sub_frame_school_info2, text="1~4차시 프로그램", width=13)
lbl_program_1.pack(side="left", padx=5, pady=5)

# 프로그램 콤보박스
opt_program = ["수상한스튜디오", "어나더랜드", "취업조작단", "어나더비밀상담소", "코드5", "AI오피스"]
cmb_program_1 = ttk.Combobox(sub_frame_school_info2, state="readonly", values=opt_program, width=15, height=5)
cmb_program_1.pack(side="left", padx=5, pady=5)

# 5~6차시 프로그램 옵션
lbl_program_1_1 = Label(sub_frame_school_info2, text="5~6차시 프로그램", width=13)

# 프로그램 콤보박스
opt_program_1 = ["코드5", "DISC", "취업조작단1~6차시"]
cmb_program_1_1 = ttk.Combobox(sub_frame_school_info2, state="readonly", values=opt_program_1, width=15, height=1)



# 학교 정보 서브 프레임 3
sub_frame_school_info3 = Frame(frame_school_info)

# 학년 레이블
lbl_grade_2 = Label(sub_frame_school_info3, text="학년", width=5)
lbl_grade_2.pack(side="left", padx=5, pady=5)

# 학년 콤보박스
cmb_grade_2 = ttk.Combobox(sub_frame_school_info3, state="readonly", values=opt_grade, width=5)
cmb_grade_2.pack(side="left", padx=5, pady=5)

# 반 레이블
lbl_ban_2 = Label(sub_frame_school_info3, text="반", width=2)
lbl_ban_2.pack(side="left", padx=5, pady=5)

# 반 콤보박스
cmb_ban_2 = ttk.Combobox(sub_frame_school_info3, state="readonly", values=opt_ban, width=3)
cmb_ban_2.pack(side="left", padx=5, pady=5)

# 차시 레이블
lbl_class_2 = Label(sub_frame_school_info3, text="차시", width=3)
lbl_class_2.pack(side="left", padx=5, pady=5)

# 차시 콤보박스
cmb_class_2 = ttk.Combobox(sub_frame_school_info3, state="readonly", values=opt_class, width=5)
cmb_class_2.pack(side="left", padx=5, pady=5)
cmb_class_2.bind("<<ComboboxSelected>>", show_frame_3)

# 반 당 금액 레이블
lbl_price_2 = Label(sub_frame_school_info3, text="반 당 금액", width=7)
lbl_price_2.pack(side="left", padx=5, pady=5)

# 반 당 금액 텍스트박스
txt_price_2 = Text(sub_frame_school_info3, width=10, height=1)
txt_price_2.pack(side="left", padx=5, pady=5)
txt_price_2.bind("<KeyRelease>", format_cost)

# 1~4차시 프로그램 옵션
lbl_program_2 = Label(sub_frame_school_info3, text="1~4차시 프로그램", width=13)
lbl_program_2.pack(side="left", padx=5, pady=5)

# 프로그램 콤보박스
cmb_program_2 = ttk.Combobox(sub_frame_school_info3, state="readonly", values=opt_program, width=15, height=5)
cmb_program_2.pack(side="left", padx=5, pady=5)

# 5~6차시 프로그램 옵션
lbl_program_2_1 = Label(sub_frame_school_info3, text="5~6차시 프로그램", width=13)

# 프로그램 콤보박스
cmb_program_2_1 = ttk.Combobox(sub_frame_school_info3, state="readonly", values=opt_program_1, width=15, height=1)



# 학교 정보 서브 프레임 4
sub_frame_school_info4 = Frame(frame_school_info)

# 학년 레이블
lbl_grade_3 = Label(sub_frame_school_info4, text="학년", width=5)
lbl_grade_3.pack(side="left", padx=5, pady=5)

# 학년 콤보박스
cmb_grade_3 = ttk.Combobox(sub_frame_school_info4, state="readonly", values=opt_grade, width=5)
cmb_grade_3.pack(side="left", padx=5, pady=5)

# 반 레이블
lbl_ban_3 = Label(sub_frame_school_info4, text="반", width=2)
lbl_ban_3.pack(side="left", padx=5, pady=5)

# 반 콤보박스
cmb_ban_3 = ttk.Combobox(sub_frame_school_info4, state="readonly", values=opt_ban, width=3)
cmb_ban_3.pack(side="left", padx=5, pady=5)

# 차시 레이블
lbl_class_3 = Label(sub_frame_school_info4, text="차시", width=3)
lbl_class_3.pack(side="left", padx=5, pady=5)

# 차시 콤보박스
cmb_class_3 = ttk.Combobox(sub_frame_school_info4, state="readonly", values=opt_class, width=5)
cmb_class_3.pack(side="left", padx=5, pady=5)
cmb_class_3.bind("<<ComboboxSelected>>", show_frame_4)

# 반 당 금액 레이블
lbl_price_3 = Label(sub_frame_school_info4, text="반 당 금액", width=7)
lbl_price_3.pack(side="left", padx=5, pady=5)

# 반 당 금액 텍스트박스
txt_price_3 = Text(sub_frame_school_info4, width=10, height=1)
txt_price_3.pack(side="left", padx=5, pady=5)
txt_price_3.bind("<KeyRelease>", format_cost)

# 1~4차시 프로그램 옵션
lbl_program_3 = Label(sub_frame_school_info4, text="1~4차시 프로그램", width=13)
lbl_program_3.pack(side="left", padx=5, pady=5)

# 프로그램 콤보박스
cmb_program_3 = ttk.Combobox(sub_frame_school_info4, state="readonly", values=opt_program, width=15, height=5)
cmb_program_3.pack(side="left", padx=5, pady=5)

# 5~6차시 프로그램 옵션
lbl_program_3_1 = Label(sub_frame_school_info4, text="5~6차시 프로그램", width=13)

# 프로그램 콤보박스
cmb_program_3_1 = ttk.Combobox(sub_frame_school_info4, state="readonly", values=opt_program_1, width=15, height=1)

# 양식 폴더경로 프레임
path_frame_1 = LabelFrame(root, text="양식 폴더경로")
path_frame_1.pack(fill="x", padx=5, pady=5, ipady=5)

txt_dest_path_1 = Entry(path_frame_1)
txt_dest_path_1.insert(0, load_config_1())
txt_dest_path_1.pack(side="left", fill="x", expand=True, padx=5, pady=5, ipady=4) # 높이 변경

btn_dest_path_1 = Button(path_frame_1, text="찾아보기", width=10, command=browse_dest_path_1)
btn_dest_path_1.pack(side="right", padx=5, pady=5)

# 파일 저장할 경로 프레임
path_frame_2 = LabelFrame(root, text="파일 저장할 폴더경로")
path_frame_2.pack(fill="x", padx=5, pady=5, ipady=5)

txt_dest_path_2 = Entry(path_frame_2)
txt_dest_path_2.insert(0, load_config_2())
txt_dest_path_2.pack(side="left", fill="x", expand=True, padx=5, pady=5, ipady=4) # 높이 변경

btn_dest_path_2 = Button(path_frame_2, text="찾아보기", width=10, command=browse_dest_path_2)
btn_dest_path_2.pack(side="right", padx=5, pady=5)

# 이메일 전송 여부
frame_email = LabelFrame(root, text="이메일 전송")
frame_email.pack(fill="x", padx=5, pady=5, ipady=5)

email_var = IntVar()
email_var.set(1)

email_checkBox = Checkbutton(frame_email, text="이메일 보내기", variable=email_var)
email_checkBox.pack(side="left", padx=5, pady=5)

email_label = Label(frame_email, text="받을 메일")
email_label.pack(side="left", padx=5, pady=5)

email_entry = Entry(frame_email, width=30)
email_entry.pack(side="left", padx=5, pady=5, ipady=4)

teacher_label = Label(frame_email, text="부장님 성함")
teacher_label.pack(side="left", padx=5, pady=5)

teacher_entry = Entry(frame_email, width=20)
teacher_entry.pack(side="left", padx=5, pady=5, ipady=4)

manager_label = Label(frame_email, text="매니저")
manager_label.pack(side="left", padx=5, pady=5)

manager_names = get_manager_list()

cmb_manager = ttk.Combobox(frame_email, values=manager_names, width=10)
cmb_manager.pack(side="left", padx=5, pady=5, ipady=4)

email_set_label = Label(frame_email, text="메일 형식")
email_set_label.pack(side="left", padx=5, pady=5)

email_set = ["재구매", "전학년", "신규"]

cmb_email_set = ttk.Combobox(frame_email, state="readonly", values=email_set, width=5)
cmb_email_set.pack(side="left", padx=5, pady=5, ipady=4)

# 매니저 관리
frame_manager = LabelFrame(root, text="매니저 관리")
frame_manager.pack(fill="x", padx=5, pady=5, ipady=5)

manager_name_label = Label(frame_manager, text="이름")
manager_name_label.pack(side="left", padx=5, pady=5)

manager_name_entry = Entry(frame_manager, width=20)
manager_name_entry.pack(side="left", padx=5, pady=5, ipady=4)

manager_email_label = Label(frame_manager, text="메일 주소")
manager_email_label.pack(side="left", padx=5, pady=5)

manager_email_entry = Entry(frame_manager, width=20)
manager_email_entry.pack(side="left", padx=5, pady=5, ipady=4)

manager_password_label = Label(frame_manager, text="앱 비밀번호")
manager_password_label.pack(side="left", padx=5, pady=5)

manager_password_entry = Entry(frame_manager, width=20)
manager_password_entry.pack(side="left", padx=5, pady=5, ipady=4)

btn_delete = Button(frame_manager, padx=5, pady=5, text="삭제", width=8, command=lambda: delete_manager(manager_name_entry.get(), cmb_manager))
btn_delete.pack(side="right", padx=5, pady=5)

btn_add = Button(frame_manager, padx=5, pady=5, text="추가", width=8, command=lambda: add_manager(manager_name_entry.get(), manager_email_entry.get(), manager_password_entry.get(), cmb_manager))
btn_add.pack(side="right", padx=5, pady=5)

# 실행 프레임
frame_run = Frame(root)
frame_run.pack(fill="x", padx=5, pady=5)

btn_close = Button(frame_run, padx=5, pady=5, text="닫기", width=12, command=root.quit)
btn_close.pack(side="right", padx=5, pady=5)

btn_start = Button(frame_run, padx=5, pady=5, text="시작", width=12, command=start)
btn_start.pack(side="right", padx=5, pady=5)

root.resizable(False, False)
root.mainloop()