import os
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, font

# ChosunCentennial 폰트 경로 설정 (프로젝트 폴더 내에 폰트 파일이 있어야 함)
FONT_PATH = "Pretendard-Bold.ttf"

FOLDER = r"\\192.168.0.213\nas 공유폴더\백업\2025\■ 동아리 활동 일지"
FILE_NAME = "동아리 활동.xlsx"
CLUB_LIST_NAME = "동아리 목록.xlsx"

LIST_PATH = os.path.join(FOLDER, CLUB_LIST_NAME)
SAVE_PATH = os.path.join(FOLDER, FILE_NAME)

def increment(var):
    var.set(var.get() + 1)

def decrement(var):
    if var.get() > 0:
        var.set(var.get() - 1)

def reset_fields():
    category_var.set("")
    club_name_var.set("")
    participants_var.set("")
    author_var.set("")
    activity_year_var.set(today.strftime("%Y"))
    activity_month_var.set(today.strftime("%m"))
    activity_day_var.set(today.strftime("%d"))
    start_hour_var.set("")
    start_minute_var.set("")
    end_hour_var.set("")
    end_minute_var.set("")
    activity_content_text.delete("1.0", tk.END)  # 활동 내용 초기화
    for var in age_participants_vars.values():
        var.set(0)

def show_success(message):
    messagebox.showinfo("활동일지 프로그램", message)

def show_warning(message):
    messagebox.showwarning("활동일지 프로그램 오류", message)

def format_two_digits(var):
    value = var.get()
    if value.isdigit() and len(value) == 1:
        var.set(f"0{value}")

def is_file_locked(filepath):
    if not os.path.exists(filepath):
        return False
    try:
        # 'a+' 모드로 열어보기: 실패하면 잠긴 파일
        with open(filepath, 'a+'):
            return False
    except PermissionError:
        return True

def format_24H(var):
    value = var.get()
    if value.isdigit() and len(value) == 1:
        var.set(f"0{value}")
        
    if value.isdigit():
        hour = int(value)
        if 1 <= hour <= 8:
            var.set(f"{hour + 12}")

# Tkinter에서 사용할 폰트 로드
if os.path.exists(FONT_PATH):
    ft = "Pretendard"
else:
    ft = "Arial"  # 폰트가 없을 경우 기본 Arial 사용

def get_club_activity():
    global category_var, club_name_var, participants_var, author_var
    global activity_year_var, activity_month_var, activity_day_var
    global start_hour_var, start_minute_var, end_hour_var, end_minute_var
    global age_participants_vars, today, activity_content_text

    # to-do
    # 나중에 여기에 추가
    club2category = {
        '광대': '밴드', 
        '라이트': '댄스', 
        '라인': '밴드', 
        '모먼트': '밴드', 
        '밴디름': '밴드', 
        '보히': '밴드', 
        '블리스': '댄스',
        '샛별': '밴드', 
        '세인트': '밴드', 
        '시온': '밴드', 
        '아크로체': '미술', 
        '초청': '밴드', 
        '클러스터': '과학', 
        '클로버': '댄스',
        '투핸드': '밴드', 
        '퍼스트': '밴드', 
        '포커스': '밴드', 
        '하모리': '미술', 
        '하이애니드림하이': '여가',
        '하이라이트': '밴드', 
        '히어로': '연극', 
        '혼동': '댄스', 
        '유니스': '댄스', 
        'WeCD': '밴드', 
        'S.P': '밴드'
    }

    if os.path.exists(LIST_PATH):
        try:
            _df = pd.read_excel(LIST_PATH)
            if "동아리" in _df.columns and "분야" in _df.columns:
                club2category = dict(zip(_df["동아리"], _df["분야"]))
                show_success(f"동아리 목록을 성공적으로 불러왔습니다.\n\n동아리 수: {len(club2category)}개")
            else:
                show_warning("엑셀 파일에 '동아리' 또는 '분야' 열이 없습니다. 기본값을 사용합니다.")
        except Exception as e:
            show_warning(f"엑셀 파일을 읽는 도중 오류 발생: {e}\n기본값을 사용합니다.")
    else:
        show_warning(f"동아리 목록 엑셀 파일이 없습니다: \n기본값을 사용합니다.")

    clubs = list(club2category.keys())
    categories = list(set(club2category.values()))
    
    age_groups = ['초등 남', '초등 여', '중등 남', '중등 여', '고등 남', '고등 여', '후기 남', '후기 여']

    # 기본 GUI 설정 (창 크기 1000x800)
    root = tk.Tk()
    root.title("2025년 동아리 활동일지")
    root.state("zoomed")
    root.geometry("1000x800")
    root.iconbitmap('nareum.ico')
    
    # 스크롤 가능한 영역 생성
    main_frame = tk.Frame(root)
    main_frame.pack(fill=tk.BOTH, expand=True)

    canvas = tk.Canvas(main_frame)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar = tk.Scrollbar(main_frame, orient=tk.VERTICAL, command=canvas.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    canvas.configure(yscrollcommand=scrollbar.set)
    container = tk.Frame(canvas)
    container_window = canvas.create_window((0, 0), window=container, anchor="n")

    def on_canvas_configure(event):
        canvas.itemconfig(container_window, width=event.width)
        canvas.configure(scrollregion=canvas.bbox("all"))
    canvas.bind("<Configure>", on_canvas_configure)

    def _on_mousewheel(event):
        canvas.yview_scroll(-1 * int(event.delta / 120), "units")
    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    # container 내부의 그리드 열 중앙 정렬
    container.grid_columnconfigure(0, weight=1)
    container.grid_columnconfigure(1, weight=1)

    # 기본 입력 변수들
    category_var = tk.StringVar()
    club_name_var = tk.StringVar()
    participants_var = tk.StringVar()
    author_var = tk.StringVar()
    # 날짜와 시간에 대한 변수들
    activity_year_var = tk.StringVar()
    activity_month_var = tk.StringVar()
    activity_day_var = tk.StringVar()
    start_hour_var = tk.StringVar()
    start_minute_var = tk.StringVar()
    end_hour_var = tk.StringVar()
    end_minute_var = tk.StringVar()

    # 연령대 성별별 참가자 수 변수 (0~10 선택)
    age_participants_vars = {age: tk.IntVar() for age in age_groups}

    # Row 0: 제목
    title_label = tk.Label(container, text="동아리 활동일지 작성", font=(ft, 25))
    title_label.grid(row=0, column=0, columnspan=2, pady=(20, 20))

    # Row 1: 분야 선택
    tk.Label(container, text="분야 (자동 선택):", font=(ft, 14))\
        .grid(row=1, column=0, sticky="e", padx=10, pady=10)
    
    category_label = tk.Label(container, text="", font=(ft, 14), width=22, anchor="w", relief="sunken")
    category_label.grid(row=1, column=1, sticky="w", padx=10, pady=10)
    # category_dropdown = tk.OptionMenu(container, category_var, *categories)
    # category_dropdown.config(font=(ft, 14), width=20)
    # category_dropdown.grid(row=1, column=1, sticky="w", padx=10, pady=10)

    # Row 2: 동아리 선택
    tk.Label(container, text="동아리 선택:", font=(ft, 14))\
        .grid(row=2, column=0, sticky="e", padx=10, pady=10)

    def update_category(*args):
        club = club_name_var.get()
        category = club2category.get(club, "기타")
        category_label.config(text=category)
        category_var.set(category)

    club_name_var.trace_add("write", update_category)
    update_category()

    club_dropdown = tk.OptionMenu(container, club_name_var, *clubs)
    club_dropdown.config(font=(ft, 14), width=20)
    club_dropdown.grid(row=2, column=1, sticky="w", padx=10, pady=10)

    # 날짜와 시간에 대한 변수들
    today = datetime.today()
    activity_year_var = tk.StringVar(value=today.strftime("%Y"))
    activity_month_var = tk.StringVar(value=today.strftime("%m"))
    activity_day_var = tk.StringVar(value=today.strftime("%d"))

    # Row 3: 활동 날짜 (연도-월-일)
    tk.Label(container, text="활동 날짜:", font=(ft, 14))\
        .grid(row=3, column=0, sticky="e", padx=10, pady=10)
    
    date_frame = tk.Frame(container)
    date_frame.grid(row=3, column=1, sticky="w", padx=10, pady=10)
    
    year_entry = tk.Entry(date_frame, textvariable=activity_year_var, font=(ft, 14), width=6)
    year_entry.pack(side=tk.LEFT)
    
    tk.Label(date_frame, text="-", font=(ft, 14)).pack(side=tk.LEFT)
    
    month_entry = tk.Entry(date_frame, textvariable=activity_month_var, font=(ft, 14), width=3)
    month_entry.pack(side=tk.LEFT)
    
    tk.Label(date_frame, text="-", font=(ft, 14)).pack(side=tk.LEFT)
    
    day_entry = tk.Entry(date_frame, textvariable=activity_day_var, font=(ft, 14), width=3)
    day_entry.pack(side=tk.LEFT)

    # 자동 포커스: 연도 입력 후 4자리면 월로, 월 2자리면 일로 이동
    def check_year(event):
        if len(activity_year_var.get()) >= 4:
            month_entry.focus_set()
    year_entry.bind("<KeyRelease>", check_year)
    
    def check_month(event):
        if len(activity_month_var.get()) >= 2:
            day_entry.focus_set()
    month_entry.bind("<KeyRelease>", check_month)

   # Row 4: 활동 시작 시간 (시:분)
    tk.Label(container, text="(09:00 ~ 20:00) 활동 시작 시간:", font=(ft, 14))\
        .grid(row=4, column=0, sticky="e", padx=10, pady=10)
    start_time_frame = tk.Frame(container)
    start_time_frame.grid(row=4, column=1, sticky="w", padx=10, pady=10)
    start_hour_entry = tk.Entry(start_time_frame, textvariable=start_hour_var, font=(ft, 14), width=3)
    start_hour_entry.pack(side=tk.LEFT)
    tk.Label(start_time_frame, text=":", font=(ft, 14)).pack(side=tk.LEFT)
    start_minute_entry = tk.Entry(start_time_frame, textvariable=start_minute_var, font=(ft, 14), width=3)
    start_minute_entry.pack(side=tk.LEFT)
    
    def check_start_hour(event):
        if len(start_hour_var.get()) >= 2:
            start_minute_entry.focus_set()
    start_hour_entry.bind("<KeyRelease>", check_start_hour)

    # Row 5: 활동 종료 시간 (시:분)
    tk.Label(container, text="(09:00 ~ 20:00) 활동 종료 시간:", font=(ft, 14))\
        .grid(row=5, column=0, sticky="e", padx=10, pady=10)
    end_time_frame = tk.Frame(container)
    end_time_frame.grid(row=5, column=1, sticky="w", padx=10, pady=10)
    end_hour_entry = tk.Entry(end_time_frame, textvariable=end_hour_var, font=(ft, 14), width=3)
    end_hour_entry.pack(side=tk.LEFT)
    tk.Label(end_time_frame, text=":", font=(ft, 14)).pack(side=tk.LEFT)
    end_minute_entry = tk.Entry(end_time_frame, textvariable=end_minute_var, font=(ft, 14), width=3)
    end_minute_entry.pack(side=tk.LEFT)
    
    def check_end_hour(event):
        if len(end_hour_var.get()) >= 2:
            end_minute_entry.focus_set()
    end_hour_entry.bind("<KeyRelease>", check_end_hour)

    # Row 6: 참가자 이름 (쉼표로 구분)
    tk.Label(container, text="참가자 이름 (쉼표로 구분):", font=(ft, 14))\
        .grid(row=6, column=0, sticky="e", padx=10, pady=10)
    tk.Entry(container, textvariable=participants_var, font=(ft, 14), width=30)\
        .grid(row=6, column=1, sticky="w", padx=10, pady=10, ipady=5)

    # Row 7: 연령대 성별별 참가자 수 (0-10)
    tk.Label(container, text="연령대 성별별 참가자 수:", font=(ft, 14))\
        .grid(row=7, column=0, sticky="e", padx=10, pady=10)
    counter_row = 8
    for age_group in age_groups:
        tk.Label(container, text=f"{age_group}:", font=(ft, 14)).grid(row=counter_row, column=0, sticky="e", padx=10, pady=10)
        frame = tk.Frame(container)
        frame.grid(row=counter_row, column=1, padx=10, pady=5, sticky="w")
        tk.Button(frame, text="-", command=lambda v=age_participants_vars[age_group]: decrement(v)).pack(side=tk.LEFT)
        tk.Label(frame, textvariable=age_participants_vars[age_group], font=(ft, 14), width=5).pack(side=tk.LEFT, padx=5)
        tk.Button(frame, text="+", command=lambda v=age_participants_vars[age_group]: increment(v)).pack(side=tk.LEFT)
        counter_row += 1
    
    # Row 16: 활동 내용 라벨
    tk.Label(container, text="활동 내용:", font=(ft, 14))\
        .grid(row=16, column=0, columnspan=2, pady=(10, 0))
    
    # Row 17: 활동 내용 입력 (다중행 Text와 스크롤바)
    content_frame = tk.Frame(container)
    content_frame.grid(row=17, column=0, columnspan=2, padx=10, pady=10)
    activity_content_text = tk.Text(content_frame, font=(ft, 14), width=60, height=5)
    activity_content_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    activity_scroll = tk.Scrollbar(content_frame, command=activity_content_text.yview)
    activity_scroll.pack(side=tk.RIGHT, fill=tk.Y)
    activity_content_text.configure(yscrollcommand=activity_scroll.set)

    # Row 18: 작성자
    tk.Label(container, text="작성자:", font=(ft, 14))\
        .grid(row=18, column=0, sticky="e", padx=10, pady=10)
    tk.Entry(container, textvariable=author_var, font=(ft, 14), width=30)\
        .grid(row=18, column=1, sticky="w", padx=10, pady=10, ipady=5)

    # Row 19: 저장 버튼
    def save_data():
        format_24H(start_hour_var)
        format_two_digits(start_minute_var)
        format_24H(end_hour_var)
        format_two_digits(end_minute_var)
        
        category = category_var.get()
        club_name = club_name_var.get()
        # 날짜와 시간을 결합하여 문자열로 만듦
        activity_date = activity_year_var.get() + "-" + activity_month_var.get() + "-" + activity_day_var.get()
        start_time = start_hour_var.get() + ":" + start_minute_var.get()
        end_time = end_hour_var.get() + ":" + end_minute_var.get()
        participants = participants_var.get()
        activity_content = activity_content_text.get("1.0", "end-1c")
        author = author_var.get()

        if not category:
            show_warning("분야를 선택하세요.")
            return
        if not club_name:
            show_warning("동아리 명을 선택하세요.")
            return
        if not author:
            show_warning("작성자를 입력하세요.")
            return
        
        try:
            datetime.strptime(activity_date, "%Y-%m-%d")
        except ValueError:
            show_warning("잘못된 날짜 형식입니다. 다시 입력하세요.")
            return

        try:
            datetime.strptime(start_time, "%H:%M")
            datetime.strptime(end_time, "%H:%M")
        except ValueError:
            show_warning("잘못된 시간 형식입니다. 다시 입력하세요.")
            return

        total_participants = sum(age_participants_vars[age].get() for age in age_groups)        
        data = {
            "분야": category,
            "동아리 명": club_name,
            "활동 날짜": activity_date,
            "시작 시간": start_time,
            "종료 시간": end_time,
            "참가자": participants,
            "활동 내용": activity_content,
            "작성자": author,
            "총 인원수": total_participants
        }
        for age in age_groups:
            data[age] = age_participants_vars[age].get()

        try:
            save_to_excel(data)
            show_success(f"[{category}/{club_name}]\n{activity_date} {start_time}~{end_time}\n\n{total_participants}명 활동이 기록되었습니다.")
            reset_fields()
        except PermissionError: # file is opened, on this computer or somewhere else
            if is_file_locked(SAVE_PATH):
                show_warning("파일이 열려 있습니다. 닫고 다시 시도해주세요.")
            else:
                show_warning("쓰기 권한이 없거나, 다른 곳에서 파일이 열려 있습니다.\n 관리자에게 확인을 요청해주세요.")

    save_button = tk.Button(container, text="저장", command=save_data, font=(ft, 14), width=20)
    save_button.grid(row=19, column=0, columnspan=2, pady=20)

    root.mainloop()

import os

'''
network_path = r"\\192.168.0.213\nas 공유폴더"
drive_letter = "Z:"

# 네트워크 드라이브를 할당
os.system(f'net use {drive_letter} {network_path} /persistent:yes')

FOLDER = os.path.join(drive_letter, r"백업\2025\■ 동아리 활동 일지")
'''

def save_to_excel(data):
    os.makedirs(FOLDER, exist_ok=True)  # 폴더가 없으면 생성

    # 날짜 문자열을 datetime 객체로 변환
    if isinstance(data.get("활동 날짜"), str):
        try:
            data["활동 날짜"] = datetime.strptime(data["활동 날짜"], "%Y-%m-%d")
        except ValueError:
            pass  # 날짜 형식이 아니면 변환하지 않음

    try:
        df = pd.read_excel(SAVE_PATH, engine='openpyxl')
    except (FileNotFoundError, ValueError):
        df = pd.DataFrame(columns=data.keys())

    df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)

    # 날짜 포맷 유지한 채로 저장
    with pd.ExcelWriter(SAVE_PATH, engine='openpyxl', datetime_format='YYYY-MM-DD') as writer:
        df.to_excel(writer, index=False)

if __name__ == "__main__":
    get_club_activity()
