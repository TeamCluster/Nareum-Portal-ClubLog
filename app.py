from flask import Flask, render_template, request, redirect, url_for, session, send_file, Response, flash
from io import BytesIO
import pandas as pd
import os
from datetime import datetime, date, timedelta

app = Flask(__name__)
app.secret_key = 'Skfma20601318'

# 경로 설정 (예시)
DATA_FOLDER = 'data'
CLUB_LIST_PATH = os.path.join(DATA_FOLDER, 'club_list.xlsx')
DIARY_LOG_PATH = os.path.join(DATA_FOLDER, 'club_log.xlsx')
WEEKLY_LOG_PATH = os.path.join(DATA_FOLDER, "club_weekly.xlsx")

# club_list.xlsx를 딕셔너리로 로딩
club_df = pd.read_excel(CLUB_LIST_PATH)
club_df.columns = club_df.columns.str.strip()

#print("엑셀 컬럼명:", club_df.columns.tolist())  # 디버깅용
club_dict = dict(zip(club_df['동아리 명'], club_df['동아리 분야']))

# 초기 페이지: 활동일지 작성 폼
@app.route('/', methods=['GET', 'POST'])
def main():
    club_names = club_df['동아리 명'].tolist()

    if request.method == 'POST':
        # 입력값 수집
        category = request.form.get('category')
        club_name = request.form.get('club_name')
        year = request.form.get('activity_year')
        month = request.form.get('activity_month')
        day = request.form.get('activity_day')
        start_hour = int(request.form.get('start_hour'))
        start_minute = int(request.form.get('start_minute'))
        end_hour = int(request.form.get('end_hour'))
        end_minute = int(request.form.get('end_minute'))
        participants = request.form.get('participants')
        author = request.form.get('author')
        content = request.form.get('activity_content')
        element_man = int(request.form.get('element_man'))
        element_woman = int(request.form.get('element_woman'))
        middle_man = int(request.form.get('middle_man'))
        middle_woman = int(request.form.get('middle_woman'))
        high_man = int(request.form.get('high_man'))
        high_woman = int(request.form.get('high_woman'))
        univ_man = int(request.form.get('univ_man'))
        univ_woman = int(request.form.get('univ_woman'))
        party_num = element_man + element_woman + middle_man + middle_woman + high_man + high_woman + univ_man + univ_woman
        if party_num == 0:
            return """
                <script>
                    alert("총 인원수가 0명일 수 없습니다.");
                    history.back();
                </script>
                """
        if start_hour < 9: start_hour += 12
        if end_hour < 9: end_hour += 12
        if start_hour * 60 + start_minute >= end_hour * 60 + end_minute:  # 시간 역순 체크
            return """
                <script>
                    alert("시작 시간은 종료 시간 이전이어야 합니다.");
                    history.back();
                </script>
                """

        # 저장할 DataFrame 생성
        df = pd.DataFrame([{
            '동아리 분야': category,
            '동아리 명': club_name,
            '활동 일자': date(int(year), int(month), int(day)),
            '시작 시간': f"{start_hour}:{start_minute}",
            '종료 시간': f"{end_hour}:{end_minute}",
            '참가자': participants,
            '활동 내용': content,
            '작성자': author,
            '총 인원수': party_num,
            '초등 남': element_man,
            '초등 여': element_woman,
            '중등 남': middle_man,
            '중등 여': middle_woman,
            '고등 남': high_man,
            '고등 여': high_woman,
            '후기 남': univ_man,
            '후기 여': univ_woman,
            '기록일시': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }])

        # 파일이 존재하면 이어쓰기, 없으면 새로 쓰기
        if os.path.exists(DIARY_LOG_PATH):
            with pd.ExcelWriter(DIARY_LOG_PATH, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
                startrow = writer.sheets['Sheet1'].max_row
                df.to_excel(writer, index=False, header=False, startrow=startrow)
        else:
            df.to_excel(DIARY_LOG_PATH, index=False)

        # 세션에 결과 저장
        session['club'] = club_name
        session['category'] = category
        session['date'] = f"{year}-{month}-{day}"
        session['people'] = party_num

        return redirect(url_for('success_popup'))

    return render_template('main.html', club_names=club_names, club_dict=club_dict)

def form():
    pass  # 이전 main 함수는 이제 필요 없음. 빈껍데기로 유지.

# 입력 성공 팝업
@app.route('/success')
def success_popup():
    club = session.get('club')
    category = session.get('category')
    date = session.get('date')
    people = session.get('people')

    # 세션에서 읽은 후 지우기 (선택사항)
    session.pop('club', None)
    session.pop('category', None)
    session.pop('date', None)
    session.pop('people', None)

    return render_template('success_popup.html', club=club, category=category, date=date, people=people)

# 설정 페이지
@app.route('/setting')
def setting():
    club_df = pd.read_excel(CLUB_LIST_PATH) if os.path.exists(CLUB_LIST_PATH) else pd.DataFrame()
    log_df = pd.read_excel(DIARY_LOG_PATH) if os.path.exists(DIARY_LOG_PATH) else pd.DataFrame()
    return render_template('setting.html', club_df=club_df.to_html(index=False), log_df=log_df.to_html(index=False))

# 로그 파일 다운로드 받기
@app.route("/download_log")
def download_log():
    try:
        # 엑셀 파일 읽기
        log_df = pd.read_excel(DIARY_LOG_PATH)

        # 현재 시각 기반 파일명 만들기
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"동아리 일지_{timestamp}.xlsx"

        # 메모리 버퍼에 저장
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            log_df.to_excel(writer, index=False)
        output.seek(0)  # 스트림 처음으로 이동

        # 파일 전송
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as err:
        print(err)
        return Response(str(err), status=500)
    
def load_club_list():
    if not os.path.exists(CLUB_LIST_PATH):
        # 최소 스키마 생성
        df = pd.DataFrame(columns=["동아리 명", "동아리 분야", "인증번호"])
        df.to_excel(CLUB_LIST_PATH, index=False)
    df = pd.read_excel(CLUB_LIST_PATH)
    # 인증번호 컬럼 없으면 채워줌
    if "인증번호" not in df.columns:
        df["인증번호"] = "0000"
    df["인증번호"] = df["인증번호"].fillna("0000").astype(str).str.zfill(4)
    return df

def ensure_weekly_log():
    if not os.path.exists(WEEKLY_LOG_PATH):
        df = pd.DataFrame(columns=["동아리 명", "활동 주차", "활동 여부", "작성 일시", "분야"])
        df.to_excel(WEEKLY_LOG_PATH, index=False)

def append_weekly_log(rowdict):
    ensure_weekly_log()
    df = pd.read_excel(WEEKLY_LOG_PATH)
    df = pd.concat([df, pd.DataFrame([rowdict])], ignore_index=True)
    df.to_excel(WEEKLY_LOG_PATH, index=False)

# ------- 주차 계산 (요구 스펙) -------
"""
입력 가능 기한: 토요일 00:00:00 ~ 금요일 23:59:59
이 기간에 '등록 대상 주'는 다음 규칙:
- 토/일(5/6요일): 다음 주 월~일
- 월~금(0~4요일): 이번 주 월~일
"""
def get_target_monday(now):
    wd = now.weekday()  # Mon=0 ... Sun=6
    if wd in (5, 6):  # Sat/Sun
        # 다음 주 월요일
        days = (7 - wd) % 7
        if days == 0:
            days = 7
        return (now + timedelta(days=days)).replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        # 이번 주 월요일
        return (now - timedelta(days=wd)).replace(hour=0, minute=0, second=0, microsecond=0)

def get_week_header(now):
    mon = get_target_monday(now)
    days = [mon + timedelta(days=i) for i in range(7)]  # 월~일
    # ISO 주차(월 시작)
    year, week_num, _ = days[0].isocalendar()
    return year, week_num, days

def within_submission_window(now):
    mon = get_target_monday(now)
    window_start = mon - timedelta(days=2)                       # 토요일 00:00:00
    window_end   = mon + timedelta(days=4, hours=23, minutes=59, seconds=59)  # 금요일 23:59:59
    return window_start <= now <= window_end

# ------- 주간 입력 페이지 -------
@app.route("/weekly", methods=["GET", "POST"])
def weekly():
    now = datetime.now()
    clubs_df = load_club_list()
    # 화면용: 동아리명 리스트
    club_names = clubs_df["동아리 명"].dropna().tolist()

    # 주차 헤더 데이터 (2025년 N주차 + 날짜 라벨)
    year, week_num, days = get_week_header(now)
    week_label = f"{year}년 {week_num}주차"
    day_labels = [d.strftime("%m-%d") for d in days]  # "09-01" 형식

    if request.method == "POST":
        if not within_submission_window(now):
            flash("현재는 입력 가능 기간(토요일~금요일)이 아닙니다.", "error")
            return redirect(url_for("weekly"))

        club_name = request.form.get("club_name", "").strip()
        activity_status = request.form.get("activity_status", "동아리 활동 일정 없음")
        auth_code_input = (request.form.get("club_auth", "") or "").strip()

        # 유효성: 동아리/분야 조회 & 인증번호 대조
        row = clubs_df[clubs_df["동아리 명"] == club_name]
        if row.empty:
            flash("동아리 명을 선택해 주세요.", "error")
            return redirect(url_for("weekly"))

        real_auth = str(row.iloc[0]["인증번호"]).zfill(4)
        category  = row.iloc[0]["동아리 분야"] if "동아리 분야" in row.columns else ""

        if not (auth_code_input.isdigit() and len(auth_code_input) == 4):
            flash("동아리 인증번호는 숫자 4자리여야 합니다.", "error")
            return redirect(url_for("weekly"))

        if auth_code_input != real_auth:
            flash("동아리 인증번호가 올바르지 않습니다.", "error")
            return redirect(url_for("weekly"))

        # 저장
        append_weekly_log({
            "동아리 명": club_name,
            "활동 주차": week_label,
            "활동 여부": activity_status,
            "작성 일시": now.strftime("%Y-%m-%d %H:%M:%S"),
            "분야": category
        })
        return redirect(url_for("weekly_success", week=week_label, category=category, club=club_name, activity=activity_status))

    return render_template(
        "weekly.html",
        club_names=club_names,
        week_label=week_label,
        day_labels=day_labels
    )

# ------- 성공 팝업 -------
@app.route("/weekly/success")
def weekly_success():
    week = request.args.get("week", "")
    category = request.args.get("category", "")
    club = request.args.get("club", "")
    activity = request.args.get("activity", "")
    return render_template(
        "weekly_success.html",
        week=week, category=category, club=club, activity=activity
    )

# ------- 세팅 페이지 (주간) -------
@app.route("/weekly/setting")
def weekly_setting():
    clubs_df = load_club_list().rename(columns={
        "동아리 명": "동아리 명",
        "동아리 분야": "분야"
    })
    ensure_weekly_log()
    weekly_df = pd.read_excel(WEEKLY_LOG_PATH)

    # 테이블 HTML (간단 스타일용 class 부여)
    club_table_html = clubs_df[["동아리 명", "분야", "인증번호"]].to_html(index=False, classes="table table-sm table-striped", border=0)
    weekly_table_html = weekly_df[["동아리 명", "활동 주차", "활동 여부", "작성 일시"]].to_html(index=False, classes="table table-sm table-striped", border=0)

    return render_template(
        "weekly_setting.html",
        club_table_html=club_table_html,
        weekly_table_html=weekly_table_html
    )

# 주간 활동 파일 다운로드 하기
@app.route("/download_weekly")
def download_weekly():
    try:
        # 엑셀 파일 읽기
        log_df = pd.read_excel(WEEKLY_LOG_PATH)

        # 현재 시각 기반 파일명 만들기
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"동아리 주간 활동 여부_{timestamp}.xlsx"

        # 메모리 버퍼에 저장
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            log_df.to_excel(writer, index=False)
        output.seek(0)  # 스트림 처음으로 이동

        # 파일 전송
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as err:
        print(err)
        return Response(str(err), status=500)

if __name__ == '__main__':
    app.run(debug=True)