from flask import Flask, render_template, request, redirect, url_for, session, send_file, Response
from io import BytesIO
import pandas as pd
import os
from datetime import datetime, date

app = Flask(__name__)
app.secret_key = 'Skfma20601318'

# 경로 설정 (예시)
DATA_FOLDER = 'data'
CLUB_LIST_PATH = os.path.join(DATA_FOLDER, 'club_list.xlsx')
DIARY_LOG_PATH = os.path.join(DATA_FOLDER, 'club_log.xlsx')

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
            return render_template('main.html', club_names=club_names, club_dict=club_dict, error="총 인원수가 0명일 수 없습니다.", form=request.form)
        if start_hour < 9: start_hour += 12
        if end_hour < 9: end_hour += 12
        if (start_hour > end_hour) or (start_hour == end_hour and start_minute >= end_minute):
            return render_template('main.html', club_names=club_names, club_dict=club_dict, error="시작 시간은 종료 시간 이전이여야 합니다.", form=request.form)

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

if __name__ == '__main__':
    app.run(debug=True)