"""활동일지(logs 테이블) 관련 데이터 처리.

엑셀(pandas) 기반에서 SQLite 기반으로 변경되었습니다.
평소 조회/저장은 순수 SQL, 엑셀 다운로드 시에만 pandas 를 사용합니다.

응답 dict 의 키는 모두 영문(DB 컬럼명) 입니다.
한글 헤더로의 매핑은 export_excel() 한 곳에서만 일어납니다.
"""
from datetime import date, datetime
from io import BytesIO

import pandas as pd

from db import get_db

# 엑셀 export 시 영문 컬럼 -> 한글 헤더 매핑 (컬럼 순서도 이걸로 결정)
EXCEL_HEADERS = {
    "category":      "동아리 분야",
    "club_name":     "동아리 명",
    "activity_date": "활동 일자",
    "start_time":    "시작 시간",
    "end_time":      "종료 시간",
    "participants":  "참가자",
    "content":       "활동 내용",
    "author":        "작성자",
    "total_count":   "총 인원수",
    "elem_m":        "초등 남",
    "elem_f":        "초등 여",
    "mid_m":         "중등 남",
    "mid_f":         "중등 여",
    "high_m":        "고등 남",
    "high_f":        "고등 여",
    "univ_m":        "후기 남",
    "univ_f":        "후기 여",
    "created_at":    "기록일시",
}


def _row_to_dict(row):
    """sqlite3.Row -> 일반 dict (JSON 직렬화 가능)."""
    return {k: row[k] for k in row.keys()}


# ----------------------------------------------------------------------
# 조회
# ----------------------------------------------------------------------
def get_logs():
    """활동일지 전체를 created_at 최신순 dict 리스트로 반환."""
    db = get_db()
    rows = db.execute(
        "SELECT * FROM logs ORDER BY created_at DESC"
    ).fetchall()
    return [_row_to_dict(r) for r in rows]


# ----------------------------------------------------------------------
# 작성
# ----------------------------------------------------------------------
def create_log(data: dict):
    """활동일지 한 건 작성. (성공여부, 메시지, 결과dict) 반환.

    data 는 프론트에서 보낸 JSON 으로, 아래 키를 기대합니다.
      category, club_name, activity_year/month/day,
      start_hour/minute, end_hour/minute,
      participants, author, activity_content,
      element_man/woman, middle_man/woman, high_man/woman, univ_man/woman
    """
    try:
        category = data.get("category", "")
        club_name = data.get("club_name", "")
        year = int(data["activity_year"])
        month = int(data["activity_month"])
        day = int(data["activity_day"])
        start_hour = int(data["start_hour"])
        start_minute = int(data["start_minute"])
        end_hour = int(data["end_hour"])
        end_minute = int(data["end_minute"])
        participants = data.get("participants", "")
        author = data.get("author", "")
        content = data.get("activity_content", "")

        counts = {
            "elem_m": int(data.get("element_man", 0) or 0),
            "elem_f": int(data.get("element_woman", 0) or 0),
            "mid_m":  int(data.get("middle_man", 0) or 0),
            "mid_f":  int(data.get("middle_woman", 0) or 0),
            "high_m": int(data.get("high_man", 0) or 0),
            "high_f": int(data.get("high_woman", 0) or 0),
            "univ_m": int(data.get("univ_man", 0) or 0),
            "univ_f": int(data.get("univ_woman", 0) or 0),
        }
    except (KeyError, ValueError, TypeError):
        return False, "입력값 형식이 올바르지 않습니다.", None

    total = sum(counts.values())
    if total == 0:
        return False, "총 인원수가 0명일 수 없습니다.", None

    # 오후 시간 보정 (예: '1' 입력 -> 13시). 기존 로직 유지.
    if start_hour < 9:
        start_hour += 12
    if end_hour < 9:
        end_hour += 12

    if start_hour * 60 + start_minute >= end_hour * 60 + end_minute:
        return False, "시작 시간은 종료 시간 이전이어야 합니다.", None

    try:
        activity_date = date(year, month, day).isoformat()  # 'YYYY-MM-DD'
    except ValueError:
        return False, "활동 일자가 올바르지 않습니다.", None

    start_time = f"{start_hour:02d}:{start_minute:02d}"
    end_time = f"{end_hour:02d}:{end_minute:02d}"
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    db = get_db()
    db.execute(
        """
        INSERT INTO logs (
            category, club_name, activity_date, start_time, end_time,
            participants, content, author, total_count,
            elem_m, elem_f, mid_m, mid_f, high_m, high_f, univ_m, univ_f,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            category, club_name, activity_date, start_time, end_time,
            participants, content, author, total,
            counts["elem_m"], counts["elem_f"],
            counts["mid_m"], counts["mid_f"],
            counts["high_m"], counts["high_f"],
            counts["univ_m"], counts["univ_f"],
            created_at,
        ),
    )
    db.commit()

    result = {
        "club": club_name,
        "category": category,
        "date": activity_date,
        "people": total,
    }
    return True, "활동일지가 저장되었습니다.", result


# ----------------------------------------------------------------------
# 삭제
# ----------------------------------------------------------------------
def delete_log(log_id):
    """활동일지 한 건 삭제 (id 기준). (성공여부, 메시지) 반환.

    엑셀 시절엔 created_at 문자열로 찾았지만, SQLite 에선 id 로 정확하게.
    """
    try:
        log_id = int(log_id)
    except (TypeError, ValueError):
        return False, "삭제할 항목을 식별할 수 없습니다."

    db = get_db()
    cur = db.execute("DELETE FROM logs WHERE id = ?", (log_id,))
    db.commit()
    if cur.rowcount == 0:
        return False, "해당 활동 이력을 찾을 수 없습니다."
    return True, "활동 이력이 삭제되었습니다."


# ----------------------------------------------------------------------
# 통계 (관리자 대시보드)
# ----------------------------------------------------------------------
def get_dashboard_stats(total_club_count: int):
    """관리자 대시보드용 통계 반환."""
    db = get_db()
    today_str = datetime.now().strftime("%Y-%m-%d")

    today_count = db.execute(
        "SELECT COUNT(*) AS c FROM logs WHERE activity_date = ?",
        (today_str,),
    ).fetchone()["c"]

    recent_rows = db.execute(
        "SELECT * FROM logs ORDER BY created_at DESC LIMIT 5"
    ).fetchall()
    recent = [_row_to_dict(r) for r in recent_rows]

    return {
        "today_activity_count": today_count,
        "total_club_count": total_club_count,
        "recent_logs": recent,
    }


# ----------------------------------------------------------------------
# 엑셀 다운로드
# ----------------------------------------------------------------------
def export_excel():
    """현재 로그를 엑셀 바이트로 반환. (BytesIO, 파일명).

    SQLite -> pandas DataFrame -> 한글 헤더로 rename -> xlsx.
    영문 컬럼명을 한글 헤더로 매핑하는 것이 이 함수의 유일한 책임.
    """
    db = get_db()
    columns = ", ".join(EXCEL_HEADERS.keys())
    df = pd.read_sql_query(
        f"SELECT {columns} FROM logs ORDER BY created_at DESC",
        db,
    )
    df = df.rename(columns=EXCEL_HEADERS)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"동아리 일지_{timestamp}.xlsx"

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output, filename