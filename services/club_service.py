"""동아리 목록(clubs 테이블) 관련 데이터 처리.

엑셀(pandas) 기반에서 SQLite 기반으로 변경되었습니다.
외부 인터페이스는 기존과 동일 — get_clubs / get_club_dict /
add_club / delete_club / get_clubs_with_auth.

주의: get_clubs_with_auth() 의 반환 타입이 DataFrame -> dict 리스트로
변경되었습니다. 호출처(app.py, weekly_service.py) 도 함께 조정 필요.
"""
import sqlite3

from db import get_db
from utils.sorting import sort_key_korean_first


# ----------------------------------------------------------------------
# 조회
# ----------------------------------------------------------------------
def get_clubs():
    """동아리 목록을 [{name, category}, ...] 형태로 반환 (한글 우선 정렬).

    SQLite 는 기본 콜레이션이 한글 정렬에 부적합하므로,
    Python 쪽에서 sort_key_korean_first 로 다시 정렬합니다.
    """
    db = get_db()
    rows = db.execute("SELECT name, category FROM clubs").fetchall()
    clubs = [{"name": r["name"], "category": r["category"]} for r in rows]
    clubs.sort(key=lambda c: sort_key_korean_first(c["name"]))
    return clubs


def get_club_dict():
    """{동아리명: 분야} 딕셔너리 반환. 프론트의 분야 자동 채움에 사용."""
    return {c["name"]: c["category"] for c in get_clubs()}


# ----------------------------------------------------------------------
# 변경
# ----------------------------------------------------------------------
def add_club(name: str, category: str):
    """동아리 추가. (성공여부, 메시지) 반환."""
    name = (name or "").strip()
    category = (category or "").strip()
    if not name or not category:
        return False, "동아리 이름과 분야를 모두 입력해주세요."

    db = get_db()
    try:
        db.execute(
            "INSERT INTO clubs (name, category) VALUES (?, ?)",
            (name, category),
        )
        db.commit()
    except sqlite3.IntegrityError:
        # name UNIQUE 제약 위반
        return False, "이미 존재하는 동아리입니다."

    return True, f"'{name}' 동아리가 추가되었습니다."


def delete_club(name: str):
    """동아리 삭제. (성공여부, 메시지) 반환.

    참고: logs / weekly_logs 의 club_name 은 단순 텍스트 컬럼이므로
    동아리를 삭제해도 과거 활동 이력은 그대로 보존됩니다.
    """
    db = get_db()
    cur = db.execute("DELETE FROM clubs WHERE name = ?", (name,))
    db.commit()
    if cur.rowcount == 0:
        return False, "해당 동아리를 찾을 수 없습니다."
    return True, f"'{name}' 동아리가 삭제되었습니다."