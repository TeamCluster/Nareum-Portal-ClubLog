"""동아리 목록(clubs 테이블) 관련 데이터 처리 — 멀티테넌트.

각 함수가 slug 를 받아 해당 기관의 DB(db/<slug>.sqlite3) 에 접근합니다.
slug 유효성 + places 테이블 존재 여부는 라우트 단계에서 검증되어
이 모듈에는 이미 유효한 slug 가 들어온다고 가정합니다.
"""
import sqlite3

from db import get_place_db
from utils.sorting import sort_key_korean_first


# ----------------------------------------------------------------------
# 조회
# ----------------------------------------------------------------------
def get_clubs(slug: str):
    """기관의 동아리 목록 [{name, category}, ...] (한글 우선 정렬)."""
    db = get_place_db(slug)
    rows = db.execute("SELECT name, category FROM clubs").fetchall()
    clubs = [{"name": r["name"], "category": r["category"]} for r in rows]
    clubs.sort(key=lambda c: sort_key_korean_first(c["name"]))
    return clubs


def get_club_dict(slug: str):
    """{동아리명: 분야} 딕셔너리 반환. 프론트의 분야 자동 채움에 사용."""
    return {c["name"]: c["category"] for c in get_clubs(slug)}


# ----------------------------------------------------------------------
# 변경
# ----------------------------------------------------------------------
def add_club(slug: str, name: str, category: str):
    """동아리 추가. (성공여부, 메시지) 반환."""
    name = (name or "").strip()
    category = (category or "").strip()
    if not name or not category:
        return False, "동아리 이름과 분야를 모두 입력해주세요."

    db = get_place_db(slug)
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


def delete_club(slug: str, name: str):
    """동아리 삭제. (성공여부, 메시지) 반환.

    logs.club_name 은 단순 텍스트 컬럼이라 동아리를 삭제해도 과거
    활동 이력은 그대로 보존됨.
    """
    db = get_place_db(slug)
    cur = db.execute("DELETE FROM clubs WHERE name = ?", (name,))
    db.commit()
    if cur.rowcount == 0:
        return False, "해당 동아리를 찾을 수 없습니다."
    return True, f"'{name}' 동아리가 삭제되었습니다."