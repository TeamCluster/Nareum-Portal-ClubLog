"""기관(places) CRUD — 슈퍼 관리자 페이지에서 사용.

DB:  db/super.sqlite3 의 places 테이블

이름은 두 가지로 받음:
  full_name   풀네임 (예: "나름청소년활동센터") — 일지 작성 페이지 상단
  short_name  축약별칭 (예: "나름") — 관리자 헤더

핵심 동작:
  * add_place: 검증 -> 파일 생성 -> INSERT (이 순서로 옵션 B 와 자연스럽게 맞물림)
  * delete_place: places 행만 삭제, db/<slug>.sqlite3 파일은 보존 (옵션 B).
    같은 slug 로 재추가 시 이전 데이터가 그대로 복구됨.
"""
import sqlite3
from datetime import datetime

from werkzeug.security import check_password_hash, generate_password_hash

from config import is_valid_slug
from db import get_super_db, init_place_db

MIN_PASSWORD_LENGTH = 6
MAX_FULL_NAME_LENGTH = 100
MAX_SHORT_NAME_LENGTH = 20
MAX_SYNC_URL_LENGTH = 300


# ----------------------------------------------------------------------
# 재단 동기화 URL 검증
# ----------------------------------------------------------------------
def validate_sync_url(url: str):
    """재단 동기화 URL 검증. (ok, cleaned, msg) 반환.

    빈값은 허용 — '동기화 안 함' 을 의미한다. 값이 있으면 Apps Script 웹앱
    배포 주소 형태(https://script.google.com/macros/.../exec)만 받는다.
    """
    url = (url or "").strip()
    if not url:
        return True, "", ""  # 빈값 = 동기화 비활성
    if len(url) > MAX_SYNC_URL_LENGTH:
        return False, url, f"URL은 {MAX_SYNC_URL_LENGTH}자 이내여야 합니다."
    if not url.startswith("https://"):
        return False, url, "URL은 https:// 로 시작해야 합니다."
    if "script.google.com/macros/" not in url or not url.endswith("/exec"):
        return False, url, (
            "Google Apps Script 웹앱 배포 주소여야 합니다. "
            "(https://script.google.com/macros/s/.../exec 형식)"
        )
    return True, url, ""


# ----------------------------------------------------------------------
# 조회
# ----------------------------------------------------------------------
def get_places():
    """기관 목록 [{id, slug, full_name, short_name, foundation_sync_url,
    created_at}, ...] (최신순). 비밀번호 해시는 응답에 포함하지 않음."""
    rows = get_super_db().execute(
        "SELECT id, slug, full_name, short_name, foundation_sync_url, created_at"
        " FROM places ORDER BY created_at DESC"
    ).fetchall()
    return [dict(r) for r in rows]


def get_place(slug: str):
    """단건 조회. 없으면 None."""
    row = get_super_db().execute(
        "SELECT id, slug, full_name, short_name, foundation_sync_url, created_at"
        " FROM places WHERE slug = ?",
        (slug,),
    ).fetchone()
    return dict(row) if row else None


# ----------------------------------------------------------------------
# 추가
# ----------------------------------------------------------------------
def add_place(slug: str, full_name: str, short_name: str, password: str,
              foundation_sync_url: str = ""):
    """기관 추가. (성공여부, 메시지, 결과dict) 반환.

    흐름: 검증 -> init_place_db() (파일 생성) -> INSERT.
    foundation_sync_url 은 선택값(빈값 허용 = 동기화 안 함).
    """
    # --- 슬러그 ---
    if not is_valid_slug(slug):
        return False, (
            "슬러그(영문이름)는 영문 소문자로 시작하는 2~30자여야 합니다. "
            "허용 문자: 소문자, 숫자, 하이픈(-), 언더스코어(_). "
            "예약어(super/api/admin/static/public)는 사용 불가."
        ), None

    # --- 풀네임 / 축약별칭 ---
    full_name = (full_name or "").strip()
    short_name = (short_name or "").strip()
    if not full_name:
        return False, "기관 풀네임을 입력해주세요.", None
    if not short_name:
        return False, "기관 축약별칭을 입력해주세요.", None
    if len(full_name) > MAX_FULL_NAME_LENGTH:
        return False, f"기관 풀네임은 {MAX_FULL_NAME_LENGTH}자 이내여야 합니다.", None
    if len(short_name) > MAX_SHORT_NAME_LENGTH:
        return False, f"기관 축약별칭은 {MAX_SHORT_NAME_LENGTH}자 이내여야 합니다.", None

    # --- 비밀번호 ---
    if not isinstance(password, str) or len(password) < MIN_PASSWORD_LENGTH:
        return False, (
            f"기관 비밀번호는 {MIN_PASSWORD_LENGTH}자 이상이어야 합니다."
        ), None

    # --- 재단 동기화 URL (선택) ---
    ok, sync_url, msg = validate_sync_url(foundation_sync_url)
    if not ok:
        return False, msg, None

    # --- DB 파일 + 스키마 보장 (멱등) ---
    init_place_db(slug)

    # --- places 행 추가 ---
    db = get_super_db()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        db.execute(
            "INSERT INTO places"
            " (slug, full_name, short_name, password_hash, foundation_sync_url, created_at)"
            " VALUES (?, ?, ?, ?, ?, ?)",
            (slug, full_name, short_name, generate_password_hash(password),
             sync_url, now),
        )
        db.commit()
    except sqlite3.IntegrityError:
        return False, f"슬러그 '{slug}' 는 이미 사용 중입니다.", None

    return True, f"기관 '{full_name}' 이(가) 추가되었습니다.", {
        "slug": slug,
        "full_name": full_name,
        "short_name": short_name,
        "foundation_sync_url": sync_url,
        "created_at": now,
    }


# ----------------------------------------------------------------------
# 삭제 (옵션 B — DB 파일 보존)
# ----------------------------------------------------------------------
def delete_place(slug: str):
    """기관 삭제. places 행만 삭제, db/<slug>.sqlite3 파일은 보존."""
    db = get_super_db()
    cur = db.execute("DELETE FROM places WHERE slug = ?", (slug,))
    db.commit()
    if cur.rowcount == 0:
        return False, "해당 기관을 찾을 수 없습니다."
    return True, (
        f"기관 '{slug}' 이(가) 삭제되었습니다. "
        f"데이터 파일은 보존되어 있어 같은 슬러그로 재추가하면 복구됩니다."
    )


# ----------------------------------------------------------------------
# 비밀번호 변경 / 검증
# ----------------------------------------------------------------------
def update_place_password(slug: str, new_password: str):
    if not isinstance(new_password, str) or len(new_password) < MIN_PASSWORD_LENGTH:
        return False, f"비밀번호는 {MIN_PASSWORD_LENGTH}자 이상이어야 합니다."

    db = get_super_db()
    cur = db.execute(
        "UPDATE places SET password_hash = ? WHERE slug = ?",
        (generate_password_hash(new_password), slug),
    )
    db.commit()
    if cur.rowcount == 0:
        return False, "해당 기관을 찾을 수 없습니다."
    return True, "기관 비밀번호가 변경되었습니다."


# ----------------------------------------------------------------------
# 재단 동기화 URL 등록 / 수정 / 해제
# ----------------------------------------------------------------------
def update_place_sync_url(slug: str, url: str):
    """기관의 재단 동기화 URL 을 등록/수정/해제. (성공여부, 메시지) 반환.

    빈 문자열을 보내면 동기화를 끈다(URL 제거).
    """
    ok, sync_url, msg = validate_sync_url(url)
    if not ok:
        return False, msg

    db = get_super_db()
    cur = db.execute(
        "UPDATE places SET foundation_sync_url = ? WHERE slug = ?",
        (sync_url, slug),
    )
    db.commit()
    if cur.rowcount == 0:
        return False, "해당 기관을 찾을 수 없습니다."
    if sync_url:
        return True, "재단 동기화 URL이 저장되었습니다."
    return True, "재단 동기화가 해제되었습니다. (URL 제거)"


def verify_place_password(slug: str, password: str) -> bool:
    if not isinstance(password, str) or not password:
        return False

    row = get_super_db().execute(
        "SELECT password_hash FROM places WHERE slug = ?", (slug,)
    ).fetchone()
    if not row:
        return False

    return check_password_hash(row["password_hash"], password)