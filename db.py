"""SQLite 데이터베이스 — 커넥션 관리 + 스키마 정의.

설계 요점:
  * 컬럼명은 영문(소문자_스네이크) — 쿼리/인덱스 다루기 편하도록.
    한글 헤더는 엑셀 export 시점에만 매핑한다 (services/log_service.py 참고).
  * 한 HTTP 요청 안에서는 같은 커넥션을 재사용 (flask.g).
  * 외래키는 일단 안 걸었다 — 동아리 이름이 logs/weekly_logs 의 단순 텍스트
    컬럼으로 들어가는 기존 동작과 호환 (동아리가 삭제되어도 과거 로그는
    그대로 남는 게 자연스러움).
"""
import os
import sqlite3

from flask import g

from config import DB_FOLDER, DB_PATH


# ----------------------------------------------------------------------
# 스키마 정의
# ----------------------------------------------------------------------
SCHEMA = """
CREATE TABLE IF NOT EXISTS clubs (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    name        TEXT    NOT NULL UNIQUE,
    category    TEXT    NOT NULL
);

CREATE TABLE IF NOT EXISTS logs (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    category        TEXT,
    club_name       TEXT    NOT NULL,
    activity_date   TEXT    NOT NULL,    -- 'YYYY-MM-DD'
    start_time      TEXT    NOT NULL,    -- 'HH:MM'
    end_time        TEXT    NOT NULL,    -- 'HH:MM'
    participants    TEXT    DEFAULT '',
    content         TEXT    DEFAULT '',
    author          TEXT    DEFAULT '',
    total_count     INTEGER NOT NULL,
    elem_m          INTEGER NOT NULL DEFAULT 0,
    elem_f          INTEGER NOT NULL DEFAULT 0,
    mid_m           INTEGER NOT NULL DEFAULT 0,
    mid_f           INTEGER NOT NULL DEFAULT 0,
    high_m          INTEGER NOT NULL DEFAULT 0,
    high_f          INTEGER NOT NULL DEFAULT 0,
    univ_m          INTEGER NOT NULL DEFAULT 0,
    univ_f          INTEGER NOT NULL DEFAULT 0,
    created_at      TEXT    NOT NULL     -- 'YYYY-MM-DD HH:MM:SS'
);

-- 자주 조회/정렬하는 컬럼에 인덱스
CREATE INDEX IF NOT EXISTS idx_logs_created_at  ON logs(created_at);
CREATE INDEX IF NOT EXISTS idx_logs_activity    ON logs(activity_date);
"""


# ----------------------------------------------------------------------
# 커넥션 헬퍼
# ----------------------------------------------------------------------
def _connect():
    """파일 핸들 한 개 생성. (Flask 컨텍스트 밖에서도 사용 가능)"""
    os.makedirs(DB_FOLDER, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row   # row['column_name'] 접근 가능
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def get_db():
    """현재 요청 동안 공유되는 커넥션을 반환.

    Flask 의 g 에 캐싱 -> 같은 요청 안에서 여러 번 호출해도 한 번만 연결.
    """
    if "db" not in g:
        g.db = _connect()
    return g.db


def close_db(_=None):
    """요청 종료 시 커넥션 닫기. app.teardown_appcontext 에 등록해서 사용."""
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    """테이블이 없으면 생성. 앱 시작 시 1회 호출한다."""
    conn = _connect()
    try:
        conn.executescript(SCHEMA)
        conn.commit()
    finally:
        conn.close()


# ----------------------------------------------------------------------
# 디버그용 — 단독 실행하면 DB 초기화
# ----------------------------------------------------------------------
if __name__ == "__main__":
    init_db()
    print(f"DB initialized: {DB_PATH}")