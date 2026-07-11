"""SQLite 데이터베이스 — 멀티테넌트 + 자체 부트스트랩.

구조:
  db/super.sqlite3
      places          # 기관 목록 (slug, name, password_hash)
      super_admin     # 슈퍼 관리자 비밀번호 해시 (단일 row)
      app_settings    # 시스템 설정 (SECRET_KEY 등 key-value)
  db/<slug>.sqlite3
      clubs, logs     # 기관별 실제 데이터

설계 요점:
  * 한 HTTP 요청 안에서 super DB 와 여러 기관 DB 에 동시에 접근할 수 있음.
    flask.g.dbs 에 {'super': conn, '<slug>': conn, ...} 형태로 캐싱.
  * 첫 실행 시 init_super_db() 가:
      - 슈퍼 비밀번호가 없으면 임시값을 생성해 콘솔에 1회 출력
      - SECRET_KEY 가 없으면 랜덤 생성해 app_settings 에 저장
    설정 파일(.env) 없이 동작하는 구조.
"""
import os
import secrets
import sqlite3
from datetime import datetime

from flask import g
from werkzeug.security import generate_password_hash

import config
from config import DB_FOLDER, SUPER_DB_PATH, place_db_path


# ----------------------------------------------------------------------
# 스키마
# ----------------------------------------------------------------------
SUPER_SCHEMA = """
CREATE TABLE IF NOT EXISTS places (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    slug                TEXT    NOT NULL UNIQUE,
    full_name           TEXT    NOT NULL,     -- 풀네임 (예: 나름청소년활동센터)
    short_name          TEXT    NOT NULL,     -- 축약별칭 (예: 나름)
    password_hash       TEXT    NOT NULL,
    foundation_sync_url TEXT    DEFAULT '',   -- 재단 동기화용 Apps Script /exec URL (빈값=동기화 안 함)
    created_at          TEXT    NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_places_slug ON places(slug);

-- 슈퍼 관리자는 1명뿐 -> id = 1 만 허용하는 단일 row 테이블.
CREATE TABLE IF NOT EXISTS super_admin (
    id              INTEGER PRIMARY KEY CHECK (id = 1),
    password_hash   TEXT    NOT NULL,
    updated_at      TEXT    NOT NULL
);

-- 시스템 자동 설정 (SECRET_KEY 등). 향후 key-value 로 확장 가능.
CREATE TABLE IF NOT EXISTS app_settings (
    key     TEXT PRIMARY KEY,
    value   TEXT NOT NULL
);
"""

PLACE_SCHEMA = """
CREATE TABLE IF NOT EXISTS clubs (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    name        TEXT    NOT NULL UNIQUE,
    category    TEXT    NOT NULL
);

CREATE TABLE IF NOT EXISTS logs (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    category        TEXT,
    club_name       TEXT    NOT NULL,
    activity_date   TEXT    NOT NULL,
    start_time      TEXT    NOT NULL,
    end_time        TEXT    NOT NULL,
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
    created_at      TEXT    NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_logs_created_at  ON logs(created_at);
CREATE INDEX IF NOT EXISTS idx_logs_activity    ON logs(activity_date);
"""


# ----------------------------------------------------------------------
# 커넥션 헬퍼
# ----------------------------------------------------------------------
def _connect(path: str) -> sqlite3.Connection:
    """주어진 경로의 SQLite 에 연결. Flask 컨텍스트 밖에서도 사용 가능."""
    os.makedirs(DB_FOLDER, exist_ok=True)
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def _get_cached(key: str, path: str) -> sqlite3.Connection:
    """g.dbs 에서 해당 key 의 커넥션을 꺼내거나, 없으면 새로 만들어 캐싱."""
    if "dbs" not in g:
        g.dbs = {}
    if key not in g.dbs:
        g.dbs[key] = _connect(path)
    return g.dbs[key]


def get_super_db() -> sqlite3.Connection:
    """슈퍼 DB 커넥션."""
    return _get_cached("super", SUPER_DB_PATH)


def get_place_db(slug: str) -> sqlite3.Connection:
    """기관 DB 커넥션. 슬러그 유효성/존재 여부는 라우트에서 검증."""
    return _get_cached(slug, place_db_path(slug))


def close_dbs(_=None) -> None:
    """요청 종료 시 열려있는 모든 커넥션을 닫음."""
    dbs = g.pop("dbs", None)
    if dbs:
        for conn in dbs.values():
            try:
                conn.close()
            except Exception:
                pass


# ----------------------------------------------------------------------
# 스키마 초기화 + 첫 실행 부트스트랩
# ----------------------------------------------------------------------
def init_super_db() -> None:
    """앱 시작 시 슈퍼 DB 의 테이블 + 초기값(슈퍼 비밀번호, SECRET_KEY)을 보장.

    멱등 — 이미 값이 있으면 건드리지 않음.
    옛 스키마(name 컬럼만 있는 경우) 자동 마이그레이션 포함.
    """
    conn = _connect(SUPER_DB_PATH)
    try:
        conn.executescript(SUPER_SCHEMA)
        _migrate_places_schema(conn)
        _backfill_sync_url(conn)

        # 1) 슈퍼 비밀번호가 없으면 임시값 생성 + 콘솔 출력
        row = conn.execute(
            "SELECT id FROM super_admin WHERE id = 1"
        ).fetchone()
        if not row:
            temp_password = secrets.token_urlsafe(12)
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            conn.execute(
                "INSERT INTO super_admin (id, password_hash, updated_at)"
                " VALUES (1, ?, ?)",
                (generate_password_hash(temp_password), now),
            )
            _print_initial_password_banner(temp_password)

        # 2) SECRET_KEY 가 없으면 랜덤 생성
        row = conn.execute(
            "SELECT value FROM app_settings WHERE key = 'secret_key'"
        ).fetchone()
        if not row:
            conn.execute(
                "INSERT INTO app_settings (key, value) VALUES ('secret_key', ?)",
                (secrets.token_hex(32),),
            )

        conn.commit()
    finally:
        conn.close()


def _migrate_places_schema(conn: sqlite3.Connection) -> None:
    """옛 스키마에서 새 스키마로 자동 마이그레이션.

    옛 스키마: places(name) 만 있음
    새 스키마: places(full_name, short_name, foundation_sync_url)

    옛 name 값을 두 컬럼에 그대로 복사. 마이그레이션 후엔 운영자가
    슈퍼 페이지에서 short_name 을 짧게 다듬으면 됨.

    ALTER TABLE ADD COLUMN 은 기존 행/데이터를 건드리지 않으므로(새 컬럼은
    기본값으로 채워짐) 기존 데이터는 손상 없이 그대로 승계된다.
    """
    cols = {r[1] for r in conn.execute("PRAGMA table_info(places)").fetchall()}

    if "full_name" not in cols:
        conn.execute("ALTER TABLE places ADD COLUMN full_name TEXT")
    if "short_name" not in cols:
        conn.execute("ALTER TABLE places ADD COLUMN short_name TEXT")

    # 재단 동기화 URL 컬럼 추가 (값 승계는 _backfill_sync_url 이 담당).
    if "foundation_sync_url" not in cols:
        conn.execute(
            "ALTER TABLE places ADD COLUMN foundation_sync_url TEXT DEFAULT ''"
        )

    # 옛 name 컬럼 데이터 복사 (있을 때만)
    if "name" in cols:
        conn.execute(
            "UPDATE places SET full_name = name "
            "WHERE full_name IS NULL OR full_name = ''"
        )
        conn.execute(
            "UPDATE places SET short_name = name "
            "WHERE short_name IS NULL OR short_name = ''"
        )


def _backfill_sync_url(conn: sqlite3.Connection) -> None:
    """config.FOUNDATION_SYNC 의 URL 을 DB(places.foundation_sync_url)로 1회 승계.

    app_settings 의 'fsync_backfilled' 플래그로 멱등 보장 — startup 마다 또는
    컬럼이 (크래시 등으로) 먼저 생성된 경우에도 정확히 1회만 실행된다.
    값이 이미 있는 기관은 건드리지 않으므로, 운영자가 슈퍼 페이지에서 비우거나
    수정한 값은 보존된다.
    """
    done = conn.execute(
        "SELECT 1 FROM app_settings WHERE key = 'fsync_backfilled'"
    ).fetchone()
    if done:
        return

    for s, c in (getattr(config, "FOUNDATION_SYNC", {}) or {}).items():
        url = (c or {}).get("url")
        if url:
            conn.execute(
                "UPDATE places SET foundation_sync_url = ? "
                "WHERE slug = ? AND (foundation_sync_url IS NULL OR foundation_sync_url = '')",
                (url, s),
            )

    conn.execute(
        "INSERT OR REPLACE INTO app_settings (key, value)"
        " VALUES ('fsync_backfilled', '1')"
    )


def _print_initial_password_banner(password: str) -> None:
    """첫 실행 시 슈퍼 임시 비밀번호를 콘솔에 1회 출력."""
    bar = "=" * 64
    print()
    print(bar)
    print("  슈퍼 관리자 초기 비밀번호:")
    print(f"      {password}")
    print()
    print("  이 비밀번호로 첫 로그인한 뒤 즉시 변경하세요.")
    print("  이 메시지는 다시 표시되지 않습니다.")
    print(bar)
    print(flush=True)


def init_place_db(slug: str) -> None:
    """기관 추가 시 해당 기관의 DB 파일 + 스키마를 생성 (멱등)."""
    conn = _connect(place_db_path(slug))
    try:
        conn.executescript(PLACE_SCHEMA)
        conn.commit()
    finally:
        conn.close()


def place_db_exists(slug: str) -> bool:
    """기관 DB 파일이 디스크에 존재하는지."""
    return os.path.exists(place_db_path(slug))


# ----------------------------------------------------------------------
# 시스템 설정 접근자 (app_settings)
# ----------------------------------------------------------------------
def get_secret_key() -> str:
    """Flask SECRET_KEY 를 슈퍼 DB 에서 가져옴.

    init_super_db() 가 먼저 실행되어 있어야 함. 앱 시작 시 1회 호출하여
    app.secret_key 에 세팅.
    """
    conn = _connect(SUPER_DB_PATH)
    try:
        row = conn.execute(
            "SELECT value FROM app_settings WHERE key = 'secret_key'"
        ).fetchone()
        if not row:
            raise RuntimeError(
                "SECRET_KEY 가 초기화되지 않았습니다. init_super_db() 를 먼저 호출하세요."
            )
        return row["value"]
    finally:
        conn.close()


# ----------------------------------------------------------------------
# 디버그용 — 단독 실행하면 슈퍼 DB 초기화
# ----------------------------------------------------------------------
if __name__ == "__main__":
    init_super_db()
    print(f"Super DB ready: {SUPER_DB_PATH}")