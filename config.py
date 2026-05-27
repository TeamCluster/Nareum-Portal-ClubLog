"""애플리케이션 설정.

데이터 저장은 SQLite (db/clublog.sqlite3) 를 사용합니다.
운영 환경에서는 환경변수로 값을 덮어쓸 수 있습니다.
"""
import os

# backend/ 디렉터리의 절대 경로
BASE_DIR = os.path.abspath(os.path.dirname(__file__))

# --- SQLite (메인 데이터 저장소) -----------------------------------------
DB_FOLDER = os.path.join(BASE_DIR, "db")
DB_PATH = os.environ.get("DB_PATH", os.path.join(DB_FOLDER, "clublog.sqlite3"))

# --- 보안 ----------------------------------------------------------------
# 세션 서명용 비밀키 / 관리자 비밀번호
# 운영 시 반드시 환경변수(SECRET_KEY, ADMIN_PASSWORD)로 주입하세요.
SECRET_KEY = os.environ.get("SECRET_KEY", "Skfma20601318")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "skfma2013")

# --- 프론트엔드 -----------------------------------------------------------
# React(Vite) 개발 서버 주소 — CORS 허용 대상
FRONTEND_ORIGIN = os.environ.get("FRONTEND_ORIGIN", "http://localhost:5173")