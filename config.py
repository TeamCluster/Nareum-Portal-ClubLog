"""애플리케이션 설정.

데이터 구조:
  db/super.sqlite3       # 슈퍼 관리자 — places + super_admin + app_settings
  db/<slug>.sqlite3      # 기관마다 별도 파일 (예: nareum.sqlite3)

비밀번호 / 비밀키:
  - 슈퍼 관리자 비밀번호 → super.sqlite3 의 super_admin 테이블 (해시, 변경 가능)
  - 기관별 비밀번호      → super.sqlite3 의 places 테이블 (해시, 변경 가능)
  - Flask SECRET_KEY     → super.sqlite3 의 app_settings 테이블 (첫 실행 시 자동 생성)

설정 파일(.env)이 필요 없습니다. 첫 실행 시 콘솔에 슈퍼 임시 비밀번호가
1회 출력되니, 그것으로 첫 로그인 후 즉시 변경하세요.
"""
import os
import re

# backend/ 디렉터리의 절대 경로
BASE_DIR = os.path.abspath(os.path.dirname(__file__))

# --- SQLite 폴더 / 파일 경로 ---------------------------------------------
DB_FOLDER = os.path.join(BASE_DIR, "db")
SUPER_DB_PATH = os.path.join(DB_FOLDER, "super.sqlite3")


def place_db_path(slug: str) -> str:
    """기관 slug 로부터 해당 기관의 DB 파일 경로 반환."""
    return os.path.join(DB_FOLDER, f"{slug}.sqlite3")


# --- 슬러그 검증 ---------------------------------------------------------
# URL 경로 + DB 파일명에 함께 쓰이므로 안전한 문자만 허용.
SLUG_REGEX = re.compile(r"^[a-z][a-z0-9_-]{1,29}$")

# URL 경로 / 시스템 용어와 충돌할 가능성이 있는 식별자는 차단.
RESERVED_SLUGS = frozenset({"super", "api", "admin", "static", "public"})


def is_valid_slug(slug: str) -> bool:
    """슬러그가 형식 + 예약어 규칙을 모두 통과하는지 검증."""
    if not isinstance(slug, str):
        return False
    if slug in RESERVED_SLUGS:
        return False
    return bool(SLUG_REGEX.match(slug))


# --- 프론트엔드 (CORS 허용 도메인) ---------------------------------------
# 운영에서 도메인이 다르면 환경변수 FRONTEND_ORIGIN 으로 덮어쓰세요.
FRONTEND_ORIGIN = os.environ.get("FRONTEND_ORIGIN", "http://localhost:5173")


# --- 재단 실적관리 자동 동기화 (Google Apps Script) ----------------------
# 동아리 활동일지 저장 시, 아래에 등록된 기관(slug)에 한해 재단 구글시트
# (Apps Script 웹앱)로 프로그램 참가실적을 자동 전송합니다.  자세한 동작/연령
# 매핑은 services/foundation_sync.py 참고.
#
#   url       : 재단 실적관리 HTML 의 const API 값 (Apps Script /exec URL)
#   prog_name : 재단 시트의 '프로그램 선택' 에 등록된 상위(그룹형) 프로그램명
#
# 이 URL 은 곧 인증키나 다름없으니(인증 없는 엔드포인트) 운영에서는
# 환경변수 FOUNDATION_SYNC_URL 로 주입하세요. 전송을 끄려면 해당 slug 줄을
# 주석 처리하면 됩니다. (slug 미등록 = 전송 안 함)
FOUNDATION_SYNC_URL = os.environ.get(
    "FOUNDATION_SYNC_URL",
    # 나름(nareum) 실적관리 시트 — 실적관리31_나름.html 의 const API 값 (새 배포)
    "https://script.google.com/macros/s/AKfycbyjSdfLc5mAaZB44t2oyg3EgXAMktRarIchNGZBwRKalYxGDUE1H8H_uStxbUuRMcg/exec",
)

FOUNDATION_SYNC = {
    "nareum": {"url": FOUNDATION_SYNC_URL, "prog_name": "청소년동아리연합회"},
    # "didim":  {"url": FOUNDATION_SYNC_URL, "prog_name": "청소년동아리연합회"},
}