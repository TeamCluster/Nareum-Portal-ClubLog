"""나름 동아리 활동일지 — Flask JSON API.

기존 app.py 는 render_template 로 HTML 을 직접 그렸지만,
이 버전은 모든 응답을 JSON 으로 내려주는 순수 API 서버입니다.
화면(폼/모달/표)은 React(Vite) 프론트엔드가 담당합니다.

데이터 저장소: SQLite (db/clublog.sqlite3).
"""
from functools import wraps

from flask import Flask, jsonify, request, send_file, session
from flask_cors import CORS

import config
import db
from services import club_service, log_service

app = Flask(__name__)
app.secret_key = config.SECRET_KEY

# 세션 쿠키 설정 — React 개발 서버(다른 포트)에서 쿠키를 주고받기 위함.
# localhost 끼리는 same-site 라 Lax 로 충분합니다.
# 운영 환경에서 프론트/백 도메인이 다르면 SAMESITE="None" + SECURE=True 로 바꾸세요.
app.config.update(
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_HTTPONLY=True,
)

# 프론트엔드 주소만 CORS 허용 + 쿠키 동반 허용
CORS(app, origins=[config.FRONTEND_ORIGIN], supports_credentials=True)

# --- SQLite 초기화 & 요청 종료 훅 --------------------------------------
db.init_db()                              # 앱 시작 시 테이블 생성 (멱등)
app.teardown_appcontext(db.close_db)      # 요청 끝나면 커넥션 닫기


def login_required(view):
    """관리자 로그인 여부를 확인하는 데코레이터."""
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("logged_in"):
            return jsonify({"error": "로그인이 필요합니다."}), 401
        return view(*args, **kwargs)
    return wrapped


# ======================================================================
#  공개 API — 활동일지
# ======================================================================
@app.get("/api/clubs")
def api_clubs():
    """동아리 목록 반환. (분야 자동 채움용 dict 도 함께 제공)"""
    clubs = club_service.get_clubs()
    return jsonify({
        "clubs": clubs,
        "club_dict": {c["name"]: c["category"] for c in clubs},
    })


@app.post("/api/logs")
def api_create_log():
    """활동일지 작성."""
    data = request.get_json(silent=True) or {}
    ok, message, result = log_service.create_log(data)
    status = 200 if ok else 400
    return jsonify({"ok": ok, "message": message, "result": result}), status


@app.get("/api/logs/download")
def api_download_log():
    """활동일지 엑셀 다운로드."""
    try:
        output, filename = log_service.export_excel()
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as err:  # noqa: BLE001
        return jsonify({"error": str(err)}), 500


# ======================================================================
#  관리자 인증
# ======================================================================
@app.post("/api/admin/login")
def api_login():
    """관리자 로그인. 성공 시 세션에 표시."""
    data = request.get_json(silent=True) or {}
    if data.get("password") == config.ADMIN_PASSWORD:
        session["logged_in"] = True
        return jsonify({"ok": True})
    return jsonify({"ok": False, "message": "비밀번호가 틀렸습니다."}), 401


@app.post("/api/admin/logout")
def api_logout():
    """관리자 로그아웃."""
    session.pop("logged_in", None)
    return jsonify({"ok": True})


@app.get("/api/admin/session")
def api_session():
    """현재 로그인 상태 확인 (프론트 라우트 가드용)."""
    return jsonify({"logged_in": bool(session.get("logged_in"))})


# ======================================================================
#  관리자 API — 보호된 라우트
# ======================================================================
@app.get("/api/admin/dashboard")
@login_required
def api_dashboard():
    """관리자 대시보드 통계."""
    total = len(club_service.get_clubs())
    return jsonify(log_service.get_dashboard_stats(total))


@app.get("/api/admin/logs")
@login_required
def api_admin_logs():
    """전체 활동일지 (최신순)."""
    return jsonify({"logs": log_service.get_logs()})


@app.delete("/api/admin/logs")
@login_required
def api_delete_log():
    """활동일지 한 건 삭제 (id 기준)."""
    data = request.get_json(silent=True) or {}
    ok, message = log_service.delete_log(data.get("id", 0))
    return jsonify({"ok": ok, "message": message}), (200 if ok else 400)


@app.get("/api/admin/clubs")
@login_required
def api_admin_clubs():
    """동아리 목록 (관리자용)."""
    return jsonify({"clubs": club_service.get_clubs()})


@app.post("/api/admin/clubs")
@login_required
def api_add_club():
    """동아리 추가."""
    data = request.get_json(silent=True) or {}
    ok, message = club_service.add_club(
        data.get("name", ""), data.get("category", "")
    )
    return jsonify({"ok": ok, "message": message}), (200 if ok else 400)


@app.delete("/api/admin/clubs")
@login_required
def api_delete_club():
    """동아리 삭제."""
    data = request.get_json(silent=True) or {}
    ok, message = club_service.delete_club(data.get("name", ""))
    return jsonify({"ok": ok, "message": message}), (200 if ok else 400)


if __name__ == "__main__":
    app.run(debug=True, port=5000)