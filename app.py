"""나름 동아리 활동일지 — Flask JSON API (멀티테넌트).

URL 구조:
  /api/super/...        슈퍼 관리자 (전 기관 총괄)
  /api/<slug>/...       기관별 (공개 / 관리자)

세션 구조:
  super_logged_in: bool
  place_admins: dict[str, bool]    # 기관별 독립 로그인 상태
                                     예: {"nareum": True, "place01": False}

DB 초기화:
  db.init_super_db() 가 첫 실행 시 슈퍼 비밀번호와 SECRET_KEY 를 생성.
  기관 DB(<slug>.sqlite3) 는 place_service.add_place 가 만든다.
"""
import threading
from functools import wraps

from flask import Flask, jsonify, request, send_file, session
from flask_cors import CORS

import config
import db
from config import is_valid_slug
from services import (
    club_service,
    foundation_sync,
    log_service,
    place_service,
    super_service,
)

app = Flask(__name__)

# SQLite 초기화 (첫 실행 시 슈퍼 임시 비밀번호 콘솔에 1회 출력)
db.init_super_db()
app.secret_key = db.get_secret_key()
app.teardown_appcontext(db.close_dbs)

# 세션 쿠키 / CORS
app.config.update(
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_HTTPONLY=True,
)
CORS(app, origins=[config.FRONTEND_ORIGIN], supports_credentials=True)


# ======================================================================
#  데코레이터
# ======================================================================
def super_required(view):
    """슈퍼 관리자 로그인 필요."""
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("super_logged_in"):
            return jsonify({"error": "슈퍼 관리자 로그인이 필요합니다."}), 401
        return view(*args, **kwargs)
    return wrapped


def place_required(view):
    """URL 의 <slug> 가 유효하고 places 테이블에 존재해야 함.
    그렇지 않으면 404. 로그인 여부는 검증하지 않음 (공개 라우트용)."""
    @wraps(view)
    def wrapped(slug, *args, **kwargs):
        if not is_valid_slug(slug):
            return jsonify({"error": "잘못된 기관 식별자입니다."}), 404
        if not place_service.get_place(slug):
            return jsonify({"error": "기관을 찾을 수 없습니다."}), 404
        return view(slug, *args, **kwargs)
    return wrapped


def place_admin_required(view):
    """기관 검증 + 해당 기관 관리자 로그인 필요.
    place_required + 로그인 확인을 한 곳에서 (데코레이터 합성 대신 명시적으로)."""
    @wraps(view)
    def wrapped(slug, *args, **kwargs):
        if not is_valid_slug(slug):
            return jsonify({"error": "잘못된 기관 식별자입니다."}), 404
        if not place_service.get_place(slug):
            return jsonify({"error": "기관을 찾을 수 없습니다."}), 404
        if not session.get("place_admins", {}).get(slug):
            return jsonify({"error": "기관 관리자 로그인이 필요합니다."}), 401
        return view(slug, *args, **kwargs)
    return wrapped


# ======================================================================
#  슈퍼 관리자 — 인증
# ======================================================================
@app.post("/api/super/login")
def api_super_login():
    data = request.get_json(silent=True) or {}
    if super_service.verify_super_password(data.get("password", "")):
        session["super_logged_in"] = True
        return jsonify({"ok": True})
    return jsonify({"ok": False, "message": "비밀번호가 틀렸습니다."}), 401


@app.post("/api/super/logout")
def api_super_logout():
    session.pop("super_logged_in", None)
    return jsonify({"ok": True})


@app.get("/api/super/session")
def api_super_session():
    return jsonify({"logged_in": bool(session.get("super_logged_in"))})


@app.post("/api/super/password")
@super_required
def api_super_password():
    data = request.get_json(silent=True) or {}
    ok, msg = super_service.update_super_password(data.get("new_password", ""))
    return jsonify({"ok": ok, "message": msg}), (200 if ok else 400)


# ======================================================================
#  슈퍼 관리자 — 기관 CRUD
# ======================================================================
@app.get("/api/super/places")
@super_required
def api_super_places_list():
    return jsonify({"places": place_service.get_places()})


@app.post("/api/super/places")
@super_required
def api_super_places_add():
    data = request.get_json(silent=True) or {}
    ok, msg, result = place_service.add_place(
        data.get("slug", ""),
        data.get("full_name", ""),
        data.get("short_name", ""),
        data.get("password", ""),
        data.get("foundation_sync_url", ""),
    )
    return jsonify({"ok": ok, "message": msg, "result": result}), (200 if ok else 400)


@app.delete("/api/super/places/<slug>")
@super_required
def api_super_places_delete(slug):
    ok, msg = place_service.delete_place(slug)
    return jsonify({"ok": ok, "message": msg}), (200 if ok else 404)


@app.post("/api/super/places/<slug>/password")
@super_required
def api_super_place_password(slug):
    data = request.get_json(silent=True) or {}
    ok, msg = place_service.update_place_password(slug, data.get("new_password", ""))
    return jsonify({"ok": ok, "message": msg}), (200 if ok else 400)


@app.post("/api/super/places/<slug>/sync-url")
@super_required
def api_super_place_sync_url(slug):
    """기관의 재단 동기화 URL 등록/수정/해제. 빈 문자열 = 동기화 해제."""
    data = request.get_json(silent=True) or {}
    ok, msg = place_service.update_place_sync_url(
        slug, data.get("foundation_sync_url", "")
    )
    return jsonify({"ok": ok, "message": msg}), (200 if ok else 400)


# ======================================================================
#  기관 — 인증
# ======================================================================
@app.post("/api/<slug>/admin/login")
@place_required
def api_place_login(slug):
    data = request.get_json(silent=True) or {}
    if place_service.verify_place_password(slug, data.get("password", "")):
        # 세션의 place_admins dict 에 해당 slug 만 추가
        admins = session.get("place_admins", {})
        admins[slug] = True
        session["place_admins"] = admins
        session.modified = True  # dict 내부 변경은 Flask 가 자동 감지 못함
        return jsonify({"ok": True})
    return jsonify({"ok": False, "message": "비밀번호가 틀렸습니다."}), 401


@app.post("/api/<slug>/admin/logout")
@place_required
def api_place_logout(slug):
    admins = session.get("place_admins", {})
    admins.pop(slug, None)
    session["place_admins"] = admins
    session.modified = True
    return jsonify({"ok": True})


@app.get("/api/<slug>/admin/session")
@place_required
def api_place_session(slug):
    logged_in = bool(session.get("place_admins", {}).get(slug))
    return jsonify({"logged_in": logged_in})


# ======================================================================
#  기관 — 공개 (일지 작성 페이지)
# ======================================================================
@app.get("/api/<slug>/info")
@place_required
def api_place_info(slug):
    """기관 공개 정보 — 로그인 불필요. 일지 작성 페이지 / 헤더에서 사용."""
    p = place_service.get_place(slug)
    return jsonify({
        "slug": p["slug"],
        "full_name": p["full_name"],
        "short_name": p["short_name"],
    })


@app.get("/api/<slug>/clubs")
@place_required
def api_place_clubs(slug):
    clubs = club_service.get_clubs(slug)
    return jsonify({
        "clubs": clubs,
        "club_dict": {c["name"]: c["category"] for c in clubs},
    })


@app.post("/api/<slug>/logs")
@place_required
def api_place_logs_create(slug):
    data = request.get_json(silent=True) or {}
    ok, msg, result = log_service.create_log(slug, data)
    if ok:
        # 재단 구글시트로 비동기(베스트에포트) 전송 — 활동일지 저장 응답을 막지 않음.
        _spawn_foundation_sync(slug, data, result)
    return jsonify({"ok": ok, "message": msg, "result": result}), (200 if ok else 400)


def _spawn_foundation_sync(slug, data, result):
    """재단 동기화를 데몬 스레드로 던진다(결과는 콘솔 로깅).

    동기화 URL 은 기관 DB(places.foundation_sync_url)에서 읽는다. DB 접근은
    Flask 요청 컨텍스트가 필요하므로 스레드를 띄우기 전에 여기서 읽어서
    값(문자열)만 넘긴다. prog_name 은 config 의 기관별 설정(없으면 기본값).
    """
    place = place_service.get_place(slug) or {}
    url = (place.get("foundation_sync_url") or "").strip()
    prog_name = (
        getattr(config, "FOUNDATION_SYNC", {}).get(slug, {}).get("prog_name")
        or "청소년동아리연합회"
    )

    def _run():
        sync_ok, sync_msg = foundation_sync.sync_log(
            slug, data, result, url, prog_name
        )
        if sync_ok is None:
            return  # 미설정 기관 — 조용히 건너뜀
        tag = "OK" if sync_ok else "FAIL"
        app.logger.info("[재단동기화 %s] %s/%s: %s",
                        tag, slug, result.get("club", ""), sync_msg)

    threading.Thread(target=_run, daemon=True).start()


@app.get("/api/<slug>/logs/download")
@place_required
def api_place_logs_download(slug):
    try:
        output, filename = log_service.export_excel(slug)
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as err:  # noqa: BLE001
        return jsonify({"error": str(err)}), 500


# ======================================================================
#  기관 — 관리자 (보호)
# ======================================================================
@app.get("/api/<slug>/admin/dashboard")
@place_admin_required
def api_place_dashboard(slug):
    total = len(club_service.get_clubs(slug))
    return jsonify(log_service.get_dashboard_stats(slug, total))


@app.get("/api/<slug>/admin/logs")
@place_admin_required
def api_place_admin_logs(slug):
    return jsonify({"logs": log_service.get_logs(slug)})


@app.delete("/api/<slug>/admin/logs")
@place_admin_required
def api_place_admin_logs_delete(slug):
    data = request.get_json(silent=True) or {}
    ok, msg = log_service.delete_log(slug, data.get("id", 0))
    return jsonify({"ok": ok, "message": msg}), (200 if ok else 400)


@app.get("/api/<slug>/admin/clubs")
@place_admin_required
def api_place_admin_clubs(slug):
    return jsonify({"clubs": club_service.get_clubs(slug)})


@app.post("/api/<slug>/admin/clubs")
@place_admin_required
def api_place_admin_clubs_add(slug):
    data = request.get_json(silent=True) or {}
    ok, msg = club_service.add_club(slug, data.get("name", ""), data.get("category", ""))
    return jsonify({"ok": ok, "message": msg}), (200 if ok else 400)


@app.delete("/api/<slug>/admin/clubs")
@place_admin_required
def api_place_admin_clubs_delete(slug):
    data = request.get_json(silent=True) or {}
    ok, msg = club_service.delete_club(slug, data.get("name", ""))
    return jsonify({"ok": ok, "message": msg}), (200 if ok else 400)


if __name__ == "__main__":
    app.run(debug=True, port=5015)