"""재단 실적관리(Google Apps Script) 자동 동기화.

동아리 활동일지가 저장되면, 해당 기관(slug)의 재단 동기화 URL
(슈퍼 페이지에서 등록, places.foundation_sync_url)이 설정돼 있을 때
재단 구글시트로 '프로그램 참가 실적'(saveProg 액션)을 전송한다.

설계 메모
---------
* 서버 → 서버 호출이라 브라우저 CORS 제약이 없다. (프론트에서 직접 쏘면 막힐 수 있음)
* 베스트에포트: 전송 실패가 활동일지 저장을 막지 않도록 모든 예외를 삼키고,
  결과는 (ok, message) 로만 돌려준다. 호출부는 이 결과를 로깅/표시만 한다.
* 외부 의존성을 늘리지 않으려고 표준 라이브러리 urllib 만 사용한다.

연령 구간 매핑(중요)
--------------------
우리 앱은 학교급(초/중/고/후기)으로, 재단은 나이 구간으로 집계한다.
1:1 로 떨어지지 않아 아래처럼 근사 매핑한다.

  초등(elem) → 청소년 9~13 (T1)     ※ 8세인 초등 저학년도 여기로 들어감(근사 한계)
  중등(mid)  → 청소년 14~16 (T2)
  고등(high) → 청소년 17~19 (T3)
  후기(univ) → 청소년 20~24 (T4)
  아동(8세↓, A) / 성인(25↑, Ad) → 우리 앱에 데이터 없음 = 0

재단 분야(category)는 재단 시트에 등록된 하위항목(subItem=동아리명) 기준으로
재단 측에서 자동으로 붙으므로 여기서 보내지 않는다. 따라서 전송 대상 동아리는
재단 그룹형 프로그램의 하위항목으로 미리 등록돼 있어야 분야가 정상 표기된다.
"""
import json
import urllib.request

# 재단 saveProg 가 기대하는 연령 필드 <- 우리 앱(요청 JSON) 필드
_AGE_MAP = {
    "mT1": "element_man",  "fT1": "element_woman",
    "mT2": "middle_man",   "fT2": "middle_woman",
    "mT3": "high_man",     "fT3": "high_woman",
    "mT4": "univ_man",     "fT4": "univ_woman",
}
# 우리 앱에 대응값이 없어 항상 0 으로 보내는 재단 필드
_ZERO_FIELDS = ("mA", "fA", "mAd", "fAd")


def _to_int(v):
    try:
        return int(v or 0)
    except (TypeError, ValueError):
        return 0


def build_payload(prog_name: str, data: dict, result: dict) -> dict:
    """우리 앱의 활동일지 입력값을 재단 saveProg payload 로 변환.

    ``data`` 는 프론트가 보낸 원본 요청 JSON(연령별 인원 포함),
    ``result`` 는 create_log 가 돌려준 정규화 결과(date 등)이다.
    """
    payload = {
        "date": result.get("date", ""),
        "progName": prog_name,
        "subItem": data.get("club_name", ""),   # 하위항목 = 동아리명
        "session": _to_int(data.get("session")) or 1,
        "memo": data.get("activity_content", ""),
    }
    for fac_field, app_field in _AGE_MAP.items():
        payload[fac_field] = _to_int(data.get(app_field))
    for f in _ZERO_FIELDS:
        payload[f] = 0
    return payload


def sync_log(slug: str, data: dict, result: dict, url: str, prog_name: str,
             timeout: int = 15):
    """활동일지 한 건을 재단 시트로 전송. (ok, message) 반환.

    url/prog_name 은 호출부(요청 컨텍스트)에서 기관 DB 를 읽어 넘겨준다.
    이 함수는 데몬 스레드에서 실행될 수 있으므로 Flask g / DB 에 직접
    접근하지 않는다. url 이 비어있으면 (None, '미설정') 으로 건너뛴다.
    예외는 모두 흡수하여 호출부(활동일지 저장)에 영향을 주지 않는다.
    """
    if not url:
        return None, "재단 동기화 미설정(URL 없음)"

    payload = build_payload(prog_name or "", data, result)
    body = json.dumps({"action": "saveProg", "data": payload}).encode("utf-8")

    # fetch() 기본동작과 동일하게 text/plain 으로 보낸다.
    # Apps Script 는 e.postData.contents 로 본문을 그대로 읽는다.
    req = urllib.request.Request(
        url,
        data=body,
        method="POST",
        headers={"Content-Type": "text/plain;charset=utf-8"},
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read().decode("utf-8", "replace")
        try:
            res = json.loads(raw)
        except json.JSONDecodeError:
            return False, f"재단 응답 파싱 실패: {raw[:200]}"
        if res.get("success"):
            return True, "재단 시트 전송 완료"
        return False, f"재단 저장 실패: {res.get('message', raw[:200])}"
    except Exception as err:  # noqa: BLE001
        return False, f"재단 전송 오류: {err}"
