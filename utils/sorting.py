"""동아리 명 정렬 헬퍼.

기존 app.py 의 is_korean / sort_key_korean_first 를 그대로 옮겨왔습니다.
한글로 시작하는 이름을 먼저, 그다음 영문/기타 순으로 정렬합니다.
"""
import re

_KOREAN_RE = re.compile(r"^[가-힣]$")


def is_korean(char: str) -> bool:
    """주어진 문자가 한글(가-힣)인지 확인합니다."""
    return _KOREAN_RE.match(char) is not None


def sort_key_korean_first(club_name):
    """정렬 키 생성: (그룹, 이름). 그룹 0=한글, 1=기타, 2=예외."""
    if not club_name or not isinstance(club_name, str):
        return (2, str(club_name))
    return (0, club_name) if is_korean(club_name[0]) else (1, club_name)