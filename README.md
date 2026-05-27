# 나름 동아리 활동일지 — Backend

Flask JSON API + SQLite. React(Vite) 프론트엔드가 호출하는 API 서버입니다.

## 디렉터리 구조

```
backend/
├── app.py                  Flask 라우트 (JSON API)
├── config.py               경로/비밀번호/CORS 설정
├── db.py                   SQLite 커넥션 + 스키마 정의
├── requirements.txt
├── db/                     SQLite 파일 (자동 생성)
│   └── clublog.sqlite3
├── services/               비즈니스 로직
│   ├── club_service.py     동아리 CRUD
│   └── log_service.py      활동일지 CRUD + 엑셀 export
└── utils/
    └── sorting.py          한글 우선 정렬
```

## 실행 방법

```bash
cd backend
pip install -r requirements.txt
python app.py
```

- 첫 실행 시 `db/clublog.sqlite3` 가 자동 생성되고 빈 테이블이 만들어집니다.
- 서버는 `http://localhost:5000` 에서 동작합니다.
- React 개발 서버(`http://localhost:5173`)가 같이 떠 있어야 CORS 가 통과됩니다.

## 첫 사용 흐름 (빈 DB 부트스트랩)

1. 관리자 로그인 — `POST /api/admin/login` body `{"password": "skfma2013"}`
2. 동아리 추가 — `POST /api/admin/clubs` body `{"name": "광대", "category": "밴드"}` (여러 번 호출)
3. 그 다음부터 일반 사용자가 `POST /api/logs` 로 활동일지 작성

운영 환경에서는 환경변수로 비밀번호를 덮어쓰세요:

```bash
export SECRET_KEY="<랜덤한 긴 문자열>"
export ADMIN_PASSWORD="<관리자 비밀번호>"
```

## API 목록

| 메서드 | 경로 | 인증 | 설명 |
|---|---|---|---|
| GET | `/api/clubs` | — | 동아리 목록 + 분야 매핑 dict |
| POST | `/api/logs` | — | 활동일지 작성 |
| GET | `/api/logs/download` | — | 활동일지 엑셀 다운로드 |
| POST | `/api/admin/login` | — | 로그인 |
| POST | `/api/admin/logout` | — | 로그아웃 |
| GET | `/api/admin/session` | — | 로그인 상태 확인 |
| GET | `/api/admin/dashboard` | 🔒 | 대시보드 통계 |
| GET | `/api/admin/logs` | 🔒 | 전체 활동일지 (최신순) |
| DELETE | `/api/admin/logs` | 🔒 | 활동일지 1건 삭제 (body: `{"id": N}`) |
| GET | `/api/admin/clubs` | 🔒 | 동아리 목록 (관리자용) |
| POST | `/api/admin/clubs` | 🔒 | 동아리 추가 (body: `{"name", "category"}`) |
| DELETE | `/api/admin/clubs` | 🔒 | 동아리 삭제 (body: `{"name"}`) |

🔒 = 관리자 세션 쿠키 필요. 프론트는 `fetch(..., { credentials: 'include' })` 로 호출.

## DB 스키마

```sql
CREATE TABLE clubs (
    id        INTEGER PRIMARY KEY AUTOINCREMENT,
    name      TEXT NOT NULL UNIQUE,
    category  TEXT NOT NULL
);

CREATE TABLE logs (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    category        TEXT,
    club_name       TEXT NOT NULL,
    activity_date   TEXT NOT NULL,    -- 'YYYY-MM-DD'
    start_time      TEXT NOT NULL,    -- 'HH:MM'
    end_time        TEXT NOT NULL,
    participants    TEXT DEFAULT '',
    content         TEXT DEFAULT '',
    author          TEXT DEFAULT '',
    total_count     INTEGER NOT NULL,
    elem_m, elem_f, mid_m, mid_f,
    high_m, high_f, univ_m, univ_f   INTEGER NOT NULL DEFAULT 0,
    created_at      TEXT NOT NULL     -- 'YYYY-MM-DD HH:MM:SS'
);
```

API 응답은 영문 컬럼명 그대로. 엑셀 export 시점에만 한글 헤더로 변환됩니다.