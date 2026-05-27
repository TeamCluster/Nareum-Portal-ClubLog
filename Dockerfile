# Synology DS220+ (x86_64) 운영용 — Python 3.12 슬림 베이스
FROM python:3.12-slim

# 시간대 — 컨테이너 내부의 datetime 이 KST 로 동작 (created_at 등)
ENV TZ=Asia/Seoul
# 로그가 즉시 stdout 으로 (Docker logs 에서 바로 보이게)
ENV PYTHONUNBUFFERED=1
# .pyc 파일 안 만듦 (이미지 작게)
ENV PYTHONDONTWRITEBYTECODE=1

WORKDIR /app

# 의존성 먼저 설치 — 코드 변경 시 이 레이어가 캐시되도록 분리
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 앱 코드 전체 복사
COPY . .

# gunicorn 이 듣는 포트 (gunicorn.conf.py 의 bind 와 일치)
EXPOSE 5015

# 운영 — gunicorn (개발용 python app.py 가 아님)
# 설정은 gunicorn.conf.py 에서 (workers=2, threads=4, gthread, etc)
CMD ["gunicorn", "-c", "gunicorn.conf.py", "app:app"]