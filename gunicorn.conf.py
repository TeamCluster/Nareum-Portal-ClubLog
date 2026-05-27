"""Gunicorn 설정 — Synology DS220+ 운영 환경 기준.

사용법:
    gunicorn -c gunicorn.conf.py app:app

DS220+ 사양 (참고):
  - CPU: Intel Celeron J4025 (2 cores / 2 threads, 2.0~2.9GHz)
  - RAM: 2GB (기본) ~ 6GB
RAM 이 2GB 라 보수적으로 워커 2개 / 워커당 스레드 4개로 설정.
RAM 을 늘렸다면 GUNICORN_WORKERS 환경변수로 늘릴 수 있습니다.
"""
import os

# --- 바인딩 -------------------------------------------------------------
# 0.0.0.0 -> NAS 외부에서도 접속 가능. 포트는 app.py 와 동일하게 5015.
bind = os.environ.get("GUNICORN_BIND", "0.0.0.0:5015")

# --- 워커 / 스레드 -------------------------------------------------------
# sync 대신 gthread — SQLite/파일 IO 위주라 스레드가 효율적.
# 동시 처리 가능량 = workers * threads = 2 * 4 = 8.
workers = int(os.environ.get("GUNICORN_WORKERS", "2"))
worker_class = "gthread"
threads = int(os.environ.get("GUNICORN_THREADS", "4"))

# --- 타임아웃 ------------------------------------------------------------
# 엑셀 export 가 데이터 많을 때 시간 걸릴 수 있어 넉넉히.
timeout = 60
graceful_timeout = 30
keepalive = 5

# --- 워커 리사이클링 -----------------------------------------------------
# 1000 요청마다 워커를 재시작 -> 메모리 누수 방지 (pandas 가 leaky 한 편).
# jitter 로 모든 워커가 동시에 재시작되지 않도록 분산.
max_requests = 1000
max_requests_jitter = 100

# --- 로깅 ----------------------------------------------------------------
# stdout/stderr 로 출력 -> Docker 또는 DSM Task Scheduler 가 캡쳐.
accesslog = "-"
errorlog = "-"
loglevel = os.environ.get("GUNICORN_LOG_LEVEL", "info")

# --- 기타 ----------------------------------------------------------------
proc_name = "clublog-api"   # ps / top 에서 식별 용이
preload_app = True          # 워커 fork 전 앱 로드 -> 메모리 공유로 RAM 절약