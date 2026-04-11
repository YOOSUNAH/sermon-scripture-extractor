#!/bin/bash
# 설교문 처리기 서버 시작 스크립트
# 포트 80: sudo ./start.sh
# 포트 8080 (sudo 없이): ./start.sh 8080

cd "$(dirname "$0")"

PORT=${1:-80}

if [ "$PORT" -eq 80 ] && [ "$(id -u)" -ne 0 ]; then
  echo "포트 80은 sudo 권한이 필요합니다."
  echo "  sudo bash start.sh      # 포트 80"
  echo "  bash start.sh 8080      # 포트 8080"
  exit 1
fi

echo "서버 시작: http://0.0.0.0:$PORT/"
python3 app.py $PORT
