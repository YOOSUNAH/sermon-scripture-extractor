#!/usr/bin/env bash
# EC2 Ubuntu 최초 1회 실행용 셋업 스크립트
# 사용: bash deploy/setup.sh
set -euo pipefail

REPO_URL="https://github.com/YOOSUNAH/sermon-scripture-extractor.git"
APP_DIR="/opt/sermon"

echo "[1/6] 패키지 설치"
sudo apt update
sudo apt install -y \
    python3 python3-venv python3-pip \
    git nginx libreoffice \
    build-essential libxml2-dev libxslt1-dev python3-dev

echo "[2/6] 스왑 2GB 생성 (t3.micro 메모리 보강)"
if [ ! -f /swapfile ]; then
    sudo fallocate -l 2G /swapfile
    sudo chmod 600 /swapfile
    sudo mkswap /swapfile
    sudo swapon /swapfile
    echo '/swapfile none swap sw 0 0' | sudo tee -a /etc/fstab
fi

echo "[3/6] 앱 클론"
sudo mkdir -p "$APP_DIR"
sudo chown "$USER:$USER" "$APP_DIR"
if [ ! -d "$APP_DIR/.git" ]; then
    git clone "$REPO_URL" "$APP_DIR"
fi
cd "$APP_DIR"

echo "[4/6] venv + 의존성"
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt

echo "[5/6] systemd 등록"
sudo cp deploy/sermon.service /etc/systemd/system/sermon.service
sudo systemctl daemon-reload
sudo systemctl enable --now sermon

echo "[6/6] nginx 설정"
sudo cp deploy/nginx.conf /etc/nginx/sites-available/sermon
sudo ln -sf /etc/nginx/sites-available/sermon /etc/nginx/sites-enabled/sermon
sudo rm -f /etc/nginx/sites-enabled/default
sudo nginx -t
sudo systemctl reload nginx

echo
echo "✅ 셋업 완료"
echo "   - 상태 확인: sudo systemctl status sermon"
echo "   - 로그:      sudo journalctl -u sermon -f"
echo "   - 접속:      http://<EC2-퍼블릭-IP>/"
