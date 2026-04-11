#!/bin/bash
# 부팅 시 자동 시작 설정 (관리자 비밀번호 필요)
# 한 번만 실행하면 됩니다.

echo "=== 설교문 처리기 자동시작 설정 ==="

PLIST_SRC="$(dirname "$0")/com.sermon.processor.plist"
PLIST_DST="/Library/LaunchDaemons/com.sermon.processor.plist"

sudo cp "$PLIST_SRC" "$PLIST_DST"
sudo chown root:wheel "$PLIST_DST"
sudo chmod 644 "$PLIST_DST"
sudo launchctl load -w "$PLIST_DST"

echo ""
echo "✅ 완료! 서버가 시작되었고, Mac 재부팅 시에도 자동으로 켜집니다."
echo "   접속 주소: http://192.168.35.80:8000/"
echo ""
echo "서버 중지: sudo launchctl unload $PLIST_DST"
echo "서버 재시작: sudo launchctl kickstart -k system/com.sermon.processor"
