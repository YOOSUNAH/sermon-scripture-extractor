# 설교문 말씀 추출기 (Sermon Scripture Extractor)

DOCX 형식의 설교문에서 **성경 본문**과 **보조 본문**을 자동으로 감지하여 색상 표시하고, 참조 구절을 정리해 주는 웹 툴입니다.

> ⚠️ **저작권 안내**: 이 도구는 설교문의 내용을 보여주거나 저장하지 않습니다.  
> 설교문 내의 **성경 말씀 구절 참조(예: 28:1, 마 27:51)** 만 추출합니다.  
> 예시 설교문 파일은 저작권 보호를 위해 저장소에 포함하지 않습니다.

---

## 기능

| 단락 패턴 | 처리 방식 | 분류 |
|---|---|---|
| `28:1`, `28:5-7` 등 — **숫자:숫자** 로 시작 | 🔴 빨간색 | 본문 |
| `마 27:51`, `고전 15:4` 등 — **책이름 숫자:숫자** 로 시작 | 🟡 노란 하이라이트 + 밑줄 | 보조본문 |

### 출력 파일 (ZIP)
- `파일명_output.docx` — 서식 적용된 Word 파일
- `파일명_output.pdf` — LibreOffice 기반 PDF 변환본
- `파일명_refs.md` — 본문 / 보조본문 구절 목록 (복사용)

**refs.md 예시 출력:**
```
본문
28:1
28:2-4
28:5-7
28:8-10
28:11-15
28:16-17
28:18-20

보조본문
마 27:51-52
마 16:21
마 17:22-23
마 20:18-19
마 26:31-32
고전 15:4-8
요 20:17
요 20:19-20
요 12:24
요 15:5
요 15:16
```

---

## 설치 및 실행

### 필수 환경
- Python 3.10+
- LibreOffice (PDF 변환용)

```bash
# macOS
brew install --cask libreoffice

# Python 패키지
pip3 install flask python-docx
```

### 서버 실행

```bash
cd sermon_processor

# 기본 포트 5000
python3 app.py

# 포트 지정
python3 app.py 8000
```

접속: `http://localhost:5000/` 또는 `http://[내_IP]:5000/`

### Mac 부팅 시 자동 시작 (선택)

```bash
bash setup_autostart.sh
```

---

## 파일 구조

```
sermon_processor/
├── app.py                        # Flask 서버 + DOCX 처리 로직
├── templates/
│   └── index.html                # 웹 UI
├── com.sermon.processor.plist    # macOS LaunchDaemon 설정
├── setup_autostart.sh            # 자동시작 등록 스크립트
└── start.sh                      # 수동 실행 스크립트
```

---

## 사용 방법

1. 브라우저에서 서버 주소 접속
2. DOCX 설교문 파일을 드래그하거나 클릭하여 업로드
3. **처리 시작** 클릭
4. ZIP 파일 다운로드 → 내부에 DOCX / PDF / MD 파일 포함
