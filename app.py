import re
import os
import io
import zipfile
import tempfile
import subprocess
from flask import Flask, request, send_file, render_template, jsonify
from docx import Document
from docx.shared import RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

# ── 패턴 ──────────────────────────────────────────────────────────────
RED_PATTERN = re.compile(r'^\s*\d+:\d+')           # 28:1  28:2  ...
YELLOW_PATTERN = re.compile(r'^\s*[가-힣A-Za-z]+\s+\d+:\d+(?![\d-])')  # 마 27:51  고전 15:4  ... (범위형 제목 제외)

LIBREOFFICE = '/Applications/LibreOffice.app/Contents/MacOS/soffice'


def extract_reference(text: str, is_yellow: bool) -> str | None:
    """단락에서 성경 참조를 추출하여 범위(예: 28:2-4, 마 27:51-52)로 반환."""
    refs = re.findall(r'(\d+):(\d+)', text)
    if not refs:
        return None

    first_ch, first_v = refs[0][0], int(refs[0][1])

    # 첫 번째 챕터 기준으로 같은 챕터의 모든 절 수집
    same_ch_verses = sorted(int(v) for c, v in refs if c == first_ch)

    if len(same_ch_verses) == 1:
        ref_part = f"{first_ch}:{same_ch_verses[0]}"
    else:
        ref_part = f"{first_ch}:{same_ch_verses[0]}-{same_ch_verses[-1]}"

    if is_yellow:
        m = re.match(r'^\s*([가-힣A-Za-z]+)\s+', text)
        book = m.group(1) if m else ''
        return f"{book} {ref_part}".strip()
    return ref_part


def _get_or_add_rPr(run):
    rPr = run._r.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        run._r.insert(0, rPr)
    return rPr


def set_paragraph_red(para):
    """단락의 모든 텍스트를 빨간색으로."""
    for run in para.runs:
        run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)


def set_paragraph_yellow_underline(para):
    """단락의 모든 텍스트에 노란색 하이라이트 + 밑줄."""
    for run in para.runs:
        run.font.underline = True
        rPr = _get_or_add_rPr(run)
        # 기존 highlight 제거 후 재삽입
        for existing in rPr.findall(qn('w:highlight')):
            rPr.remove(existing)
        hl = OxmlElement('w:highlight')
        hl.set(qn('w:val'), 'yellow')
        rPr.append(hl)


def process_document(input_bytes: bytes):
    """
    DOCX 바이트를 처리하여 (Document, red_refs, yellow_refs) 반환.
    """
    doc = Document(io.BytesIO(input_bytes))
    red_refs = []
    yellow_refs = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        if RED_PATTERN.match(text):
            set_paragraph_red(para)
            ref = extract_reference(text, is_yellow=False)
            if ref:
                red_refs.append(ref)
        elif YELLOW_PATTERN.match(text):
            set_paragraph_yellow_underline(para)
            ref = extract_reference(text, is_yellow=True)
            if ref:
                yellow_refs.append(ref)

    return doc, red_refs, yellow_refs


def docx_to_pdf(docx_path: str, out_dir: str) -> str | None:
    """LibreOffice로 DOCX → PDF 변환. 실패 시 None."""
    if not os.path.exists(LIBREOFFICE):
        return None
    try:
        result = subprocess.run(
            [LIBREOFFICE, '--headless', '--convert-to', 'pdf',
             '--outdir', out_dir, docx_path],
            capture_output=True, timeout=90
        )
        if result.returncode == 0:
            pdf_name = os.path.splitext(os.path.basename(docx_path))[0] + '.pdf'
            pdf_path = os.path.join(out_dir, pdf_name)
            if os.path.exists(pdf_path):
                return pdf_path
    except Exception:
        pass
    return None


def build_text_content(red_refs: list, yellow_refs: list) -> str:
    lines = ['본문']
    lines.extend(red_refs)
    lines.append('')
    lines.append('보조본문')
    lines.extend(yellow_refs)
    return '\n'.join(lines)


# ── 라우트 ────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/process', methods=['POST'])
def process():
    if 'file' not in request.files:
        return jsonify({'error': '파일이 없습니다.'}), 400

    f = request.files['file']
    if not f.filename.lower().endswith('.docx'):
        return jsonify({'error': 'DOCX 파일만 업로드 가능합니다.'}), 400

    # 원본 파일명에서 stem 추출
    base_name = os.path.splitext(f.filename)[0]
    input_bytes = f.read()

    doc, red_refs, yellow_refs = process_document(input_bytes)

    with tempfile.TemporaryDirectory() as tmp:
        # 1) DOCX 저장
        out_docx_path = os.path.join(tmp, f'{base_name}_output.docx')
        doc.save(out_docx_path)

        # 2) PDF 변환
        pdf_path = docx_to_pdf(out_docx_path, tmp)

        # 3) 텍스트 파일
        text_content = build_text_content(red_refs, yellow_refs)
        text_path = os.path.join(tmp, f'{base_name}_refs.md')
        with open(text_path, 'w', encoding='utf-8') as tf:
            tf.write(text_content)

        # 4) ZIP으로 묶기
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.write(out_docx_path, f'{base_name}_output.docx')
            if pdf_path and os.path.exists(pdf_path):
                zf.write(pdf_path, f'{base_name}_output.pdf')
            else:
                # PDF 없으면 안내 메모 추가
                zf.writestr('PDF_변환실패_안내.txt',
                            'LibreOffice가 설치되지 않아 PDF를 생성할 수 없었습니다.\n'
                            'brew install --cask libreoffice 후 재시도하세요.')
            zf.write(text_path, f'{base_name}_refs.md')
        zip_buf.seek(0)

    return send_file(
        zip_buf,
        mimetype='application/zip',
        as_attachment=True,
        download_name=f'{base_name}_result.zip'
    )


if __name__ == '__main__':
    import sys
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 5000
    print(f'서버 시작 → http://0.0.0.0:{port}/')
    app.run(host='0.0.0.0', port=port, debug=False)
