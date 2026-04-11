import re, os, io, uuid, time, subprocess, shutil, tempfile
from pathlib import Path
from flask import Flask, request, send_file, render_template, jsonify
from docx import Document
from docx.shared import RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from ppt_generator import (generate_ppt, parse_title_paragraph,
                            extract_verses, split_to_lines, split_to_slides)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

LIBREOFFICE = '/Applications/LibreOffice.app/Contents/MacOS/soffice'
PPT_TEMPLATE = os.path.join(os.path.dirname(__file__), 'ppt_template.pptx')

# 세션 임시 저장소
SESSIONS: dict[str, dict] = {}
SESSION_TTL = 3600  # 1시간


# ── 패턴 ──────────────────────────────────────────────────────────────
RED_PATTERN    = re.compile(r'^\s*\d+:\d+')
YELLOW_PATTERN = re.compile(r'^\s*[가-힣A-Za-z]+\s+\d+:\d+(?![\d-])')


# ── DOCX 처리 ─────────────────────────────────────────────────────────

def _get_or_add_rPr(run):
    rPr = run._r.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        run._r.insert(0, rPr)
    return rPr


def set_paragraph_red(para):
    for run in para.runs:
        run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)


def set_paragraph_yellow_underline(para):
    for run in para.runs:
        run.font.underline = True
        rPr = _get_or_add_rPr(run)
        for ex in rPr.findall(qn('w:highlight')):
            rPr.remove(ex)
        hl = OxmlElement('w:highlight')
        hl.set(qn('w:val'), 'yellow')
        rPr.append(hl)


def extract_reference(text: str, is_yellow: bool) -> str | None:
    refs = re.findall(r'(\d+):(\d+)', text)
    if not refs:
        return None
    first_ch, _ = refs[0]
    same_ch_verses = sorted(int(v) for c, v in refs if c == first_ch)
    if len(same_ch_verses) == 1:
        ref = f'{first_ch}:{same_ch_verses[0]}'
    else:
        ref = f'{first_ch}:{same_ch_verses[0]}-{same_ch_verses[-1]}'
    if is_yellow:
        m = re.match(r'^\s*([가-힣A-Za-z]+)\s+', text)
        book = m.group(1) if m else ''
        return f'{book} {ref}'.strip()
    return ref


def process_document(input_bytes: bytes):
    doc = Document(io.BytesIO(input_bytes))
    red_refs, yellow_refs, yellow_para_texts = [], [], []
    title_text = ''

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        if i == 0 and not title_text:
            title_text = text
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
            yellow_para_texts.append(text)

    return doc, red_refs, yellow_refs, yellow_para_texts, title_text


def build_refs_text(title_text: str, red_refs: list, yellow_refs: list) -> str:
    passage, sermon_title = parse_title_paragraph(title_text)
    heading = f'{passage} "{sermon_title}"' if sermon_title else title_text

    lines = [heading, '']
    lines += ['본문'] + red_refs + ['']
    lines += ['보조본문'] + yellow_refs
    return '\n'.join(lines)


# ── PDF 변환 ──────────────────────────────────────────────────────────

def docx_to_pdf(docx_path: str, out_dir: str) -> str | None:
    # 1) docx2pdf (Word 사용, macOS에서 품질 우수)
    try:
        from docx2pdf import convert
        pdf_path = os.path.join(out_dir, Path(docx_path).stem + '.pdf')
        convert(docx_path, pdf_path)
        if os.path.exists(pdf_path):
            return pdf_path
    except Exception:
        pass

    # 2) LibreOffice
    if os.path.exists(LIBREOFFICE):
        try:
            subprocess.run(
                [LIBREOFFICE, '--headless', '--convert-to', 'pdf',
                 '--outdir', out_dir, docx_path],
                capture_output=True, timeout=90
            )
            pdf_path = os.path.join(out_dir, Path(docx_path).stem + '.pdf')
            if os.path.exists(pdf_path):
                return pdf_path
        except Exception:
            pass
    return None


# ── 세션 정리 ──────────────────────────────────────────────────────────

def _cleanup_old_sessions():
    now = time.time()
    for sid in list(SESSIONS.keys()):
        if now - SESSIONS[sid]['created'] > SESSION_TTL:
            d = SESSIONS[sid].get('dir')
            if d and os.path.isdir(d):
                shutil.rmtree(d, ignore_errors=True)
            del SESSIONS[sid]


# ── 라우트 ────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/process', methods=['POST'])
def process():
    _cleanup_old_sessions()

    if 'file' not in request.files:
        return jsonify({'error': '파일이 없습니다.'}), 400
    f = request.files['file']
    if not f.filename.lower().endswith('.docx'):
        return jsonify({'error': 'DOCX 파일만 업로드 가능합니다.'}), 400

    base_name = Path(f.filename).stem   # 예: 2801_설교문_input
    input_bytes = f.read()

    doc, red_refs, yellow_refs, yellow_para_texts, title_text = process_document(input_bytes)

    tmp_dir = tempfile.mkdtemp()
    session_id = str(uuid.uuid4())

    # 1) DOCX 저장
    docx_out_name = f'{base_name}_output.docx'
    docx_path = os.path.join(tmp_dir, docx_out_name)
    doc.save(docx_path)

    # 2) PDF (입력 파일명 유지: base_name.pdf)
    pdf_name = f'{base_name}.pdf'
    pdf_path = docx_to_pdf(docx_path, tmp_dir)
    if pdf_path:
        final_pdf = os.path.join(tmp_dir, pdf_name)
        if pdf_path != final_pdf:
            os.rename(pdf_path, final_pdf)
        pdf_path = final_pdf

    # 3) PPT
    passage, sermon_title = parse_title_paragraph(title_text)
    ppt_name = f'{base_name}_output.pptx'
    ppt_path = None
    try:
        ppt_bytes = generate_ppt(PPT_TEMPLATE, passage, sermon_title, yellow_para_texts)
        ppt_path = os.path.join(tmp_dir, ppt_name)
        with open(ppt_path, 'wb') as pf:
            pf.write(ppt_bytes)
    except Exception as e:
        print(f'PPT 생성 오류: {e}')

    # 4) 텍스트
    refs_text = build_refs_text(title_text, red_refs, yellow_refs)

    # 세션 저장
    SESSIONS[session_id] = {
        'created': time.time(),
        'dir': tmp_dir,
        'docx': docx_path if os.path.exists(docx_path) else None,
        'pdf':  pdf_path  if pdf_path and os.path.exists(pdf_path) else None,
        'ppt':  ppt_path  if ppt_path and os.path.exists(ppt_path) else None,
        'docx_name': docx_out_name,
        'pdf_name':  pdf_name,
        'ppt_name':  ppt_name,
    }

    return jsonify({
        'session_id': session_id,
        'refs_text':  refs_text,
        'has_pdf':    bool(SESSIONS[session_id]['pdf']),
        'has_ppt':    bool(SESSIONS[session_id]['ppt']),
        'docx_name':  docx_out_name,
        'pdf_name':   pdf_name,
        'ppt_name':   ppt_name,
    })


@app.route('/download/<session_id>/<file_type>')
def download(session_id, file_type):
    sess = SESSIONS.get(session_id)
    if not sess:
        return '세션이 만료됐습니다.', 404

    file_map = {
        'docx': (sess['docx'], sess['docx_name'], 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
        'pdf':  (sess['pdf'],  sess['pdf_name'],  'application/pdf'),
        'ppt':  (sess['ppt'],  sess['ppt_name'],  'application/vnd.openxmlformats-officedocument.presentationml.presentation'),
    }
    if file_type not in file_map:
        return '잘못된 요청', 400

    path, name, mime = file_map[file_type]
    if not path or not os.path.exists(path):
        return '파일 없음', 404

    return send_file(path, mimetype=mime, as_attachment=True, download_name=name)


if __name__ == '__main__':
    import sys
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8000
    print(f'서버 시작 → http://0.0.0.0:{port}/')
    app.run(host='0.0.0.0', port=port, debug=False)
