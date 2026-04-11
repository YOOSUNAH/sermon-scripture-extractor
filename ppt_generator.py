"""
PPT 생성기 — 보조본문 슬라이드 + 설교제목 슬라이드
"""
import re
import copy
import io
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

# ── 폰트/색상 상수 ────────────────────────────────────────────────────
FONT_NAME = '나눔고딕 ExtraBold'
BLUE  = RGBColor(0x00, 0xB0, 0xF0)
GREEN = RGBColor(0x00, 0xB0, 0x50)

# ── 줄 분리 설정 ──────────────────────────────────────────────────────
CHARS_PER_LINE = 22   # 한 줄 최대 글자 수 (한국어 기준)
LINES_PER_SLIDE = 2   # 슬라이드당 최대 줄 수

# ── 템플릿 슬라이드 인덱스 (ppt_template.pptx 기준) ──────────────────
IDX_VERSE_NORMAL = 0   # 보조본문 일반 (하단 참조 있음)
IDX_VERSE_SLASH  = 1   # 보조본문 마지막 (// 있음)
IDX_TITLE        = 2   # 설교제목


# ── 텍스트 분리 유틸 ──────────────────────────────────────────────────

def split_to_lines(text: str, max_chars: int = CHARS_PER_LINE) -> list[str]:
    """단어 경계를 유지하며 텍스트를 줄로 분리."""
    words = text.split()
    lines, cur, cur_len = [], [], 0
    for word in words:
        wl = len(word)
        test = cur_len + (1 if cur else 0) + wl
        if cur and test > max_chars:
            lines.append(' '.join(cur))
            cur, cur_len = [word], wl
        else:
            cur.append(word)
            cur_len = test
    if cur:
        lines.append(' '.join(cur))
    return lines or ['']


def split_to_slides(lines: list[str], max_lines: int = LINES_PER_SLIDE) -> list[list[str]]:
    """줄 목록을 슬라이드 단위로 그룹화."""
    if not lines:
        return [['']]
    return [lines[i:i+max_lines] for i in range(0, len(lines), max_lines)]


# ── DOCX 파싱 ─────────────────────────────────────────────────────────

def parse_title_paragraph(text: str) -> tuple[str, str]:
    """
    '마 28:1-20 "가서 제자 삼으라"...' → (passage='마 28:1-20', title='가서 제자 삼으라')
    """
    passage_m = re.match(r'^([가-힣A-Za-z]+\s+\d+:\d+(?:-\d+)?)', text.strip())
    passage = passage_m.group(1).strip() if passage_m else ''

    # 한국어/유니코드 따옴표 모두 처리
    title_m = re.search(r'["\u201c\u201d\u0022\uff02](.+?)["\u201c\u201d\u0022\uff02]', text)
    if not title_m:
        title_m = re.search(r'[""](.+?)[""]', text)
    title = title_m.group(1).strip() if title_m else ''
    return passage, title


def extract_verses(para_text: str) -> tuple[str, str, list[tuple[str, str]]]:
    """
    '마 27:51 이에 성소... 27:52 무덤들이...'
    → (book='마', chapter='27', [(verse_num, verse_text), ...])
    """
    book_m = re.match(r'^([가-힣A-Za-z]+)\s+', para_text)
    book = book_m.group(1) if book_m else ''

    tokens = re.split(r'(\d+:\d+)\s*', para_text)
    chapter, verses = None, []
    i = 0
    while i < len(tokens):
        tok = tokens[i].strip()
        if re.match(r'^\d+:\d+$', tok):
            ch, v = tok.split(':')
            if chapter is None:
                chapter = ch
            verse_text = tokens[i+1].strip() if i+1 < len(tokens) else ''
            verses.append((v, verse_text))
            i += 2
        else:
            i += 1

    return book, chapter or '', verses


# ── PPT 슬라이드 생성 ─────────────────────────────────────────────────

def _add_slide_from_template(out_prs: Presentation, template_slide) -> object:
    """템플릿 슬라이드의 도형들을 복사해 새 슬라이드를 추가."""
    layout = out_prs.slide_layouts[0]
    new_slide = out_prs.slides.add_slide(layout)
    sp_tree = new_slide.shapes._spTree

    # 레이아웃에서 자동으로 추가된 도형 제거
    for child in list(sp_tree):
        if child.tag not in (qn('p:nvGrpSpPr'), qn('p:grpSpPr')):
            sp_tree.remove(child)

    # 템플릿 도형 복사
    for shape in template_slide.shapes:
        sp_tree.append(copy.deepcopy(shape.element))

    return new_slide


def _get_shape(slide, name: str):
    for sh in slide.shapes:
        if sh.name == name:
            return sh
    return None


def _set_tf(tf, lines: list[str], font_size: int = None, color: RGBColor = None):
    """텍스트 프레임의 내용을 교체."""
    tf.clear()
    for i, line in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        run = para.add_run()
        run.text = line
        run.font.name = FONT_NAME
        if font_size:
            run.font.size = Pt(font_size)
        if color:
            run.font.color.rgb = color


def _make_verse_slide(out_prs, template_normal, template_slash,
                      verse_num: str, content_lines: list[str],
                      bottom_text: str | None, use_slash_pos: bool):
    """
    보조본문 슬라이드 생성.
    bottom_text=None → 하단 참조 없음 (중간 연속 슬라이드)
    use_slash_pos=True → // 높이 위치 (살짝 위)
    """
    tmpl = template_slash if use_slash_pos and bottom_text else template_normal
    slide = _add_slide_from_template(out_prs, tmpl)

    # 절 번호 (TextBox 7)
    tb7 = _get_shape(slide, 'TextBox 7')
    if tb7 and tb7.has_text_frame:
        _set_tf(tb7.text_frame, [verse_num], font_size=36)

    # 절 내용 (TextBox 6)
    tb6 = _get_shape(slide, 'TextBox 6')
    if tb6 and tb6.has_text_frame:
        _set_tf(tb6.text_frame, content_lines, font_size=40)

    # 하단 참조 (TextBox 2)
    tb2 = _get_shape(slide, 'TextBox 2')
    if bottom_text is None:
        # 중간 연속 슬라이드: 하단 텍스트박스 제거
        if tb2:
            tb2.element.getparent().remove(tb2.element)
    else:
        if tb2 and tb2.has_text_frame:
            _set_tf(tb2.text_frame, [bottom_text], font_size=80, color=BLUE)

    return slide


def _make_title_slide(out_prs, template_title, passage: str, sermon_title: str):
    """설교제목 슬라이드 생성."""
    slide = _add_slide_from_template(out_prs, template_title)

    tb3 = _get_shape(slide, 'TextBox 3')
    if tb3 and tb3.has_text_frame:
        _set_tf(tb3.text_frame, [sermon_title], font_size=40)

    tb4 = _get_shape(slide, 'TextBox 4')
    if tb4 and tb4.has_text_frame:
        _set_tf(tb4.text_frame, [passage], font_size=30)

    # TextBox 1 ("설교제목", 초록색)는 템플릿에서 그대로 유지

    return slide


# ── 메인 생성 함수 ────────────────────────────────────────────────────

def generate_ppt(template_path: str,
                 passage: str,
                 sermon_title: str,
                 yellow_para_texts: list[str]) -> bytes:
    """
    PPT 바이트 생성.

    template_path: ppt_template.pptx 경로
    passage: '마 28:1-20'
    sermon_title: '가서 제자 삼으라'
    yellow_para_texts: 보조본문 단락 텍스트 목록 (순서대로)
    """
    template_prs = Presentation(template_path)
    t_normal = template_prs.slides[IDX_VERSE_NORMAL]
    t_slash  = template_prs.slides[IDX_VERSE_SLASH]
    t_title  = template_prs.slides[IDX_TITLE]

    # 출력 프레젠테이션 (템플릿 기반으로 마스터/테마 유지)
    out_prs = Presentation(template_path)
    # 기존 슬라이드 제거
    slide_id_lst = out_prs.slides._sldIdLst
    for sldId in list(slide_id_lst):
        rId = sldId.get(qn('r:id'))
        slide_id_lst.remove(sldId)
        try:
            out_prs.slides.part.drop_rel(rId)
        except Exception:
            pass

    # ── 슬라이드 빌드 ──────────────────────────────────────────────
    # 구조: [GROUP1 SLIDES] [설교제목] [GROUP2 SLIDES] [설교제목] ...

    for group_idx, para_text in enumerate(yellow_para_texts):
        book, chapter, verses = extract_verses(para_text)
        if not verses:
            continue

        # 그룹 전체 범위 참조 (예: "마 27:51-52")
        all_verse_nums = [int(v) for v, _ in verses]
        v_min, v_max = min(all_verse_nums), max(all_verse_nums)
        if v_min == v_max:
            group_ref = f'({book} {chapter}:{v_min})'
        else:
            group_ref = f'({book} {chapter}:{v_min}-{v_max})'

        # 이 그룹의 모든 슬라이드 정보를 먼저 계산
        group_slides = []  # list of dict
        for vi, (verse_num, verse_text) in enumerate(verses):
            is_last_verse = (vi == len(verses) - 1)
            lines = split_to_lines(verse_text)
            chunks = split_to_slides(lines)

            for ci, chunk in enumerate(chunks):
                is_first_chunk = (ci == 0)
                is_last_chunk = (ci == len(chunks) - 1)
                is_last_of_group = is_last_verse and is_last_chunk

                # 마지막 슬라이드의 마지막 줄에 (group_ref) 추가
                content = list(chunk)
                if is_last_of_group:
                    content[-1] = content[-1] + ' ' + group_ref

                group_slides.append({
                    'verse_num': verse_num,
                    'content': content,
                    'is_first_chunk': is_first_chunk,
                    'is_last_of_group': is_last_of_group,
                    'book': book,
                    'chapter': chapter,
                })

        # 슬라이드 생성
        for si, sd in enumerate(group_slides):
            verse_num = sd['verse_num']
            content = sd['content']
            is_first = sd['is_first_chunk']
            is_last_of_group = sd['is_last_of_group']
            bk = sd['book']
            ch = sd['chapter']

            if is_last_of_group:
                # 그룹 마지막: // 스타일
                if is_first:
                    bottom_text = f'{bk} {ch}:{verse_num} //'
                else:
                    bottom_text = '//'
                use_slash = True
            elif is_first:
                # 절의 첫 슬라이드: 참조 표시
                bottom_text = f'{bk} {ch}:{verse_num}'
                use_slash = False
            else:
                # 중간 연속 슬라이드: 참조 없음
                bottom_text = None
                use_slash = False

            _make_verse_slide(out_prs, t_normal, t_slash,
                              verse_num, content, bottom_text, use_slash)

        # 그룹 뒤 설교제목 슬라이드
        _make_title_slide(out_prs, t_title, passage, sermon_title)

    # bytes로 반환
    buf = io.BytesIO()
    out_prs.save(buf)
    buf.seek(0)
    return buf.read()
