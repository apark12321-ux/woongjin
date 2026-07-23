#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
hwpx_eqn.py — PDF/이미지 수학문제 → 한글 HWPX 변환 엔진 (핵심 모듈)
================================================================
한컴 공식 수식 명령어 표준(help.hancom.com) 기반.

[사용 흐름]
1. 각 문제를 Problem 객체로 구성 (텍스트와 수식을 «...» 마크업으로 작성)
2. build_document()로 .hwpx 생성
3. 한글에서 열면 수식 편집기 없이 즉시 수식 렌더링

[핵심 설계 원칙]
- 텍스트와 수식이 한 문장에서 정확한 위치에 배치됨 (인라인)
- 수식은 hp:equation 객체로 삽입 → 진짜 수식으로 렌더링
- 수식 문자열은 normalize_eqn()으로 한컴 표준에 맞게 정규화
"""
import re
from hwpx.document import HwpxDocument
from lxml import etree

HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"


# =====================================================================
# 1) 수식 정규화 — 한컴 공식 표준에 맞게 명령어 교정
# =====================================================================

# 글자 장식 명령어: 첫 글자 대문자 표기가 가장 안전 (공식 문서 권고)
# 소문자도 대부분 동작하지만, 호환성을 위해 표준 표기로 정규화
DECORATION_CMDS = {
    "bar": "bar", "vec": "vec", "hat": "hat", "tilde": "tilde",
    "acute": "acute", "grave": "grave", "dot": "dot", "ddot": "ddot",
    "dyad": "dyad", "under": "under", "check": "check", "arch": "arch",
}

# LaTeX → 한컴 매핑 (Mathpix 등 LaTeX 출력을 한컴으로 변환할 때 사용)
LATEX_TO_HWP = {
    r"\frac": "OVER_FRAC",   # 특수 처리 필요 (\frac{a}{b} → {a} over {b})
    r"\sqrt": "sqrt",
    r"\times": "times",
    r"\div": "div",
    r"\pm": "+-",
    r"\mp": "-+",
    r"\cdot": "cdot",
    r"\cdots": "cdots",
    r"\ldots": "...",
    r"\leq": "<=",
    r"\geq": ">=",
    r"\neq": "!=",
    r"\infty": "inf",
    r"\alpha": "alpha", r"\beta": "beta", r"\gamma": "gamma",
    r"\delta": "delta", r"\theta": "theta", r"\pi": "pi",
    r"\lambda": "lambda", r"\mu": "mu", r"\sigma": "sigma",
    r"\sum": "sum", r"\int": "int", r"\lim": "lim",
    r"\left": "", r"\right": "",
    r"\overline": "bar",
    r"\bar": "bar",
}


def latex_frac_to_hwp(s):
    r"""\frac{a}{b} → {a} over {b} 변환 (중첩 대응)"""
    def find_brace(text, start):
        depth = 0
        for i in range(start, len(text)):
            if text[i] == "{":
                depth += 1
            elif text[i] == "}":
                depth -= 1
                if depth == 0:
                    return i
        return -1

    while r"\frac" in s:
        idx = s.find(r"\frac")
        b1 = s.index("{", idx)
        b1e = find_brace(s, b1)
        num = s[b1+1:b1e]
        b2 = s.index("{", b1e)
        b2e = find_brace(s, b2)
        den = s[b2+1:b2e]
        repl = "{" + num + "} over {" + den + "}"
        s = s[:idx] + repl + s[b2e+1:]
    return s


def latex_to_hwp(latex):
    """LaTeX 수식 → 한컴 수식 문자열 변환"""
    s = latex.strip().strip("$")
    s = latex_frac_to_hwp(s)
    for tex, hwp in LATEX_TO_HWP.items():
        if tex == r"\frac":
            continue
        s = s.replace(tex, " " + hwp + " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def normalize_eqn(script, add_trailing_space=True):
    r"""
    한컴 수식 문자열 정규화 (최종 완성 파일 분석 기반).
    - LaTeX 잔재(\) 감지 시 변환
    - 연속 공백 정리 (중괄호 구조는 보존)
    - 끝에 백틱(`) 자동 추가: 수식 뒤 글자와의 간격 확보 (완성 파일의 98%가 이 방식)
      백틱은 한컴 수식에서 1/4 빈칸. 겹침도 밀림도 막아줌.
    """
    if "\\" in script:  # LaTeX 잔재가 있으면 변환 시도
        script = latex_to_hwp(script)
    # 연속 공백 정리
    script = re.sub(r"[ \t]+", " ", script).strip()
    # 끝 백틱 추가 (이미 ` 로 끝나면 생략)
    if add_trailing_space and script and not script.endswith("`"):
        script = script + "`"
    return script


# =====================================================================
# 2) HWPX 빌더 — 텍스트/수식을 정확한 위치에 삽입
# =====================================================================

# 한컴 길이 단위: 1mm ≈ 283 HwpUnit, baseUnit 1000 = 약 10pt
_EQ_CMDS = ['over', 'sqrt', 'bar', 'vec', 'hat', 'tilde', 'times', 'cdot',
            'cdots', 'sum', 'int', 'lim', 'inf', 'alpha', 'beta', 'gamma',
            'delta', 'theta', 'pi', 'sigma', 'neq', 'TIMES', 'leq', 'geq']


def estimate_eqn_size(script):
    """
    수식 문자열 → (width, height) HwpUnit.
    완성 파일(단원체크.hwpx) 792개 수식에서 역산한 실측 글자폭 기반.
    한글이 hwpx의 width를 그대로 신뢰하므로 정확도가 핵심.
    """
    # 실측 글자폭 테이블 (HwpUnit)
    CW = {
        '(': 460, ')': 501, '+': 1015, ',': 434, '-': 961, '.': 253, '/': 460,
        '0': 564, '1': 561, '2': 561, '3': 595, '4': 588, '5': 585,
        '6': 599, '7': 607, '8': 576, '9': 615, '=': 1153,
        'a': 603, 'b': 491, 'c': 472, 'd': 530, 'e': 500, 'f': 535,
        'g': 530, 'h': 530, 'i': 280, 'j': 280, 'k': 600, 'l': 367,
        'm': 1012, 'n': 560, 'o': 530, 'p': 530, 'q': 530, 'r': 400,
        's': 460, 't': 360, 'u': 560, 'v': 530, 'w': 780, 'x': 664,
        'y': 605, 'z': 500, '|': 218, '℃': 1042, '`': 58, ' ': 0,
        '<': 1153, '>': 1153, '!': 280, '*': 560,
    }
    DEFAULT_CW = 560
    FRAC_PAD = 560      # 분수선 좌우 여백
    SUP_RATIO = 0.65    # 위/아래첨자 축소비

    def text_width(s):
        return sum(CW.get(c, DEFAULT_CW) for c in s)

    def measure(s):
        """구조를 고려한 폭 계산"""
        s = s.strip()
        if not s:
            return 0
        total = 0
        i = 0
        while i < len(s):
            # 최상위 over 탐지
            if s[i:i+4] == 'over' and (i == 0 or not s[i-1].isalnum()) \
                    and (i+4 >= len(s) or not s[i+4].isalnum()):
                # 이 위치 기준 좌/우가 분자/분모 (단순형)
                # 보통 {a}over{b} 형태는 아래 그룹 처리에서 다룸
                i += 4
                continue
            if s[i] == '{':
                d = 0; j = i
                while j < len(s):
                    if s[j] == '{': d += 1
                    elif s[j] == '}':
                        d -= 1
                        if d == 0: break
                    j += 1
                inner = s[i+1:j]
                # 다음이 over 인지 확인 → 분수
                rest = s[j+1:].lstrip()
                if rest.startswith('over'):
                    # {분자} over {분모 또는 토큰}
                    after = s[j+1:]
                    k = after.find('over') + 4
                    # 분모 토큰 추출
                    den_str, consumed = _read_token(after[k:])
                    num_w = measure(inner)
                    den_w = measure(den_str)
                    total += max(num_w, den_w) + FRAC_PAD
                    i = j + 1 + k + consumed
                    continue
                else:
                    total += measure(inner); i = j + 1; continue
            if s[i] in '^_':
                tok, consumed = _read_token(s[i+1:])
                total += int(measure(tok) * SUP_RATIO)
                i += 1 + consumed; continue
            # 단순 'NoverM' (중괄호 없는 분수): 앞 토큰 over 뒤 토큰
            m = re.match(r'([0-9a-zA-Z]+)over([0-9a-zA-Z]+)', s[i:])
            if m:
                num_w = text_width(m.group(1))
                den_w = text_width(m.group(2))
                total += max(num_w, den_w) + FRAC_PAD
                i += m.end(); continue
            # 명령어 (rm/ita는 뒤 글자와 붙어 있으므로 우선 분리)
            if s[i:i+2] == 'rm' and (i+2 >= len(s) or not s[i+2] in 'aeiou' or True):
                # rm 뒤 단위는 로만체로 좁게 (실측 글자당 약 575)
                i += 2
                unit = re.match(r'[a-zA-Z]+', s[i:])
                if unit:
                    total += len(unit.group(0)) * 575
                    i += unit.end()
                continue
            if s[i:i+3] == 'ita':
                i += 3
                continue
            mc = re.match(r'[a-zA-Z]+', s[i:])
            if mc:
                word = mc.group(0)
                CMD = {'sqrt': 700, 'times': 800, 'cdot': 300, 'cdots': 900,
                       'left': 200, 'right': 200,
                       'bar': 0, 'vec': 0, 'hat': 0, 'alpha': 600, 'beta': 600,
                       'pi': 600, 'theta': 600, 'sum': 800, 'int': 700,
                       'lim': 900, 'inf': 700, 'triangle': 1100, 'cases': 0,
                       'neq': 1153, 'leq': 1153, 'geq': 1153}
                if word in CMD:
                    total += CMD[word]
                else:
                    total += text_width(word)
                i += mc.end(); continue
            total += CW.get(s[i], DEFAULT_CW); i += 1
        return total

    width = max(int(measure(script)), 300)

    # 높이 + baseLine: 구조 기반 (완성파일 실측 매핑)
    if 'cases' in script or '#' in script:
        height, base_line = 4650, 58
    elif 'over' in script or 'sqrt' in script:
        height, base_line = 2250, 66
    elif '^' in script or '_' in script:
        height, base_line = 1175, 88
    else:
        height, base_line = 1000, 86
    return width, height, base_line


def _read_token(s):
    """문자열 앞에서 한 토큰(그룹 또는 단어/문자) 읽기 → (토큰내용, 소비길이)"""
    s2 = s.lstrip()
    lead = len(s) - len(s2)
    if not s2:
        return "", lead
    if s2[0] == '{':
        d = 0
        for j in range(len(s2)):
            if s2[j] == '{': d += 1
            elif s2[j] == '}':
                d -= 1
                if d == 0:
                    return s2[1:j], lead + j + 1
        return s2[1:], lead + len(s2)
    m = re.match(r'[0-9a-zA-Z]+', s2)
    if m:
        return m.group(0), lead + m.end()
    return s2[0], lead + 1


def _add_equation_element(run, script, eq_id, width, height, base_line=86):
    """run 안에 한글 표준 수식 객체를 추가 (완성파일 단원체크.hwpx와 동일 구조).

    완성 파일 분석 결과:
    - font 속성 없음 (생략)
    - width/height는 실측 글자폭 기반 정확값 (한글이 그대로 신뢰)
    - baseLine은 height에 따라 다름: 1줄=86, 첨자=88, 분수=66, cases=58
    - outMargin left/right=56
    """
    eq = etree.SubElement(run, f"{{{HP}}}equation")
    eq.set("id", str(eq_id))
    eq.set("zOrder", "0")
    eq.set("numberingType", "EQUATION")
    eq.set("textWrap", "SQUARE")
    eq.set("textFlow", "BOTH_SIDES")
    eq.set("lock", "0")
    eq.set("version", "Equation Version 60")
    eq.set("baseLine", str(base_line))
    eq.set("textColor", "#000000")
    eq.set("baseUnit", "1000")
    eq.set("lineMode", "CHAR")
    sz = etree.SubElement(eq, f"{{{HP}}}sz")
    sz.set("width", str(width)); sz.set("widthRelTo", "ABSOLUTE")
    sz.set("height", str(height)); sz.set("heightRelTo", "ABSOLUTE")
    sz.set("protect", "0")
    pos = etree.SubElement(eq, f"{{{HP}}}pos")
    pos.set("treatAsChar", "1")
    pos.set("affectLSpacing", "0")
    pos.set("flowWithText", "1")
    pos.set("allowOverlap", "0")
    pos.set("holdAnchorAndSO", "0")
    pos.set("vertRelTo", "PARA"); pos.set("horzRelTo", "PARA")
    pos.set("vertAlign", "TOP"); pos.set("horzAlign", "LEFT")
    pos.set("vertOffset", "0"); pos.set("horzOffset", "0")
    om = etree.SubElement(eq, f"{{{HP}}}outMargin")
    om.set("left", "56"); om.set("right", "56")
    om.set("top", "0"); om.set("bottom", "0")
    sc = etree.SubElement(eq, f"{{{HP}}}shapeComment")
    sc.text = "수식입니다."
    s = etree.SubElement(eq, f"{{{HP}}}script")
    s.text = script


def _clear_para(para):
    for child in list(para.element):
        if child.tag == f"{{{HP}}}run":
            para.element.remove(child)


_EQ_COUNTER = [0]


# ---- 글자 장식 스타일(밑줄 __..__, 볼드 **..**) ----
_STYLE = {"u": "0", "b": "0", "bu": "0"}   # build_document에서 문서별로 채움

def _register_styles(doc):
    """밑줄/볼드/볼드+밑줄 charPr 등록(라이브러리 ensure_run_style)."""
    _STYLE["u"]  = str(doc.ensure_run_style(underline=True))
    _STYLE["b"]  = str(doc.ensure_run_style(bold=True))
    _STYLE["bu"] = str(doc.ensure_run_style(bold=True, underline=True))

def _charpr_for(bold, ul):
    if bold and ul: return _STYLE["bu"]
    if bold:        return _STYLE["b"]
    if ul:          return _STYLE["u"]
    return "0"

def _parse_deco(text):
    """**볼드**, __밑줄__ 토글 파싱 → [(piece, bold, ul), ...]"""
    out = []; i = 0; bold = False; ul = False; buf = ""
    def flush():
        nonlocal buf
        if buf:
            out.append((buf, bold, ul)); buf = ""
    while i < len(text):
        if text[i:i+2] == "**":
            flush(); bold = not bold; i += 2
        elif text[i:i+2] == "__":
            flush(); ul = not ul; i += 2
        else:
            buf += text[i]; i += 1
    flush()
    return out

def _emit_segments(para_el, segments):
    """문단 element에 segments 배치. ('t',텍스트)=장식 파싱, ('e',식)=수식 객체."""
    if not segments:
        run = etree.SubElement(para_el, f"{{{HP}}}run"); run.set("charPrIDRef", "0")
        etree.SubElement(run, f"{{{HP}}}t"); return
    for kind, content in segments:
        if kind == "t":
            for piece, b, u in _parse_deco(content):
                if piece == "":
                    continue
                run = etree.SubElement(para_el, f"{{{HP}}}run")
                run.set("charPrIDRef", _charpr_for(b, u))
                t = etree.SubElement(run, f"{{{HP}}}t"); t.text = piece
        elif kind == "e":
            run = etree.SubElement(para_el, f"{{{HP}}}run"); run.set("charPrIDRef", "0")
            _EQ_COUNTER[0] += 1
            script = normalize_eqn(content)
            width, height, base_line = estimate_eqn_size(script)
            _add_equation_element(run, script, _EQ_COUNTER[0], width, height, base_line)


def add_paragraph_segments(doc, segments):
    """한 줄(문단)을 doc에 추가. segments=[('t',..),('e',..)]. 빈 리스트=빈 문단."""
    para = doc.add_paragraph("")
    _clear_para(para)
    _emit_segments(para.element, segments)
    return para


def add_bogi_box(doc, items):
    """〈보기〉 박스: 가운데 '〈보기〉' 라벨 + 테두리 1칸 표(글자취급)에 항목 배치.
    items: 보기 항목 줄 리스트(«»수식·__밑줄__·**볼드** 지원)."""
    # 라벨(가운데정렬은 후처리에서 paraPrIDRef=100 부여)
    lbl = add_paragraph_segments(doc, [("t", "〈보기〉")])
    lbl.element.set("paraPrIDRef", "100")
    # 1칸 표
    tbl = doc.add_table(1, 1)
    cell = tbl.cell(0, 0)
    # 셀 기본 문단 비우고 첫 항목, 이후 항목은 새 문단
    base_ps = cell.paragraphs
    first_el = base_ps[0].element if base_ps else cell.add_paragraph("").element
    _clear_para_el(first_el)
    _emit_segments(first_el, parse_markup(latinize_line(items[0])) if items else [])
    for it in items[1:]:
        pel = cell.add_paragraph("").element
        _clear_para_el(pel)
        _emit_segments(pel, parse_markup(latinize_line(it)))
    return tbl


def _clear_para_el(para_el):
    for child in list(para_el):
        if child.tag == f"{{{HP}}}run":
            para_el.remove(child)


_LATIN_RUN = re.compile(r'[A-Za-z□]+')

def latinize_line(line):
    """«» 밖의 영문(및 □) 런을 모두 «»수식으로 감싸고, 인접 수식끼리 병합.
    예: 'A«(1, 4)», B«(5, 4)»' → '«A(1, 4)», «B(5, 4)»'
        '□ABCD는' → '«□ABCD»는' / 'a, b, c의' → '«a», «b», «c»의'"""
    if "@@" in line:           # @@BOGI@@ 등 마커는 건드리지 않음
        return line
    out = []; i = 0; n = len(line)
    while i < n:
        if line[i] == "«":     # 기존 수식 구간은 그대로 보존
            j = line.find("»", i)
            if j == -1:
                out.append(line[i:]); break
            out.append(line[i:j+1]); i = j + 1
        else:
            k = line.find("«", i)
            seg = line[i:(k if k != -1 else n)]
            seg = _LATIN_RUN.sub(lambda m: "«" + m.group(0) + "»", seg)
            out.append(seg)
            i = (k if k != -1 else n)
    res = "".join(out)
    res = res.replace("»«", "")   # 라벨+좌표 등 인접 수식 병합
    return res


def parse_markup(line):
    """
    «...»로 감싼 부분 = 수식, 나머지 = 텍스트.
    텍스트와 수식이 원본 순서대로 정확히 배치됨.
    """
    segs = []
    buf = ""
    i = 0
    while i < len(line):
        if line[i] == "«":
            if buf:
                segs.append(("t", buf)); buf = ""
            j = line.find("»", i)
            if j == -1:  # 닫는 기호 없음 → 그냥 텍스트로
                buf += line[i]; i += 1; continue
            segs.append(("e", line[i+1:j].strip()))
            i = j + 1
        else:
            buf += line[i]; i += 1
    if buf:
        segs.append(("t", buf))
    return segs


# =====================================================================
# 3) 문서 빌더 — 문제 리스트 → 완성된 .hwpx
# =====================================================================

class Problem:
    """한 문제 = 번호 + 줄(line) 리스트. 줄 안의 «...»는 수식."""
    def __init__(self, number, lines):
        self.number = number
        self.lines = lines


def _postprocess_header(hwpx_path, font_name=None, left_align=True, font_size_pt=10):
    """header.xml 후처리: 폰트 교체, (옵션)왼쪽맞춤, 〈보기〉 라벨용 CENTER 문단모양(id=100) 생성.
    글자 크기는 기본 템플릿이 10pt이므로 그대로 둔다(양쪽정렬은 left_align=False)."""
    import zipfile, shutil, re as _re
    tmp = hwpx_path + ".tmp"
    with zipfile.ZipFile(hwpx_path, "r") as zin:
        names = zin.namelist()
        data = {n: zin.read(n) for n in names}
    header = data["Contents/header.xml"].decode("utf-8")
    if font_name:
        header = _re.sub(r'face="[^"]+"', f'face="{font_name}"', header)
    if left_align:
        header = header.replace('horizontal="JUSTIFY"', 'horizontal="LEFT"')
    # 가운데정렬 문단모양(id=100) 1회 생성 (〈보기〉 라벨용)
    if 'id="100"' not in header:
        m = _re.search(r'(<hh:paraPr\b[^>]*id="1"[^>]*>.*?</hh:paraPr>)', header, _re.S)
        if m:
            clone = m.group(0).replace('id="1"', 'id="100"', 1)
            if 'horizontal="CENTER"' not in clone:
                clone = _re.sub(r'horizontal="[^"]+"', 'horizontal="CENTER"', clone, count=1)
            header = header.replace(m.group(0), m.group(0) + clone, 1)
            header = _re.sub(r'(<hh:paraPrs[^>]*itemCnt=")(\d+)(")',
                             lambda x: f'{x.group(1)}{int(x.group(2)) + 1}{x.group(3)}',
                             header, count=1)
    data["Contents/header.xml"] = header.encode("utf-8")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            if n == "mimetype":
                zout.writestr(n, data[n], compress_type=zipfile.ZIP_STORED)
            else:
                zout.writestr(n, data[n])
    shutil.move(tmp, hwpx_path)


def build_document(problems, out_path, title=None, subtitle=None,
                   font="나눔바른고딕 옛한글", answers_vertical=True,
                   left_align=True, bracket_number=True, font_size_pt=10):
    """
    problems: Problem 객체 리스트
    out_path: 저장할 .hwpx 경로
    font: 폰트명 / font_size_pt: 글자 크기(pt)
    answers_vertical: 선택지(①~⑤) 세로 배열
    left_align: True=왼쪽맞춤, False=양쪽맞춤(JUSTIFY)
    bracket_number: True면 [번호], False면 번호 그대로
    본문 줄 안에서 «»=수식, __..__=밑줄, **..**=볼드.
    '@@BOGI@@' ~ '@@/BOGI@@' 사이 줄들은 〈보기〉 박스(표)로.
    """
    doc = HwpxDocument.new()
    _register_styles(doc)
    if title:
        add_paragraph_segments(doc, [("t", title)])
    if subtitle:
        add_paragraph_segments(doc, [("t", subtitle)])
    if title or subtitle:
        add_paragraph_segments(doc, [])

    for p in problems:
        header = f"[{p.number}]" if bracket_number else str(p.number)
        add_paragraph_segments(doc, parse_markup(latinize_line(header)))
        lines = list(p.lines)
        i = 0
        while i < len(lines):
            line = lines[i]
            if line.strip() == "@@BOGI@@":
                items = []
                i += 1
                while i < len(lines) and lines[i].strip() != "@@/BOGI@@":
                    if lines[i] != "":
                        items.append(lines[i])
                    i += 1
                add_bogi_box(doc, items)
                i += 1  # @@/BOGI@@ 스킵
                continue
            if line == "":
                add_paragraph_segments(doc, [])
            elif answers_vertical and _is_choice_line(line):
                for choice in _split_choices(line):
                    add_paragraph_segments(doc, parse_markup(latinize_line(choice)))
            else:
                add_paragraph_segments(doc, parse_markup(latinize_line(line)))
            i += 1
        add_paragraph_segments(doc, [])  # 문제 간 간격

    doc.save_to_path(out_path)
    _postprocess_header(out_path, font_name=font, left_align=left_align,
                        font_size_pt=font_size_pt)
    # 통계
    import zipfile
    with zipfile.ZipFile(out_path) as z:
        content = z.read('Contents/section0.xml').decode('utf-8')
    eq_count = content.count('<hp:equation')
    return {"problems": len(problems), "equations": eq_count, "path": out_path}


# 선택지 원문자
_CHOICE_MARKS = "①②③④⑤⑥⑦⑧⑨⑩"


def insert_images(hwpx_path, images):
    """
    생성된 hwpx에 문항 이미지를 삽입한다.
    images: [(token, png_path, disp_w, disp_h), ...]
      - token: 본문 마크업에 넣어둔 자리표시 문자열 (예: '@IMG:0544@')
      - png_path: 삽입할 PNG 파일 경로
      - disp_w, disp_h: 표시 크기(HwpUnit). 미지정 시 원본 비율로 자동.
    배치 규칙: 토큰이 든 문단을 '가운데 정렬'로 만들고 그 자리에 그림을 넣는다.
    (마크업에서 토큰은 '문제 본문 다음 빈 줄 다음 줄'에 두면 그 위치에 가운데 배치됨)
    """
    import zipfile, shutil, re as _re
    from PIL import Image as _PILImage

    tmp = hwpx_path + ".tmp"
    with zipfile.ZipFile(hwpx_path) as zin:
        data = {n: zin.read(n) for n in zin.namelist()}

    header = data['Contents/header.xml'].decode('utf-8')
    hpf = data['Contents/content.hpf'].decode('utf-8')
    sec = data['Contents/section0.xml'].decode('utf-8')

    # hc 네임스페이스 보장
    if 'xmlns:hc=' not in sec[:sec.find('>') + 1]:
        sec = sec.replace('<hs:sec ',
            '<hs:sec xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" ', 1)

    # 가운데정렬 문단모양(CENTER) 1회만 추가 (id=100)
    if 'id="100"' not in header:
        m = _re.search(r'(<hh:paraPr\b[^>]*id="1"[^>]*>.*?</hh:paraPr>)', header, _re.S)
        if m:
            clone = m.group(0).replace('id="1"', 'id="100"', 1)
            clone = clone.replace('horizontal="LEFT"', 'horizontal="CENTER"', 1)
            header = header.replace(m.group(0), m.group(0) + clone, 1)
            header = _re.sub(r'(<hh:paraPrs[^>]*itemCnt=")(\d+)(")',
                             lambda x: f'{x.group(1)}{int(x.group(2)) + 1}{x.group(3)}',
                             header, count=1)

    pic_id = 2000000001
    for i, spec in enumerate(images):
        token, png_path = spec[0], spec[1]
        disp_w = spec[2] if len(spec) > 2 and spec[2] else None
        disp_h = spec[3] if len(spec) > 3 and spec[3] else None
        # 표시 크기 자동: 원본 비율 유지, 가로 기준 약 50000 HwpUnit(≈ 17.6mm*?) 적정값
        if not (disp_w and disp_h):
            im = _PILImage.open(png_path)
            pw, ph = im.size
            disp_w = 38000
            disp_h = int(disp_w * ph / pw)
        img_id = f"imgC{i+1}"
        data[f"BinData/{img_id}.png"] = open(png_path, "rb").read()
        item = (f'<opf:item id="{img_id}" href="BinData/{img_id}.png" '
                f'media-type="image/png" isEmbeded="1"/>')
        hpf = hpf.replace('</opf:manifest>', item + '</opf:manifest>')

        pic = (
            f'<hp:pic id="{pic_id}" zOrder="0" numberingType="PICTURE" textWrap="SQUARE" '
            f'textFlow="BOTH_SIDES" lock="0" groupLevel="0" instid="{pic_id}" reverse="0">'
            f'<hp:offset x="0" y="0"/><hp:orgSz width="{disp_w}" height="{disp_h}"/>'
            f'<hp:curSz width="{disp_w}" height="{disp_h}"/><hp:flip horizontal="0" vertical="0"/>'
            f'<hp:rotationInfo angle="0" centerX="{disp_w//2}" centerY="{disp_h//2}"/>'
            f'<hp:renderingInfo><hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>'
            f'<hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>'
            f'<hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/></hp:renderingInfo>'
            f'<hp:imgRect><hc:pt0 x="0" y="0"/><hc:pt1 x="{disp_w}" y="0"/>'
            f'<hc:pt2 x="{disp_w}" y="{disp_h}"/><hc:pt3 x="0" y="{disp_h}"/></hp:imgRect>'
            f'<hp:imgClip left="0" right="0" top="0" bottom="0"/>'
            f'<hp:inMargin left="0" right="0" top="0" bottom="0"/>'
            f'<hc:img binaryItemIDRef="{img_id}" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/>'
            f'<hp:effects/><hp:sz width="{disp_w}" widthRelTo="ABSOLUTE" height="{disp_h}" '
            f'heightRelTo="ABSOLUTE" protect="0"/>'
            f'<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" '
            f'holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" '
            f'horzAlign="CENTER" vertOffset="0" horzOffset="0"/>'
            f'<hp:outMargin left="0" right="0" top="0" bottom="0"/>'
            f'<hp:shapeComment>그림입니다.</hp:shapeComment></hp:pic>'
        )
        pic_id += 1

        esc = _re.escape(token)
        def repl_para(mm):
            para = mm.group(0)
            para = _re.sub(r'paraPrIDRef="\d+"', 'paraPrIDRef="100"', para, count=1)
            para = _re.sub(rf'<hp:t>{esc}</hp:t>', '<hp:t></hp:t>' + pic, para)
            return para
        sec = _re.sub(
            rf'<hp:p\b[^>]*>(?:(?!</hp:p>).)*?{esc}(?:(?!</hp:p>).)*?</hp:p>',
            repl_para, sec, count=1, flags=_re.S)

    data['Contents/header.xml'] = header.encode('utf-8')
    data['Contents/content.hpf'] = hpf.encode('utf-8')
    data['Contents/section0.xml'] = sec.encode('utf-8')

    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        if 'mimetype' in data:
            zout.writestr('mimetype', data.pop('mimetype'), zipfile.ZIP_STORED)
        for n, b in data.items():
            zout.writestr(n, b)
    shutil.move(tmp, hwpx_path)
    return len(images)


def _is_choice_line(line):
    """① 이 2개 이상 들어있으면 선택지 줄로 판단."""
    cnt = sum(line.count(m) for m in _CHOICE_MARKS)
    return cnt >= 2


def _split_choices(line):
    """'① a  ② b  ③ c' → ['① a', '② b', '③ c'] (원문자 기준 분리)"""
    result = []
    cur = ""
    for ch in line:
        if ch in _CHOICE_MARKS:
            if cur.strip():
                result.append(cur.strip())
            cur = ch
        else:
            cur += ch
    if cur.strip():
        result.append(cur.strip())
    return result


if __name__ == "__main__":
    # 자체 테스트
    probs = [
        Problem("0001", [
            "두 복소수 «z_1 =(1-i)^2 », «z_2 = {3- sqrt {3} i} over {3+ sqrt {3} i} »에 대하여 «z_1 z_2 »의 값은?",
            "",
            "① «-2»  ② «-1»  ③ «0»  ④ «1»  ⑤ «2»",
        ]),
        Problem("0002", [
            "켤레복소수 테스트: «z bar {z} =0»이면 «z=0»이다. 또 «bar { alpha } + bar { beta } ».",
        ]),
    ]
    result = build_document(probs, "/tmp/test_engine.hwpx",
                            title="엔진 테스트", subtitle="2문항")
    print("결과:", result)
