#!/usr/bin/env python3
"""DOCX(구글 문서에서 내려받은 .docx 포함) → HWPX 변환기  v12

사용법:
    pip install python-docx "python-hwpx==6.1.0"
    python docx2hwpx.py 입력.docx [출력.hwpx]
    python docx2hwpx.py 폴더/          # 폴더 안 모든 .docx 일괄 변환

지원
  - 제목(Title / Heading 1~6 / 제목 n), 본문 문단, 글머리 목록(• 기호)
  - 글자 서식: 굵게, 기울임, 밑줄, 취소선, 글자 색, 형광펜(강조색), 글자 크기, 글꼴 이름
    (문단·문자 스타일과 문서 기본값에서 물려받은 서식도 반영. 한글 글꼴(eastAsia) 우선)
  - 문단 서식: 정렬(왼쪽·가운데·오른쪽·양쪽·배분), 왼쪽/오른쪽 들여쓰기,
    첫 줄 들여쓰기, 내어쓰기, 줄 간격(배수·고정, '최소'는 고정/배수로 환산), 문단 앞/뒤 간격
    (본문과 표 셀 안 문단 모두, 문서 기본값·스타일에서 물려받은 값 포함)
  - 표: 가로 병합(gridSpan), 세로 병합(vMerge), 셀 배경색, 셀 안 글자·문단 서식,
    표 안의 표(여러 단계 중첩 가능, 안쪽 표가 잘리지 않게 셀 높이 자동 조정)
  - 머리글·바닥글: 글자 서식, 정렬, 쪽 번호(PAGE)·전체 쪽 수(NUMPAGES) 필드, 머리글 안 표(한 줄 텍스트)
  - 각주·미주: 한글 각주/미주로 변환(또는 옵션으로 본문 [1] 표시 + 문서 끝 목록)
    각주·미주 안 글자 서식(굵게·기울임·밑줄·취소선·색·형광펜·크기·글꼴), 여러 문단,
    이미지(글자처럼 취급), 표(셀 병합·배경색·표 안의 표 포함, 글자 9pt)
  - 쪽 종류별 머리글·바닥글: 홀수 쪽/짝수 쪽 다르게, 첫 쪽 다르게
    (첫 쪽 머리글이 비어 있으면 '첫 쪽 감추기', 내용이 있으면 첫 쪽 전용 구역으로 분리)
  - 여러 구역 문서: 구역마다 머리글·바닥글, '이전 구역에 연결' 상속, 구역별 첫 쪽 다르게,
    쪽 번호 새로 시작, 홀수/짝수 쪽에서 시작 → 한글 구역 나누기로 변환
  - 구역별 용지 크기·방향(세로/가로)·여백(위·아래·왼쪽·오른쪽·머리말·꼬리말·제본), 다단
    (단 수·간격·구분선·단별 폭, 같은 쪽 안에서 단이 바뀌는 '이어서' 구역 포함)
    표·이미지 폭은 그 구역의 단 폭에 맞춤
  - 구역별 세로쓰기(영문 눕힘/세움), 줄 번호(시작 번호·간격·본문과 거리·쪽/구역마다/이어서),
    쪽 테두리(변마다 선 종류·굵기·색, 종이/본문 기준 간격, 모든 쪽/첫 쪽만/첫 쪽 빼고)
  - 문단별 줄 번호 건너뛰기(suppressLineNumbers), 그림 쪽 테두리(무늬에 가까운 한글 선으로 근사),
    쪽 배경색(디자인 → 페이지 색)
  - 인라인 이미지(png/jpg/gif/bmp)
미지원: 그림 쪽 테두리의 무늬 자체(선으로 근사), 쪽 배경 그림·그라데이션(기본 색만), 각주·미주 안 이미지 위치(글 뒤/앞, 떠 있는 개체 → 글자처럼)
근사: 이어서(continuous) 구역 나누기는 머리글이 바뀌면 한글에서 새 쪽부터 시작
"""
import re
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from hwpx.document import HwpxDocument

EMU_PER_MM = 36000
IMG_EXT = {"image/png": "png", "image/jpeg": "jpg", "image/gif": "gif", "image/bmp": "bmp"}
HEADING_PT = {1: 20, 2: 16, 3: 14, 4: 12, 5: 11, 6: 10}

# Word 형광펜(w:highlight) 이름 → 색상 코드
HIGHLIGHT = {
    "yellow": "#FFFF00", "green": "#00FF00", "cyan": "#00FFFF", "magenta": "#FF00FF",
    "blue": "#0000FF", "red": "#FF0000", "darkBlue": "#000080", "darkCyan": "#008080",
    "darkGreen": "#008000", "darkMagenta": "#800080", "darkRed": "#800000",
    "darkYellow": "#808000", "darkGray": "#808080", "lightGray": "#C0C0C0",
    "black": "#000000", "white": "#FFFFFF",
}
HEX6 = re.compile(r"^[0-9A-Fa-f]{6}$")
TWIPS_PER_MM = 1440 / 25.4          # Word 길이 단위(twip, 1/20pt) → mm

# Word 정렬(w:jc) → 한글 정렬
ALIGN = {
    "left": "LEFT", "start": "LEFT",
    "center": "CENTER",
    "right": "RIGHT", "end": "RIGHT",
    "both": "JUSTIFY", "distribute": "DISTRIBUTE",
    "lowKashida": "JUSTIFY", "mediumKashida": "JUSTIFY", "highKashida": "JUSTIFY",
}


# ───────────────────────── 공통 도우미 ─────────────────────────
def _hex(val):
    """'FF0000' → '#FF0000'. auto/빈 값/잘못된 값은 None."""
    return f"#{val.upper()}" if val and HEX6.match(val) else None


def _style_chain(style):
    while style is not None:
        yield style
        style = style.base_style


def _rpr_val(rpr, tag):
    """rPr 안의 on/off 속성(w:b, w:i 등) → True/False/None"""
    if rpr is None:
        return None
    el = rpr.find(qn(tag))
    if el is None:
        return None
    v = el.get(qn("w:val"))
    return v not in ("0", "false", "none")


def _style_rpr(style):
    return style.element.rPr if style is not None and style.element is not None else None


def iter_blocks(doc):
    """본문의 문단·표를 문서 순서대로 돌려준다."""
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)


def heading_level(par):
    name = (par.style.name or "").lower() if par.style is not None else ""
    if name == "title":
        return 1
    m = re.match(r"(heading|제목)\s*(\d)", name)
    return int(m.group(2)) if m else 0


def heading_size(par, level, keep_format=True):
    """제목 글자 크기(pt): 서식 유지 시 docx 제목 스타일에 지정된 크기를 쓰고, 없으면 기본표."""
    if keep_format:
        for st in _style_chain(par.style) if par.style is not None else []:
            if st.font is not None and st.font.size is not None:
                return round(st.font.size.pt, 1)
    return HEADING_PT[min(level, 6)]


def is_list(par):
    ppr = par._p.pPr
    name = (par.style.name or "").lower() if par.style is not None else ""
    return (ppr is not None and ppr.numPr is not None) or "list" in name



# ───────────────────────── 글꼴 ─────────────────────────
def _theme_fonts(doc):
    """테마(theme1.xml)의 본문/제목 글꼴 → {"minor": {"latin":..,"ea":..,"hang":..}, "major": {...}}"""
    out = {"minor": {}, "major": {}}
    try:
        theme = next(p for p in doc.part.package.parts if p.partname.endswith("/theme1.xml"))
    except StopIteration:
        return out
    from lxml import etree
    root = etree.fromstring(theme.blob)
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    for kind, tag in (("minor", "minorFont"), ("major", "majorFont")):
        fe = root.find(f".//a:{tag}", ns)
        if fe is None:
            continue
        for child, key in (("latin", "latin"), ("ea", "ea")):
            el = fe.find(f"a:{child}", ns)
            if el is not None and el.get("typeface"):
                out[kind][key] = el.get("typeface")
        hang = fe.find("a:font[@script='Hang']", ns)
        if hang is not None and hang.get("typeface"):
            out[kind]["hang"] = hang.get("typeface")
    return out


def _font_from_rfonts(rfonts, theme):
    """w:rFonts 하나 → 글꼴 이름. 한글 문서이므로 eastAsia(한글 글꼴) → ascii 순으로 고른다."""
    if rfonts is None:
        return None
    for attr in ("w:eastAsia", "w:ascii", "w:hAnsi", "w:cs"):
        v = rfonts.get(qn(attr))
        if v:
            return v
    for attr in ("w:eastAsiaTheme", "w:asciiTheme", "w:hAnsiTheme"):
        v = rfonts.get(qn(attr))
        if v:
            kind = "major" if v.startswith("major") else "minor"
            t = theme.get(kind, {})
            name = t.get("hang") or t.get("ea") if "eastAsia" in v or "EastAsia" in v else t.get("latin")
            if name:
                return name
    return None


class FontResolver:
    """런에 실제로 적용되는 글꼴 이름을 구한다: 런 → 문자 스타일 → 문단 스타일 → 문서 기본값."""

    def __init__(self, doc):
        self.theme = _theme_fonts(doc)
        self.default = None
        styles_el = doc.styles.element
        rpr_default = styles_el.find(qn("w:docDefaults") + "/" + qn("w:rPrDefault") + "/" + qn("w:rPr"))
        if rpr_default is not None:
            self.default = _font_from_rfonts(rpr_default.find(qn("w:rFonts")), self.theme)

    def resolve(self, run, par):
        rpr = run._r.rPr
        sources = [rpr]
        if run.style is not None:
            sources += [_style_rpr(s) for s in _style_chain(run.style)]
        if par.style is not None:
            sources += [_style_rpr(s) for s in _style_chain(par.style)]
        for src in sources:
            if src is None:
                continue
            name = _font_from_rfonts(src.find(qn("w:rFonts")), self.theme)
            if name:
                return name
        return self.default


# ───────────────────────── 문단 서식 ─────────────────────────
def _ppr_chain(par):
    """문단에 직접 지정된 pPr → 문단 스타일 pPr (상속 순)"""
    chain = [par._p.pPr]
    if par.style is not None:
        chain += [s.element.pPr for s in _style_chain(par.style) if s.element is not None]
    return [p for p in chain if p is not None]


def paragraph_format(par):
    """docx 문단의 정렬·들여쓰기 → {"alignment": "CENTER", "left_mm":.., "right_mm":.., "first_mm":..}

    first_mm 이 음수면 내어쓰기(hanging)."""
    fmt = {}
    chain = _ppr_chain(par)
    for ppr in chain:                                   # 정렬
        jc = ppr.find(qn("w:jc"))
        if jc is not None and jc.get(qn("w:val")) in ALIGN:
            fmt["alignment"] = ALIGN[jc.get(qn("w:val"))]
            break

    def ind_attr(*names):
        for ppr in chain:
            ind = ppr.find(qn("w:ind"))
            if ind is None:
                continue
            for n in names:
                v = ind.get(qn(n))
                if v not in (None, ""):
                    try:
                        return int(float(v))
                    except ValueError:
                        pass
        return None

    left = ind_attr("w:left", "w:start")
    right = ind_attr("w:right", "w:end")
    first = ind_attr("w:firstLine")
    hanging = ind_attr("w:hanging")
    if left:
        fmt["left_mm"] = round(left / TWIPS_PER_MM, 2)
    if right:
        fmt["right_mm"] = round(right / TWIPS_PER_MM, 2)
    if hanging:
        fmt["first_mm"] = -round(hanging / TWIPS_PER_MM, 2)
    elif first:
        fmt["first_mm"] = round(first / TWIPS_PER_MM, 2)
    fmt.update(paragraph_spacing(par, chain))
    if line_numbers_suppressed(par):
        fmt["no_lnnum"] = True
    return fmt


def line_numbers_suppressed(par):
    """문단 '줄 번호 표시 안 함'(w:suppressLineNumbers, 문단 직접 지정 또는 스타일)"""
    for ppr in _ppr_chain(par):
        el = ppr.find(qn("w:suppressLineNumbers"))
        if el is not None:
            return el.get(qn("w:val")) not in ("0", "false", "off")
    return False


def contextual_spacing(par):
    """'같은 스타일 문단 사이에 간격 넣지 않기'(w:contextualSpacing) 설정 여부"""
    for ppr in _ppr_chain(par):
        el = ppr.find(qn("w:contextualSpacing"))
        if el is not None:
            return el.get(qn("w:val")) not in ("0", "false", "off")
    return False


def _style_id(par):
    return par.style.style_id if par.style is not None else None


def apply_contextual(fmt, par, prev_par, next_par):
    """Word의 contextualSpacing 규칙을 앞/뒤 간격에 미리 반영한다.
    (한글에는 같은 기능이 없어서, 변환할 때 간격 값을 0으로 바꿔 같은 결과를 만든다)"""
    if not contextual_spacing(par):
        return fmt
    fmt = dict(fmt)
    sid = _style_id(par)
    if prev_par is not None and _style_id(prev_par) == sid:
        fmt["before_pt"] = 0.0
    if next_par is not None and _style_id(next_par) == sid:
        fmt["after_pt"] = 0.0
    return fmt


def _doc_default_ppr(par):
    """문서 기본 문단 서식(w:docDefaults/w:pPrDefault/w:pPr)"""
    try:
        styles_el = par.part.styles.element
    except Exception:
        return None
    return styles_el.find(qn("w:docDefaults") + "/" + qn("w:pPrDefault") + "/" + qn("w:pPr"))


def _table_style_ppr(par):
    """표 셀 안 문단이면 표 스타일의 문단 서식(예: Table Grid는 줄 간격 1줄, 뒤 간격 0)"""
    el = par._p.getparent()
    while el is not None and el.tag != qn("w:tbl"):
        el = el.getparent()
    if el is None:
        return []
    tpr = el.find(qn("w:tblPr"))
    sid = tpr.find(qn("w:tblStyle")) if tpr is not None else None
    if sid is None:
        return []
    try:
        style = par.part.styles.element.get_by_id(sid.get(qn("w:val")))
    except Exception:
        style = None
    out = []
    while style is not None:
        if style.pPr is not None:
            out.append(style.pPr)
        based = style.find(qn("w:basedOn"))
        style = par.part.styles.element.get_by_id(based.get(qn("w:val"))) if based is not None else None
    return out


def _para_font_pt(par):
    """문단 첫 글자의 글자 크기(pt). 못 찾으면 문서 기본값, 그것도 없으면 10pt."""
    for r in par.runs:
        if r.font.size is not None:
            return r.font.size.pt
    for st in _style_chain(par.style) if par.style is not None else []:
        if st.font is not None and st.font.size is not None:
            return st.font.size.pt
    try:
        sz = par.part.styles.element.find(
            qn("w:docDefaults") + "/" + qn("w:rPrDefault") + "/" + qn("w:rPr") + "/" + qn("w:sz"))
        if sz is not None:
            return int(sz.get(qn("w:val"))) / 2
    except Exception:
        pass
    return 10.0


def paragraph_spacing(par, chain):
    """줄 간격·문단 앞뒤 간격 → {"line": ("PERCENT", 115) | ("FIXED", pt),
                                  "before_pt": .., "after_pt": ..}

    상속 순서: 문단 직접 지정 → 문단 스타일 → 표 스타일 → 문서 기본값.
    Word는 속성(줄 간격, 앞, 뒤)마다 따로 상속되므로 각각 찾는다."""
    sources = list(chain) + _table_style_ppr(par)
    default = _doc_default_ppr(par)
    if default is not None:
        sources.append(default)

    def attr(name):
        for ppr in sources:
            sp = ppr.find(qn("w:spacing"))
            if sp is not None and sp.get(qn(name)) not in (None, ""):
                return sp
        return None

    out = {}
    sp = attr("w:line")
    if sp is not None:
        line = int(float(sp.get(qn("w:line"))))
        rule = sp.get(qn("w:lineRule"), "auto")
        if rule == "auto":                       # 배수: 240 = 1줄
            out["line"] = ("PERCENT", round(line / 240 * 100))
        elif rule == "exact":                    # 고정: twip(1/20pt)
            out["line"] = ("FIXED", round(line / 20, 1))
        elif rule == "atLeast":
            # '최소' 줄 간격: 한글 AT_LEAST는 확인 가능한 렌더러(LibreOffice)에서 값이 무시되어
            # 결과를 검증할 수 없으므로, 글자 기준 1줄(≈ 글자 크기 × 1.2)보다 크면 고정 간격,
            # 작거나 같으면 1줄(배수 100%)로 바꾼다. 일반 본문에서 Word와 같은 결과가 나온다.
            pt = line / 20
            size = _para_font_pt(par)
            out["line"] = ("FIXED", round(pt, 1)) if pt > size * 1.2 else ("PERCENT", 100)

    for key, name, auto_name in (("before_pt", "w:before", "w:beforeAutospacing"),
                                 ("after_pt", "w:after", "w:afterAutospacing")):
        sp = attr(auto_name)
        if sp is not None and sp.get(qn(auto_name)) in ("1", "true", "on"):
            out[key] = 14.0                      # HTML식 자동 간격 ≈ 14pt
            continue
        sp = attr(name)
        if sp is not None:
            out[key] = round(int(float(sp.get(qn(name)))) / 20, 1)
    # 단위 '줄'(beforeLines/afterLines, 100 = 1줄)로 지정된 경우
    for key, name in (("before_pt", "w:beforeLines"), ("after_pt", "w:afterLines")):
        sp = attr(name)
        if sp is not None:
            out[key] = round(int(float(sp.get(qn(name)))) / 100 * 12, 1)
    return out


# Word '배수' 줄 간격 → 한글 '글자에 따라(%)' 환산 배율.
# Word 1줄은 글꼴 높이의 약 1.2배, 한글 100%는 글자 크기의 1.0배로 줄 높이를 잡기 때문에
# 같은 숫자를 그대로 쓰면 한글 쪽이 촘촘해진다. 그래서 대략 1.2~1.4배를 곱한다
# (Word 1.0줄 ≈ 한글 130%, 1.15줄 ≈ 160%, 1.5줄 ≈ 200%, 2.0줄 ≈ 260%).
WORD_TO_HWP_LINE = 1.35


class ParaStyles:
    """문단 서식 → 한글 paraPr id 캐시 (같은 서식은 하나의 paraPr을 공유)."""

    def __init__(self, doc):
        self.header = doc.parts.headers[0]
        self.cache = {}

    def apply(self, hwpx_par, fmt):
        if not fmt:
            return False
        key = (hwpx_par.para_pr_id_ref, tuple(sorted(fmt.items())))
        if key not in self.cache:
            # 한글 문단 여백 값은 HWPUNIT의 2배로 저장된다(값 200 = 1pt).
            # python-hwpx는 이 배율을 적용하지 않아 그대로 쓰면 절반 크기로 나온다.
            mm = lambda v: round(v * 7200 / 25.4 * 2)
            left = fmt.get("left_mm", 0.0)
            first = fmt.get("first_mm", 0.0)
            if first < 0:
                # Word 내어쓰기: 첫 줄 = 왼쪽 여백 - h, 둘째 줄부터 = 왼쪽 여백
                # 한글 내어쓰기: 첫 줄 = 왼쪽 여백,     둘째 줄부터 = 왼쪽 여백 + h
                left = max(left + first, 0.0)
            margins = {}
            if left:
                margins["left"] = mm(left)
            if "right_mm" in fmt:
                margins["right"] = mm(fmt["right_mm"])
            if first:
                margins["intent"] = mm(first)
            # 문단 앞/뒤 간격: 한글 값 = pt × 200 (여백과 같은 2배 규칙, 렌더링으로 확인)
            if "before_pt" in fmt:
                margins["prev"] = round(fmt["before_pt"] * 200)
            if "after_pt" in fmt:
                margins["next"] = round(fmt["after_pt"] * 200)

            line = fmt.get("line")
            percent = None
            if line and line[0] == "PERCENT":
                percent = max(round(line[1] * WORD_TO_HWP_LINE), 50)
            pid = self.header.ensure_paragraph_format(
                base_para_pr_id=hwpx_par.para_pr_id_ref,
                alignment=fmt.get("alignment"),
                margins=margins or None,
                line_spacing_percent=percent,
            )
            if line and line[0] == "FIXED":
                # 고정/최소 값도 여백과 같은 2배 규칙: pt × 200 (렌더링 측정으로 확인)
                pid = self._with_line_spacing(pid, line[0], round(line[1] * 200))
            if fmt.get("no_lnnum"):
                pid = self._with_attr(pid, "suppressLineNumbers", "1")
            self.cache[key] = pid
        hwpx_par.para_pr_id_ref = self.cache[key]
        return True

    def _with_attr(self, base_id, attr, value):
        """paraPr을 복제해 속성 하나만 바꾼다(예: suppressLineNumbers="1" = 줄 번호 건너뜀)."""
        from copy import deepcopy
        ck = ("attr", base_id, attr, value)
        if ck in self.cache:
            return self.cache[ck]
        container = self.header._para_properties_element(create=True)
        base = next(e for e in container if e.get("id") == str(base_id))
        if base.get(attr) == value:
            return base_id
        new = deepcopy(base)
        new.set(attr, value)
        new_id = self.header._allocate_ref_id(container, base.tag)
        new.set("id", new_id)
        container.append(new)
        if container.get("itemCnt") is not None:
            container.set("itemCnt", str(sum(1 for e in container if e.tag == base.tag)))
        self.header.mark_dirty() if hasattr(self.header, "mark_dirty") else None
        self.cache[ck] = new_id
        return new_id

    def _with_line_spacing(self, base_id, kind, value):
        """고정/최소 줄 간격 paraPr을 만든다. python-hwpx는 %만 지원하므로 복제 후 직접 지정.
        값: pt × 200. 같은 설정은 하나만 만든다."""
        from copy import deepcopy
        ck = ("line", base_id, kind, value)
        if ck in self.cache:
            return self.cache[ck]
        container = self.header._para_properties_element(create=True)
        base = next(e for e in container if e.get("id") == str(base_id))
        new = deepcopy(base)
        found = False
        for ls in new.iter():
            if ls.tag.endswith("}lineSpacing"):
                ls.set("type", kind)
                ls.set("value", str(value))
                ls.set("unit", "HWPUNIT")
                found = True
        if not found:
            return base_id
        new_id = self.header._allocate_ref_id(container, base.tag)
        new.set("id", new_id)
        container.append(new)
        if container.get("itemCnt") is not None:
            container.set("itemCnt", str(sum(1 for e in container if e.tag == base.tag)))
        self.header.mark_dirty() if hasattr(self.header, "mark_dirty") else None
        self.cache[ck] = new_id
        return new_id

# ───────────────────────── 글자 서식 ─────────────────────────
def run_format(run, par, fonts=None):
    """docx 런의 실제 적용 서식(스타일 상속 포함)을 python-hwpx add_run 인자로 만든다."""
    rpr = run._r.rPr
    char_styles = list(_style_chain(run.style)) if run.style is not None else []
    para_styles = list(_style_chain(par.style)) if par.style is not None else []

    def flag(tag):
        v = _rpr_val(rpr, tag)                       # 1) 런에 직접 지정
        if v is not None:
            return v
        for st in char_styles + para_styles:         # 2) 문자 스타일 → 문단 스타일 순
            v = _rpr_val(_style_rpr(st), tag)
            if v is not None:
                return v
        return False

    def child_val(tag):
        for src in [rpr] + [_style_rpr(s) for s in char_styles + para_styles]:
            if src is not None and src.find(qn(tag)) is not None:
                return src.find(qn(tag))
        return None

    fmt = {}
    if flag("w:b"):
        fmt["bold"] = True
    if flag("w:i"):
        fmt["italic"] = True
    if flag("w:strike") or flag("w:dstrike"):
        fmt["strike"] = True

    u = child_val("w:u")
    if u is not None and u.get(qn("w:val"), "single") != "none":
        fmt["underline"] = True

    c = child_val("w:color")
    if c is not None:
        color = _hex(c.get(qn("w:val")))
        if color and color != "#000000":
            fmt["color"] = color

    # 형광펜: w:highlight(이름) 우선, 없으면 런 음영 w:shd(구글 문서가 쓰는 방식)
    hl = rpr.find(qn("w:highlight")) if rpr is not None else None
    if hl is not None and hl.get(qn("w:val")) in HIGHLIGHT:
        fmt["highlight"] = HIGHLIGHT[hl.get(qn("w:val"))]
    else:
        shd = rpr.find(qn("w:shd")) if rpr is not None else None
        fill = _hex(shd.get(qn("w:fill"))) if shd is not None else None
        if fill and fill != "#FFFFFF":
            fmt["highlight"] = fill

    if run.font.size is not None:
        fmt["size"] = round(run.font.size.pt, 1)
    if fonts is not None and (name := fonts.resolve(run, par)):
        fmt["font"] = name
    return fmt


def runs_of(par):
    """하이퍼링크·변경 내용 안의 런까지 포함해 문단의 런을 순서대로."""
    for r in par._p.iter(qn("w:r")):
        yield Run(r, par)


class CharStyles:
    """글자 서식 → 한글 charPr id 캐시.

    python-hwpx의 ensure_run은 지정하지 않은 속성을 '상관없음'으로 보고 비슷한 기존 서식을
    재사용한다(예: 20pt 굵은 제목 서식을 본문 굵게에 재사용). 그래서 모든 속성을 빠짐없이
    명시해 서식마다 정확히 하나의 charPr을 만들고, 그 id를 직접 런에 지정한다.
    """
    # highlight 기본값을 "none"이 아닌 흰색으로 둔다: "none"이면 앞 런의 형광펜이 뒤 글자·문단까지
    # 번져 보이는 현상이 있어(LibreOffice 렌더링에서 확인), 형광펜이 없는 글자도 배경을 명시한다.
    DEFAULT = {"bold": False, "italic": False, "underline": False, "strike": False,
               "color": "#000000", "highlight": "#FFFFFF", "size": 10}

    def __init__(self, doc):
        self.doc, self.cache = doc, {}

    def id_for(self, fmt):
        full = {**self.DEFAULT, **fmt}
        key = tuple(sorted((k, v) for k, v in full.items() if v is not None))
        if key not in self.cache:
            if full.get("font"):
                # 글꼴을 먼저 등록해야 ensure_run이 그 글꼴을 가리키는 새 서식을 만든다
                # (등록하지 않으면 글꼴이 무시되고 기존 서식이 재사용됨)
                self.doc.styles.ensure_font(full["font"])
            else:
                full.pop("font", None)
            self.cache[key] = self.doc.styles.ensure_run(**full)
        return self.cache[key]


class Ctx:
    """변환 중에 함께 쓰는 상태 묶음"""

    def __init__(self, src, out, keep_format, footnote_mode):
        self.src, self.out, self.keep_format = src, out, keep_format
        self.styles = CharStyles(out)
        self.fonts = FontResolver(src) if keep_format else None
        self.paras = ParaStyles(out)
        self.footnotes = read_notes(src, "footnote", self.fonts)
        self.endnotes = read_notes(src, "endnote", self.fonts)
        self.footnote_mode = footnote_mode          # "note" = 한글 각주, "text" = 문서 끝 목록
        self.text_notes = []                       # footnote_mode="text"일 때 모아 둘 각주
        self.text_endnotes = []                    # footnote_mode="text"일 때 모아 둘 미주
        self.sec_starts = {}                       # {본문 순번: 그 문단부터 쓸 한글 구역 번호}
        self.col_changes = {}                      # {본문 순번: (단 설정 layout, 균형)} 같은 쪽 안 단 변경
        self.body_w, self.col_w = {}, {}           # 한글 구역별 본문 폭 · 단 폭 (HWPUNIT)
        self.cur_col_w = None                      # 지금 위치의 단 폭(구역 안 단 변경 반영)
        self.page_fill = page_background(src) if keep_format else None   # 문서 배경색
        self.sec = 0                               # 지금 글을 쓰는 한글 구역 번호
        self.stats = {"heading": 0, "para": 0, "table": 0, "nested_table": 0, "image": 0,
                      "merged": 0, "styled_runs": 0, "para_fmt": 0, "footnote": 0,
                      "endnote": 0, "header": 0, "footer": 0}


# ───────────────────────── 각주 ─────────────────────────
NOTE_IMG_MAX_MM = 150   # 각주·미주 안 이미지 최대 폭(mm)
NOTE_SIZE = 9          # 각주·미주 기본 글자 크기(pt): 한글 각주/미주 스타일 기본값과 같음
NOTE_SKIP = (qn("w:footnoteRef"), qn("w:endnoteRef"), qn("w:separator"), qn("w:continuationSeparator"))


def note_images(p_el, part):
    """각주·미주 문단 안 인라인 이미지 → [(bytes, ext, 폭 mm, 높이 mm)]
    이미지 관계(r:embed)는 본문이 아니라 footnotes.xml/endnotes.xml 파트 기준으로 찾는다."""
    out = []
    for blip in p_el.iter(qn("a:blip")):
        rid = blip.get(qn("r:embed"))
        if not rid or rid not in part.related_parts:
            continue
        img = part.related_parts[rid]
        ext = IMG_EXT.get(img.content_type)
        if not ext:
            continue
        w_mm = h_mm = None
        drawing = blip
        while drawing is not None and drawing.tag not in (qn("wp:inline"), qn("wp:anchor")):
            drawing = drawing.getparent()
        if drawing is not None and (e := drawing.find(qn("wp:extent"))) is not None:
            w_mm, h_mm = int(e.get("cx")) / EMU_PER_MM, int(e.get("cy")) / EMU_PER_MM
            if w_mm > NOTE_IMG_MAX_MM:                     # 쪽 폭을 넘지 않게 비율 유지 축소
                h_mm, w_mm = h_mm * NOTE_IMG_MAX_MM / w_mm, NOTE_IMG_MAX_MM
        out.append((img.blob, ext, w_mm, h_mm))
    return out


def read_notes(doc, kind="footnote", fonts=None):
    """word/footnotes.xml 또는 word/endnotes.xml → {주석 id: [블록, …]}

    블록 (원본 순서 그대로)
      {"type": "p", "segs": [(글자, 서식 dict), …], "images": [(bytes, ext, 폭mm, 높이mm), …]}
      {"type": "tbl", "el": w:tbl 요소}        ← 본문 표와 같은 build_table로 변환
    글자 서식은 본문과 같은 run_format으로 읽는다(런 → 문자 스타일 → 각주/미주 문단 스타일).
    맨 앞 번호 표시 런(footnoteRef/endnoteRef)과 그 뒤 공백은 빼고, 글자·이미지가 없는 문단은 건너뛴다."""
    notes = {}
    part = next((p for p in doc.part.package.parts if p.partname.endswith(f"/{kind}s.xml")), None)
    if part is None:
        return notes
    # python-docx의 parse_xml로 읽어야 w:r/w:p가 CT_R/CT_P가 되어 Run.text 등이 동작한다
    from docx.oxml import parse_xml
    root = parse_xml(part.blob)
    for fn in root.iter(qn(f"w:{kind}")):
        if fn.get(qn("w:type")) in ("separator", "continuationSeparator", "continuationNotice"):
            continue
        blocks, first_text = [], True
        for child in fn.iterchildren():                # 직속 자식만: 표 안 문단은 표로 처리
            if child.tag == qn("w:tbl"):
                blocks.append({"type": "tbl", "el": child})
                continue
            if child.tag != qn("w:p"):
                continue
            # 각주 파트는 python-docx가 직접 감싸 주지 않는다. 스타일은 문서 공용이므로
            # 본문(_body)을 부모로 두면 par.style(각주 문단 스타일)이 그대로 풀린다.
            par = Paragraph(child, doc._body)
            segs = []
            for r in child.iter(qn("w:r")):
                if any(r.find(t) is not None for t in NOTE_SKIP):
                    continue
                run = Run(r, par)
                text = (run.text or "").replace("\t", "    ")
                if text:
                    segs.append([text, run_format(run, par, fonts)])
            if segs and first_text:
                segs[0][0] = segs[0][0].lstrip()        # 번호 뒤 공백 정리
            segs = [(t, f) for t, f in segs if t]
            images = note_images(child, part)
            if segs or images:
                if segs:
                    first_text = False
                blocks.append({"type": "p", "segs": segs, "images": images})
        if blocks:
            notes[fn.get(qn("w:id"))] = blocks
    return notes


def note_plain(blocks):
    """블록 → 평문 한 줄 (라이브러리에 각주 틀을 만들 때만 씀)"""
    return " ".join("".join(t for t, _ in b["segs"]) for b in blocks if b["type"] == "p")


NOTE_REFS = (qn("w:footnoteReference"), qn("w:endnoteReference"))


def has_note_ref(el):
    return any(el.find(".//" + tag) is not None for tag in NOTE_REFS)


# ───────────────────────── 런 쓰기 ─────────────────────────
def write_runs(ctx, dst_par, src_par, prefix="", base=None):
    """src_par(docx)의 런들을 dst_par(hwpx)에 서식과 함께 옮긴다. 각주 참조도 이 위치에 넣는다."""
    base = base or {}
    styles = ctx.styles
    if prefix:
        dst_par.add_run(prefix, char_pr_id_ref=styles.id_for(base))
    for run in runs_of(src_par):
        ref = run._r.find(qn("w:footnoteReference"))
        if ref is not None:
            add_footnote(ctx, dst_par, ref.get(qn("w:id")))
            continue
        ref = run._r.find(qn("w:endnoteReference"))
        if ref is not None:
            add_endnote(ctx, dst_par, ref.get(qn("w:id")))
            continue
        text = run.text
        if not text:
            continue
        fmt = dict(base)
        if ctx.keep_format:
            f = run_format(run, src_par, ctx.fonts)
            if base.get("size"):          # 제목은 제목 크기를 우선
                f.pop("size", None)
            fmt.update(f)
            if any(k in f for k in ("bold", "italic", "underline", "strike", "color", "highlight")):
                ctx.stats["styled_runs"] += 1
        dst_par.add_run(text.replace("\t", "    "), char_pr_id_ref=styles.id_for(fmt))


def note_fmt(ctx, fmt):
    """각주·미주 조각 서식 → charPr 인자. 서식 끄기 옵션이면 기본 크기만, 크기 미지정이면 9pt."""
    if not ctx.keep_format:
        return {"size": NOTE_SIZE}
    f = dict(fmt)
    f.setdefault("size", NOTE_SIZE)
    return f


def _seg_runs(ctx, p_el, segs, hp):
    """hp:p 요소에 서식 조각마다 hp:run을 붙인다."""
    from lxml import etree
    for text, fmt in segs:
        r = etree.SubElement(p_el, f"{hp}run", {"charPrIDRef": str(ctx.styles.id_for(note_fmt(ctx, fmt)))})
        etree.SubElement(r, f"{hp}t").text = text
        if ctx.keep_format and any(k in fmt for k in ("bold", "italic", "underline", "strike", "color", "highlight")):
            ctx.stats["note_styled"] = ctx.stats.get("note_styled", 0) + 1


def _add_note_picture(ctx, para, blob, ext, w_mm, h_mm):
    """문단(HwpxOxmlParagraph)에 이미지를 글자처럼 넣는다. 크기를 모르면 폭 60mm, PNG면 비율 유지."""
    from hwpx._document.media import add_image, _png_dimensions
    bin_id = str(add_image(ctx.out, blob, ext))
    w = round((w_mm or 60) * 7200 / 25.4)
    if h_mm:
        h = round(h_mm * 7200 / 25.4)
    else:
        dims = _png_dimensions(blob)
        h = round(w * dims[1] / dims[0]) if dims else w
    para.add_picture(bin_id, width=w, height=h)
    ctx.stats["note_image"] = ctx.stats.get("note_image", 0) + 1


def fill_note(ctx, note, blocks):
    """python-hwpx가 만든 한글 각주/미주(평문 한 런)에 서식 조각·이미지·표를 채운다.

    라이브러리 구조(한컴 실제 파일 모양):
      hp:footNote/hp:subList/hp:p/hp:run[ctrl(autoNum) + t]
    - 첫 문단: autoNum 런(번호)만 두고, 글자는 서식 조각마다 새 런으로 붙인다.
    - 다음 블록: 첫 문단의 문단 모양·스타일을 복사한 새 hp:p를 subList에 차례로 추가.
      이미지는 그 문단 안에 글자처럼(treatAsChar), 표는 새 문단 안에 본문 표와 같은 방식으로 넣는다."""
    from lxml import etree
    from hwpx.oxml import HwpxOxmlParagraph
    el = note.element
    hp = el.tag[: el.tag.index("}") + 1]
    section = note.paragraph.section
    first_p = el.find(f"{hp}subList/{hp}p")
    num_run = first_p.find(f"{hp}run")
    for t in num_run.findall(f"{hp}t"):
        num_run.remove(t)
    num_run.set("charPrIDRef", str(ctx.styles.id_for({"size": NOTE_SIZE})))

    state = {"prev": first_p, "used_first": False}

    def next_p():
        """다음 블록을 쓸 hp:p. 첫 블록이 글자 문단이면 번호가 있는 첫 문단을 그대로 쓴다."""
        p_el = etree.Element(f"{hp}p", dict(first_p.attrib))
        state["prev"].addnext(p_el)
        state["prev"] = p_el
        return p_el

    for i, b in enumerate(blocks):
        if b["type"] == "p":
            if i == 0:
                p_el = first_p
            else:
                p_el = next_p()
            _seg_runs(ctx, p_el, b["segs"], hp)
            for blob, ext, w_mm, h_mm in b["images"]:
                _add_note_picture(ctx, HwpxOxmlParagraph(p_el, section), blob, ext, w_mm, h_mm)
        else:                                         # 표
            host = HwpxOxmlParagraph(next_p(), section)
            made = build_table(ctx, b["el"], ctx.src._body, host.add_table, base={"size": NOTE_SIZE})
            if made is not None:
                ctx.stats["note_table"] = ctx.stats.get("note_table", 0) + 1
    section.mark_dirty()


def add_footnote(ctx, dst_par, note_id):
    _add_note(ctx, dst_par, note_id, "footnote")


def add_endnote(ctx, dst_par, note_id):
    """미주: 한글 미주(문서 끝에 모임)로 넣거나, text 모드면 본문에 [i] 표시 후 문서 끝 '미주' 목록"""
    _add_note(ctx, dst_par, note_id, "endnote")


def _add_note(ctx, dst_par, note_id, kind):
    blocks = (ctx.footnotes if kind == "footnote" else ctx.endnotes).get(note_id)
    if not blocks:
        return
    ctx.stats[kind] += 1
    if ctx.footnote_mode == "text":
        items = ctx.text_notes if kind == "footnote" else ctx.text_endnotes
        items.append(blocks)
        n = len(items)
        mark, color = (f"[{n}]", "#1F4E79") if kind == "footnote" else (f"[{to_roman(n)}]", "#7F3F00")
        dst_par.add_run(mark, char_pr_id_ref=ctx.styles.id_for({"size": 8, "color": color}))
        return
    note = (dst_par.add_footnote if kind == "footnote" else dst_par.add_endnote)(note_plain(blocks) or " ")
    fill_note(ctx, note, blocks)


def write_note_list(ctx, add_par, make_table):
    """footnote_mode="text": 문서 끝에 각주([1]) · 미주([i]) 목록 — 서식·이미지·표 유지"""
    from hwpx.oxml import HwpxOxmlParagraph
    for title, items, label in (("각주", ctx.text_notes, str), ("미주", ctx.text_endnotes, to_roman)):
        if not items:
            continue
        dp = add_par()
        dp.add_run(title, char_pr_id_ref=ctx.styles.id_for({"bold": True, "size": 11}))
        for i, blocks in enumerate(items, 1):
            numbered = False
            for b in blocks:
                if b["type"] == "tbl":
                    if not numbered:                  # 표로 시작하는 주석: 번호만 먼저
                        add_par().add_run(f"[{label(i)}]", char_pr_id_ref=ctx.styles.id_for({"size": NOTE_SIZE}))
                        numbered = True
                    if build_table(ctx, b["el"], ctx.src._body, make_table, base={"size": NOTE_SIZE}) is not None:
                        ctx.stats["note_table"] = ctx.stats.get("note_table", 0) + 1
                    continue
                dp = add_par()
                prefix = f"[{label(i)}] " if not numbered else "      "
                numbered = True
                dp.add_run(prefix, char_pr_id_ref=ctx.styles.id_for({"size": NOTE_SIZE}))
                for text, fmt in b["segs"]:
                    dp.add_run(text, char_pr_id_ref=ctx.styles.id_for(note_fmt(ctx, fmt)))
                for blob, ext, w_mm, h_mm in b["images"]:
                    _add_note_picture(ctx, dp, blob, ext, w_mm, h_mm)


def to_roman(n):
    """미주 번호: 각주([1], [2])와 구별되도록 i, ii, iii … (Word 미주 기본 번호 형식과 같음)"""
    vals = [(10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]
    out = ""
    for v, sym in vals:
        while n >= v:
            out += sym
            n -= v
    return out


def write_paragraph(ctx, dst_par, src_par, base=None, prev_par=None, next_par=None):
    """문단 하나: 글머리 기호 + 런 + 정렬/들여쓰기/간격"""
    prefix = "• " if is_list(src_par) else ""
    write_runs(ctx, dst_par, src_par, prefix=prefix, base=base)
    if ctx.keep_format:
        fmt = apply_contextual(paragraph_format(src_par), src_par, prev_par, next_par)
        if ctx.paras.apply(dst_par, fmt):
            ctx.stats["para_fmt"] += 1
    else:
        fmt = {"no_lnnum": True} if line_numbers_suppressed(src_par) else {}
        ctx.paras.apply(dst_par, fmt)                 # 서식을 꺼도 줄 번호 건너뛰기는 유지
    if fmt.get("no_lnnum"):
        ctx.stats["lnnum_skip"] = ctx.stats.get("lnnum_skip", 0) + 1


# ───────────────────────── 이미지 ─────────────────────────
def paragraph_images(par, src):
    """문단 안 인라인 이미지 → [(bytes, ext, width_mm)]"""
    out = []
    for blip in par._p.iter(qn("a:blip")):
        rid = blip.get(qn("r:embed"))
        if not rid or rid not in src.part.related_parts:
            continue
        part = src.part.related_parts[rid]
        ext = IMG_EXT.get(part.content_type)
        if not ext:
            continue
        width_mm = None
        drawing = blip
        while drawing is not None and drawing.tag not in (qn("wp:inline"), qn("wp:anchor")):
            drawing = drawing.getparent()
        if drawing is not None:
            ext_el = drawing.find(qn("wp:extent"))
            if ext_el is not None:
                width_mm = min(int(ext_el.get("cx")) / EMU_PER_MM, 160)  # A4 본문 폭 이내
        out.append((part.blob, ext, width_mm))
    return out


# ───────────────────────── 표 ─────────────────────────
def table_grid(tbl):
    """docx 표 XML → (행 수, 열 수, 병합 정보가 담긴 셀 목록)

    셀 목록 항목: {"r","c","rowspan","colspan","tc"}  (병합된 셀은 왼쪽 위 칸 하나로 표현)
    """
    t = tbl._tbl
    grid = t.find(qn("w:tblGrid"))
    n_cols = len(grid.findall(qn("w:gridCol"))) if grid is not None else 0
    cells, owner = [], {}            # owner[(r,c)] = 그 칸을 덮는 셀
    rows = t.findall(qn("w:tr"))
    for r, tr in enumerate(rows):
        trpr = tr.find(qn("w:trPr"))
        gb = trpr.find(qn("w:gridBefore")) if trpr is not None else None
        c = int(gb.get(qn("w:val"))) if gb is not None else 0
        for tc in tr.findall(qn("w:tc")):
            tcpr = tc.find(qn("w:tcPr"))
            span_el = tcpr.find(qn("w:gridSpan")) if tcpr is not None else None
            span = int(span_el.get(qn("w:val"))) if span_el is not None else 1
            vm = tcpr.find(qn("w:vMerge")) if tcpr is not None else None
            if vm is not None and vm.get(qn("w:val")) != "restart" and (r - 1, c) in owner:
                cell = owner[(r - 1, c)]          # 세로 병합 이어짐 → 위 셀 확장
                cell["rowspan"] = r - cell["r"] + 1
            else:
                cell = {"r": r, "c": c, "rowspan": 1, "colspan": span, "tc": tc}
                cells.append(cell)
            for k in range(span):
                owner[(r, c + k)] = cell
            c += span
        n_cols = max(n_cols, c)
    return len(rows), n_cols, cells


def cell_fill(tc):
    tcpr = tc.find(qn("w:tcPr"))
    shd = tcpr.find(qn("w:shd")) if tcpr is not None else None
    fill = _hex(shd.get(qn("w:fill"))) if shd is not None else None
    return fill if fill and fill != "#FFFFFF" else None


def _cell_sz(cell_el):
    return next((e for e in cell_el.iter() if e.tag.endswith("}cellSz")), None)


def _direct_cells(tbl_el):
    """이 표에 직접 속한 셀(안쪽 표의 셀 제외) → [(tc 요소, rowAddr, rowSpan)]"""
    out = []
    for tr in tbl_el:
        if not tr.tag.endswith("}tr"):
            continue
        for tc in tr:
            if not tc.tag.endswith("}tc"):
                continue
            addr = next(e for e in tc if e.tag.endswith("}cellAddr"))
            span = next(e for e in tc if e.tag.endswith("}cellSpan"))
            out.append((tc, int(addr.get("rowAddr")), int(span.get("rowSpan"))))
    return out


def table_height(tbl_el):
    """행별 가장 높은 셀 높이의 합(HWPUNIT)"""
    rows = {}
    for tc, r, rs in _direct_cells(tbl_el):
        if rs == 1:
            rows[r] = max(rows.get(r, 0), int(_cell_sz(tc).get("height")))
    return sum(rows.values())


def _cell_text_height(tc_xml, col_width_hwp):
    """docx 셀 내용으로 필요한 높이(HWPUNIT)를 어림한다: 문단별 줄 수 × 줄 높이 + 문단 간격."""
    total = 0
    usable = max(col_width_hwp - 1020, 2000)          # 셀 좌우 안쪽 여백(510×2) 제외
    for p in tc_xml.findall(qn("w:p")):              # 직속 문단만 (안쪽 표는 따로 계산)
        text = "".join(t.text or "" for t in p.iter(qn("w:t")))
        size_pt = 10.0
        sz = p.find(".//" + qn("w:sz"))
        if sz is not None:
            size_pt = int(sz.get(qn("w:val"))) / 2
        # 한글 1글자 ≈ 글자 크기, 영문/숫자 ≈ 0.55배
        width = sum(size_pt * (1.0 if ord(ch) > 0x2E80 else 0.55) for ch in text) * 100
        lines = max(1, -(-int(width) // int(usable)))
        total += int(lines * size_pt * 1.6 * 100)    # 줄 높이 ≈ 글자 × 160%
    return max(total, 1000)


def compact_rows(tbl_el, block):
    """행 높이를 내용에 맞춘다.

    python-hwpx 기본 행 높이(약 12.7mm)는 한 줄짜리 행보다 훨씬 커서 표가 길어지고,
    반대로 너무 작게 두면 일부 프로그램이 행을 충분히 늘리지 않아 글자가 잘린다.
    그래서 docx의 행 높이(w:trHeight)와 셀 글자 양으로 계산한 높이 중 큰 값을 넣는다."""
    PAD = 282 + 400                                   # 셀 위아래 안쪽 여백 + 여유
    trs = block._tbl.findall(qn("w:tr"))
    cols = {}                                         # (r,c) → (tc xml, 열 너비)
    hw_cells = _direct_cells(tbl_el)
    for tc, r, rs in hw_cells:
        addr = next(e for e in tc if e.tag.endswith("}cellAddr"))
        cols[(r, int(addr.get("colAddr")))] = int(_cell_sz(tc).get("width"))

    heights = []
    for r, tr in enumerate(trs):
        trpr = tr.find(qn("w:trPr"))
        h = trpr.find(qn("w:trHeight")) if trpr is not None else None
        preset = int(int(h.get(qn("w:val"), "0")) / 20 * 100) if h is not None else 0  # twip → HWPUNIT
        need, c = 0, 0
        for tc in tr.findall(qn("w:tc")):
            tcpr = tc.find(qn("w:tcPr"))
            span_el = tcpr.find(qn("w:gridSpan")) if tcpr is not None else None
            span = int(span_el.get(qn("w:val"))) if span_el is not None else 1
            vm = tcpr.find(qn("w:vMerge")) if tcpr is not None else None
            if vm is None or vm.get(qn("w:val")) == "restart":
                width = sum(cols.get((r, c + k), 0) for k in range(span)) or 10000
                if vm is None:                       # 세로 병합 셀은 여러 행에 나뉘므로 제외
                    need = max(need, _cell_text_height(tc, width) + PAD)
            c += span
        heights.append(max(preset, need, 1500))

    for tc, r, rs in hw_cells:
        sz = _cell_sz(tc)
        if sz is not None and r < len(heights):
            sz.set("height", str(sum(heights[r:r + rs])))
    tsz = next((e for e in tbl_el if e.tag.endswith("}sz")), None)
    if tsz is not None:
        tsz.set("height", str(sum(heights)))


def grow_cell_for_nested(tbl_el, cell_el, n_text_paras, nested_heights):
    """안쪽 표가 잘리지 않도록 바깥 셀(과 같은 행 셀)과 바깥 표 높이를 늘린다.

    한글 셀 높이는 고정값으로 저장되므로, 안쪽 표가 들어간 셀은 직접 키워 줘야 한다."""
    need = n_text_paras * 1700 + sum(nested_heights) + 600 * len(nested_heights) + 600
    sz = _cell_sz(cell_el)
    if need <= int(sz.get("height")):
        return
    row = int(next(e for e in cell_el if e.tag.endswith("}cellAddr")).get("rowAddr"))
    for tc, r, rs in _direct_cells(tbl_el):
        if r == row:
            _cell_sz(tc).set("height", str(need))
    tsz = next(e for e in tbl_el if e.tag.endswith("}sz"))
    tsz.set("height", str(table_height(tbl_el)))


def build_table(ctx, docx_tbl_el, parent, make_table, base=None):
    """docx 표(w:tbl) → 한글 표. make_table(rows, cols)가 빈 표를 만들어 돌려준다.
    셀 안에 표가 있으면 재귀적으로 안쪽 표를 만든다.
    base: 셀 글자에 공통으로 줄 서식(각주·미주 안 표는 {"size": 9})."""
    block = Table(docx_tbl_el, parent)
    n_rows, n_cols, cells = table_grid(block)
    if not n_rows or not n_cols:
        return None
    tbl = make_table(n_rows, n_cols)

    # 1) 병합 먼저 (병합 후에 내용을 채워야 왼쪽 위 칸에 들어간다)
    for cell in cells:
        if cell["rowspan"] > 1 or cell["colspan"] > 1:
            tbl.merge_cells(cell["r"], cell["c"],
                            cell["r"] + cell["rowspan"] - 1, cell["c"] + cell["colspan"] - 1)
            ctx.stats["merged"] += 1

    compact_rows(tbl.element, block)                 # 병합 후 행 높이를 내용에 맞춤

    # 2) 내용과 배경색
    for cell in cells:
        r, c, tc = cell["r"], cell["c"], cell["tc"]
        fill = cell_fill(tc) if ctx.keep_format else None
        if fill:
            tbl.set_cell_shading(r, c, fill)
        cell_base = dict(base or {})
        if fill:
            cell_base["highlight"] = fill                    # 배경색 셀: 글자 배경도 셀 색으로
        dst_cell = tbl.cell(r, c)
        first = True
        n_text, nested_heights = 0, []
        for child in tc:                                   # 셀의 직속 자식만 (안쪽 표는 재귀)
            if child.tag == qn("w:p"):
                sp = Paragraph(child, block)
                prv, nxt = child.getprevious(), child.getnext()
                prev_par = Paragraph(prv, block) if prv is not None and prv.tag == qn("w:p") else None
                next_par = Paragraph(nxt, block) if nxt is not None and nxt.tag == qn("w:p") else None
                if not sp.text.strip() and not has_note_ref(child):
                    continue
                dp = dst_cell.paragraphs[0] if first else dst_cell.add_paragraph("")
                if first:
                    dp.clear_text()
                    first = False
                write_paragraph(ctx, dp, sp, base=cell_base, prev_par=prev_par, next_par=next_par)
                n_text += 1
            elif child.tag == qn("w:tbl"):
                inner = build_table(ctx, child, block, dst_cell.add_table, base=base)
                if inner is not None:
                    first = False
                    ctx.stats["nested_table"] += 1
                    nested_heights.append(table_height(inner.element))
        if nested_heights:
            grow_cell_for_nested(tbl.element, dst_cell.element, n_text, nested_heights)
    return tbl


# ───────────────────────── 머리글 · 바닥글 ─────────────────────────
NBSP = "\u00a0"


def _field_code(el):
    """w:fldSimple / w:instrText 의 필드 코드 → "PAGE" | "NUMPAGES" | None"""
    instr = el.get(qn("w:instr")) if el.tag == qn("w:fldSimple") else (el.text or "")
    word = instr.strip().split(" ")[0].upper() if instr.strip() else ""
    return word if word in ("PAGE", "NUMPAGES", "SECTIONPAGES") else None


def hf_paragraph_items(ctx, par):
    """머리글/바닥글 문단 → [("text", 글자, 서식) | ("page", "PAGE"|"NUMPAGES")]

    쪽 번호 필드(단순 필드 w:fldSimple, 복합 필드 w:fldChar+w:instrText) 모두 인식한다."""
    items, in_field, field_shown = [], False, False
    for el in par._p.iter():
        if el.tag == qn("w:fldSimple"):
            code = _field_code(el)
            if code:
                items.append(("page", code))
                field_shown = True              # 필드 안 표시용 글자(캐시된 쪽 번호)는 건너뜀
        elif el.tag == qn("w:fldChar"):
            kind = el.get(qn("w:fldCharType"))
            if kind == "begin":
                in_field, field_shown = True, False
            elif kind == "end":
                in_field = False
        elif el.tag == qn("w:instrText") and in_field:
            code = _field_code(el)
            if code and not field_shown:
                items.append(("page", code))
                field_shown = True
        elif el.tag == qn("w:r"):
            if in_field or el.getparent().tag == qn("w:fldSimple"):
                continue                          # 필드 결과 글자(예: 캐시된 "1")는 쓰지 않음
            run = Run(el, par)
            if run.text:
                fmt = run_format(run, par, ctx.fonts) if ctx.keep_format else {}
                items.append(("text", run.text.replace("\t", "    "), fmt))
    return items


def hf_blocks(ctx, hf):
    """docx 머리글/바닥글 → 문단 목록 [(정렬, 항목들)]. 머리글 안 표는 칸을 ' | '로 이어 한 줄로."""
    out = []
    for child in hf._element:
        if child.tag == qn("w:p"):
            par = Paragraph(child, hf)
            items = hf_paragraph_items(ctx, par)
            if items:
                align = paragraph_format(par).get("alignment") if ctx.keep_format else None
                out.append((align, items))
        elif child.tag == qn("w:tbl"):
            for tr in child.iter(qn("w:tr")):
                texts = ["".join(t.text or "" for t in tc.iter(qn("w:t"))).strip()
                         for tc in tr.findall(qn("w:tc"))]
                line = " | ".join(t for t in texts if t)
                if line:
                    out.append((None, [("text", line, {})]))
    return out


def _resync_header_footer_controls(props, tag):
    """머리글/바닥글 컨트롤을 쪽 종류(BOTH/ODD/EVEN)마다 하나씩 다시 만든다.

    python-hwpx는 머리글을 추가할 때마다 같은 종류(header)의 컨트롤을 모두 지우고 마지막 것만
    남겨서, 홀수·짝수 머리글을 함께 넣으면 하나가 사라진다. 그래서 구역 정의에 있는 머리글
    전부를 컨트롤로 다시 복사한다."""
    from copy import deepcopy
    HP = props.element.tag.split("}")[0] + "}"
    run = props._header_footer_control_run()
    for c in list(run.findall(HP + "ctrl")):
        if c.find(HP + tag) is not None:
            run.remove(c)
    for el in props.element.findall(HP + tag):
        c = run.makeelement(HP + "ctrl", {})
        run.append(c)
        c.append(deepcopy(el))
    props.section.mark_dirty()


def write_header_footer(ctx, kind, blocks, page_type="BOTH", sec_index=0):
    """한글 머리글/바닥글 작성. page_type: BOTH(모든 쪽) / ODD(홀수 쪽) / EVEN(짝수 쪽)

    python-hwpx로 문단 틀을 만든 뒤, 한글이 실제로 쓰는 머리글 복사본(구역 정의 + 컨트롤)
    모두에 글자 서식과 쪽 번호 종류를 맞춰 넣는다."""
    if not blocks:
        return
    sec = ctx.out.sections[sec_index]
    spec, run_plan = [], []                 # run_plan: 생성될 hp:run 순서대로 (종류, 값)
    for align, items in blocks:
        children = []
        for item in items:
            if item[0] == "text":
                text = item[1]
                # 머리글/바닥글에서는 앞뒤 공백이 사라지므로 줄바꿈 없는 공백으로 바꾼다
                text = re.sub(r"^ +| +$", lambda m: NBSP * len(m.group(0)), text)
                children.append({"type": "run", "text": text})
                run_plan.append(("text", item[2]))
            else:
                children.append({"type": "page_number", "page_number": "page"})
                run_plan += [("autonum", item[1]), ("pagenum", None)]
        spec.append({"align": align or "LEFT", "children": children})

    props = sec.properties
    if kind == "header":
        hf = props.set_header_content(spec, page_type=page_type)
    else:
        hf = props.set_footer_content(spec, page_type=page_type)
    _resync_header_footer_controls(props, kind)
    hid = hf.element.get("id")
    tag = "}header" if kind == "header" else "}footer"
    for copy_el in [e for e in sec.element.iter() if e.tag.endswith(tag) and e.get("id") == hid]:
        runs = [r for r in copy_el.iter() if r.tag.endswith("}run")]
        for run_el, (what, val) in zip(runs, run_plan):
            if what == "text":
                run_el.set("charPrIDRef", ctx.styles.id_for(val))
            elif what == "autonum":
                auto = next(e for e in run_el.iter() if e.tag.endswith("}autoNum"))
                # NUMPAGES(전체 쪽 수) → 한글 TOTAL_PAGE
                auto.set("numType", "PAGE" if val == "PAGE" else "TOTAL_PAGE")
            elif what == "pagenum":
                # pageNum 컨트롤은 쪽 번호를 쪽 모서리에 따로 한 번 더 그리므로 제거하고,
                # 글자 흐름 안의 autoNum만 남긴다 ("- 1 -", "1 / 3" 같은 모양을 유지)
                if run_el.getparent() is not None:
                    run_el.getparent().remove(run_el)
    ctx.stats[kind] += 1


def _docx_hf_variants(section, kind):
    """docx 구역의 머리글(또는 바닥글) 3종: 기본(default), 짝수(even), 첫 쪽(first).
    각 항목은 내용이 있을 때만 python-docx 객체, 없으면 None."""
    base = section.header if kind == "header" else section.footer
    even = section.even_page_header if kind == "header" else section.even_page_footer
    first = section.first_page_header if kind == "header" else section.first_page_footer
    pick = lambda hf: None if hf.is_linked_to_previous else hf
    return pick(base), pick(even), pick(first)


def _set_visibility(ctx, sec_index=0, **flags):
    """구역 정의의 hp:visibility(첫 쪽 머리글 감추기 등) 속성 설정"""
    sec = ctx.out.sections[sec_index]
    vis = next((e for e in sec.properties.element.iter() if e.tag.endswith("}visibility")), None)
    if vis is None:
        return False
    for k, v in flags.items():
        vis.set(k, "1" if v else "0")
    sec.properties.section.mark_dirty()
    return True


def _break_kind(el):
    if el.tag == qn("w:br") and el.get(qn("w:type")) == "page":
        return "page"
    if el.tag == qn("w:lastRenderedPageBreak"):
        return "rendered"
    return None


def page_breaks(p_el):
    """문단의 쪽 나눔 → (앞에서 나눔, 뒤에서 나눔). 문단 중간의 쪽 나눔은 문단 끝으로 근사."""
    before = False
    ppr = p_el.find(qn("w:pPr"))
    if ppr is not None:
        pbb = ppr.find(qn("w:pageBreakBefore"))
        before = pbb is not None and pbb.get(qn("w:val")) not in ("0", "false")
    text_seen, after = False, False
    for el in p_el.iter():
        if el.tag == qn("w:t") and (el.text or "").strip():
            text_seen = True
        elif el.tag == qn("w:br") and el.get(qn("w:type")) == "page":
            if text_seen:
                after = True
            else:
                before = True
    return before, after


def first_page_end(children):
    """원본 docx에서 첫 쪽이 끝나는 위치(본문 자식 순번). 그 앞까지가 첫 쪽.

    1) 쪽 나눔(w:br type=page, '앞에서 쪽 나누기')  2) Word가 저장 때 기록한 쪽 경계
    (w:lastRenderedPageBreak)  3) 구역 나눔 순서로 찾는다. 못 찾으면 None."""
    if hasattr(children, "iterchildren"):      # body 요소를 넘긴 경우
        children = [c for c in children.iterchildren() if c.tag in (qn("w:p"), qn("w:tbl"))]
    for i, child in enumerate(children):
        if child.tag == qn("w:tbl"):
            if any(_break_kind(e) for e in child.iter()) and i > 0:
                return i
            continue
        ppr = child.find(qn("w:pPr"))
        if ppr is not None:
            pbb = ppr.find(qn("w:pageBreakBefore"))
            if pbb is not None and pbb.get(qn("w:val")) not in ("0", "false") and i > 0:
                return i
        before = ""
        for el in child.iter():
            if el.tag == qn("w:t"):
                before += el.text or ""
            elif _break_kind(el):
                if before.strip():
                    return i + 1          # 문단 중간 경계 → 문단 단위로 근사(문단 끝에서 나눔)
                if i > 0:
                    return i
        if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
            return i + 1
    return None


def _section_ranges(body):
    """본문 자식(문단·표) 순번을 Word 구역별 범위로 나눈다.
    Word는 구역 끝 문단의 pPr에 sectPr을 두고, 마지막 구역은 body 직속 sectPr로 끝난다.
    → [(시작, 끝(미포함), sectPr), …]  (python-docx의 doc.sections 순서와 같다)"""
    children = [c for c in body.iterchildren() if c.tag in (qn("w:p"), qn("w:tbl"))]
    ranges, start = [], 0
    for i, c in enumerate(children):
        ppr = c.find(qn("w:pPr")) if c.tag == qn("w:p") else None
        sp = ppr.find(qn("w:sectPr")) if ppr is not None else None
        if sp is not None:
            ranges.append((start, i + 1, sp))
            start = i + 1
    ranges.append((start, len(children), body.find(qn("w:sectPr"))))
    return children, ranges


def _sect_info(sect_pr):
    """구역 시작 방식(nextPage/continuous/oddPage/evenPage)과 쪽 번호 새로 시작 값"""
    if sect_pr is None:
        return "nextPage", None
    t = sect_pr.find(qn("w:type"))
    start_type = t.get(qn("w:val"), "nextPage") if t is not None else "nextPage"
    pg = sect_pr.find(qn("w:pgNumType"))
    restart = pg.get(qn("w:start")) if pg is not None else None
    return start_type, int(restart) if restart and restart.isdigit() else None


def _resolve_hf_blocks(ctx, sections):
    """구역마다 머리글·바닥글 3종(기본, 짝수, 첫 쪽)의 내용을 정한다.
    '이전 구역에 연결'(linkToPrevious = 해당 headerReference 없음)이면 앞 구역 것을 물려받는다.
    python-docx의 연결 해석은 첫 구역에서 빈 정의를 새로 만들어 버리므로 직접 따라간다."""
    resolved, prev = [], {"header": ([], [], []), "footer": ([], [], [])}
    for section in sections:
        cur = {}
        for kind in ("header", "footer"):
            own = _docx_hf_variants(section, kind)
            cur[kind] = tuple(hf_blocks(ctx, hf) if hf is not None else prev[kind][j]
                              for j, hf in enumerate(own))
        resolved.append(cur)
        prev = cur
    return resolved


def _write_section_hf(ctx, sec_index, blocks, odd_even, which):
    """한글 구역 하나에 머리글·바닥글을 넣는다. which: "main"(기본·홀짝) 또는 "first"(첫 쪽)"""
    for kind in ("header", "footer"):
        base_b, even_b, first_b = blocks[kind]
        if which == "first":
            write_header_footer(ctx, kind, first_b, "BOTH", sec_index=sec_index)
        elif odd_even:
            write_header_footer(ctx, kind, base_b, "ODD", sec_index=sec_index)
            write_header_footer(ctx, kind, even_b, "EVEN", sec_index=sec_index)
            if base_b or even_b:
                ctx.stats["odd_even"] = 1
        else:
            write_header_footer(ctx, kind, base_b, "BOTH", sec_index=sec_index)


def _set_section_start(ctx, sec_index, page=None, starts_on=None):
    """한글 구역의 쪽 번호 시작값과 시작 쪽(홀수/짝수).

    - 구역 정의 hp:startNum page: 0 = 앞 구역에서 이어짐, n = n부터 새로 시작.
      새 구역은 앞 구역 정의를 복사해 만들어지므로 '이어짐'(0)도 반드시 명시한다.
    - 새로 시작할 때는 구역 첫 문단에 '새 번호로 시작'(hp:newNum PAGE) 컨트롤도 넣는다.
      한글이 실제로 쓰는 방식이며, startNum만 읽지 않는 프로그램에서도 번호가 바뀐다."""
    sec = ctx.out.sections[sec_index]
    kw = {"page": page if page is not None else 0,
          "page_starts_on": starts_on or "BOTH"}
    sec.properties.set_start_numbering(**kw)
    if page is not None:
        sec.paragraphs[0].add_new_num(number=page, kind="PAGE")


def plan_sections(ctx, header_footer=True):
    """Word 구역마다 머리글·바닥글을 한글 구역으로 옮긴다.

    Word 설정                        → 한글
    - 구역 나누기(새 쪽/홀수 쪽/짝수 쪽) → 한글 구역 나누기(구역마다 머리글·바닥글 따로)
      · '이전 구역에 연결'            → 앞 구역의 머리글을 그대로 물려받음
      · 쪽 번호 새로 시작(pgNumType start) → 한글 구역 '쪽 번호 시작' 값
      · 홀수/짝수 쪽에서 시작          → 한글 구역 시작 쪽(pageStartsOn) ODD/EVEN
      · 이어서(continuous) 구역        → 머리글·쪽 번호가 앞 구역과 같으면 한 구역으로 합치고,
                                          다르면 새 한글 구역(한글에서는 새 쪽에서 시작)
    - 홀짝 다르게(evenAndOddHeaders, 문서 전체) → ODD = 기본, EVEN = 짝수 머리글
    - 첫 쪽 다르게(titlePg, 구역별)  → 그 구역의 첫 쪽만 담는 한글 구역을 따로 만들어
                                       첫 쪽 머리글을 넣는다(비어 있으면 머리글 없음 = 감추기).
                                       첫 쪽 끝 위치를 못 찾으면 한글 '첫 쪽 감추기' 설정으로 대신.
    - 용지 크기·방향·여백이 다른 구역      → 새 한글 구역(구역마다 용지·여백 따로)
    - 다단(w:cols)                          → 한글 구역의 단 정의(단 수·간격·구분선·단별 폭)
      · 같은 쪽에서 단만 바뀌는 '이어서' 구역 → 구역을 나누지 않고 그 문단에 단 정의 컨트롤
    header_footer=False면 머리글·바닥글은 쓰지 않고 용지·여백·다단만 옮긴다.
    결과: ctx.sec_starts = {본문 순번: 한글 구역 번호}, ctx.col_changes = {본문 순번: 단 설정}
    """
    src = ctx.src
    sections = list(src.sections)
    odd_even = bool(src.settings.odd_and_even_pages_header_footer)
    children, ranges = _section_ranges(src.element.body)
    if len(ranges) != len(sections):           # 비정상 문서: 첫 구역 기준으로만 처리
        ranges = [(0, len(children), None)]
        sections = sections[:1]
    empty = {"header": ([], [], []), "footer": ([], [], [])}
    resolved = _resolve_hf_blocks(ctx, sections) if header_footer else [empty] * len(sections)
    layouts = [section_layout(sp) for _, _, sp in ranges]
    ctx.sec_starts = {}
    hwp_sec = -1                               # 지금까지 쓴 마지막 한글 구역
    prev_key = None
    restarts = 0
    first_info = []

    def new_hwp_section():
        nonlocal hwp_sec
        hwp_sec += 1
        if hwp_sec > 0:
            ctx.out.add_section()
        return hwp_sec

    prev_cols = None
    for k, (section, (start, end, sect_pr), blocks) in enumerate(zip(sections, ranges, resolved)):
        if start >= end and k > 0:             # 내용 없는 구역은 건너뜀
            continue
        lay = layouts[k]
        title_pg = header_footer and bool(section.different_first_page_header_footer)
        start_type, restart = _sect_info(sect_pr)
        key = (repr(blocks), title_pg, _page_key(lay))
        # 이어서 구역: 머리글·용지·여백이 같고 쪽 번호를 새로 시작하지 않으면 앞 한글 구역에 합친다.
        # 단 수만 다르면 그 문단에 단 정의 컨트롤을 넣어 같은 쪽에서 단을 바꾼다.
        if k > 0 and start_type == "continuous" and key == prev_key and restart is None:
            if _col_key(lay) != prev_cols:
                nxt = _sect_info(ranges[k + 1][2])[0] if k + 1 < len(ranges) else None
                ctx.col_changes[start] = (lay, nxt == "continuous")   # 뒤가 이어서 구역이면 단 균형
                prev_cols = _col_key(lay)
            continue
        prev_key = key
        prev_cols = _col_key(lay)
        # Word는 '이어서' 구역 나누기 바로 앞의 다단을 단 길이가 같게(균형) 맞춘다
        nxt_cont = k + 1 < len(ranges) and _sect_info(ranges[k + 1][2])[0] == "continuous"
        starts_on = {"oddPage": "ODD", "evenPage": "EVEN"}.get(start_type)

        any_main = any(blocks[t][0] or blocks[t][1] for t in blocks)
        first_has = title_pg and any(blocks[t][2] for t in blocks)
        split = None
        if title_pg and (first_has or any_main):
            rel = first_page_end(children[start:end])
            if rel is not None and rel > 0:
                split = start + rel
        if split is not None:
            s0 = new_hwp_section()             # 첫 쪽 전용 구역
            ctx.sec_starts[start] = s0
            apply_section_layout(ctx, s0, lay, nxt_cont and split >= end)
            _write_section_hf(ctx, s0, blocks, odd_even, "first")
            _set_section_start(ctx, s0, restart, starts_on)
            first_info.append("전용 머리글" if first_has else "감추기")
            if split < end:                    # 둘째 쪽부터 = 다음 구역(쪽 번호 이어짐)
                s1 = new_hwp_section()
                ctx.sec_starts[split] = s1
                apply_section_layout(ctx, s1, lay, nxt_cont)
                _write_section_hf(ctx, s1, blocks, odd_even, "main")
                _set_section_start(ctx, s1)    # 쪽 번호 이어짐
        else:
            s = new_hwp_section()
            ctx.sec_starts[start] = s
            apply_section_layout(ctx, s, lay, nxt_cont)
            _write_section_hf(ctx, s, blocks, odd_even, "main")
            _set_section_start(ctx, s, restart, starts_on)
            if title_pg:                       # 첫 쪽 끝을 못 찾음: 한글 '첫 쪽 감추기' 설정
                hide = {flag: True for t, flag in (("header", "hideFirstHeader"), ("footer", "hideFirstFooter"))
                        if blocks[t][0] or blocks[t][1]}
                if hide and _set_visibility(ctx, s, **hide):
                    first_info.append("감추기(첫 쪽 끝 위치를 찾지 못함)" if first_has else "감추기")
        if restart is not None:
            restarts += 1

    if first_info:
        ctx.stats["first_page"] = first_info[0] if len(set(first_info)) == 1 else ", ".join(first_info)
    word_secs = len(sections)
    if word_secs > 1:
        ctx.stats["sections"] = word_secs
    if restarts:
        ctx.stats["page_restart"] = restarts
    ctx.stats["hwp_sections"] = hwp_sec + 1
    lays = [l for l in layouts if l]
    if any(l["page"][2] for l in lays):
        ctx.stats["landscape"] = sum(1 for l in lays if l["page"][2])
    if any(l["cols"]["num"] > 1 for l in lays):
        ctx.stats["columns"] = max(l["cols"]["num"] for l in lays)
    if len({(l["page"], tuple(sorted(l["margin"].items()))) for l in lays}) > 1:
        ctx.stats["page_setup"] = len({(l["page"], tuple(sorted(l["margin"].items()))) for l in lays})


# ───────────────────────── 용지 · 여백 · 다단 ─────────────────────────
TWIP_TO_HWP = 5          # 1 twip = 1/1440 inch, 1 HWPUNIT = 1/7200 inch
HWP_PER_MM = 7200 / 25.4


def _twip(el, attr, default=None):
    if el is None:
        return default
    v = el.get(qn(f"w:{attr}"))
    try:
        return int(float(v)) if v is not None else default
    except ValueError:
        return default


def section_layout(sect_pr):
    """Word 구역(w:sectPr)의 용지·여백·다단 → 한글 단위(HWPUNIT) dict. 없으면 None.

    page   : (폭, 높이) 세로 기준 크기, landscape(가로 여부)
    margin : 한글 여백. Word와 한글은 머리글 여백의 기준이 달라서 바꿔 계산한다.
               Word  : 위쪽 여백 = 종이 끝 → 본문,   머리글 = 종이 끝 → 머리글
               한글  : 위쪽 여백 = 종이 끝 → 머리말, 머리말 = 머리말 영역 높이(→ 본문)
             그래서 한글 위쪽 = Word 머리글 거리, 한글 머리말 = Word 위쪽 − 머리글 거리.
             (본문이 시작하는 위치가 Word와 같아진다. 아래쪽·꼬리말도 같은 방식)
    cols   : {"num", "gap", "sep", "widths": [(폭, 간격), …] | None}
    """
    if sect_pr is None:
        return None
    pg = sect_pr.find(qn("w:pgSz"))
    mar = sect_pr.find(qn("w:pgMar"))
    cols = sect_pr.find(qn("w:cols"))
    w, h = _twip(pg, "w", 11906), _twip(pg, "h", 16838)          # 기본 A4
    landscape = (pg is not None and pg.get(qn("w:orient")) == "landscape") or w > h
    top, bottom = abs(_twip(mar, "top", 1440)), abs(_twip(mar, "bottom", 1440))
    hdr, ftr = _twip(mar, "header", 720), _twip(mar, "footer", 720)
    hdr, ftr = min(hdr, top), min(ftr, bottom)
    lay = {
        "page": (min(w, h) * TWIP_TO_HWP, max(w, h) * TWIP_TO_HWP, landscape),
        "margin": {
            "left": _twip(mar, "left", 1440) * TWIP_TO_HWP,
            "right": _twip(mar, "right", 1440) * TWIP_TO_HWP,
            "top": hdr * TWIP_TO_HWP, "header": (top - hdr) * TWIP_TO_HWP,
            "bottom": ftr * TWIP_TO_HWP, "footer": (bottom - ftr) * TWIP_TO_HWP,
            "gutter": _twip(mar, "gutter", 0) * TWIP_TO_HWP,
        },
        "cols": {"num": 1, "gap": 0, "sep": False, "widths": None},
        "deco": section_decoration(sect_pr),                # 세로쓰기·줄 번호·쪽 테두리
    }
    if cols is not None:
        num = max(1, _twip(cols, "num", 1) or 1)
        widths = None
        col_els = cols.findall(qn("w:col"))
        if cols.get(qn("w:equalWidth")) in ("0", "false") and col_els:
            widths = [(_twip(c, "w", 0) * TWIP_TO_HWP, _twip(c, "space", 0) * TWIP_TO_HWP) for c in col_els]
            num = len(widths)
        lay["cols"] = {"num": num, "gap": _twip(cols, "space", 720) * TWIP_TO_HWP,
                       "sep": cols.get(qn("w:sep")) in ("1", "true", "on"), "widths": widths}
    return lay


def _page_key(lay):
    """구역을 따로 나눠야 하는 쪽 설정(용지·여백·세로쓰기·줄 번호·쪽 테두리)"""
    return None if lay is None else (lay["page"], tuple(sorted(lay["margin"].items())), repr(lay.get("deco")))


def _col_key(lay):
    if lay is None:
        return (1,)
    c = lay["cols"]
    return (c["num"], c["gap"], c["sep"], tuple(c["widths"] or ()))


def _col_width(lay, body_w):
    """다단일 때 한 단의 폭(HWPUNIT). 표·이미지 폭 상한으로 쓴다."""
    c = lay["cols"] if lay else {"num": 1}
    if c["num"] <= 1:
        return body_w
    if c.get("widths"):
        return min(w for w, _ in c["widths"]) or body_w
    return max((body_w - c["gap"] * (c["num"] - 1)) // c["num"], 1000)


def _fill_col_pr(col_pr, cols, balanced=False):
    """hp:colPr 요소에 단 개수·간격·구분선·단별 폭을 채운다."""
    hp = col_pr.tag.split("}")[0] + "}"
    for ch in list(col_pr):
        col_pr.remove(ch)
    col_pr.set("type", "BALANCED_NEWSPAPER" if balanced and cols["num"] > 1 else "NEWSPAPER")
    col_pr.set("layout", "LEFT")
    col_pr.set("colCount", str(cols["num"]))
    if cols.get("widths"):
        col_pr.set("sameSz", "0")
        col_pr.set("sameGap", "0")
    else:
        col_pr.set("sameSz", "1")
        col_pr.set("sameGap", str(cols["gap"] if cols["num"] > 1 else 0))
    if cols["num"] > 1 and cols.get("sep"):
        col_pr.append(col_pr.makeelement(hp + "colLine", {"type": "SOLID", "width": "0.12 mm", "color": "#000000"}))
    if cols.get("widths"):
        for w, g in cols["widths"]:
            col_pr.append(col_pr.makeelement(hp + "colSz", {"width": str(w), "gap": str(g)}))


def apply_section_layout(ctx, sec_index, lay, balanced=False):
    """한글 구역에 용지 크기·방향·여백·다단을 넣고, 그 구역의 본문 폭·단 폭을 기록한다."""
    sec = ctx.out.sections[sec_index]
    props = sec.properties
    if lay is None:
        apply_section_decoration(ctx, sec_index, None)   # 앞 구역 설정 지움(+ 쪽 배경색)
    if lay is not None:
        w, h, landscape = lay["page"]
        # 한글 파일은 용지 크기를 늘 세로 기준(폭 < 높이)으로 쓰고 방향은 landscape 속성으로 표시한다.
        # 세로 = WIDELY, 가로 = NARROWLY (한글 고유 표기; 실제 한글 파일과 같은 방식)
        props.set_page_size(width=w, height=h, orientation="NARROWLY" if landscape else "WIDELY")
        props.set_page_margins(**lay["margin"])
        apply_section_decoration(ctx, sec_index, lay.get("deco"))
        col_pr = next((e for e in sec.element.iter() if e.tag.endswith("}colPr")), None)
        if col_pr is not None:
            _fill_col_pr(col_pr, lay["cols"], balanced)
        sec.mark_dirty()
    size, m = props.page_size, props.page_margins
    page_w = size.height if size.orientation == "NARROWLY" else size.width
    body_w = page_w - m.left - m.right - m.gutter
    ctx.body_w[sec_index] = body_w
    ctx.col_w[sec_index] = _col_width(lay, body_w)


def add_column_change(ctx, dst_par, lay, balanced):
    """같은 쪽 안에서 단 수가 바뀌는 곳(Word '이어서' 구역 나누기): 문단 맨 앞에 단 정의 컨트롤."""
    ctrl = dst_par.add_column_definition(max(1, lay["cols"]["num"]) if lay else 1)
    col_pr = next(e for e in ctrl.element.iter() if e.tag.endswith("}colPr"))
    _fill_col_pr(col_pr, lay["cols"] if lay else {"num": 1, "gap": 0, "sep": False, "widths": None}, balanced)
    run = ctrl.element.getparent()                 # 컨트롤이 든 런을 문단 맨 앞으로
    p_el = run.getparent()
    p_el.remove(run)
    p_el.insert(0, run)
    ctx.stats["column_change"] = ctx.stats.get("column_change", 0) + 1


def current_width_mm(ctx):
    """지금 쓰는 위치의 단 폭(mm). 이미지·표 폭 상한."""
    w = ctx.cur_col_w or ctx.col_w.get(ctx.sec) or 42520
    return w / HWP_PER_MM


# ───────────────────────── 세로쓰기 · 줄 번호 · 쪽 테두리 ─────────────────────────
# Word w:textDirection(구역) → 한글 secPr textDirection
#   tbRl  = 위→아래, 오른쪽→왼쪽 (동아시아 세로쓰기; 한글/한자는 세우고 영문은 눕힘) → VERTICAL(영문 눕힘)
#   tbRlV / tbLrV = 세로쓰기 + 영문도 세움 → VERTICALALL(영문 세움)
#   btLr / lrTb 등 = 가로쓰기 취급
TEXT_DIR = {"tbRl": "VERTICAL", "tbRlV": "VERTICALALL", "tbLrV": "VERTICALALL"}

# Word w:lnNumType restart → 한글 lineNumberShape restartType
#   (hwpxlib LineNumberRestartType: 1 쪽마다 새 번호, 2 구역마다 새 번호, 3 이전 번호에 이어서)
LN_RESTART = {"newPage": 1, "newSection": 2, "continuous": 3}

# Word 쪽 테두리 선 종류(w:val) → 한글 선 종류
BORDER_TYPE = {
    "single": "SOLID", "thick": "SOLID", "double": "DOUBLE_SLIM", "dotted": "DOT", "dashed": "DASH",
    "dashSmallGap": "DASH", "dotDash": "DASH_DOT", "dotDotDash": "DASH_DOT_DOT", "triple": "SLIM_THICK_SLIM",
    "thinThickSmallGap": "SLIM_THICK", "thinThickMediumGap": "SLIM_THICK", "thinThickLargeGap": "SLIM_THICK",
    "thickThinSmallGap": "THICK_SLIM", "thickThinMediumGap": "THICK_SLIM", "thickThinLargeGap": "THICK_SLIM",
    "thinThickThinSmallGap": "SLIM_THICK_SLIM", "thinThickThinMediumGap": "SLIM_THICK_SLIM",
    "thinThickThinLargeGap": "SLIM_THICK_SLIM",
    # 한글 테두리 선 종류(hwpxlib LineType2)에는 물결선이 없다 → 가장 비슷한 선
    "wave": "SOLID", "doubleWave": "DOUBLE_SLIM",
    "dashDotStroked": "DASH_DOT", "threeDEmboss": "SOLID", "threeDEngrave": "SOLID",
    "outset": "SOLID", "inset": "SOLID",
}
# 한글 선 굵기는 정해진 값 중 하나만 쓸 수 있다(hwpxlib LineWidth)
def art_border_type(val):
    """Word 그림 테두리(사과·별·점 무늬 등 160여 종) → 무늬 이름으로 가장 비슷한 한글 선.
    한글 쪽 테두리는 그림 무늬를 지원하지 않으므로 선 모양만 근사한다."""
    v = (val or "").lower()
    if "dash" in v:
        return "DASH"
    if "dot" in v or "circle" in v or "ring" in v or "pearl" in v:
        return "CIRCLE" if "white" in v or "circle" in v or "ring" in v else "DOT"
    if "thinlines" in v or "double" in v or "twisted" in v:
        return "DOUBLE_SLIM"
    if "wide" in v or "outline" in v or "inline" in v or "midline" in v:
        return "SLIM_THICK_SLIM"
    if "square" in v or "diamond" in v or "zigzag" in v or "triangle" in v:
        return "LONG_DASH"
    return "SOLID"                                   # 그림(사과·별·나무 등)은 실선


HWP_LINE_WIDTHS = (0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0)
BORDER_DISPLAY = {"allPages": "SHOW_ALL", "firstPage": "SHOW_FIRST", "notFirstPage": "HIDE_FIRST"}


def _line_width_mm(sz_eighth_pt):
    """Word 선 굵기(1/8 pt) → 가장 가까운 한글 선 굵기 문자열('0.5 mm')"""
    mm = sz_eighth_pt / 8 * 25.4 / 72
    best = min(HWP_LINE_WIDTHS, key=lambda w: abs(w - mm))
    return f"{best:g} mm" if best != int(best) else f"{best:.1f} mm"


def section_decoration(sect_pr):
    """Word 구역의 세로쓰기·줄 번호·쪽 테두리 → dict (없는 항목은 None)"""
    deco = {"dir": None, "lnnum": None, "border": None}
    if sect_pr is None:
        return deco
    td = sect_pr.find(qn("w:textDirection"))
    if td is not None:
        deco["dir"] = TEXT_DIR.get(td.get(qn("w:val")))

    ln = sect_pr.find(qn("w:lnNumType"))
    count_by = _twip(ln, "countBy", 0) if ln is not None else 0
    if ln is not None and count_by > 0:
        # Word의 w:start는 0부터 센다(없으면 1번부터): start=4 → 5번부터
        start = _twip(ln, "start", 0) + 1
        dist = _twip(ln, "distance", None)
        deco["lnnum"] = {
            "restart": LN_RESTART.get(ln.get(qn("w:restart"), "newPage"), 1),
            "count_by": count_by,
            # Word 기본 간격('자동')은 약 0.25인치(360twip)
            "distance": (dist if dist is not None else 360) * TWIP_TO_HWP,
            "start": start,
        }

    pb = sect_pr.find(qn("w:pgBorders"))
    if pb is not None:
        sides = {}
        for side in ("top", "left", "bottom", "right"):
            el = pb.find(qn(f"w:{side}"))
            val = el.get(qn("w:val")) if el is not None else None
            if el is None or val in (None, "nil", "none"):
                continue
            color = el.get(qn("w:color"), "auto")
            art = val not in BORDER_TYPE
            sz = _twip(el, "sz", 4)
            sides[side] = {
                # 일반 선: 굵기 단위 1/8pt / 그림 테두리: 굵기 단위 pt (OOXML 규칙)
                "type": art_border_type(val) if art else BORDER_TYPE[val],
                # 그림 테두리는 무늬 크기(pt)라 그대로 선 굵기로 쓰면 너무 굵다 → 무늬 크기의 1/8 정도 선
                "width": _line_width_mm(sz if art else sz),
                "art": val if art else None,
                "color": "#000000" if color in ("auto", None) else "#" + color.upper(),
                "space": _twip(el, "space", 0 if pb.get(qn("w:offsetFrom")) != "page" else 24),
            }
        if sides:
            from_page = pb.get(qn("w:offsetFrom")) == "page"
            arts = {s["art"] for s in sides.values() if s["art"]}
            deco["border"] = {
                "sides": sides,
                "from_page": from_page,
                "display": BORDER_DISPLAY.get(pb.get(qn("w:display"), "allPages"), "SHOW_ALL"),
                "art": sorted(arts),
            }
    return deco


def page_background(doc):
    """문서 배경색(디자인 → 페이지 색, w:background/@w:color) → '#RRGGBB' 또는 None.
    흰색·자동은 배경 없음으로 본다. 그림·그라데이션 배경은 기본 색(w:color)만 쓴다."""
    bg = doc.element.find(qn("w:background"))
    if bg is None:
        return None
    c = bg.get(qn("w:color"))
    if not c or c.lower() in ("auto", "ffffff"):
        return None
    return "#" + c.upper()


def _border_fill_for_sides(ctx, sides, fill=None):
    """네 변의 선 종류·굵기·색(+ 배경색) → 한글 테두리/배경(borderFill) id.
    네 변이 같으면 라이브러리의 같은 정의를 재사용하고, 다르면 새 정의를 만들어 변마다 바꾼다."""
    from copy import deepcopy
    root = ctx.out._root
    if not sides:                                   # 테두리 없이 배경색만
        return root.ensure_border_fill(fill_color=fill, active_borders=[])
    first = next(iter(sides.values()))
    same = len(sides) == 4 and all((s["type"], s["width"], s["color"]) ==
                                    (first["type"], first["width"], first["color"]) for s in sides.values())
    bid = root.ensure_border_fill(border_color=first["color"], border_width=first["width"],
                                  border_type=first["type"], active_borders=list(sides), fill_color=fill)
    if same:
        return bid
    header = root.headers[0]
    holder = next(e for e in header.element.iter() if e.tag.endswith("}borderFills"))
    base = next(e for e in holder if e.get("id") == str(bid))
    new = deepcopy(base)
    new.set("id", str(max(int(e.get("id", 0)) for e in holder) + 1))
    for side, el_name in (("left", "leftBorder"), ("right", "rightBorder"),
                          ("top", "topBorder"), ("bottom", "bottomBorder")):
        el = next((c for c in new if c.tag.endswith("}" + el_name)), None)
        if el is None:
            continue
        s = sides.get(side)
        if s is None:
            el.set("type", "NONE")
        else:
            el.set("type", s["type"]); el.set("width", s["width"]); el.set("color", s["color"])
    holder.append(new)
    holder.set("itemCnt", str(len([e for e in holder if e.tag.endswith("}borderFill")])))
    header.mark_dirty()
    return new.get("id")


def _reset_section_decoration(sec):
    """새 구역은 앞 구역 정의를 복사해 만들어지므로 세로쓰기·줄 번호·쪽 테두리를 기본값으로 되돌린다."""
    props = sec.properties
    props.set_text_direction("HORIZONTAL")
    props.set_line_number_shape(restart_type=0, count_by=0, distance=0, start_number=0)
    for e in props.element.iter():
        name = e.tag.split("}")[-1]
        if name == "visibility":
            e.set("showLineNumber", "0")
            e.set("border", "SHOW_ALL")
        elif name == "pageBorderFill":
            e.set("borderFillIDRef", "1"); e.set("textBorder", "PAPER")
            e.set("headerInside", "0"); e.set("footerInside", "0"); e.set("fillArea", "PAPER")
            off = next((c for c in e if c.tag.endswith("}offset")), None)
            if off is not None:
                for k in ("left", "right", "top", "bottom"):
                    off.set(k, "1417")
    sec.mark_dirty()


def _fix_page_border_bools(props):
    """라이브러리는 headerInside/footerInside를 true/false로 쓰지만 한글 형식은 0/1"""
    for e in props.element.iter():
        if e.tag.endswith("}pageBorderFill"):
            for k in ("headerInside", "footerInside"):
                e.set(k, "1" if e.get(k) in ("1", "true") else "0")


def apply_section_decoration(ctx, sec_index, deco):
    """한글 구역에 세로쓰기·줄 번호·쪽 테두리를 넣는다."""
    sec = ctx.out.sections[sec_index]
    _reset_section_decoration(sec)
    deco = deco or {"dir": None, "lnnum": None, "border": None}
    props = sec.properties
    if deco["dir"]:
        props.set_text_direction(deco["dir"])
        ctx.stats["vertical"] = ctx.stats.get("vertical", 0) + 1
    if deco["lnnum"]:
        ln = deco["lnnum"]
        props.set_line_number_shape(restart_type=ln["restart"], count_by=ln["count_by"],
                                    distance=ln["distance"], start_number=ln["start"])
        _set_visibility(ctx, sec_index, showLineNumber=True)
        ctx.stats["line_number"] = ctx.stats.get("line_number", 0) + 1
    fill = ctx.page_fill
    if fill and not deco["border"]:                 # 테두리 없이 쪽 배경색만
        bid = _border_fill_for_sides(ctx, {}, fill)
        for page_type in ("BOTH", "EVEN", "ODD"):
            props.set_page_border_fill(page_type=page_type, border_fill_id_ref=bid, fill_area="PAPER")
        _fix_page_border_bools(props)
        ctx.stats["page_fill"] = fill
    if deco["border"]:
        b = deco["border"]
        bid = _border_fill_for_sides(ctx, b["sides"], fill)
        if fill:
            ctx.stats["page_fill"] = fill
        if b.get("art"):                            # 구역마다 쓴 그림 테두리 이름을 모두 모은다
            seen = [x for x in ctx.stats.get("art_border", "").split(", ") if x]
            ctx.stats["art_border"] = ", ".join(dict.fromkeys(seen + b["art"]))
        sp = {k: v["space"] for k, v in b["sides"].items()}
        dflt = max(sp.values()) if sp else 0
        # Word 간격 단위: pt (1pt = 100 HWPUNIT).
        #   쪽 기준(offsetFrom=page) → 한글 '종이' 기준, 종이 끝에서 간격
        #   본문 기준(기본)          → 한글 '쪽' 기준(본문 둘레), 머리말·꼬리말 포함
        off = {k: sp.get(k, dflt) * 100 for k in ("left", "right", "top", "bottom")}
        for page_type in ("BOTH", "EVEN", "ODD"):
            props.set_page_border_fill(
                page_type=page_type, border_fill_id_ref=bid,
                text_border="PAPER" if b["from_page"] else "CONTENT",
                header_inside=not b["from_page"], footer_inside=not b["from_page"],
                fill_area="PAPER",
                offset_left=off["left"], offset_right=off["right"],
                offset_top=off["top"], offset_bottom=off["bottom"])
        _fix_page_border_bools(props)
        for e in props.element.iter():
            if e.tag.endswith("}visibility"):
                e.set("border", b["display"])
        sec.mark_dirty()
        ctx.stats["page_border"] = ctx.stats.get("page_border", 0) + 1


# ───────────────────────── 본체 ─────────────────────────
def convert(src_path, dst_path, keep_format=True, footnote_mode="note", header_footer=True):
    """docx → hwpx

    keep_format   : 글자·문단 서식 유지 여부
    footnote_mode : "note" = 한글 각주로 넣기 / "text" = 본문에 [1] 표시 + 문서 끝에 각주 목록
    header_footer : 머리글·바닥글 옮기기 여부
    """
    src = Document(src_path)
    out = HwpxDocument.new()
    ctx = Ctx(src, out, keep_format, footnote_mode)
    stats = ctx.stats

    plan_sections(ctx, header_footer)        # 구역 나누기 + 용지·여백·다단 (+ 머리글·바닥글)
    ctx.cur_col_w = ctx.col_w.get(0)

    def add_par():
        return out.add_paragraph("", inherit_style=False, include_run=False, section_index=ctx.sec)

    idx = -1
    pending_break = False                     # 앞 문단 끝에 쪽 나눔이 있었음
    for child in src.element.body.iterchildren():
        if child.tag not in (qn("w:p"), qn("w:tbl")):
            continue
        idx += 1
        new_section = idx in ctx.sec_starts and ctx.sec_starts[idx] != ctx.sec
        if new_section:
            ctx.sec, pending_break = ctx.sec_starts[idx], False   # 새 구역은 새 쪽에서 시작
            ctx.cur_col_w = ctx.col_w.get(ctx.sec)
        col_change = ctx.col_changes.get(idx)
        if col_change is not None:
            lay_c, balanced = col_change
            add_column_change(ctx, add_par(), lay_c, balanced)    # 단 정의만 담은 빈 문단
            ctx.cur_col_w = _col_width(lay_c, ctx.body_w.get(ctx.sec, 42520))

        if child.tag == qn("w:tbl"):
            if build_table(ctx, child, src,
                           lambda r, c: out.add_table(r, c, section_index=ctx.sec,
                                                      width=ctx.cur_col_w)) is not None:
                stats["table"] += 1
            continue
        par = Paragraph(child, src)
        prv, nxt = child.getprevious(), child.getnext()
        prev_par = Paragraph(prv, src) if prv is not None and prv.tag == qn("w:p") else None
        next_par = Paragraph(nxt, src) if nxt is not None and nxt.tag == qn("w:p") else None
        for blob, ext, w in paragraph_images(par, src):
            w = min(w, current_width_mm(ctx)) if w else None      # 단 폭을 넘지 않게
            kw = {"width_mm": w} if w else {}
            out.add_picture(blob, image_format=ext, section_index=ctx.sec, **kw)
            stats["image"] += 1
        brk_before, brk_after = page_breaks(child)
        has_note = has_note_ref(child)
        if not par.text.strip() and not has_note:
            pending_break = pending_break or brk_before or brk_after
            continue

        # 한글 기본 '개요' 스타일은 자동 번호가 뒤 문단까지 번지므로 쓰지 않고,
        # 일반 문단 + 굵고 큰 글자 런으로 제목을 표현한다.
        dp = add_par()
        if (pending_break or brk_before) and not new_section:
            dp.element.set("pageBreak", "1")    # 원본의 쪽 나눔 유지
            stats["page_break"] = stats.get("page_break", 0) + 1
        pending_break = brk_after
        level = heading_level(par)
        if level:
            write_runs(ctx, dp, par, base={"bold": True, "size": heading_size(par, level, keep_format)})
            if keep_format:
                fmt = apply_contextual(paragraph_format(par), par, prev_par, next_par)
                if ctx.paras.apply(dp, fmt):
                    stats["para_fmt"] += 1
            else:
                fmt = {"no_lnnum": True} if line_numbers_suppressed(par) else {}
                ctx.paras.apply(dp, fmt)
            if fmt.get("no_lnnum"):
                stats["lnnum_skip"] = stats.get("lnnum_skip", 0) + 1
            stats["heading"] += 1
        else:
            write_paragraph(ctx, dp, par, prev_par=prev_par, next_par=next_par)
            stats["para"] += 1

    # footnote_mode="text": 문서 끝에 각주([1]) · 미주([i]) 목록 — 서식·이미지·표 유지
    write_note_list(ctx, add_par, lambda r, c: out.add_table(r, c, section_index=ctx.sec))

    if stats.get("endnote") and footnote_mode == "note":
        set_endnote_format(out)
    sync_bin_manifest(out)
    out.save_to_path(dst_path)
    return stats


def sync_bin_manifest(out):
    """python-hwpx 6.1.0 버그 보정: 구역을 추가한 뒤에 넣은 이미지가 content.hpf 목록에서 빠진다.

    구역 추가 때 문서 루트가 목록(manifest)을 따로 들고 '변경됨'으로 표시하는데, 이미지 등록은
    패키지 쪽 목록에만 기록된다. 저장할 때 루트 쪽 옛 목록이 덮어써서 이미지 항목이 사라지고
    한글/LibreOffice에서 그림이 보이지 않는다. 저장 직전에 빠진 BinData 항목을 루트 목록에 더한다."""
    root = getattr(out, "_root", None)
    manifest = getattr(root, "_manifest", None)
    if manifest is None:
        return
    opf = "http://www.idpf.org/2007/opf/"
    m_el = next((e for e in manifest.iter() if e.tag.endswith("}manifest")), None)
    if m_el is None:
        return
    have = {e.get("id") for e in m_el if e.tag.endswith("}item")}
    added = False
    for item in out._package._manifest_items():
        iid = item.get("id")
        if iid and iid not in have and str(item.get("href", "")).startswith("BinData/"):
            ns = m_el.tag.split("}")[0] + "}" if "}" in m_el.tag else "{" + opf + "}"
            m_el.append(m_el.makeelement(ns + "item", dict(item.attrib)))
            have.add(iid)
            added = True
    if added:
        root._manifest_dirty = True


def set_endnote_format(out):
    """미주 번호를 Word 기본값처럼 로마 숫자 소문자(i, ii, iii)로 하고, 문서 끝에 모은다.
    각주(1, 2, 3)와 번호가 섞이지 않게 한다. 모든 구역에 같은 설정을 넣는다."""
    for sec in out.sections:
        for pr in sec.properties.element.iter():
            if not pr.tag.endswith("}endNotePr"):
                continue
            for el in pr:
                name = el.tag.split("}")[1]
                if name == "autoNumFormat":
                    el.set("type", "ROMAN_SMALL")
                    el.set("suffixChar", "")
                elif name == "placement":
                    el.set("place", "END_OF_DOCUMENT")
        sec.properties.section.mark_dirty()


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    src = Path(sys.argv[1])
    jobs = ([(p, p.with_suffix(".hwpx")) for p in sorted(src.glob("*.docx"))] if src.is_dir()
            else [(src, Path(sys.argv[2]) if len(sys.argv) > 2 else src.with_suffix(".hwpx"))])
    for s, d in jobs:
        try:
            st = convert(str(s), str(d))
            print(f"[OK] {s.name} → {d.name}  {st}")
        except Exception as e:
            print(f"[실패] {s.name}: {type(e).__name__}: {e}")


if __name__ == "__main__":
    main()
