# -*- coding: utf-8 -*-
"""
HWPX Core — hwpx_writer에서 추출한 공통 HWPX 생성 유틸리티

문자열 기반 XML 생성으로 한글 오피스 호환 HWPX 문서를 빌드한다.
"""

import base64
import re
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape as xml_escape


def _log(msg: str) -> None:
    print(msg, file=sys.stderr)


# ---------------------------------------------------------------------------
# 네임스페이스 상수
# ---------------------------------------------------------------------------
ALL_NS = (
    'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" '
    'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" '
    'xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" '
    'xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" '
    'xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" '
    'xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" '
    'xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" '
    'xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" '
    'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" '
    'xmlns:dc="http://purl.org/dc/elements/1.1/" '
    'xmlns:opf="http://www.idpf.org/2007/opf/" '
    'xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" '
    'xmlns:epub="http://www.idpf.org/2007/ops" '
    'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"'
)

# 빈 1x1 PNG (Preview/PrvImage.png용)
_EMPTY_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4"
    "nGNgYPgPAAEDAQAIicLsAAAABElFTkSuQmCC"
)


# ---------------------------------------------------------------------------
# 단위 변환
# ---------------------------------------------------------------------------
def pt_to_height(pt: float) -> int:
    """포인트 → charPr height (1pt = 100)"""
    return int(pt * 100)

def pt_to_hwpunit(pt: float) -> int:
    """포인트 → HWPUNIT (1pt = 50)"""
    return int(pt * 50)

def hwpunit_to_pt(hu: int) -> float:
    """HWPUNIT → 포인트"""
    return hu / 50.0

def height_to_pt(h: int) -> float:
    """charPr height → 포인트"""
    return h / 100.0

def mm_to_hwpunit(mm: float) -> int:
    """밀리미터 → HWPUNIT (1mm ≈ 283.46 HWPUNIT, 1pt=50, 72pt=1in, 1in=25.4mm)"""
    return int(mm * 7200 / 25.4 / 2)


# ---------------------------------------------------------------------------
# 마커 파싱
# ---------------------------------------------------------------------------
def parse_markers(text: str) -> list[tuple[str, str | None]]:
    """{{color:텍스트}} 마커를 파싱. [(text, color_or_None), ...]"""
    text = re.sub(r'<[^>]+>', '', text)
    text = text.replace('&nbsp;', ' ').replace('&lt;', '<').replace('&gt;', '>')
    text = text.replace('&amp;', '&').replace('&quot;', '"')

    segments = []
    pos = 0
    for m in re.finditer(r'\{\{(\w+):([^}]+)\}\}', text):
        if m.start() > pos:
            segments.append((text[pos:m.start()], None))
        segments.append((m.group(2), m.group(1)))
        pos = m.end()
    if pos < len(text):
        segments.append((text[pos:], None))
    return segments if segments else [(text, None)]


# ---------------------------------------------------------------------------
# HWPXBuilder — HWPX 문서를 빌드하기 위한 핵심 클래스
# ---------------------------------------------------------------------------
class HWPXBuilder:
    """HWPX 문서의 XML 요소를 생성하고 ZIP으로 패키징하는 빌더."""

    def __init__(self, line_spacing: int = 160):
        self.line_spacing = line_spacing

        # CharPr / ParaPr 관리
        self._charpr_list: list[tuple] = []
        self._parapr_list: list[tuple] = []
        self._next_charpr_id = 1
        self._next_parapr_id = 1
        self._charpr_cache: dict[tuple, int] = {}
        self._parapr_cache: dict[tuple, int] = {}

        # 표 셀 전용 ParaPr
        self._table_parapr_list: list[tuple] = []

        # 폰트 관리
        self._fonts: list[tuple[int, str]] = []
        self._font_cache: dict[str, int] = {}

        # 색상 매핑
        self.colors: dict[str, str] = {
            "red": "#DC2626", "green": "#16A34A", "blue": "#2563EB",
            "yellow": "#EAB308", "black": "#000000",
        }

    # -- 폰트 --
    @staticmethod
    def font_family_type(face: str) -> str:
        myeongjo = ['바탕', '명조', 'Batang', 'Myeongjo', 'Gungsuh', 'Serif']
        return "FCAT_MYEONGJO" if any(k in face for k in myeongjo) else "FCAT_GOTHIC"

    def register_font(self, face: str) -> int:
        if face in self._font_cache:
            return self._font_cache[face]
        fid = len(self._fonts)
        self._fonts.append((fid, face))
        self._font_cache[face] = fid
        return fid

    # -- CharPr --
    def get_charpr_id(self, height: int, text_color: str, font_name: str,
                      bold: bool = False) -> int:
        font_id = self.register_font(font_name)
        key = (height, text_color.upper(), font_id, bold)
        if key in self._charpr_cache:
            return self._charpr_cache[key]
        cid = self._next_charpr_id
        self._next_charpr_id += 1
        self._charpr_list.append((cid, height, text_color.upper(), font_id, bold))
        self._charpr_cache[key] = cid
        return cid

    # -- ParaPr --
    def get_parapr_id(self, left_margin: int = 0, space_before: int = 0,
                      space_after: int = 0, align: str = "JUSTIFY",
                      indent: int = 0, line_spacing: int | None = None) -> int:
        ls = line_spacing if line_spacing is not None else self.line_spacing
        key = (left_margin, space_before, space_after, align, indent, ls)
        if key in self._parapr_cache:
            return self._parapr_cache[key]
        pid = self._next_parapr_id
        self._next_parapr_id += 1
        self._parapr_list.append((pid, left_margin, space_before, space_after, align, indent, ls))
        self._parapr_cache[key] = pid
        return pid

    def get_table_parapr_id(self, align: str = "LEFT", line_spacing: int = 160) -> int:
        key = (align, line_spacing)
        for tpid, a, ls in self._table_parapr_list:
            if (a, ls) == key:
                return tpid
        pid = self._next_parapr_id
        self._next_parapr_id += 1
        self._table_parapr_list.append((pid, align, line_spacing))
        return pid

    def resolve_color(self, name: str) -> str:
        return self.colors.get(name.lower(), "#000000").upper()

    # ==================================================================
    # XML 빌더
    # ==================================================================

    # -- run / paragraph --
    def run_xml(self, text: str, charpr_id: int) -> str:
        return (
            f'<hp:run charPrIDRef="{charpr_id}">'
            f'<hp:t>{xml_escape(text)}</hp:t>'
            f'</hp:run>'
        )

    def paragraph_xml(self, runs_xml: str, parapr_id: int = 0,
                      page_break: str = "0") -> str:
        return (
            f'<hp:p id="0" paraPrIDRef="{parapr_id}"'
            f' styleIDRef="0" pageBreak="{page_break}"'
            f' columnBreak="0" merged="0">'
            f'{runs_xml}'
            f'</hp:p>'
        )

    def text_paragraph(self, text: str, font_name: str, font_size_pt: float,
                       left_margin_pt: float = 0, space_before_pt: float = 0,
                       space_after_pt: float = 0, align: str = "JUSTIFY",
                       hanging_indent_pt: float = 0, bold: bool = False,
                       color: str = "#000000",
                       line_spacing: int | None = None) -> str:
        """마커 색상을 지원하는 텍스트 paragraph"""
        height = pt_to_height(font_size_pt)
        actual_left = pt_to_hwpunit(left_margin_pt)
        indent_val = -pt_to_hwpunit(hanging_indent_pt) if hanging_indent_pt else 0
        parapr_id = self.get_parapr_id(
            actual_left,
            pt_to_hwpunit(space_before_pt),
            pt_to_hwpunit(space_after_pt),
            align, indent_val,
            line_spacing,
        )

        segments = parse_markers(text)
        runs = ""
        for seg_text, seg_marker in segments:
            is_bold = bold or (seg_marker == "bold")
            if seg_marker and seg_marker != "bold":
                color_hex = self.resolve_color(seg_marker)
            else:
                color_hex = color
            cid = self.get_charpr_id(height, color_hex, font_name, bold=is_bold)
            runs += self.run_xml(seg_text, cid)

        return self.paragraph_xml(runs, parapr_id)

    def empty_paragraph(self) -> str:
        """빈 문단 (한글 오피스 호환 — linesegarray 포함)"""
        return (
            '<hp:p id="0" paraPrIDRef="0" styleIDRef="0"'
            ' pageBreak="0" columnBreak="0" merged="0">'
            '<hp:run charPrIDRef="0"/>'
            '<hp:linesegarray>'
            '<hp:lineseg textpos="0" vertpos="0" vertsize="1000"'
            ' textheight="1000" baseline="850" spacing="600"'
            ' horzpos="0" horzsize="42520" flags="393216"/>'
            '</hp:linesegarray>'
            '</hp:p>'
        )

    def page_break_paragraph(self) -> str:
        """페이지 나눔 문단"""
        return self.paragraph_xml(
            '<hp:run charPrIDRef="0"><hp:t></hp:t></hp:run>',
            parapr_id=0, page_break="1",
        )

    def separator_line_paragraph(self, width_hwpunit: int = 42520) -> str:
        """구분선 문단 (가로선)"""
        parapr_id = self.get_parapr_id(0, pt_to_hwpunit(4), pt_to_hwpunit(4), "LEFT")
        line_w = width_hwpunit
        return (
            f'<hp:p id="0" paraPrIDRef="{parapr_id}"'
            f' styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
            f'<hp:run charPrIDRef="0">'
            f'<hp:drawingObjectGroup>'
            f'<hp:line id="0" zOrder="0" numberingType="NONE"'
            f' textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0"'
            f' dropcapstyle="None" href="" groupLevel="0">'
            f'<hp:sz width="{line_w}" widthRelTo="ABSOLUTE"'
            f' height="0" heightRelTo="ABSOLUTE" protect="0"/>'
            f'<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1"'
            f' allowOverlap="0" holdAnchorAndSO="0"'
            f' vertRelTo="PARA" horzRelTo="PARA"'
            f' vertAlign="TOP" horzAlign="LEFT"'
            f' vertOffset="0" horzOffset="0"/>'
            f'<hp:outMargin left="0" right="0" top="0" bottom="0"/>'
            f'<hp:lineShape>'
            f'<hp:startPt x="0" y="0"/>'
            f'<hp:endPt x="{line_w}" y="0"/>'
            f'<hc:lineAttr width="0.12 mm" headStyle="NONE" tailStyle="NONE"'
            f' headSz="MEDIUM_MEDIUM" tailSz="MEDIUM_MEDIUM"'
            f' outlineStyle="NORMAL" cap="FLAT">'
            f'<hc:lineType type="SOLID"/>'
            f'<hc:color value="#CCCCCC"/>'
            f'</hc:lineAttr>'
            f'</hp:lineShape>'
            f'</hp:line>'
            f'</hp:drawingObjectGroup>'
            f'</hp:run>'
            f'</hp:p>'
        )

    # -- 표 --
    def table_cell_xml(self, text: str, col_idx: int, row_idx: int,
                       cell_width: int, font_name: str = "맑은 고딕",
                       font_size: float = 10, is_header: bool = False,
                       align: str | None = None,
                       line_spacing: int = 160) -> str:
        """표 셀 XML"""
        segments = parse_markers(text)
        cell_runs = ""
        for seg_text, seg_marker in segments:
            is_bold = is_header or (seg_marker == "bold")
            if seg_marker and seg_marker != "bold":
                color_hex = self.resolve_color(seg_marker)
            else:
                color_hex = "#000000"
            cid = self.get_charpr_id(
                pt_to_height(font_size), color_hex, font_name, bold=is_bold)
            cell_runs += self.run_xml(seg_text, cid)

        cell_align = align or ("CENTER" if is_header else "LEFT")
        tbl_parapr = self.get_table_parapr_id(cell_align, line_spacing)
        bf_id = "4" if is_header else "3"

        font_height = int(font_size * 100)
        baseline = int(font_height * 0.85)
        spacing = int(font_height * 0.6)
        vert_size = int(font_height * line_spacing / 100)
        cell_height = vert_size + 282
        inner_width = max(cell_width - 1020, 1000)

        return (
            f'<hp:tc name="" header="0" hasMargin="0" protect="0"'
            f' editable="0" dirty="0" borderFillIDRef="{bf_id}">'
            f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK"'
            f' vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0"'
            f' textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
            f'<hp:p id="0" paraPrIDRef="{tbl_parapr}" styleIDRef="0"'
            f' pageBreak="0" columnBreak="0" merged="0">'
            f'{cell_runs}'
            f'<hp:linesegarray>'
            f'<hp:lineseg textpos="0" vertpos="0" vertsize="{vert_size}"'
            f' textheight="{font_height}" baseline="{baseline}" spacing="{spacing}"'
            f' horzpos="0" horzsize="{inner_width}" flags="393216"/>'
            f'</hp:linesegarray>'
            f'</hp:p>'
            f'</hp:subList>'
            f'<hp:cellAddr colAddr="{col_idx}" rowAddr="{row_idx}"/>'
            f'<hp:cellSpan colSpan="1" rowSpan="1"/>'
            f'<hp:cellSz width="{cell_width}" height="{cell_height}"/>'
            f'<hp:cellMargin left="510" right="510" top="141" bottom="141"/>'
            f'</hp:tc>'
        )

    def table_xml(self, headers: list[str], rows: list[list[str]],
                  total_width: int = 42520, font_name: str = "맑은 고딕",
                  font_size: float = 10, line_spacing: int = 160) -> str:
        """표 전체 XML"""
        col_count = len(headers) if headers else (len(rows[0]) if rows else 1)
        has_header = any(h.strip() for h in headers) if headers else False
        row_count = len(rows) + (1 if has_header else 0)
        cell_width = total_width // col_count

        font_height = int(font_size * 100)
        vert_size = int(font_height * line_spacing / 100)
        cell_height = vert_size + 282

        header_row = ""
        if has_header:
            for ci, h in enumerate(headers):
                header_row += self.table_cell_xml(
                    h, ci, 0, cell_width, font_name, font_size,
                    is_header=True, line_spacing=line_spacing)
            header_row = f'<hp:tr>{header_row}</hp:tr>'

        data_rows = ""
        for ri, row in enumerate(rows):
            cells = ""
            row_idx = ri + (1 if has_header else 0)
            for ci, cell in enumerate(row):
                cells += self.table_cell_xml(
                    cell, ci, row_idx, cell_width, font_name, font_size,
                    line_spacing=line_spacing)
            data_rows += f'<hp:tr>{cells}</hp:tr>'

        table_h = cell_height * row_count

        return (
            f'<hp:tbl id="0" zOrder="0" numberingType="TABLE"'
            f' textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0"'
            f' dropcapstyle="None" pageBreak="CELL" repeatHeader="1"'
            f' rowCnt="{row_count}" colCnt="{col_count}"'
            f' cellSpacing="0" borderFillIDRef="3" noAdjust="0">'
            f'<hp:sz width="{total_width}" widthRelTo="ABSOLUTE"'
            f' height="{table_h}" heightRelTo="ABSOLUTE" protect="0"/>'
            f'<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1"'
            f' allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA"'
            f' horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT"'
            f' vertOffset="0" horzOffset="0"/>'
            f'<hp:outMargin left="283" right="283" top="283" bottom="283"/>'
            f'<hp:inMargin left="510" right="510" top="141" bottom="141"/>'
            f'{header_row}{data_rows}'
            f'</hp:tbl>'
        )

    def table_paragraph_xml(self, headers: list[str], rows: list[list[str]],
                            total_width: int = 42520, **kwargs) -> str:
        """표를 담는 paragraph XML"""
        tbl = self.table_xml(headers, rows, total_width, **kwargs)
        return (
            f'<hp:p id="0" paraPrIDRef="0" styleIDRef="0"'
            f' pageBreak="0" columnBreak="0" merged="0">'
            f'<hp:run charPrIDRef="0">'
            f'{tbl}'
            f'<hp:t/>'
            f'</hp:run>'
            f'<hp:linesegarray>'
            f'<hp:lineseg textpos="0" vertpos="0" vertsize="1000"'
            f' textheight="1000" baseline="850" spacing="600"'
            f' horzpos="0" horzsize="0" flags="393216"/>'
            f'</hp:linesegarray>'
            f'</hp:p>'
        )

    # ==================================================================
    # header.xml 빌더
    # ==================================================================
    def _build_fontfaces_xml(self) -> str:
        langs = ["HANGUL", "LATIN", "HANJA", "JAPANESE", "OTHER", "SYMBOL", "USER"]
        font_cnt = len(self._fonts)
        font_elems = ""
        for fid, face in self._fonts:
            family_type = self.font_family_type(face)
            font_elems += (
                f'<hh:font id="{fid}" face="{xml_escape(face)}" type="TTF" isEmbedded="0">'
                f'<hh:typeInfo familyType="{family_type}" weight="6" proportion="4"'
                f' contrast="0" strokeVariation="1" armStyle="1"'
                f' letterform="1" midline="1" xHeight="1"/>'
                f'</hh:font>'
            )
        faces = ""
        for lang in langs:
            faces += f'<hh:fontface lang="{lang}" fontCnt="{font_cnt}">{font_elems}</hh:fontface>'
        return f'<hh:fontfaces itemCnt="{len(langs)}">{faces}</hh:fontfaces>'

    def _build_charpr_xml(self, cid, height, text_color, font_id, bold=False) -> str:
        fid = str(font_id)
        bold_attr = ' bold="1"' if bold else ''
        return (
            f'<hh:charPr id="{cid}" height="{height}" textColor="{text_color}"'
            f' shadeColor="none" useFontSpace="0" useKerning="0"'
            f' symMark="NONE" borderFillIDRef="2"{bold_attr}>'
            f'<hh:fontRef hangul="{fid}" latin="{fid}" hanja="{fid}"'
            f' japanese="{fid}" other="{fid}" symbol="{fid}" user="{fid}"/>'
            f'<hh:ratio hangul="100" latin="100" hanja="100"'
            f' japanese="100" other="100" symbol="100" user="100"/>'
            f'<hh:spacing hangul="0" latin="0" hanja="0"'
            f' japanese="0" other="0" symbol="0" user="0"/>'
            f'<hh:relSz hangul="100" latin="100" hanja="100"'
            f' japanese="100" other="100" symbol="100" user="100"/>'
            f'<hh:offset hangul="0" latin="0" hanja="0"'
            f' japanese="0" other="0" symbol="0" user="0"/>'
            f'</hh:charPr>'
        )

    def _build_parapr_xml(self, pid, left_margin, space_before, space_after,
                          align, indent=0, line_spacing=None) -> str:
        ls = line_spacing if line_spacing is not None else self.line_spacing
        return (
            f'<hh:paraPr id="{pid}" tabPrIDRef="0" condense="0"'
            f' fontLineHeight="0" snapToGrid="1"'
            f' suppressLineNumbers="0" checked="0">'
            f'<hh:align horizontal="{align}" vertical="BASELINE"/>'
            f'<hh:heading type="NONE" idRef="0" level="0"/>'
            f'<hh:breakSetting breakLatinWord="KEEP_WORD"'
            f' breakNonLatinWord="KEEP_WORD" widowOrphan="0"'
            f' keepWithNext="0" keepLines="1" pageBreakBefore="0"'
            f' lineWrap="BREAK"/>'
            f'<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>'
            f'<hh:margin>'
            f'<hc:intent value="{indent}" unit="HWPUNIT"/>'
            f'<hc:left value="{left_margin}" unit="HWPUNIT"/>'
            f'<hc:right value="0" unit="HWPUNIT"/>'
            f'<hc:prev value="{space_before}" unit="HWPUNIT"/>'
            f'<hc:next value="{space_after}" unit="HWPUNIT"/>'
            f'</hh:margin>'
            f'<hh:lineSpacing type="PERCENT" value="{ls}" unit="HWPUNIT"/>'
            f'<hh:border borderFillIDRef="2" offsetLeft="0" offsetRight="0"'
            f' offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/>'
            f'</hh:paraPr>'
        )

    def _build_table_parapr_xml(self, pid, align, line_spacing) -> str:
        return (
            f'<hh:paraPr id="{pid}" tabPrIDRef="0" condense="0"'
            f' fontLineHeight="1" snapToGrid="1"'
            f' suppressLineNumbers="0" checked="0">'
            f'<hh:align horizontal="{align}" vertical="BASELINE"/>'
            f'<hh:heading type="NONE" idRef="0" level="0"/>'
            f'<hh:breakSetting breakLatinWord="HYPHENATION"'
            f' breakNonLatinWord="BREAK_ALL" widowOrphan="0"'
            f' keepWithNext="0" keepLines="1" pageBreakBefore="0"'
            f' lineWrap="BREAK"/>'
            f'<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>'
            f'<hh:margin>'
            f'<hc:intent value="0" unit="HWPUNIT"/>'
            f'<hc:left value="0" unit="HWPUNIT"/>'
            f'<hc:right value="0" unit="HWPUNIT"/>'
            f'<hc:prev value="0" unit="HWPUNIT"/>'
            f'<hc:next value="0" unit="HWPUNIT"/>'
            f'</hh:margin>'
            f'<hh:lineSpacing type="PERCENT" value="{line_spacing}" unit="HWPUNIT"/>'
            f'<hh:border borderFillIDRef="2" offsetLeft="0" offsetRight="0"'
            f' offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/>'
            f'</hh:paraPr>'
        )

    def _build_borderfills_xml(self) -> str:
        bfs = (
            '<hh:borderFill id="1" threeD="0" shadow="0"'
            ' centerLine="NONE" breakCellSeparateLine="0">'
            '<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
            '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
            '<hh:leftBorder type="NONE" width="0.1 mm" color="#000000"/>'
            '<hh:rightBorder type="NONE" width="0.1 mm" color="#000000"/>'
            '<hh:topBorder type="NONE" width="0.1 mm" color="#000000"/>'
            '<hh:bottomBorder type="NONE" width="0.1 mm" color="#000000"/>'
            '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>'
            '</hh:borderFill>'
            '<hh:borderFill id="2" threeD="0" shadow="0"'
            ' centerLine="NONE" breakCellSeparateLine="0">'
            '<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
            '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
            '<hh:leftBorder type="NONE" width="0.1 mm" color="#000000"/>'
            '<hh:rightBorder type="NONE" width="0.1 mm" color="#000000"/>'
            '<hh:topBorder type="NONE" width="0.1 mm" color="#000000"/>'
            '<hh:bottomBorder type="NONE" width="0.1 mm" color="#000000"/>'
            '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>'
            '<hc:fillBrush>'
            '<hc:winBrush faceColor="none" hatchColor="#999999" alpha="0"/>'
            '</hc:fillBrush>'
            '</hh:borderFill>'
            '<hh:borderFill id="3" threeD="0" shadow="0"'
            ' centerLine="NONE" breakCellSeparateLine="0">'
            '<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
            '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
            '<hh:leftBorder type="SOLID" width="0.12 mm" color="#000000"/>'
            '<hh:rightBorder type="SOLID" width="0.12 mm" color="#000000"/>'
            '<hh:topBorder type="SOLID" width="0.12 mm" color="#000000"/>'
            '<hh:bottomBorder type="SOLID" width="0.12 mm" color="#000000"/>'
            '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>'
            '</hh:borderFill>'
            '<hh:borderFill id="4" threeD="0" shadow="0"'
            ' centerLine="NONE" breakCellSeparateLine="0">'
            '<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
            '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
            '<hh:leftBorder type="SOLID" width="0.12 mm" color="#000000"/>'
            '<hh:rightBorder type="SOLID" width="0.12 mm" color="#000000"/>'
            '<hh:topBorder type="SOLID" width="0.12 mm" color="#000000"/>'
            '<hh:bottomBorder type="SOLID" width="0.12 mm" color="#000000"/>'
            '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>'
            '<hc:fillBrush>'
            '<hc:winBrush faceColor="#D9D9D9" hatchColor="#999999" alpha="0"/>'
            '</hc:fillBrush>'
            '</hh:borderFill>'
        )
        return f'<hh:borderFills itemCnt="4">{bfs}</hh:borderFills>'

    def build_header_xml(self) -> str:
        fontfaces = self._build_fontfaces_xml()

        charpr_default = self._build_charpr_xml(0, 1000, "#000000", 0)
        charprs = charpr_default
        for cid, height, color, fid, bold in self._charpr_list:
            charprs += self._build_charpr_xml(cid, height, color, fid, bold)
        charpr_cnt = 1 + len(self._charpr_list)

        parapr_default = self._build_parapr_xml(0, 0, 0, 0, "JUSTIFY")
        paraprs = parapr_default
        for item in self._parapr_list:
            paraprs += self._build_parapr_xml(*item)
        for tpid, align, ls in self._table_parapr_list:
            paraprs += self._build_table_parapr_xml(tpid, align, ls)
        parapr_cnt = 1 + len(self._parapr_list) + len(self._table_parapr_list)

        borderfills = self._build_borderfills_xml()

        tab_xml = (
            '<hh:tabProperties itemCnt="3">'
            '<hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"/>'
            '<hh:tabPr id="1" autoTabLeft="1" autoTabRight="0"/>'
            '<hh:tabPr id="2" autoTabLeft="0" autoTabRight="1"/>'
            '</hh:tabProperties>'
        )
        numberings = (
            '<hh:numberings itemCnt="1">'
            '<hh:numbering id="1" start="0">'
            + ''.join(
                f'<hh:paraHead start="1" level="{i}" align="LEFT" useInstWidth="1"'
                f' autoIndent="1" widthAdjust="0" textOffsetType="PERCENT"'
                f' textOffset="50" numFormat="DIGIT" charPrIDRef="4294967295"'
                f' checkable="0">^{i}.</hh:paraHead>'
                for i in range(1, 11)
            )
            + '</hh:numbering></hh:numberings>'
        )
        styles = (
            '<hh:styles itemCnt="22">'
            '<hh:style id="0" type="PARA" name="바탕글" engName="Normal"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="0" langID="1042" lockForm="0"/>'
            '<hh:style id="1" type="PARA" name="본문" engName="Body"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="1" langID="1042" lockForm="0"/>'
            + ''.join(
                f'<hh:style id="{i}" type="PARA" name="개요 {i-1}" engName="Outline {i-1}"'
                f' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="{i}" langID="1042" lockForm="0"/>'
                for i in range(2, 12)
            )
            + '<hh:style id="12" type="CHAR" name="쪽 번호" engName="Page Number"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="0" langID="1042" lockForm="0"/>'
            '<hh:style id="13" type="PARA" name="머리말" engName="Header"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="13" langID="1042" lockForm="0"/>'
            '<hh:style id="14" type="PARA" name="각주" engName="Footnote"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="14" langID="1042" lockForm="0"/>'
            '<hh:style id="15" type="PARA" name="미주" engName="Endnote"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="15" langID="1042" lockForm="0"/>'
            '<hh:style id="16" type="PARA" name="메모" engName="Memo"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="16" langID="1042" lockForm="0"/>'
            '<hh:style id="17" type="PARA" name="차례 제목" engName="TOC Heading"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="17" langID="1042" lockForm="0"/>'
            '<hh:style id="18" type="PARA" name="차례 1" engName="TOC 1"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="18" langID="1042" lockForm="0"/>'
            '<hh:style id="19" type="PARA" name="차례 2" engName="TOC 2"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="19" langID="1042" lockForm="0"/>'
            '<hh:style id="20" type="PARA" name="차례 3" engName="TOC 3"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="20" langID="1042" lockForm="0"/>'
            '<hh:style id="21" type="PARA" name="캡션" engName="Caption"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="21" langID="1042" lockForm="0"/>'
            '</hh:styles>'
        )

        return (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
            f'<hh:head {ALL_NS} version="1.2" secCnt="1">'
            f'<hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/>'
            f'<hh:refList>'
            f'{fontfaces}'
            f'{borderfills}'
            f'<hh:charProperties itemCnt="{charpr_cnt}">{charprs}</hh:charProperties>'
            f'{tab_xml}'
            f'{numberings}'
            f'<hh:paraProperties itemCnt="{parapr_cnt}">{paraprs}</hh:paraProperties>'
            f'{styles}'
            f'</hh:refList>'
            f'<hh:compatibleDocument targetProgram="HWP201X">'
            f'<hh:layoutCompatibility/>'
            f'</hh:compatibleDocument>'
            f'<hh:docOption>'
            f'<hh:linkinfo path="" pageInherit="0" footnoteInherit="0"/>'
            f'</hh:docOption>'
            f'<hh:trackchageConfig flags="56"/>'
            f'</hh:head>'
        )

    # ==================================================================
    # secPr (페이지 설정)
    # ==================================================================
    def build_secpr_xml(self, width: int = 59528, height: int = 84186,
                        left: int = 8504, right: int = 8504,
                        top: int = 5668, bottom: int = 4252) -> str:
        return (
            f'<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134"'
            f' tabStop="8000" outlineShapeIDRef="1" memoShapeIDRef="0"'
            f' textVerticalWidthHead="0" masterPageCnt="0">'
            f'<hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/>'
            f'<hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/>'
            f'<hp:visibility hideFirstHeader="0" hideFirstFooter="0"'
            f' hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL"'
            f' hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/>'
            f'<hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/>'
            f'<hp:pagePr landscape="WIDELY" width="{width}" height="{height}" gutterType="LEFT_ONLY">'
            f'<hp:margin header="4252" footer="4252" gutter="0"'
            f' left="{left}" right="{right}" top="{top}" bottom="{bottom}"/>'
            f'</hp:pagePr>'
            f'<hp:footNotePr>'
            f'<hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/>'
            f'<hp:noteLine length="-1" type="SOLID" width="0.12 mm" color="#000000"/>'
            f'<hp:noteSpacing betweenNotes="283" belowLine="567" aboveLine="850"/>'
            f'<hp:numbering type="CONTINUOUS" newNum="1"/>'
            f'<hp:placement place="EACH_COLUMN" beneathText="0"/>'
            f'</hp:footNotePr>'
            f'<hp:endNotePr>'
            f'<hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/>'
            f'<hp:noteLine length="14692344" type="SOLID" width="0.12 mm" color="#000000"/>'
            f'<hp:noteSpacing betweenNotes="0" belowLine="567" aboveLine="850"/>'
            f'<hp:numbering type="CONTINUOUS" newNum="1"/>'
            f'<hp:placement place="END_OF_DOCUMENT" beneathText="0"/>'
            f'</hp:endNotePr>'
            f'<hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER"'
            f' headerInside="0" footerInside="0" fillArea="PAPER">'
            f'<hp:offset left="1417" right="1417" top="1417" bottom="1417"/>'
            f'</hp:pageBorderFill>'
            f'<hp:pageBorderFill type="EVEN" borderFillIDRef="1" textBorder="PAPER"'
            f' headerInside="0" footerInside="0" fillArea="PAPER">'
            f'<hp:offset left="1417" right="1417" top="1417" bottom="1417"/>'
            f'</hp:pageBorderFill>'
            f'<hp:pageBorderFill type="ODD" borderFillIDRef="1" textBorder="PAPER"'
            f' headerInside="0" footerInside="0" fillArea="PAPER">'
            f'<hp:offset left="1417" right="1417" top="1417" bottom="1417"/>'
            f'</hp:pageBorderFill>'
            f'</hp:secPr>'
        )

    # ==================================================================
    # HWPX 패키징
    # ==================================================================
    def pack_hwpx(self, output_path: str, header_xml: str, section_xml: str):
        """HWPX ZIP 패키지 생성"""
        version_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
            '<hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version"'
            ' tagetApplication="WORDPROCESSOR" major="5" minor="1" micro="0"'
            ' buildNumber="1" os="1" xmlVersion="1.2"'
            ' application="Hancom Office Hangul" appVersion="11, 0, 0, 2129 WIN32LEWindows_8"/>'
        )
        container_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
            '<ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container"'
            ' xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf">'
            '<ocf:rootfiles>'
            '<ocf:rootfile full-path="Contents/content.hpf"'
            ' media-type="application/hwpml-package+xml"/>'
            '<ocf:rootfile full-path="Preview/PrvText.txt"'
            ' media-type="text/plain"/>'
            '<ocf:rootfile full-path="META-INF/container.rdf"'
            ' media-type="application/rdf+xml"/>'
            '</ocf:rootfiles>'
            '</ocf:container>'
        )
        container_rdf = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
            '<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">'
            '<rdf:Description rdf:about="">'
            '<ns0:hasPart xmlns:ns0="http://www.hancom.co.kr/hwpml/2016/meta/pkg#"'
            ' rdf:resource="Contents/header.xml"/>'
            '</rdf:Description>'
            '<rdf:Description rdf:about="Contents/header.xml">'
            '<rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#HeaderFile"/>'
            '</rdf:Description>'
            '<rdf:Description rdf:about="">'
            '<ns0:hasPart xmlns:ns0="http://www.hancom.co.kr/hwpml/2016/meta/pkg#"'
            ' rdf:resource="Contents/section0.xml"/>'
            '</rdf:Description>'
            '<rdf:Description rdf:about="Contents/section0.xml">'
            '<rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#SectionFile"/>'
            '</rdf:Description>'
            '<rdf:Description rdf:about="">'
            '<rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#Document"/>'
            '</rdf:Description>'
            '</rdf:RDF>'
        )
        manifest_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
            '<odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>'
        )
        settings_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
            '<ha:HWPApplicationSetting'
            ' xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app"'
            ' xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0">'
            '<ha:CaretPosition listIDRef="0" paraIDRef="0" pos="0"/>'
            '</ha:HWPApplicationSetting>'
        )
        content_hpf = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
            f'<opf:package {ALL_NS}'
            f' version="" unique-identifier="" id="">'
            f'<opf:metadata>'
            f'<opf:title/>'
            f'<opf:language>ko</opf:language>'
            f'</opf:metadata>'
            f'<opf:manifest>'
            f'<opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>'
            f'<opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/>'
            f'<opf:item id="settings" href="settings.xml" media-type="application/xml"/>'
            f'</opf:manifest>'
            f'<opf:spine>'
            f'<opf:itemref idref="header" linear="yes"/>'
            f'<opf:itemref idref="section0" linear="yes"/>'
            f'</opf:spine>'
            f'</opf:package>'
        )

        with zipfile.ZipFile(output_path, "w") as zf:
            mime_info = zipfile.ZipInfo("mimetype")
            mime_info.compress_type = zipfile.ZIP_STORED
            zf.writestr(mime_info, "application/hwp+zip")

            def _w(name, content):
                zf.writestr(name, content.encode("utf-8"),
                            compress_type=zipfile.ZIP_DEFLATED)

            _w("version.xml", version_xml)
            _w("META-INF/container.xml", container_xml)
            _w("META-INF/container.rdf", container_rdf)
            _w("META-INF/manifest.xml", manifest_xml)
            _w("settings.xml", settings_xml)
            _w("Contents/content.hpf", content_hpf)
            _w("Contents/header.xml", header_xml)
            _w("Contents/section0.xml", section_xml)
            _w("Preview/PrvText.txt", "")

            prv_info = zipfile.ZipInfo("Preview/PrvImage.png")
            prv_info.compress_type = zipfile.ZIP_DEFLATED
            zf.writestr(prv_info, _EMPTY_PNG)
