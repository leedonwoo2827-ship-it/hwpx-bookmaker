# -*- coding: utf-8 -*-
"""
HWPX 스타일 스펙 추출기 — 참고 문제집 HWPX를 파싱하여 exam-spec.json 생성

Stage 1: Template Analysis
- HWPX ZIP 해체 → XML 파싱 → 스타일 패턴 분류 → exam-spec.json
"""

import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from exam_models import ExamSpec, TextStyle, PageLayout

# HWPX 네임스페이스 매핑
NS = {
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
    "ha": "http://www.hancom.co.kr/hwpml/2011/app",
}

# 보기 번호 패턴
CHOICE_PATTERN = re.compile(r'^[①②③④⑤⑥⑦⑧⑨⑩]\s')
QUESTION_NUM_PATTERN = re.compile(r'^\d{1,3}[\.\)]\s')


def _hwpunit_to_pt(v: int) -> float:
    return v / 50.0

def _height_to_pt(v: int) -> float:
    return v / 100.0


class HWPXSpecExtractor:
    """HWPX 파일에서 시험 문제집 스타일 스펙을 추출한다."""

    def __init__(self, hwpx_path: str):
        self.hwpx_path = Path(hwpx_path)
        self._fonts: dict[str, str] = {}      # font_id → face name
        self._charpr: dict[str, dict] = {}     # charpr_id → {height, bold, color, font_id}
        self._parapr: dict[str, dict] = {}     # parapr_id → {align, leftMargin, ...}

    def extract(self) -> ExamSpec:
        """HWPX 파일을 분석하여 ExamSpec을 반환"""
        with zipfile.ZipFile(self.hwpx_path, "r") as zf:
            # 1. header.xml — 폰트, charPr, paraPr 정의
            header_xml = self._read_xml(zf, "Contents/header.xml")
            if header_xml is not None:
                self._parse_header(header_xml)

            # 2. section0.xml — 실제 문단 내용으로 패턴 분류
            section_xml = self._read_xml(zf, "Contents/section0.xml")
            paragraphs = []
            if section_xml is not None:
                paragraphs = self._extract_paragraphs(section_xml)

            # 3. 페이지 설정 추출
            page_layout = PageLayout()
            if section_xml is not None:
                page_layout = self._extract_page_layout(section_xml)

        # 4. 문단 패턴 분류 및 스타일 매핑
        return self._classify_and_build_spec(paragraphs, page_layout)

    def _read_xml(self, zf: zipfile.ZipFile, name: str) -> ET.Element | None:
        try:
            data = zf.read(name)
            return ET.fromstring(data)
        except (KeyError, ET.ParseError):
            return None

    def _parse_header(self, root: ET.Element):
        """header.xml에서 폰트, charPr, paraPr 추출"""
        # 폰트 매핑 (fontface → font 엘리먼트)
        for fontface in root.iter(f'{{{NS["hh"]}}}fontface'):
            if fontface.get("lang") == "HANGUL":
                for font in fontface.iter(f'{{{NS["hh"]}}}font'):
                    fid = font.get("id", "")
                    face = font.get("face", "")
                    self._fonts[fid] = face

        # CharPr
        for cp in root.iter(f'{{{NS["hh"]}}}charPr'):
            cid = cp.get("id", "")
            height = int(cp.get("height", "1000"))
            bold = cp.get("bold", "0") == "1"
            color = cp.get("textColor", "#000000")
            font_ref = cp.find(f'{{{NS["hh"]}}}fontRef')
            font_id = font_ref.get("hangul", "0") if font_ref is not None else "0"
            self._charpr[cid] = {
                "height": height, "bold": bold, "color": color,
                "font_id": font_id,
                "font_name": self._fonts.get(font_id, "바탕"),
            }

        # ParaPr
        for pp in root.iter(f'{{{NS["hh"]}}}paraPr'):
            pid = pp.get("id", "")
            align_el = pp.find(f'{{{NS["hh"]}}}align')
            align = align_el.get("horizontal", "JUSTIFY") if align_el is not None else "JUSTIFY"
            margin_el = pp.find(f'{{{NS["hh"]}}}margin')
            left_margin = 0
            space_before = 0
            space_after = 0
            indent = 0
            if margin_el is not None:
                left_el = margin_el.find(f'{{{NS["hc"]}}}left')
                if left_el is not None:
                    left_margin = int(left_el.get("value", "0"))
                prev_el = margin_el.find(f'{{{NS["hc"]}}}prev')
                if prev_el is not None:
                    space_before = int(prev_el.get("value", "0"))
                next_el = margin_el.find(f'{{{NS["hc"]}}}next')
                if next_el is not None:
                    space_after = int(next_el.get("value", "0"))
                intent_el = margin_el.find(f'{{{NS["hc"]}}}intent')
                if intent_el is not None:
                    indent = int(intent_el.get("value", "0"))
            ls_el = pp.find(f'{{{NS["hh"]}}}lineSpacing')
            line_spacing = int(ls_el.get("value", "160")) if ls_el is not None else 160

            self._parapr[pid] = {
                "align": align, "left_margin": left_margin,
                "space_before": space_before, "space_after": space_after,
                "indent": indent, "line_spacing": line_spacing,
            }

    def _extract_paragraphs(self, section_root: ET.Element) -> list[dict]:
        """section0.xml에서 문단 정보를 추출"""
        paragraphs = []
        for p in section_root.iter(f'{{{NS["hp"]}}}p'):
            parapr_id = p.get("paraPrIDRef", "0")

            # 텍스트 추출
            text_parts = []
            for run in p.iter(f'{{{NS["hp"]}}}run'):
                charpr_id = run.get("charPrIDRef", "0")
                for t in run.iter(f'{{{NS["hp"]}}}t'):
                    if t.text:
                        text_parts.append((t.text, charpr_id))

            full_text = "".join(t for t, _ in text_parts)
            first_charpr = text_parts[0][1] if text_parts else "0"

            paragraphs.append({
                "text": full_text,
                "charpr_id": first_charpr,
                "parapr_id": parapr_id,
            })
        return paragraphs

    def _extract_page_layout(self, section_root: ET.Element) -> PageLayout:
        """section0.xml에서 페이지 레이아웃 추출"""
        layout = PageLayout()
        for secpr in section_root.iter(f'{{{NS["hp"]}}}secPr'):
            page_pr = secpr.find(f'{{{NS["hp"]}}}pagePr')
            if page_pr is not None:
                w = int(page_pr.get("width", "59528"))
                h = int(page_pr.get("height", "84186"))
                layout.width_mm = round(w * 25.4 / 7200, 1)
                layout.height_mm = round(h * 25.4 / 7200, 1)
                margin = page_pr.find(f'{{{NS["hp"]}}}margin')
                if margin is not None:
                    layout.margin_top_mm = round(int(margin.get("top", "5668")) * 25.4 / 7200, 1)
                    layout.margin_bottom_mm = round(int(margin.get("bottom", "4252")) * 25.4 / 7200, 1)
                    layout.margin_left_mm = round(int(margin.get("left", "8504")) * 25.4 / 7200, 1)
                    layout.margin_right_mm = round(int(margin.get("right", "8504")) * 25.4 / 7200, 1)
            break
        return layout

    def _classify_and_build_spec(self, paragraphs: list[dict],
                                  page_layout: PageLayout) -> ExamSpec:
        """문단 텍스트 패턴으로 분류하여 ExamSpec 생성"""
        spec = ExamSpec(page=page_layout)

        # 패턴별 스타일 수집
        question_styles: list[dict] = []
        choice_styles: list[dict] = []
        body_styles: list[dict] = []
        title_styles: list[dict] = []

        for para in paragraphs:
            text = para["text"].strip()
            if not text:
                continue

            charpr = self._charpr.get(para["charpr_id"], {})
            parapr = self._parapr.get(para["parapr_id"], {})

            style_info = {
                "font": charpr.get("font_name", "바탕"),
                "size_pt": _height_to_pt(charpr.get("height", 1000)),
                "bold": charpr.get("bold", False),
                "align": parapr.get("align", "JUSTIFY"),
                "line_spacing": parapr.get("line_spacing", 160),
                "left_indent_pt": _hwpunit_to_pt(parapr.get("left_margin", 0)),
                "space_before_pt": _hwpunit_to_pt(parapr.get("space_before", 0)),
                "space_after_pt": _hwpunit_to_pt(parapr.get("space_after", 0)),
            }

            # 분류
            if QUESTION_NUM_PATTERN.match(text):
                question_styles.append(style_info)
            elif CHOICE_PATTERN.match(text):
                choice_styles.append(style_info)
            elif _height_to_pt(charpr.get("height", 1000)) >= 13:
                title_styles.append(style_info)
            else:
                body_styles.append(style_info)

        # 가장 빈번한 스타일을 대표값으로 사용
        if question_styles:
            rep = self._most_common_style(question_styles)
            spec.question_number = TextStyle(**rep)
            spec.question_text = TextStyle(**{**rep, "bold": False})
        if choice_styles:
            rep = self._most_common_style(choice_styles)
            spec.choice = TextStyle(**rep)
            # 보기 번호 형식 감지
            for para in paragraphs:
                if CHOICE_PATTERN.match(para["text"].strip()):
                    spec.choice_numbering = "circled"
                    break
        if title_styles:
            rep = self._most_common_style(title_styles)
            spec.exam_title = TextStyle(**rep)
            spec.section_header = TextStyle(**rep)

        # 페이지당 문제 수 추정
        q_count = len(question_styles)
        if q_count > 0 and paragraphs:
            total_paras = len([p for p in paragraphs if p["text"].strip()])
            estimated_pages = max(1, total_paras // 30)
            spec.questions_per_page = max(1, q_count // estimated_pages)

        return spec

    @staticmethod
    def _most_common_style(styles: list[dict]) -> dict:
        """가장 빈번한 스타일 조합을 반환"""
        if not styles:
            return {}
        # 폰트+크기+bold 조합으로 빈도 계산
        from collections import Counter
        keys = Counter()
        for s in styles:
            key = (s["font"], s["size_pt"], s["bold"])
            keys[key] += 1
        most_common = keys.most_common(1)[0][0]
        # 해당 조합의 첫 번째 스타일 반환 (나머지 속성 포함)
        for s in styles:
            if (s["font"], s["size_pt"], s["bold"]) == most_common:
                return s
        return styles[0]


def extract_spec(hwpx_path: str) -> ExamSpec:
    """편의 함수: HWPX 파일에서 ExamSpec 추출"""
    return HWPXSpecExtractor(hwpx_path).extract()


def extract_and_save(hwpx_path: str, output_path: str) -> str:
    """HWPX 파일을 분석하여 exam-spec.json으로 저장"""
    spec = extract_spec(hwpx_path)
    spec_dict = spec.to_dict()
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(spec_dict, f, ensure_ascii=False, indent=2)
    return output_path
