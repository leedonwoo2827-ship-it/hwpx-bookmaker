# -*- coding: utf-8 -*-
"""
HWPX 텍스트 추출기 — 법령/교재 HWPX에서 구조화된 텍스트 추출

법제처 HWPX, 교재 HWPX 등 다양한 HWPX 파일에서 텍스트를 추출하여
조/항/호/목 체계를 보존한 구조화 텍스트로 변환한다.
"""

import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
}

# 조문 번호 패턴
ARTICLE_PATTERN = re.compile(r'^제(\d+)조(?:의\d+)?(?:\(([^)]+)\))?')
PARAGRAPH_PATTERN = re.compile(r'^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]')
SUBPARAGRAPH_PATTERN = re.compile(r'^\d+\.')
ITEM_PATTERN = re.compile(r'^[가나다라마바사아자차카타파하]\.')


def extract_text_from_hwpx(hwpx_path: str) -> str:
    """HWPX 파일에서 전체 텍스트를 추출 (줄바꿈 유지)"""
    path = Path(hwpx_path)
    if not path.exists():
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {hwpx_path}")

    paragraphs = []
    with zipfile.ZipFile(path, "r") as zf:
        # section 파일들을 순서대로 처리
        section_files = sorted(
            [n for n in zf.namelist() if n.startswith("Contents/section") and n.endswith(".xml")]
        )
        if not section_files:
            raise ValueError(f"HWPX 파일에 section XML이 없습니다: {hwpx_path}")

        for section_file in section_files:
            data = zf.read(section_file)
            root = ET.fromstring(data)
            for p in root.iter(f'{{{NS["hp"]}}}p'):
                text = _extract_paragraph_text(p)
                if text.strip():
                    paragraphs.append(text)

    return "\n".join(paragraphs)


def extract_structured_law(hwpx_path: str) -> list[dict]:
    """법령 HWPX에서 조문 체계를 보존한 구조화 텍스트를 추출

    Returns:
        [
            {
                "type": "chapter", "title": "제1장 총칙",
                "articles": [
                    {
                        "type": "article", "number": 1,
                        "title": "목적",
                        "text": "이 법은 경비업의...",
                        "paragraphs": [
                            {"number": "①", "text": "..."},
                        ]
                    }
                ]
            }
        ]
    """
    raw_text = extract_text_from_hwpx(hwpx_path)
    lines = raw_text.split("\n")

    result = []
    current_chapter = None
    current_article = None

    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        # 장 제목 감지 (제N장 ...)
        chapter_match = re.match(r'^제(\d+)장\s+(.+)', stripped)
        if chapter_match:
            current_chapter = {
                "type": "chapter",
                "number": int(chapter_match.group(1)),
                "title": f"제{chapter_match.group(1)}장 {chapter_match.group(2)}",
                "articles": [],
            }
            result.append(current_chapter)
            current_article = None
            continue

        # 조문 감지 (제N조(제목))
        article_match = ARTICLE_PATTERN.match(stripped)
        if article_match:
            current_article = {
                "type": "article",
                "number": int(article_match.group(1)),
                "title": article_match.group(2) or "",
                "text": stripped,
                "paragraphs": [],
            }
            if current_chapter:
                current_chapter["articles"].append(current_article)
            else:
                # 장 없이 바로 조문이 오는 경우
                if not result or result[-1].get("type") != "chapter":
                    current_chapter = {
                        "type": "chapter", "number": 0,
                        "title": "", "articles": [],
                    }
                    result.append(current_chapter)
                current_chapter["articles"].append(current_article)
            continue

        # 항 감지 (①②③...)
        if PARAGRAPH_PATTERN.match(stripped) and current_article:
            current_article["paragraphs"].append({
                "number": stripped[0],
                "text": stripped,
            })
            continue

        # 호/목 또는 기타 텍스트 → 현재 조문에 추가
        if current_article:
            if current_article["paragraphs"]:
                current_article["paragraphs"][-1]["text"] += "\n" + stripped
            else:
                current_article["text"] += "\n" + stripped

    return result


def format_structured_law(data: list[dict]) -> str:
    """구조화된 법령 데이터를 읽기 좋은 텍스트로 포맷"""
    lines = []
    for chapter in data:
        if chapter["title"]:
            lines.append(f"\n{'='*60}")
            lines.append(f"  {chapter['title']}")
            lines.append(f"{'='*60}\n")

        for article in chapter["articles"]:
            lines.append(f"\n{article['text']}")
            for para in article["paragraphs"]:
                lines.append(f"  {para['text']}")

    return "\n".join(lines)


def _extract_paragraph_text(p_element: ET.Element) -> str:
    """paragraph 엘리먼트에서 텍스트 추출"""
    parts = []
    for run in p_element.iter(f'{{{NS["hp"]}}}run'):
        for t in run.iter(f'{{{NS["hp"]}}}t'):
            if t.text:
                parts.append(t.text)
            if t.tail:
                parts.append(t.tail)
    return "".join(parts)
