# -*- coding: utf-8 -*-
"""
데이터 모델 — 시험 문제집 생성에 사용되는 모든 데이터 구조 정의
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


# ---------------------------------------------------------------------------
# Exam Spec (Stage 1 출력 — 템플릿 분석 결과)
# ---------------------------------------------------------------------------

@dataclass
class TextStyle:
    font: str = "바탕"
    size_pt: float = 10.0
    bold: bool = False
    italic: bool = False
    color: str = "#000000"
    align: str = "JUSTIFY"          # JUSTIFY, LEFT, CENTER, RIGHT
    line_spacing: int = 160         # 퍼센트 (160 = 160%)
    left_indent_pt: float = 0.0
    hanging_indent_pt: float = 0.0
    space_before_pt: float = 0.0
    space_after_pt: float = 0.0


@dataclass
class PageLayout:
    width_mm: float = 210.0         # A4
    height_mm: float = 297.0
    margin_top_mm: float = 20.0
    margin_bottom_mm: float = 15.0
    margin_left_mm: float = 20.0
    margin_right_mm: float = 20.0


@dataclass
class ExamSpec:
    """시험 문제집의 전체 서식 스펙. analyze_template의 출력."""
    page: PageLayout = field(default_factory=PageLayout)

    # 각 요소별 스타일
    exam_title: TextStyle = field(default_factory=lambda: TextStyle(
        font="맑은 고딕", size_pt=16, bold=True, align="CENTER",
        space_before_pt=20, space_after_pt=10
    ))
    section_header: TextStyle = field(default_factory=lambda: TextStyle(
        font="맑은 고딕", size_pt=14, bold=True, align="CENTER",
        space_before_pt=15, space_after_pt=8
    ))
    exam_info: TextStyle = field(default_factory=lambda: TextStyle(
        font="맑은 고딕", size_pt=10, bold=False, align="CENTER",
        space_before_pt=2, space_after_pt=2
    ))
    question_number: TextStyle = field(default_factory=lambda: TextStyle(
        font="맑은 고딕", size_pt=11, bold=True, align="LEFT",
        space_before_pt=8, space_after_pt=2
    ))
    question_text: TextStyle = field(default_factory=lambda: TextStyle(
        font="바탕", size_pt=10, align="JUSTIFY", line_spacing=160,
        space_before_pt=0, space_after_pt=2
    ))
    choice: TextStyle = field(default_factory=lambda: TextStyle(
        font="바탕", size_pt=10, align="LEFT", line_spacing=150,
        left_indent_pt=10, space_before_pt=0, space_after_pt=0
    ))
    explanation_header: TextStyle = field(default_factory=lambda: TextStyle(
        font="맑은 고딕", size_pt=11, bold=True, align="LEFT",
        space_before_pt=6, space_after_pt=2
    ))
    explanation_text: TextStyle = field(default_factory=lambda: TextStyle(
        font="바탕", size_pt=9.5, align="JUSTIFY", line_spacing=150,
        space_before_pt=0, space_after_pt=2
    ))
    answer_table: TextStyle = field(default_factory=lambda: TextStyle(
        font="맑은 고딕", size_pt=10, align="CENTER"
    ))

    # 문제 블록 레이아웃
    choice_numbering: str = "circled"       # circled (①②③④⑤) 또는 number (1. 2. 3.)
    choice_count: int = 5                   # 보기 개수 (4 또는 5)
    questions_per_page: int = 5             # 페이지당 문제 수 (0=자동)
    separator: str = "line"                 # "line", "space", "none"

    def to_dict(self) -> dict:
        """JSON 직렬화용 딕셔너리"""
        def _style_dict(s: TextStyle) -> dict:
            return {
                "font": s.font, "size_pt": s.size_pt, "bold": s.bold,
                "italic": s.italic, "color": s.color, "align": s.align,
                "line_spacing": s.line_spacing,
                "left_indent_pt": s.left_indent_pt,
                "hanging_indent_pt": s.hanging_indent_pt,
                "space_before_pt": s.space_before_pt,
                "space_after_pt": s.space_after_pt,
            }
        def _page_dict(p: PageLayout) -> dict:
            return {
                "width_mm": p.width_mm, "height_mm": p.height_mm,
                "margin_top_mm": p.margin_top_mm,
                "margin_bottom_mm": p.margin_bottom_mm,
                "margin_left_mm": p.margin_left_mm,
                "margin_right_mm": p.margin_right_mm,
            }
        return {
            "page": _page_dict(self.page),
            "exam_title": _style_dict(self.exam_title),
            "section_header": _style_dict(self.section_header),
            "exam_info": _style_dict(self.exam_info),
            "question_number": _style_dict(self.question_number),
            "question_text": _style_dict(self.question_text),
            "choice": _style_dict(self.choice),
            "explanation_header": _style_dict(self.explanation_header),
            "explanation_text": _style_dict(self.explanation_text),
            "answer_table": _style_dict(self.answer_table),
            "choice_numbering": self.choice_numbering,
            "choice_count": self.choice_count,
            "questions_per_page": self.questions_per_page,
            "separator": self.separator,
        }

    @classmethod
    def from_dict(cls, d: dict) -> ExamSpec:
        """JSON 딕셔너리에서 ExamSpec 복원"""
        def _parse_style(sd: dict) -> TextStyle:
            return TextStyle(**{k: v for k, v in sd.items() if k in TextStyle.__dataclass_fields__})
        def _parse_page(pd: dict) -> PageLayout:
            return PageLayout(**{k: v for k, v in pd.items() if k in PageLayout.__dataclass_fields__})

        spec = cls()
        if "page" in d:
            spec.page = _parse_page(d["page"])
        for key in ["exam_title", "section_header", "exam_info", "question_number",
                     "question_text", "choice", "explanation_header",
                     "explanation_text", "answer_table"]:
            if key in d:
                setattr(spec, key, _parse_style(d[key]))
        for key in ["choice_numbering", "choice_count", "questions_per_page", "separator"]:
            if key in d:
                setattr(spec, key, d[key])
        return spec


# ---------------------------------------------------------------------------
# Exam Content (Stage 2 출력 — 문제 데이터)
# ---------------------------------------------------------------------------

@dataclass
class ExamQuestion:
    """시험 문제 1개"""
    number: int = 0
    text: str = ""
    choices: list[str] = field(default_factory=list)
    correct_answer: int = 0             # 1-based index
    explanation: str = ""
    source_reference: str = ""          # 관련 조문/출처
    difficulty: str = "중"              # 상, 중, 하


@dataclass
class ExamSection:
    """시험 섹션 (과목 구분)"""
    title: str = ""
    questions: list[ExamQuestion] = field(default_factory=list)


@dataclass
class ExamData:
    """시험 전체 데이터"""
    exam_title: str = ""
    subject: str = ""
    total_questions: int = 0
    time_limit_minutes: int = 0
    sections: list[ExamSection] = field(default_factory=list)

    def all_questions(self) -> list[ExamQuestion]:
        """모든 섹션의 문제를 순서대로 반환"""
        result = []
        for section in self.sections:
            result.extend(section.questions)
        return result

    def to_dict(self) -> dict:
        return {
            "exam_title": self.exam_title,
            "subject": self.subject,
            "total_questions": self.total_questions,
            "time_limit_minutes": self.time_limit_minutes,
            "sections": [
                {
                    "title": s.title,
                    "questions": [
                        {
                            "number": q.number, "text": q.text,
                            "choices": q.choices,
                            "correct_answer": q.correct_answer,
                            "explanation": q.explanation,
                            "source_reference": q.source_reference,
                            "difficulty": q.difficulty,
                        }
                        for q in s.questions
                    ]
                }
                for s in self.sections
            ]
        }

    @classmethod
    def from_dict(cls, d: dict) -> ExamData:
        sections = []
        for sd in d.get("sections", []):
            questions = []
            for qd in sd.get("questions", []):
                questions.append(ExamQuestion(
                    number=qd.get("number", 0),
                    text=qd.get("text", ""),
                    choices=qd.get("choices", []),
                    correct_answer=qd.get("correct_answer", 0),
                    explanation=qd.get("explanation", ""),
                    source_reference=qd.get("source_reference", ""),
                    difficulty=qd.get("difficulty", "중"),
                ))
            sections.append(ExamSection(title=sd.get("title", ""), questions=questions))
        return cls(
            exam_title=d.get("exam_title", ""),
            subject=d.get("subject", ""),
            total_questions=d.get("total_questions", 0),
            time_limit_minutes=d.get("time_limit_minutes", 0),
            sections=sections,
        )
