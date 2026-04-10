# -*- coding: utf-8 -*-
"""
시험 문제집 HWPX 생성기

Stage 3: Document Build
- ExamData + ExamSpec → 문제지 HWPX + 해설지 HWPX
"""

import json
from pathlib import Path

from exam_models import ExamData, ExamQuestion, ExamSpec, TextStyle
from hwpx_core import HWPXBuilder, ALL_NS, pt_to_hwpunit, mm_to_hwpunit, _log

# 보기 번호 (원문자)
CIRCLED_NUMBERS = ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩"]
PLAIN_NUMBERS = ["1.", "2.", "3.", "4.", "5.", "6.", "7.", "8.", "9.", "10."]


class ExamHWPXGenerator:
    """시험 문제집 HWPX 생성기"""

    def __init__(self, spec: ExamSpec):
        self.spec = spec

    def generate_exam(self, data: ExamData, output_path: str) -> str:
        """문제지 HWPX 생성 (문제 + 보기만, 해설 없음)"""
        builder = HWPXBuilder(line_spacing=self.spec.question_text.line_spacing)
        self._register_fonts(builder)

        body = self._build_first_paragraph(builder, with_title=True, data=data)

        # 문제 블록 생성
        questions = data.all_questions()
        for i, q in enumerate(questions):
            body += self._question_block_xml(builder, q)

            # 페이지 나눔 제어
            if self.spec.questions_per_page > 0 and (i + 1) % self.spec.questions_per_page == 0:
                if i + 1 < len(questions):  # 마지막이 아닌 경우만
                    body += builder.page_break_paragraph()
            elif self.spec.separator == "line" and i + 1 < len(questions):
                body += builder.separator_line_paragraph()

        body += builder.empty_paragraph()

        return self._finalize(builder, body, output_path)

    def generate_answer_key(self, data: ExamData, output_path: str) -> str:
        """해설지 HWPX 생성 (정답표 + 해설)"""
        builder = HWPXBuilder(line_spacing=self.spec.explanation_text.line_spacing)
        self._register_fonts(builder)

        body = self._build_first_paragraph(builder, with_title=True, data=data,
                                           title_suffix=" - 정답 및 해설")

        questions = data.all_questions()

        # 정답표
        body += self._answer_table_xml(builder, questions)

        # 페이지 나눔
        body += builder.page_break_paragraph()

        # 개별 해설
        body += self._styled_text(builder, "문제별 해설", self.spec.section_header)

        for q in questions:
            body += self._explanation_block_xml(builder, q)

        body += builder.empty_paragraph()

        return self._finalize(builder, body, output_path)

    # ==================================================================
    # 내부 빌더
    # ==================================================================

    def _register_fonts(self, builder: HWPXBuilder):
        """스펙에서 사용하는 모든 폰트 등록"""
        fonts = set()
        for attr_name in ["exam_title", "section_header", "exam_info",
                          "question_number", "question_text", "choice",
                          "explanation_header", "explanation_text", "answer_table"]:
            style: TextStyle = getattr(self.spec, attr_name)
            fonts.add(style.font)
        for font in sorted(fonts):
            builder.register_font(font)

    def _build_first_paragraph(self, builder: HWPXBuilder, with_title: bool,
                                data: ExamData, title_suffix: str = "") -> str:
        """첫 문단 (secPr + colPr + 선택적 제목)"""
        page = self.spec.page
        secpr = builder.build_secpr_xml(
            width=mm_to_hwpunit(page.width_mm) * 2,
            height=mm_to_hwpunit(page.height_mm) * 2,
            left=mm_to_hwpunit(page.margin_left_mm) * 2,
            right=mm_to_hwpunit(page.margin_right_mm) * 2,
            top=mm_to_hwpunit(page.margin_top_mm) * 2,
            bottom=mm_to_hwpunit(page.margin_bottom_mm) * 2,
        )
        colpr = (
            '<hp:ctrl>'
            '<hp:colPr id="" type="NEWSPAPER" layout="LEFT"'
            ' colCount="1" sameSz="1" sameGap="0"/>'
            '</hp:ctrl>'
        )

        if with_title and data.exam_title:
            ts = self.spec.exam_title
            charpr_id = builder.get_charpr_id(
                int(ts.size_pt * 100), ts.color, ts.font, bold=ts.bold)
            parapr_id = builder.get_parapr_id(
                pt_to_hwpunit(ts.left_indent_pt),
                pt_to_hwpunit(ts.space_before_pt),
                pt_to_hwpunit(ts.space_after_pt),
                ts.align,
                line_spacing=ts.line_spacing,
            )
            title_text = data.exam_title + title_suffix
            runs = builder.run_xml(title_text, charpr_id)
            body = (
                f'<hp:p id="0" paraPrIDRef="{parapr_id}"'
                f' styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
                f'<hp:run charPrIDRef="0">{secpr}{colpr}</hp:run>'
                f'{runs}'
                f'</hp:p>'
            )

            # 시험 정보
            if data.subject or data.time_limit_minutes:
                info_parts = []
                if data.subject:
                    info_parts.append(f"시험과목: {data.subject}")
                if data.total_questions:
                    info_parts.append(f"문항 수: {data.total_questions}문항")
                if data.time_limit_minutes:
                    info_parts.append(f"시험시간: {data.time_limit_minutes}분")
                info_text = "  |  ".join(info_parts)
                body += self._styled_text(builder, info_text, self.spec.exam_info)

            return body
        else:
            return (
                f'<hp:p id="0" paraPrIDRef="0" styleIDRef="0"'
                f' pageBreak="0" columnBreak="0" merged="0">'
                f'<hp:run charPrIDRef="0">{secpr}{colpr}</hp:run>'
                f'</hp:p>'
            )

    def _styled_text(self, builder: HWPXBuilder, text: str,
                     style: TextStyle) -> str:
        """TextStyle을 적용한 텍스트 문단"""
        return builder.text_paragraph(
            text,
            font_name=style.font,
            font_size_pt=style.size_pt,
            left_margin_pt=style.left_indent_pt,
            space_before_pt=style.space_before_pt,
            space_after_pt=style.space_after_pt,
            align=style.align,
            hanging_indent_pt=style.hanging_indent_pt,
            bold=style.bold,
            color=style.color,
            line_spacing=style.line_spacing,
        )

    def _question_block_xml(self, builder: HWPXBuilder,
                             question: ExamQuestion) -> str:
        """문제 1개의 XML 블록"""
        xml = ""

        # 문제 번호 + 본문
        qn_style = self.spec.question_number
        qt_style = self.spec.question_text

        # 번호 문단: "1. 경비업법상 ..."
        q_text = f"{question.number}. {question.text}"
        # 번호 부분은 bold, 나머지는 일반
        num_str = f"{question.number}. "
        height_num = int(qn_style.size_pt * 100)
        height_text = int(qt_style.size_pt * 100)

        charpr_num = builder.get_charpr_id(height_num, qn_style.color, qn_style.font, bold=True)
        charpr_text = builder.get_charpr_id(height_text, qt_style.color, qt_style.font)

        parapr_q = builder.get_parapr_id(
            pt_to_hwpunit(qn_style.left_indent_pt),
            pt_to_hwpunit(qn_style.space_before_pt),
            pt_to_hwpunit(qn_style.space_after_pt),
            qn_style.align,
            line_spacing=qt_style.line_spacing,
        )

        runs = builder.run_xml(num_str, charpr_num) + builder.run_xml(question.text, charpr_text)
        xml += builder.paragraph_xml(runs, parapr_q)

        # 보기
        numbering = CIRCLED_NUMBERS if self.spec.choice_numbering == "circled" else PLAIN_NUMBERS
        ch_style = self.spec.choice
        for ci, choice_text in enumerate(question.choices):
            prefix = numbering[ci] if ci < len(numbering) else f"{ci+1}."
            full_choice = f"{prefix} {choice_text}"
            xml += builder.text_paragraph(
                full_choice,
                font_name=ch_style.font,
                font_size_pt=ch_style.size_pt,
                left_margin_pt=ch_style.left_indent_pt,
                space_before_pt=ch_style.space_before_pt,
                space_after_pt=ch_style.space_after_pt,
                align=ch_style.align,
                line_spacing=ch_style.line_spacing,
            )

        return xml

    def _answer_table_xml(self, builder: HWPXBuilder,
                           questions: list[ExamQuestion]) -> str:
        """정답표 테이블"""
        at_style = self.spec.answer_table

        # 제목
        xml = self._styled_text(builder, "정 답 표", self.spec.section_header)

        # 10열 테이블 구성
        cols_per_row = 10
        headers = ["번호", "정답"] * (cols_per_row // 2)

        rows = []
        for start in range(0, len(questions), cols_per_row // 2):
            row = []
            for offset in range(cols_per_row // 2):
                idx = start + offset
                if idx < len(questions):
                    q = questions[idx]
                    row.append(str(q.number))
                    if self.spec.choice_numbering == "circled" and 0 < q.correct_answer <= len(CIRCLED_NUMBERS):
                        row.append(CIRCLED_NUMBERS[q.correct_answer - 1])
                    else:
                        row.append(str(q.correct_answer))
                else:
                    row.extend(["", ""])
            rows.append(row)

        xml += builder.table_paragraph_xml(
            headers, rows,
            font_name=at_style.font,
            font_size=at_style.size_pt,
            line_spacing=160,
        )

        return xml

    def _explanation_block_xml(self, builder: HWPXBuilder,
                                question: ExamQuestion) -> str:
        """해설 1개의 XML 블록"""
        xml = ""

        eh_style = self.spec.explanation_header
        et_style = self.spec.explanation_text

        # 문제 번호 + 정답
        numbering = CIRCLED_NUMBERS if self.spec.choice_numbering == "circled" else PLAIN_NUMBERS
        if 0 < question.correct_answer <= len(numbering):
            answer_str = numbering[question.correct_answer - 1]
        else:
            answer_str = str(question.correct_answer)

        header_text = f"{question.number}번  정답: {answer_str}"
        xml += self._styled_text(builder, header_text, eh_style)

        # 해설 텍스트
        if question.explanation:
            xml += self._styled_text(builder, question.explanation, et_style)

        # 관련 조문
        if question.source_reference:
            ref_text = f"[관련 조문] {question.source_reference}"
            ref_style = TextStyle(
                font=et_style.font, size_pt=et_style.size_pt - 0.5,
                color="#666666", align="LEFT",
                space_before_pt=1, space_after_pt=4,
                line_spacing=et_style.line_spacing,
            )
            xml += self._styled_text(builder, ref_text, ref_style)

        return xml

    def _finalize(self, builder: HWPXBuilder, body: str, output_path: str) -> str:
        """header.xml + section0.xml 조합 후 HWPX 패키징"""
        _log("[Build] Building header.xml...")
        header_xml = builder.build_header_xml()

        _log("[Build] Building section0.xml...")
        section_xml = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
            f'<hs:sec {ALL_NS}>'
            f'{body}'
            f'</hs:sec>'
        )

        _log("[Build] Packing HWPX...")
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        builder.pack_hwpx(output_path, header_xml, section_xml)
        _log(f"[Build] Done: {output_path}")
        return output_path


def generate_from_json(questions_path: str, spec_path: str,
                       output_dir: str) -> tuple[str, str]:
    """JSON 파일에서 읽어서 문제지 + 해설지 생성

    Returns:
        (문제지_경로, 해설지_경로)
    """
    with open(questions_path, "r", encoding="utf-8") as f:
        data = ExamData.from_dict(json.load(f))
    with open(spec_path, "r", encoding="utf-8") as f:
        spec = ExamSpec.from_dict(json.load(f))

    gen = ExamHWPXGenerator(spec)

    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # 파일명 생성
    base_name = data.exam_title or "모의고사"
    base_name = base_name.replace(" ", "_")
    exam_path = str(out_dir / f"{base_name}.hwpx")
    answer_path = str(out_dir / f"{base_name}_해설.hwpx")

    gen.generate_exam(data, exam_path)
    gen.generate_answer_key(data, answer_path)

    return exam_path, answer_path
