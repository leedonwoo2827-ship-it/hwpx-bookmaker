# -*- coding: utf-8 -*-
"""
시험 문제집 HWPX 생성기 MCP Server

Claude Desktop에서 사용하는 MCP 서버.
참고 문제집 분석, 법령 텍스트 추출, 문제 저장, HWPX 문서 생성을 수행합니다.

실행: python server.py
"""

import json
import os
import sys
from pathlib import Path

BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR / "src"))

from mcp.server.fastmcp import FastMCP
from exam_models import ExamData, ExamSpec
from spec_extractor import extract_spec
from law_parser import extract_text_from_hwpx, extract_structured_law, format_structured_law
from exam_generator import ExamHWPXGenerator

mcp = FastMCP("hwpx-bookmaker")


# ---------------------------------------------------------------------------
# 유틸리티
# ---------------------------------------------------------------------------

def _resolve_path(file_path: str, project_dir: str = "") -> Path:
    """절대경로 변환. 상대경로면 project_dir 기준."""
    p = Path(file_path)
    if p.is_absolute():
        return p
    if project_dir:
        return Path(project_dir) / file_path
    return Path.home() / "Documents" / file_path


def _get_spec_path(project_dir: str) -> Path:
    return Path(project_dir) / "exam-spec.json"


def _load_spec(project_dir: str) -> ExamSpec:
    spec_path = _get_spec_path(project_dir)
    if spec_path.exists():
        with open(spec_path, "r", encoding="utf-8") as f:
            return ExamSpec.from_dict(json.load(f))
    # 기본 스펙 사용
    default_spec = BASE_DIR / "exam-spec-default.json"
    if default_spec.exists():
        with open(default_spec, "r", encoding="utf-8") as f:
            return ExamSpec.from_dict(json.load(f))
    return ExamSpec()


# ---------------------------------------------------------------------------
# MCP 도구
# ---------------------------------------------------------------------------

@mcp.tool()
def analyze_template(
    hwpx_path: str,
    project_dir: str = "",
) -> str:
    """참고 시험 문제집 HWPX 파일을 분석하여 스타일 스펙(exam-spec.json)을 생성합니다.

    참고 문제집의 폰트, 크기, 간격, 보기 형식 등을 자동으로 추출합니다.
    여러 파일을 분석하면 마지막 결과가 저장됩니다.

    Args:
        hwpx_path: 분석할 HWPX 파일 경로 (reference/exams/ 폴더의 파일)
        project_dir: 프로젝트 폴더 경로 (예: C:/Users/ubion/Documents/testmaker/260410-1)

    Returns:
        추출된 스타일 스펙 요약
    """
    resolved = _resolve_path(hwpx_path, project_dir)
    if not resolved.exists():
        return f"오류: 파일을 찾을 수 없습니다: {resolved}"

    try:
        spec = extract_spec(str(resolved))
        spec_dict = spec.to_dict()

        # 저장
        if project_dir:
            out_path = _get_spec_path(project_dir)
            out_path.parent.mkdir(parents=True, exist_ok=True)
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(spec_dict, f, ensure_ascii=False, indent=2)
            save_msg = f"\n저장: {out_path}"
        else:
            save_msg = "\n(project_dir 미지정 — 파일 저장 안 됨)"

        # 요약
        summary = (
            f"분석 완료: {resolved.name}\n"
            f"페이지: {spec.page.width_mm}×{spec.page.height_mm}mm\n"
            f"문제번호: {spec.question_number.font} {spec.question_number.size_pt}pt"
            f" {'굵게' if spec.question_number.bold else ''}\n"
            f"문제본문: {spec.question_text.font} {spec.question_text.size_pt}pt\n"
            f"보기: {spec.choice.font} {spec.choice.size_pt}pt, "
            f"번호형식={spec.choice_numbering}\n"
            f"페이지당 문제: {spec.questions_per_page}개\n"
            f"{save_msg}"
        )
        return summary

    except Exception as e:
        return f"오류: {type(e).__name__}: {e}"


@mcp.tool()
def extract_source_text(
    hwpx_path: str,
    project_dir: str = "",
    structured: bool = True,
) -> str:
    """HWPX 파일에서 텍스트를 추출합니다. 법령, 교재 등 다양한 HWPX 파일에 사용 가능합니다.

    법제처에서 다운로드한 HWPX 파일의 경우 조/항/호/목 체계를 보존하여 추출합니다.

    Args:
        hwpx_path: 추출할 HWPX 파일 경로
        project_dir: 프로젝트 폴더 경로
        structured: True이면 법령 구조화 추출 (장/조/항 체계), False이면 단순 텍스트

    Returns:
        추출된 텍스트
    """
    resolved = _resolve_path(hwpx_path, project_dir)
    if not resolved.exists():
        return f"오류: 파일을 찾을 수 없습니다: {resolved}"

    try:
        if structured:
            data = extract_structured_law(str(resolved))
            text = format_structured_law(data)
            chapter_count = len(data)
            article_count = sum(len(ch.get("articles", [])) for ch in data)
            return (
                f"추출 완료: {resolved.name}\n"
                f"장: {chapter_count}개, 조문: {article_count}개\n"
                f"{'='*60}\n\n{text}"
            )
        else:
            text = extract_text_from_hwpx(str(resolved))
            line_count = len(text.split("\n"))
            return (
                f"추출 완료: {resolved.name} ({line_count}줄)\n"
                f"{'='*60}\n\n{text}"
            )

    except Exception as e:
        return f"오류: {type(e).__name__}: {e}"


@mcp.tool()
def get_exam_spec(
    project_dir: str,
) -> str:
    """현재 프로젝트의 exam-spec.json 스타일 스펙을 반환합니다.

    Args:
        project_dir: 프로젝트 폴더 경로

    Returns:
        현재 스타일 스펙 JSON
    """
    spec = _load_spec(project_dir)
    return json.dumps(spec.to_dict(), ensure_ascii=False, indent=2)


@mcp.tool()
def update_exam_spec(
    updates_json: str,
    project_dir: str,
) -> str:
    """exam-spec.json의 스타일 스펙을 부분 업데이트합니다.

    수정하고 싶은 속성만 JSON으로 전달하면 해당 부분만 업데이트됩니다.

    updates_json 예시:
    {
      "question_number": { "font": "맑은 고딕", "size_pt": 12, "bold": true },
      "choice": { "left_indent_pt": 15 },
      "questions_per_page": 4
    }

    Args:
        updates_json: 수정할 스타일 속성 (JSON 문자열)
        project_dir: 프로젝트 폴더 경로

    Returns:
        업데이트 결과
    """
    try:
        updates = json.loads(updates_json)
    except json.JSONDecodeError as e:
        return f"오류: 유효하지 않은 JSON: {e}"

    spec = _load_spec(project_dir)
    spec_dict = spec.to_dict()

    # 딥 머지
    for key, value in updates.items():
        if isinstance(value, dict) and key in spec_dict and isinstance(spec_dict[key], dict):
            spec_dict[key].update(value)
        else:
            spec_dict[key] = value

    # 저장
    spec_path = _get_spec_path(project_dir)
    spec_path.parent.mkdir(parents=True, exist_ok=True)
    with open(spec_path, "w", encoding="utf-8") as f:
        json.dump(spec_dict, f, ensure_ascii=False, indent=2)

    return f"스펙 업데이트 완료: {spec_path}\n\n{json.dumps(spec_dict, ensure_ascii=False, indent=2)}"


@mcp.tool()
def save_questions(
    questions_json: str,
    project_dir: str,
    filename: str = "bank.json",
) -> str:
    """시험 문제 데이터를 JSON 파일로 저장합니다.

    Claude가 생성한 문제를 구조화된 JSON으로 저장합니다.
    기존 파일이 있으면 덮어쓰기 또는 병합할 수 있습니다.

    questions_json 형식:
    {
      "exam_title": "경비지도사 경비업법 최종모의고사 제1회",
      "subject": "경비업법",
      "total_questions": 40,
      "time_limit_minutes": 50,
      "sections": [
        {
          "title": "경비업법",
          "questions": [
            {
              "number": 1,
              "text": "경비업법의 목적으로 가장 적절한 것은?",
              "choices": ["선지1", "선지2", "선지3", "선지4", "선지5"],
              "correct_answer": 1,
              "explanation": "경비업법 제1조에 따르면...",
              "source_reference": "경비업법 제1조",
              "difficulty": "하"
            }
          ]
        }
      ]
    }

    Args:
        questions_json: 문제 데이터 (JSON 문자열)
        project_dir: 프로젝트 폴더 경로
        filename: 저장할 파일명 (기본: bank.json)

    Returns:
        저장 결과
    """
    try:
        data_dict = json.loads(questions_json)
    except json.JSONDecodeError as e:
        return f"오류: 유효하지 않은 JSON: {e}"

    data = ExamData.from_dict(data_dict)

    out_dir = Path(project_dir) / "questions"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / filename

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data.to_dict(), f, ensure_ascii=False, indent=2)

    q_count = len(data.all_questions())
    return (
        f"저장 완료: {out_path}\n"
        f"시험: {data.exam_title}\n"
        f"과목: {data.subject}\n"
        f"문제 수: {q_count}개\n"
        f"섹션: {len(data.sections)}개"
    )


@mcp.tool()
def generate_exam(
    project_dir: str,
    questions_file: str = "bank.json",
    output_prefix: str = "",
) -> str:
    """저장된 문제 데이터와 스타일 스펙을 사용하여 HWPX 파일을 생성합니다.

    문제지(.hwpx)와 해설지(_해설.hwpx) 두 개의 파일이 생성됩니다.

    Args:
        project_dir: 프로젝트 폴더 경로
        questions_file: 문제 데이터 파일명 (questions/ 폴더 내, 기본: bank.json)
        output_prefix: 출력 파일명 접두사 (생략 시 exam_title 사용)

    Returns:
        생성된 파일 경로
    """
    proj = Path(project_dir)

    # 문제 데이터 로드
    q_path = proj / "questions" / questions_file
    if not q_path.exists():
        return f"오류: 문제 파일을 찾을 수 없습니다: {q_path}"

    with open(q_path, "r", encoding="utf-8") as f:
        data = ExamData.from_dict(json.load(f))

    if not data.all_questions():
        return "오류: 문제가 없습니다."

    # 스펙 로드
    spec = _load_spec(project_dir)

    # 생성
    gen = ExamHWPXGenerator(spec)
    out_dir = proj / "output"
    out_dir.mkdir(parents=True, exist_ok=True)

    prefix = output_prefix or (data.exam_title or "모의고사").replace(" ", "_")
    exam_path = str(out_dir / f"{prefix}.hwpx")
    answer_path = str(out_dir / f"{prefix}_해설.hwpx")

    try:
        gen.generate_exam(data, exam_path)
        gen.generate_answer_key(data, answer_path)

        exam_size = Path(exam_path).stat().st_size
        answer_size = Path(answer_path).stat().st_size

        return (
            f"생성 완료!\n\n"
            f"문제지: {exam_path} ({exam_size:,} bytes)\n"
            f"해설지: {answer_path} ({answer_size:,} bytes)\n\n"
            f"문제 수: {len(data.all_questions())}개\n"
            f"스펙: {_get_spec_path(project_dir)}"
        )

    except Exception as e:
        return f"오류: {type(e).__name__}: {e}"


# ---------------------------------------------------------------------------
# 진입점
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    mcp.run()
