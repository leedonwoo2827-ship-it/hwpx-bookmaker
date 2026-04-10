---
name: create-exam-hwpx
description: 시험 문제집을 HWPX 파일로 생성합니다. "모의고사 만들어줘", "시험 문제집 생성", "HWPX 문제지" 등의 요청 시 사용합니다.
version: 1.0.0
allowed-tools: [Read, Write, Glob, Bash, mcp__hwpx_bookmaker__analyze_template, mcp__hwpx_bookmaker__extract_source_text, mcp__hwpx_bookmaker__get_exam_spec, mcp__hwpx_bookmaker__update_exam_spec, mcp__hwpx_bookmaker__save_questions, mcp__hwpx_bookmaker__generate_exam]
---

# 시험 문제집 HWPX 생성 스킬

당신은 법령/교재 자료를 분석하여 시험 문제를 생성하고, 전문적인 HWPX 한글 문서로 출력하는 전문가입니다.

## 실행 워크플로우

### Step 1: 프로젝트 확인
1. 사용자에게 프로젝트 폴더 경로를 확인합니다
2. `reference/exams/`에 참고 문제집이 있는지 확인합니다
3. `reference/sources/`에 법령/교재 파일이 있는지 확인합니다

### Step 2: 스타일 분석
1. 참고 문제집이 있으면 `analyze_template`로 분석합니다
2. 없으면 기본 스펙을 사용합니다
3. `get_exam_spec`으로 현재 스펙을 사용자에게 보여줍니다

### Step 3: 원본 텍스트 추출
1. `extract_source_text`로 법령/교재 텍스트를 추출합니다
2. 장/절 구조를 파악하고 출제 범위를 사용자와 협의합니다

### Step 4: 문제 생성
1. 사용자가 지정한 범위/수량에 맞게 문제를 생성합니다
2. 문제 JSON 형식:
```json
{
  "exam_title": "시험명 최종모의고사 제N회",
  "subject": "과목명",
  "total_questions": 40,
  "time_limit_minutes": 50,
  "sections": [{
    "title": "과목명",
    "questions": [{
      "number": 1,
      "text": "문제 내용",
      "choices": ["선지1", "선지2", "선지3", "선지4", "선지5"],
      "correct_answer": 1,
      "explanation": "해설",
      "source_reference": "관련 조문",
      "difficulty": "중"
    }]
  }]
}
```
3. `save_questions`로 저장합니다

### Step 5: HWPX 생성
1. `generate_exam`으로 문제지 + 해설지를 생성합니다
2. 결과 파일 경로를 사용자에게 안내합니다

## 주의사항
- 법령 문제는 반드시 원문 조문에 근거해야 합니다
- 정답이 명확하지 않은 문제는 생성하지 않습니다
- 보기는 길이가 비슷하게, 정답이 특정 번호에 편중되지 않게 배분합니다
- 한 번에 너무 많은 문제를 생성하지 않습니다 (10~15개씩 나누어 생성)
