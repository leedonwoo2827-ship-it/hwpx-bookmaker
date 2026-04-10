# hwpx-bookmaker

Claude Desktop에서 시험 문제집을 작성하고 HWPX 한글 문서로 생성하는 MCP 서버입니다.
[hwpx_writer](https://github.com/leedonwoo2827-ship-it/hwpx_writer)의 HWPX 생성 기술을 기반으로, 시험 문제집에 특화된 3단계 파이프라인을 구현합니다.

## 파이프라인

```
1단계  analyze_template               참고 문제집 HWPX 분석 → 스타일 스펙 추출
       (MCP 도구)                      → exam-spec.json
                  │
                  ▼  폰트, 크기, 간격, 보기 형식

2단계  Claude Desktop 대화             AI가 법령/교재 기반으로 문제 생성
       (Human-in-the-Loop)             → questions/bank.json
                  │
                  ▼  구조화된 문제 JSON

3단계  generate_exam                   문제 JSON → HWPX 문서 빌드
       (MCP 도구)                      → 문제지.hwpx + 해설지.hwpx
```

| 단계 | 처리 | 도구 | 방식 |
|------|------|------|------|
| **1단계** 템플릿 분석 | 참고 HWPX의 XML 파싱 → 스타일 스펙 | MCP (자동) | 배치 |
| **2단계** 문제 생성 | 법령/교재 기반 문제 작성 | Claude Desktop | 대화형 (HITL) |
| **3단계** 문서 빌드 | JSON → HWPX 패키징 | MCP (자동) | 배치 |

## 왜 2단계는 대화형(Interactive)인가

시험 문제 생성은 **비결정적(non-deterministic) 출력**입니다. 같은 법령 텍스트를 줘도 매번 다른 문제가 나옵니다.

```
법령 추출 → 범위 협의 → 문제 초안 → 사용자 검토 → 수정/보강 → 저장
     ↑                                                         ↓
     └──────── 추가 범위 / 난이도 조정 / 오류 수정 ←───────────┘
```

| 상황 | 배치 (one-shot) | Claude Desktop 대화형 |
|------|-----------------|----------------------|
| 특정 조문에서 더 출제 | 전체 재생성 | "제4조에서 3문제 더 만들어줘" |
| 난이도 편향 | 파라미터 수정 후 재실행 | "좀 더 어렵게 바꿔줘" |
| 정답 오류 발견 | 수동 JSON 편집 | "3번 정답이 ②가 아니라 ④야" |
| 선지 품질 개선 | 불가 | "5번 보기가 너무 뻔해, 매력적 오답으로 바꿔" |

2단계를 CLI로 만들면 "한 번에 완벽한 프롬프트"를 짜야 하는데, 시험 문제 수준의 정확성이 필요한 작업에서는 현실적으로 불가능합니다.

## 설치

### 마켓플레이스에서 설치 (권장)

Claude Desktop에서 바로 설치할 수 있습니다.

1. Claude Desktop 좌측 하단 **사용자 지정** 클릭
2. 개인 플러그인 옆 **+** 버튼 → **마켓플레이스 추가** 선택
3. URL 입력란에 아래 주소를 붙여넣고 **동기화** 클릭:
   ```
   https://github.com/leedonwoo2827-ship-it/hwpx-bookmaker
   ```
4. 개인 플러그인 목록에 **hwpx-bookmaker**가 나타나면 설치 완료

### 수동 설치

마켓플레이스를 사용하지 않는 경우, 직접 다운로드하여 연결할 수 있습니다.

```bash
git clone https://github.com/leedonwoo2827-ship-it/hwpx-bookmaker.git
cd hwpx-bookmaker
pip install -r requirements.txt
```

`claude_desktop_config.json`에 아래를 추가합니다.

```json
{
  "mcpServers": {
    "hwpx-bookmaker": {
      "command": "python",
      "args": ["-X", "utf8", "server.py"],
      "cwd": "설치경로/hwpx-bookmaker"
    }
  }
}
```

## 사용 방법

### 사전 준비

프로젝트 폴더를 만들고 자료를 넣습니다:

```
C:\Users\ubion\Documents\testmaker\260410-1\
├── reference/
│   ├── exams/       ← 참고 시험 문제집 HWP/HWPX (2~5개)
│   └── sources/     ← 법령/교재 HWPX (법제처 다운로드)
```

**법제처 다운로드 방법:**
1. [법제처](https://law.go.kr) → 해당 법령 검색 → 본문 탭
2. 저장 버튼 → `HWPX파일` 라디오 버튼 선택 → 저장

**참고 문제집 (최소 2개):**
| 우선순위 | 파일 종류 | 용도 |
|---------|----------|------|
| 1 | 기출문제 최근 1~2회분 | 출제 패턴, 난이도 파악 |
| 2 | 모의고사/교재 샘플 | 레이아웃(폰트, 간격) 참고 |
| 3 | 추가 기출/모의 | 다양한 패턴 학습 (선택) |

### 워크플로우 (Claude Desktop 대화)

```
사용자: 프로젝트 폴더는 C:\Users\ubion\Documents\testmaker\260410-1 이야.
       reference/exams/ 에 있는 문제집을 분석해줘.
→ Claude: analyze_template 호출 → exam-spec.json 생성

사용자: reference/sources/ 에 있는 경비업법 텍스트를 추출해줘.
→ Claude: extract_source_text 호출 → 법령 텍스트 반환

사용자: 제1장 총칙 범위에서 문제 10개 만들어줘.
→ Claude: 문제 JSON 작성 → save_questions 호출

사용자: 40문제로 제1회 모의고사 만들어줘.
→ Claude: generate_exam 호출 → 문제지.hwpx + 해설지.hwpx
```

## MCP 도구

| 도구 | 설명 |
|------|------|
| `analyze_template` | 참고 HWPX → 스타일 스펙(exam-spec.json) 추출 |
| `extract_source_text` | 법령/교재 HWPX → 구조화된 텍스트 추출 |
| `get_exam_spec` | 현재 스타일 스펙 조회 |
| `update_exam_spec` | 스타일 스펙 부분 수정 |
| `save_questions` | 문제 데이터 JSON 저장 |
| `generate_exam` | 문제지 + 해설지 HWPX 생성 |

**공통 파라미터:**
- `project_dir`: 프로젝트 폴더 경로 (모든 도구에서 사용)

## 문제 JSON 형식

Claude가 생성하고 `save_questions`로 저장하는 문제 데이터입니다.

```json
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
          "choices": [
            "경비업의 육성 및 발전과 그 체계적 관리",
            "경비원의 근로조건 개선과 복지 향상",
            "국가중요시설의 경비 강화",
            "민간경비산업의 국제 경쟁력 강화",
            "범죄의 예방과 수사의 효율성 증진"
          ],
          "correct_answer": 1,
          "explanation": "경비업법 제1조에 따르면...",
          "source_reference": "경비업법 제1조",
          "difficulty": "하"
        }
      ]
    }
  ]
}
```

## 출력 결과

`generate_exam` 실행 시 두 개의 HWPX 파일이 생성됩니다:

| 파일 | 내용 |
|------|------|
| `모의고사_제1회.hwpx` | 문제 + 보기만 (해설 없음) — 시험용 |
| `모의고사_제1회_해설.hwpx` | 정답표 + 문제별 해설 — 학습용 |

한글(Hancom Office)에서 열어 인쇄하거나 PDF로 변환할 수 있습니다.

## 프로젝트 폴더 구조

```
프로젝트폴더/
├── reference/
│   ├── exams/          ← 참고 시험 문제집 HWP/HWPX
│   └── sources/        ← 법령/교재 HWPX (법제처 다운로드)
├── exam-spec.json      ← analyze_template로 생성된 스타일 스펙
├── questions/
│   └── bank.json       ← 문제 은행
└── output/
    ├── 모의고사_제1회.hwpx
    └── 모의고사_제1회_해설.hwpx
```

> 다른 시험 준비 시 새 프로젝트 폴더를 만들어 동일 구조로 사용합니다.
> `project_dir` 파라미터로 경로를 지정하면 됩니다.

## 기술 스택

- **HWPX 생성**: [hwpx_writer](https://github.com/leedonwoo2827-ship-it/hwpx_writer) 기반 문자열 XML 생성
- **스타일 추출**: HWPX ZIP → XML 파싱 (Vision API 불필요)
- **MCP 프레임워크**: FastMCP (Python)
- **번들 폰트**: Noto Sans KR, Noto Serif KR (OFL 라이선스)
