# 📋 Hanwha BNCP 주간 보안 보고서 자동 생성기

Hanwha BNCP 이라크 현장 주간 보안 보고서(Weekly Security Report)를 자동으로 생성하는 Python GUI 프로그램입니다.

---

## 📌 프로젝트 개요

| 항목 | 내용 |
|------|------|
| **프로젝트** | Hanwha BNCP Iraq Site 주간 보안 보고서 자동화 |
| **버전** | v3.0 (Claude AI API 연동) |
| **개발 언어** | Python 3.12 |
| **GUI** | tkinter / ttk |
| **AI 번역** | Anthropic Claude API (claude-sonnet-4-20250514) |

---

## ✅ 주요 기능

### 1. 한국어 → 영어 AI 자동 번역
- GUI에 **한국어로 입력**하면 Claude API가 **보안 보고서 스타일의 전문 영어**로 자동 번역
- 일괄 번역(Batch Translation)으로 API 호출 최소화

### 2. 반복 업무 자동 입력
- **매일**: `Daily check - PSD & Static Guard: NSTR` 자동 입력
- **매주 금요일**: Meal Request to Hanwha, Weekly Weapon & Ammunition Status 자동 추가
- **매주 화요일**: Vehicle Mileage, Weekly Security Report 제출 자동 추가

### 3. SSG 근무 교대 자동 반영
- 4일마다 교대하는 SSG(Static Security Guard) 근무조 자동 계산
- 기준일과 조 번호 입력 시 해당 주의 교대 일정 자동 반영

### 4. 재무(Finance) 자동 서식
- 금액 입력 시 자동으로 `IQD` 통화 표시 및 천 단위 콤마 포맷팅

### 5. PSD 팀 기본값
- PSD 팀은 기본적으로 `H01`, `H02`로 설정

### 6. Client Feedback 자동 구조화
- 자유 텍스트 입력 시 **Issue / Summary / Actions** 형태로 자동 분류

### 7. 보고서 .docx 생성
- 원본 템플릿의 서식을 **정확히 유지**하면서 내용만 교체
- XML 직접 조작 방식으로 서식 손실 없음

---

## 📂 파일 구조

```
WEEK-1/
├── weekly_report_generator.py   # 메인 실행 파일 (v3.0)
├── weekly_report_config.json    # 설정 파일 (API 키, 차량 마일리지) ※ .gitignore 처리
├── .gitignore                   # Git 업로드 제외 목록
└── README.md                    # 프로젝트 설명서 (현재 파일)
```

### 별도 경로 파일 (로컬 전용)
- **템플릿**: `Hanwha BNCP Weekly Report from 11 Mar 2025 to 17 Mar 2026(양식).docx`
- **출력 폴더**: `AI Output\Weekly Report\`

---

## 🚀 실행 방법

### 1. 필수 패키지 설치
```bash
pip install anthropic
```

### 2. 프로그램 실행
```bash
python weekly_report_generator.py
```

### 3. API 키 설정
1. 프로그램 상단의 **API Key** 입력란에 Anthropic API 키 입력
2. **"Save & Test"** 버튼 클릭하여 연결 확인
3. 키는 `weekly_report_config.json`에 저장되어 다음 실행 시 자동 로드

### 4. 보고서 작성
1. **기준 날짜** 선택 (자동으로 수~화 기간 계산)
2. 각 탭에서 **한국어로** 내용 입력
3. **"보고서 생성"** 버튼 클릭
4. AI가 한→영 번역 후 .docx 파일 자동 생성

---

## 🔧 개발 진행 사항

### 완료된 작업
- [x] .docx 템플릿 XML 분석 및 구조 파악
- [x] GUI 레이아웃 설계 (tkinter Notebook 탭 구조)
- [x] Weekly Summary 섹션 자동 생성
- [x] Training 테이블 자동 채우기
- [x] Issues 섹션 (5.1~5.8) 자동 생성
- [x] Vehicle Mileage 테이블 자동 채우기
- [x] Finance 섹션 IQD 자동 서식
- [x] Client Feedback 자동 구조화
- [x] PERIOD / MONTH 헤더 날짜 자동 업데이트
- [x] SSG 근무교대 4일 주기 자동 계산
- [x] 반복 업무 자동 입력 (금요일/화요일 제출 항목)
- [x] Claude API 연동 한→영 번역 (v3.0)
- [x] API 키 저장/로드 기능
- [x] GitHub 업로드

### 해결된 주요 이슈
| 이슈 | 원인 | 해결 방법 |
|------|------|----------|
| XML 파싱 오류 | 섹션 마커가 여러 `<w:t>` 태그에 분산 | 마커 텍스트 단축 + 인접 검증 방식 |
| PERIOD 날짜 미변경 | 날짜가 11개 `<w:t>` 요소에 분산 | 기존 run 삭제 후 단일 run으로 재작성 |
| Training 테이블 XML 깨짐 | 순방향 반복 시 오프셋 오류 | 역순(마지막→처음) 처리로 변경 |
| 한국어 번역 품질 불량 | 단순 사전 치환 방식의 한계 | Claude API 연동으로 전면 교체 |
| API 모델 404 오류 | 사용 불가 모델명 지정 | `claude-sonnet-4-20250514`로 변경 |

---

## ⚠️ 주의사항

- `weekly_report_config.json`에 **API 키**가 포함되어 있으므로 **절대 GitHub에 업로드하지 마세요** (`.gitignore` 처리 완료)
- 보고서 생성 시 **같은 이름의 파일이 Word에서 열려있으면** PermissionError 발생 → Word 파일 닫고 재생성
- 템플릿 파일 경로가 로컬 OneDrive 경로로 설정되어 있으므로, 다른 PC에서 사용 시 `TEMPLATE_PATH` 수정 필요

---

## 📝 라이선스

본 프로젝트는 Hanwha BNCP 내부 업무용으로 제작되었습니다.
