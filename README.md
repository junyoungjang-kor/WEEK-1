# 📋 Hanwha BNCP 주간 보안 보고서 자동 생성기

Hanwha BNCP 이라크 현장 주간 보안 보고서(Weekly Security Report)를 자동으로 생성하는 Python GUI 프로그램입니다.

---

## 📌 프로젝트 개요

| 항목 | 내용 |
|------|------|
| **프로젝트** | Hanwha BNCP Iraq Site 주간 보안 보고서 자동화 |
| **버전** | v4.0 (개조식 번역 강화 + 전면 버그 수정) |
| **개발 언어** | Python 3.12 |
| **GUI** | tkinter / ttk |
| **AI 번역** | Anthropic Claude API (`claude-sonnet-4-20250514`) |

---

## ✅ 주요 기능

### 1. 한국어 → 영어 AI 자동 번역 (개조식)
- GUI에 **한국어로 입력**하면 Claude API가 **보안 보고서 개조식(bullet-point) 영어**로 자동 번역
- 일괄 번역(Batch Translation) 1회 호출로 API 비용 최소화
- 번역 실패 시 개별 항목 자동 재시도(fallback)
- 번역 예시:
  - `"L1005 ihsan ali 병원진료로 3일간 결근"` → `"L1005 Ihsan Ali - 3 days absent, medical treatment"`
  - `"3월 21일 psd 2개팀 드라이빙 어세스먼트 실시"` → `"21 Mar - 2 PSD teams conducted Driving Assessment"`

### 2. 반복 업무 자동 입력
- **매일**: `Daily check - PSD & Static Guard: NSTR` 자동 입력
- **매주 금요일**: Meal Request to Hanwha, Weekly Weapon & Ammunition Status 자동 추가
- **매주 화요일**: Vehicle Mileage, Weekly Security Report 제출 자동 추가

### 3. SSG 근무 교대 자동 반영
- 4일마다 교대하는 SSG(Static Security Guard) 근무조 자동 계산
- 기준일과 조 번호 입력 시 해당 주의 교대 일정 자동 반영
- 월이 걸치는 주간에도 정확한 날짜 비교 (full date comparison)

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
├── weekly_report_generator.py   # 메인 실행 파일 (v4.0)
├── requirements.txt             # Python 패키지 목록 (pip freeze)
├── weekly_report_config.json    # 설정 파일 (API 키, 차량 마일리지) ※ .gitignore 처리
├── .env                         # 환경 변수 파일 ※ .gitignore 처리
├── venv/                        # Python 가상환경 ※ .gitignore 처리
├── .gitignore                   # Git 업로드 제외 목록
└── README.md                    # 프로젝트 설명서 (현재 파일)
```

### 별도 경로 파일 (로컬 전용)
- **템플릿**: `Hanwha BNCP Weekly Report from 11 Mar 2025 to 17 Mar 2026(양식).docx`
- **출력 폴더**: `AI Output\Weekly Report\`

---

## 🚀 실행 방법

### 1. 가상환경 설정 (최초 1회)

#### 가상환경 생성
```bash
cd C:\Users\user\Documents\WEEK-1
python -m venv venv
```

#### 가상환경 활성화 (PowerShell)
```powershell
.\venv\Scripts\Activate.ps1
```

> ⚠️ 보안 오류 발생 시 먼저 실행:
> ```powershell
> Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
> ```

활성화 성공 시 터미널 앞에 `(venv)` 표시됨:
```
(venv) PS C:\Users\user\Documents\WEEK-1>
```

#### 가상환경 비활성화 (나가기)
```bash
deactivate
```

### 2. 필수 패키지 설치 (가상환경 활성화 상태에서)
```bash
pip install -r requirements.txt
```

### 3. 프로그램 실행
```bash
python weekly_report_generator.py
```

### 4. API 키 설정
1. 프로그램 상단의 **API Key** 입력란에 Anthropic API 키 입력
2. **"Save & Test"** 버튼 클릭하여 연결 확인
3. 키는 `weekly_report_config.json`에 저장되어 다음 실행 시 자동 로드

### 5. 보고서 작성
1. **기준 날짜** 선택 (자동으로 수~화 기간 계산)
2. 각 탭에서 **한국어로** 내용 입력
3. **"보고서 생성"** 버튼 클릭
4. AI가 한→영 개조식 번역 후 .docx 파일 자동 생성

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
- [x] Python 가상환경(venv) 구성 및 requirements.txt 생성
- [x] **v4.0 전면 리팩토링 - 개조식 번역 강화 + 버그 9건 수정**

---

## 🔄 버전 이력

### v4.0 (2026-03-18) - 개조식 번역 강화 + 전면 버그 수정

#### 핵심 변경: 번역 품질 개선
| 항목 | v3.0 (이전) | v4.0 (현재) |
|------|------------|------------|
| **번역 스타일** | 일반 영어 문장 | **개조식(bullet-point)** 보안 보고서 문체 |
| **번역 횟수** | 2회 중복 호출 (배치 + Generator 내부) | **1회만** (배치 번역 후 Generator는 번역 안 함) |
| **API 비용** | 불필요한 2배 소모 | **절반으로 감소** |
| **번역 실패 시** | 한국어 그대로 출력 | **개별 항목 자동 재시도(fallback)** |

#### 수정된 버그 9건
| # | 등급 | 버그 | 원인 | 수정 내용 |
|---|------|------|------|----------|
| 01 | 🔴 심각 | MONTH 연도 검색 실패 | 특정 연도만 탐색 | 2024~2030 전 연도 + preserve 속성 탐색 |
| 03 | 🟡 중간 | Vehicle Mileage tc 매칭 실패 | `<w:tc>` 속성 미대응 | `<w:tc[ >]` 패턴으로 속성 포함 대응 |
| 04 | 🔴 심각 | Issues 마커에서 크래시 | `index()` 실패 시 ValueError | `find()` + 안전 fallback |
| 06 | 🟡 중간 | 배치 번역 파싱 실패 | 멀티라인 응답 미처리 | 멀티라인 파싱 + 개별 fallback |
| 07 | 🟡 중간 | **이중 번역 (API 비용 낭비)** | Generator 내부에서 재번역 | Generator 내 번역 호출 전면 제거 |
| 08 | 🟡 중간 | Finance 마커에서 크래시 | 마커 텍스트 분리 시 ValueError | 안전 탐색 + fallback |
| 09 | 🟡 중간 | 월말 Shift 오매칭 | day만 비교 (월 무시) | full_date 비교로 변경 |
| 10 | 🟢 경미 | 미사용 번역 사전 50줄+ | v3.0에서 API 전환 후 잔존 | KO_EN_DICT, KO_EN_PHRASES 삭제 |
| 12 | 🟢 경미 | Header 데드코드 | 미사용 정규식 잔존 | 정리 삭제 |

#### 기타 개선
- GUI 레이블 한국어화 (사용자 편의)
- 번역 시스템 프롬프트를 별도 상수로 분리 (유지보수 용이)
- Client Feedback 파싱에서 번역 분리 (구조 파싱 → 번역 순서 명확화)

### v3.0 - Claude API 연동
- Claude API 기반 한→영 번역 기능 추가
- API 키 입력/저장/테스트 UI 추가
- 배치 번역으로 API 호출 최소화

### v2.0 - GUI 앱 구현
- tkinter 기반 GUI 입력 인터페이스
- 4개 탭 구조 (Basic/Training/Issues/Finance)

### v1.0 - 초기 프로토타입
- .docx 템플릿 XML 분석 및 자동 생성 엔진

---

## 🐛 해결된 주요 이슈 (전체)

| 이슈 | 원인 | 해결 방법 |
|------|------|----------|
| XML 파싱 오류 | 섹션 마커가 여러 `<w:t>` 태그에 분산 | 마커 텍스트 단축 + 인접 검증 방식 |
| PERIOD 날짜 미변경 | 날짜가 11개 `<w:t>` 요소에 분산 | 기존 run 삭제 후 단일 run으로 재작성 |
| Training 테이블 XML 깨짐 | 순방향 반복 시 오프셋 오류 | 역순(마지막→처음) 처리로 변경 |
| 한국어 번역 품질 불량 | 단순 사전 치환 방식의 한계 | Claude API 연동으로 전면 교체 |
| API 모델 404 오류 | 사용 불가 모델명 지정 | `claude-sonnet-4-20250514`로 변경 |
| API 인증 401 오류 | 잘못된 API 키 입력 | 키 재발급 + Save & Test 기능 |
| 이중 번역 (v4.0 수정) | Generator 내부 재번역 | 번역 1회 호출로 통합 |
| Issues/Finance 크래시 (v4.0 수정) | XML 마커 탐색 실패 | `find()` + safe fallback |
| 월말 Shift 오류 (v4.0 수정) | day만 비교 | full_date 비교 |

---

## ⚠️ 주의사항

- `weekly_report_config.json`에 **API 키**가 포함되어 있으므로 **절대 GitHub에 업로드하지 마세요** (`.gitignore` 처리 완료)
- `.env` 파일도 민감 정보 포함 가능 → `.gitignore` 처리 완료
- 보고서 생성 시 **같은 이름의 파일이 Word에서 열려있으면** PermissionError 발생 → Word 파일 닫고 재생성
- 템플릿 파일 경로가 로컬 OneDrive 경로로 설정되어 있으므로, 다른 PC에서 사용 시 `TEMPLATE_PATH` 수정 필요

---

## 📝 라이선스

본 프로젝트는 Hanwha BNCP 내부 업무용으로 제작되었습니다.
