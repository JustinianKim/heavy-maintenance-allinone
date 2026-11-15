# 제주항공 중정비 All in One Portal — 프로젝트 인수인계 및 작업 상태 가이드

> **최종 갱신일**: 2026-11-14  
> **저장소**: [JustinianKim/heavy-maintenance-allinone](https://github.com/JustinianKim/heavy-maintenance-allinone.git) (`main` 브랜치)  
> **라이브 웹사이트**: [https://justiniankim.github.io/heavy-maintenance-allinone/](https://justiniankim.github.io/heavy-maintenance-allinone/)  
> **작업 디렉터리**: `c:\Users\user\Documents\Jeju_Air\Programming\WP_분석기`

---

## 1. 프로젝트 개요 및 아키텍처

- **목적**: Python 데스크톱 앱에서 단일 웹 포털로 전환된 제주항공 중정비(Heavy Maintenance) 종합 관리 시스템.
- **구동 환경**: 
  - 100% 클라이언트 사이드 (순수 HTML/JS/CSS).
  - 외부 CDN 및 Google Fonts 의존성 완전 제거 (중국 출장지 및 인터넷 미연결 폐쇄망에서 `실행하기.bat` 더블클릭으로 즉시 구동).
  - 로컬 벤더 스크립트: `vendor/xlsx.full.min.js` (SheetJS), `vendor/tailwind.min.js` (Tailwind CSS v3).
  - 오프라인 배포 파일: `heavy_maintenance_allinone_offline.zip` (코드 변경 시 자동 재압축 및 Git 커밋/푸시).

---

## 2. 모듈 구성 및 기능 현황

| 모듈 | 화면명 | 상태 | 주요 기능 |
| :--- | :--- | :---: | :--- |
| **모듈 1** | **Work Package 분석기** | 완료 | • 복수 WP 엑셀 드래그&드롭 일괄 분석<br>• MPD/가중치 기반 공수(M/H) 및 예상 TAT 자동 산출<br>• 실시간 분석 터미널 로그 및 분석 결과 엑셀 저장<br>• 내부 `산출 기준` 가이드 팝업 모달 탑재 |
| **모듈 2** | **중정비 업무 체크리스트** | 완료 | • 162개 전 점검 항목 수록 (입고전 36건, 입고중 115건, 완료후 11건)<br>• 기번별 독립 저장소(`localStorage`), 항목별 완료 체크 및 일시 자동 기록<br>• 인쇄/PDF 최적화 A4 양식 및 엑셀 내보내기<br>• 부서별 필터, 진행상태 필터, 실시간 키워드 검색<br>• 상단 4개 진척도 카드 클릭 시 해당 단계 즉시 필터링 |
| **신규 기능** | **공수비행 현황 (Ferry Flight)** | 완료 | • 상단 메뉴와 체크리스트 사이 슬림형 접이식(아코디언) 카드<br>• `7C*M` 편명 검색 및 Flightradar24 / ADS-B Globe 실시간 레이더 화면 내장<br>• 최근 공수비행 이착륙 기록 관리, 신규 등록 및 엑셀 다운로드<br>• 공수편 클릭 시 해당 항공기 체크리스트 즉시 자동 연동 |
| **모듈 3** | **객실 오더** | 대기 | 추후 객실 정비 오더 및 결함(Defect) 양식 구축 예정 |

---

## 3. 사용자 전용 필수 개발 규칙 (Strict Guidelines)

새로운 대화 또는 기능 추가 시 아래 원칙을 반드시 100% 준수해야 합니다.

### ① 불필요한 잡다한 요소 및 이모티콘 일절 배제 (미니멀/담백한 스타일)
- 버튼, 타이틀, 상태 배지, 푸터 등에 불필요한 이모지(✈, 📦, 📊, 🔍, ⏳, ✅ 등)를 절대 넣지 않습니다.
- 중복되는 가이드 텍스트(`전체 보기 →`, `이 단계만 필터링 →` 등)나 잡다한 부연 설명을 넣지 않고 간결한 텍스트로 처리합니다.
- 불필요한 영문 표기(`(Pending)`, `(Done)`, `Follow-up` 등)는 담백한 한글로 표기합니다.

### ② 부서 명칭 표준
- **생산계획팀** ➔ **`생산계획파트`**
- **구매팀** ➔ **`항공기재팀(구매팀)`**
- **정비지원팀** ➔ **`정비기획팀`**
- **정비팀 출장자** ➔ **`정비팀`**
- **유관 MRO 제거**: C-Check는 MRO에서 전담 수행하므로 유관부서 목록에서 `MRO` 문구는 모두 배제합니다.
- **부서 배지 수직 정렬**: 주관/유관 부서가 여러 개일 경우, `flex-col`로 줄바꿈하여 왼쪽 시작선을 일치시킵니다.

### ③ 체크리스트 하위 업무 불릿 표기
- 하위 세부 점검 항목은 **단일 불릿(`• `)** 하나로만 표기합니다.
- 엑셀 원문에 있던 `-`와 겹쳐서 `• -` 형태로 중복 표기되지 않도록 `cleanText = s.replace(/^[-•·*]\s*/, '').trim()` 처리합니다.

### ④ 대상 항공기 선택기 (Combobox) 동작 규칙
- **45대 실제 운항 기번**: `HL7213` ~ `HL8766` (공통/표준 체크리스트 가상 옵션 없음).
- **표시 포맷**:
  - 일반 리스기: `HL7213 (기령 17.2년, 반납까지 2.8년)` — 현재 날짜(`new Date()`) 기준 실시간 계산.
  - 구매기(HL8331 등) 또는 반납일 N/A 기번: `HL8331 (기령 15.1년)` — **구매기/자가보유 등 추가 꼬리표 일절 표시 안 함**.
- **선택 해제(토글 및 백스페이스) 규칙**:
  - 기번 입력창에서 백스페이스/Delete로 글자를 모두 지우면 **기존 기번으로 강제 복귀되지 않고 빈칸 유지**.
  - 드롭다운 목록에서 **이미 선택된 기번을 한 번 더 누르면 선택 해제(토글 오프)**.
  - 기번 미선택 상태에서는 체크박스 클릭이나 엑셀 저장 시 기번을 먼저 선택하도록 안전 안내 팝업 출력.

### ⑤ 공수비행 (Ferry Flight, 7C*M) 모니터링 규칙
- **편명 표준**: `7C*M` (예: `7C8032M`, `7C8061M`, `7C7213M`).
- **레이더 연동**: 외부 도메인 iframe 차단이 없는 ADS-B Exchange Globe 실시간 임베드 기본 적용 + Flightradar24 원클릭 바로가기 (`Flightradar24 ↗`, `/data/flights/`, `/data/aircraft/`).
- **체크리스트 상호 연동**: 공수비행 목록에서 편명을 클릭하거나 조회하면 대상 항공기(`HL****`) 체크리스트가 즉시 자동 선택되어 업무 연속성 보장.
- **아코디언 UI**: 기본적으로 펼쳐져 있으며, 필요 시 언제든 한 줄 요약 바로 접기(`▲/▼`) 가능.

---

## 4. GitHub 저장소 동기화 및 배포 규칙 (GitHub Sync & Deployment)

### ① GitHub에 업로드(Tracked)되는 핵심 파일 목록
- **`index.html`**: 모든 기능(WP 분석기, 체크리스트, 객실오더)이 집약된 단일 메인 웹 애플리케이션
- **`heavy_maintenance_allinone_offline.zip`**: 최신 `index.html`, `실행하기.bat`, `app_icon.png`, `vendor/` 자산이 포함된 오프라인 패키지
- **`PROJECT_STATE.md`**: 본 인수인계 및 개발 스타일 가이드 문서
- **`README.md`**: 저장소 안내 문서
- **`실행하기.bat`**: Windows 로컬 브라우저 자동 실행 배치 파일
- **`app_icon.png`**: 앱 아이콘 및 파비콘
- **`vendor/`**: 자체 호스팅된 SheetJS (`xlsx.full.min.js`) 및 Tailwind CSS (`tailwind.min.js`)

*(※ 사내 기밀 및 원본 참조 파일인 `자료/` 폴더와 `*.pptx` 파일은 `.gitignore`에 의해 GitHub에 올라가지 않고 로컬에만 안전하게 보관됩니다.)*

### ② 코드 수정 시 필수 배포 절차
코드를 수정할 때마다 AI는 항상 아래 3단계를 자동으로 일괄 수행해야 합니다:
1. `index.html` 기능 수정 및 검증
2. 최신 `index.html`을 포함하여 `heavy_maintenance_allinone_offline.zip` 압축 파일 재생성
3. `git add`, `git commit`, `git push origin main` 실행 ➔ **GitHub Pages 자동 배포 완료** (약 30초 내 라이브 반영)

---

## 5. 파일 구성

```text
WP_분석기/
├── index.html                           # 포털 단일 메인 웹 애플리케이션
├── heavy_maintenance_allinone_offline.zip # 오프라인 배포용 통합 패키지 (Git 추적)
├── 실행하기.bat                          # 오프라인 실행용 배치 파일
├── app_icon.png                         # 파비콘 및 앱 아이콘
├── PROJECT_STATE.md                     # [본 파일] 프로젝트 인수인계 및 가이드
├── README.md                            # GitHub 저장소 안내 문서
└── vendor/                              # 완전 오프라인 구동용 로컬 라이브러리
    ├── tailwind.min.js                  # Tailwind CSS
    ├── xlsx.full.min.js                 # SheetJS 엑셀 파서/빌더
    └── README_오프라인_사용방법.txt
```

---

## 6. 다음 작업 시 빠른 시작 팁

새로운 채팅창을 시작하실 때 아래 한 줄만 입력하시면 즉시 이어서 작업이 가능합니다:

```text
이 프로젝트는 제주항공 중정비 All in One 웹 포털이야. PROJECT_STATE.md 확인하고 이어서 진행하자.
```
