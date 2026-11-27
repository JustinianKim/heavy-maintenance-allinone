# 제주항공 중정비 All in One 포털 (Heavy Maintenance All-in-One Portal)

제주항공(Jeju Air) 안전정비실 정비운영팀 생산계획파트를 위한 **중정비(Heavy Maintenance / C-Check) 종합 엔지니어링 웹 포털**입니다.

---

## 1. 서비스 접속 및 실행

- **라이브 웹사이트**: [https://justiniankim.github.io/heavy-maintenance-allinone/](https://justiniankim.github.io/heavy-maintenance-allinone/)
- **로컬 및 폐쇄망(오프라인) 실행**:
  - `heavy_maintenance_allinone_offline.zip` 압축 해제 후 `실행하기.bat` 더블클릭
  - 100% 클라이언트 사이드 구동으로 외부 인터넷 연결, 별도 서버 구축, 추가 프로그램 설치 없이 즉시 실행 가능합니다.
  - 모든 작업 데이터는 사용자 PC의 브라우저 메모리 및 로컬 저장소(`localStorage`)에만 안전하게 보관됩니다.

---

## 2. 주요 모듈 구성

### 1) GMP/ICN 야간 Ground 조회
- **운항스케줄 및 GRD 현황 자동 분석**: 운항스케줄(`flightSkd`) 및 일일 GRD 현황(`LAYOVER`) 엑셀 파일을 업로드하여 야간 주기 항공기 자동 추출
- **공항별 야간 주기 기번 판별**:
  - 김포공항 (`GMP-ALM`) 야간 주기 기번 및 대수 산출
  - 인천공항 (`ICN-ALM`) 야간 주기 기번 및 대수 산출
  - 심야 및 익일 새벽 도착편(Red-eye) 자동 분리 표기
- **워드 보고서(.docx) 생성**: 분석된 야간 주기 현황을 실무 표준 양식의 Word 문서로 원클릭 다운로드

### 2) 각종 오더 (SAP 연동)
- **객실 소독/청소 오더 자동 생성**:
  - 소독 오더(Disinfection): 기종별 Task List(`10630` NG / `13870` MAX) 자동 매핑 및 릴리즈/저장
  - 청소 오더(CABIN CLEANING CHK DUE): 작업장(`ARBPL`) 자동 할당 및 릴리즈/저장
  - ZMENU, ZINV 등 SAP이 어떤 화면에 위치해 있어도 안전하게 초기화 진입하는 `session.startTransaction "IW21"` 표준 스크립트(UTF-16LE with BOM) 출력
  - 모달 팝업(`wnd[1]`) 안전 자동 통과 및 완료 시 ZIW28 자동 조회
- **기타 오더 확장 구성**:
  - ROLL SHADE
  - CARTRIDGE VALVE
  - LAP BELT
  - DOOR LATCH
  - OXYGEN BOTTLE
  - ALL CARPET REPLACEMENT

### 3) Work Package 분석기
- **일괄 배치 분석(Batch Processing)**: 복수의 Work Package 엑셀 파일을 드래그 앤 드롭으로 일괄 분석
- **최종 리비전 자동 선별**: 동일 기번에 복수 리비전 업로드 시 최고 리비전 자동 채택
- **실무 10대 가중치(주1~주10) 및 예상 TAT 자동 산출**:
  - 주1) 제작 후 14년 이상 노후기 (+50 M/H)
  - 주2) 4C 이상 Check 또는 도입 후 7년 경과 (+50 M/H)
  - 주3) GLY & LAV 장탈 점검 WET AREA (+200 M/H, 최소 TAT 12일 보정)
  - 주4) Bear Strap & Frame NDI (+40 M/H, 주3 미중복 시)
  - 주5) Wing 전체 Fuel Tank Access Panel Remove (+40 M/H)
  - 주6) 행가 외부 작업 (+50 M/H)
  - 주7) STAB' Trim ACT' 교환 (+50 M/H)
  - 주8) NGS HEAT EXCHANGER 장탈 (+10 M/H)
  - 주9) LIVERY PAINT 후 10년 경과 (+50 M/H)
  - 주10) Landing Gear 교환 (ALL L/G +400 M/H, ONLY NLG +200 M/H)
- **제주항공 45대 기단 플릿 DB 내장**: 기번별 기령(`mfg`) 및 도입경과(`deliv`) 실시간 자동 판별
- **분석 결과 엑셀(XLSX) 저장** 및 동적 테이블 정렬, 실시간 터미널 로그 지원

### 4) 중정비 업무 체크리스트
- **162개 전 점검 항목 수록**: 입고 전(36건), 입고 중(115건), 완료 후(11건) 전 과정 관리
- **기번별 독립 저장소**: 45대 기번별 개별 점검 진척도 자동 저장(`localStorage`) 및 복원
- **실시간 공수비행(Ferry Flight, 7C*M) 모니터링**: ADS-B 기반 제주항공 중정비 공수편 실시간 추적 연동
- **단계별/부서별 필터링**: 주관부서(`생산계획파트`), 유관부서(`항공기재팀(구매팀)`, `정비기획팀`, `정비팀` 등) 필터 및 키워드 검색
- **공식 인쇄/PDF 최적화 A4 양식** 및 체크리스트 엑셀 내보내기 지원

---

## 3. 기술 스택 및 아키텍처

- **Frontend**: HTML5, Tailwind CSS v3 (로컬 벤더 내장)
- **Data Engine**: SheetJS (`xlsx.full.min.js`), Docx.js
- **Automation**: SAP GUI Scripting (VBScript UTF-16LE with BOM)
- **Zero External Dependencies**: Google Fonts, 외부 CDN 완전 제거로 인터넷 미연결 폐쇄망 100% 호환
