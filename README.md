# 제주항공 중정비 All-in-One (Heavy Maintenance All-in-One)

제주항공(Jeju Air) 중정비(Heavy Maintenance / C-Check) 업무를 위한 **통합 엔지니어링 웹 포털**입니다.

---

## 🌐 서비스 접속 주소
👉 **https://justiniankim.github.io/heavy-maintenance-allinone/**

*(별도의 서버 구축이나 프로그램 설치 없이 브라우저에서 바로 사용 가능하며, 모든 데이터는 브라우저 메모리 내부에서만 안전하게 처리됩니다.)*

---

## 🛠 주요 모듈 구성

### 1. 📊 Work Package 분석기 (Work Package Analyzer)
- **일괄 배치 분석(Batch Processing)**: 여러 대의 Work Package 엑셀 파일을 한 번에 드래그 앤 드롭 분석
- **최종 리비전 자동 선별**: 동일 기번에 복수 리비전 업로드 시 최고 리비전 자동 채택
- **가중치(Manhour Factors) 자동 산출**:
  - 파일명 기반 C-Check 단계 분석 (4C 이상 시 +30 M/H)
  - 제주항공 45대 기단 기령 내장 DB 탑재 (15년 이상 시 +30 M/H 자동 판정)
  - Wet Area Access (53-380-00-01) 자동 감지 (+300 M/H, 최소 12일 TAT 보정)
  - Wing Task (13개 카드 목록, 0.7~8.5 M/H 보완치 +40 M/H)
  - 보정 항목: 990 계열(+10 M/H), Stabilizer Trim Actuator(+50 M/H), NGS Heat Exchanger Cleaning(+10 M/H)
- **결과 엑셀(XLSX) 일괄 내보내기** 및 실시간 터미널 콘솔 로그 제공

### 2. 📋 중정비 업무 체크리스트 (Heavy Maintenance Checklist)
- C-Check 단계별 필수 점검 항목 및 일정/마일스톤 관리

### 3. 🛋️ 객실 오더 관리 (Cabin Order)
- 객실 결함, 시트/갤리/화장실 점검 및 오더 정리 지원

---

## 💻 로컬 실행 방법
`index.html` 파일을 더블 클릭하여 크롬/엣지 브라우저에서 바로 실행할 수 있습니다.
