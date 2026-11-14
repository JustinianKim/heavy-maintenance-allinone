# Work Package (WP) 분석기 - Web Edition

제주항공(Jeju Air) 중정비(Heavy Maintenance) Work Package의 Man-Hour(공수) 및 가중치(Manhour Factors)를 기반으로 **Turn Around Time (TAT)**을 정밀 예측하는 웹 분석기입니다.

---

## 🌟 주요 특징

1. **설치 없는 웹 실행 (GitHub Pages 지원)**:
   - 별도 프로그램(`.exe`) 설치 없이 브라우저에서 바로 접속하여 사용 가능합니다.
   - 외부 서버로 엑셀 데이터가 전송되지 않고 **브라우저 로컬 메모리(Client-Side)에서만 처리**되므로 보안상 안전합니다.
2. **동일한 핵심 분석 로직**:
   - 파일명 기반 C-Check 단계 분석 (4C 이상 시 +30 M/H)
   - 기령 분석 (제작일 기준 15년 이상 시 +30 M/H)
   - Wet Area Access (53-380-00-01) 자동 감지 (+300 M/H, 최소 12일 보정)
   - Wing Task (지정 13개 카드 목록) 및 990 계열, Stabilizer Trim Actuator, NGS Heat Exchanger Cleaning 보정치 계산
   - 실시간 터미널 콘솔 로그 및 상세 산출 리포트 제공

---

## 🚀 GitHub Pages 배포 방법

1. 본 레포지토리의 **[Settings]** 탭으로 이동합니다.
2. 좌측 메뉴에서 **[Pages]**를 클릭합니다.
3. **Build and deployment > Branch** 항목에서:
   - Branch: `main` 선택
   - Folder: `/ (root)` 선택
4. **[Save]** 버튼을 누르면 1~2분 뒤 나만의 웹 주소가 생성됩니다:
   - `https://<깃허브_아이디>.github.io/<저장소_이름>/`

---

## 💻 로컬 실행 방법

- `index.html` 파일을 더블 클릭하여 크롬/엣지 등 브라우저에서 바로 열어서 사용할 수도 있습니다.
