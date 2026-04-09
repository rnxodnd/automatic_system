
# 🛍 Automatic Detail Page System

엑셀 + 상품 이미지를 기반으로  
온라인 쇼핑몰 상세페이지를 자동 생성하는 시스템

---

# 📌 프로젝트 개요

이 프로젝트는 다음 과정을 자동화합니다.

1. 엑셀 파일에서 상품 정보 읽기
2. 이미지와 상품 정보를 기반으로 프롬프트 생성
3. 생성형 AI를 활용해 상세페이지용 프롬프트 완성
4. 드랩아트 API를 통해 상세페이지 생성

---

# ⚙️ 시스템 구조

```text
엑셀 + 이미지
    ↓
n8n (워크플로우)
    ↓
프롬프트 생성 (JS + AI)
    ↓
드랩아트 API
    ↓
상세페이지 결과

---

# 📂 폴더 구조
automatic-program/
 ┣ workflow/
 ┃ ┗ detailpage-workflow.json
 ┣ data/
 ┃ ┗ sample/
 ┃   ┣ images/
 ┃   ┣ test.jpg
 ┃   ┗ test.xlsx
 ┣ docker/
 ┃ ┗ docker-compose.yml
 ┣ .gitignore
 ┗ README.md

---

# 🚀 실행 방법
1. Docker 실행
docker compose up -d
2. n8n 접속

브라우저에서 접속:

http://localhost:5679
3. 워크플로우 import
n8n 접속
Import 클릭
workflow/detailpage-workflow.json 선택
4. 테스트 실행

Webhook을 이용하여 테스트:

curl -X POST "http://localhost:5679/webhook-test/upload" \
  -F "image=@data/sample/test.jpg" \
  -F "excel=@data/sample/test.xlsx"

---

# 🧠 핵심 기능
1. 엑셀 데이터 처리
상품명
색상
사이즈
가격
소재
관리자 코멘트
2. 기본 프롬프트 생성
상품 정보 구조화
AI 입력용 draftPrompt 생성
3. AI 기반 프롬프트 생성
이미지 분석
상품 특성 해석
상세페이지 구성 자동 생성

---

# 📌 관리자 코멘트 가이드

관리자 코멘트는 결과 품질에 큰 영향을 줍니다.

예시:

깔끔하고 단정한 분위기 강조
핏이 잘 드러나게 구성
고급스럽고 차분한 느낌

---

# ⚠️ 주의사항
❗ API 키는 업로드 금지

다음 파일은 GitHub에 올리지 마세요:

.env
API Key
개인 설정 파일
❗ 이미지 품질 중요
정면 이미지 권장
상품이 잘 보이도록 촬영
📌 향후 확장 계획
이미지 자동 선택 로직
품질 검수(QC) 자동화
플랫폼별 상세페이지 변환
Python 기반 AI 처리 서버 연동
👥 협업 방법
GitHub 저장소 clone
docker 실행
n8n workflow import
동일 환경에서 작업

---
