# PPTX 플랫폼

PPTX 파일을 탐색하고 키워드 검색 및 텍스트 분석을 수행하는 로컬 웹 플랫폼입니다.

---

## 주요 기능

| 탭 | 기능 |
|---|---|
| 파일 스캐너 | 지정 폴더의 PPTX 파일 목록과 슬라이드 수를 엑셀로 저장 |
| 키워드 검색 | PPTX 파일 내 키워드를 검색하여 결과를 엑셀로 저장 |
| 텍스트 분석 | 형태소 분석 기반 단어 빈도·품사 분포·공출현 네트워크·워드 클라우드 시각화 |

---

## 설치 방법

### 1. 필수 프로그램 설치

- [Node.js](https://nodejs.org) LTS 버전
- [Python](https://python.org) — 설치 시 **"Add Python to PATH"** 반드시 체크
- [Java](https://www.java.com/ko/download) — 형태소 분석기(rhinoMorph) 실행에 필요

### 2. 파일 다운로드

GitHub 페이지에서 **Code → Download ZIP** 을 클릭하여 압축을 풀거나, 아래 명령어로 클론합니다.

```bash
git clone https://github.com/SukjaeChoi/pptx-platform.git
cd pptx-platform
```

### 3. 환경 체크 및 설정

다운받은 폴더에서 **`환경체크.bat`** 을 실행합니다.

- Node.js, Python, Java, 필수 패키지 설치 여부를 자동으로 확인합니다.
- 현재 사용자 계정을 감지하여 `public/index.html` 의 기본 탐색 경로를 자동으로 설정합니다.
- 미설치 항목이 있으면 설치 명령어를 안내합니다.

### 4. Python 패키지 설치

```bash
pip install python-pptx openpyxl JPype1 rhinoMorph
```

### 5. Node.js 패키지 설치

```bash
npm install
```

---

## 실행

**`PPT_플랫폼_시작.bat`** 을 더블클릭하면 서버가 시작되고 브라우저가 자동으로 열립니다.

또는 터미널에서 직접 실행할 수 있습니다.

```bash
node server.js
```

브라우저에서 `http://localhost:3000` 으로 접속합니다.

---

## 파일 구조

```
pptx-platform/
├── server.js              # Node.js 서버
├── pptx_scanner.py        # PPTX 목록 스캐너
├── pptx_search.py         # 키워드 검색
├── text_analyzer.py       # 텍스트 분석 (rhinoMorph)
├── 환경체크.bat            # 환경 확인 및 기본 경로 자동 설정
├── PPT_플랫폼_시작.bat     # 서버 시작
├── package.json
└── public/
    └── index.html         # 웹 프론트엔드
```

---

## 사용 환경

- Windows 10 / 11
- Node.js 18 이상
- Python 3.9 이상
- Java 8 이상 (rhinoMorph 실행에 필요)
