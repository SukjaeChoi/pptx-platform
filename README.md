# PPTX 플랫폼

PPTX 파일을 탐색하고 키워드 검색, 텍스트 분석, 로컬 AI 분석을 수행하는 로컬 웹 플랫폼입니다.

---

## 주요 기능

| 탭 | 배지 | 기능 |
|---|---|---|
| 📋 슬라이드 수 확인 | SCANNER | 지정 폴더의 PPTX 파일 목록과 슬라이드 수를 엑셀로 저장 |
| 🔍 키워드 검색 | SEARCH | PPTX 파일 내 키워드를 검색하여 결과를 엑셀로 저장 |
| 📊 텍스트 분석 | ANALYZE | 형태소 분석 기반 단어 빈도·품사 분포·공출현 네트워크·워드 클라우드 시각화 |
| 🤖 AI 분석 | AI | 로컬 LLM(phi4-mini)을 활용한 문서 요약·주제 추출·Q&A·비교 분석·키워드 분석 |

---

## AI 분석 탭 상세

로컬 PC에 설치된 Ollama와 phi4-mini 모델을 사용하므로 인터넷 연결 없이, 외부 서버로 데이터를 전송하지 않고 분석합니다.

### 분석 유형

| 유형 | 파일별 동작 | 종합 분석 |
|---|---|---|
| 📋 문서 요약 | 핵심 내용 3문장 이하 요약 | 전체 문서 최종 요약 |
| 🎯 핵심 주제 추출 | 핵심 주제·키워드 3문장 이하 | 공통 주제 및 분포 종합 |
| ❓ 질문 & 답변 생성 | Q:/A: 형식 핵심 질문 2개 | 종합 Q&A 작성 |
| ⚖️ 비교 분석 | 각 문서 특징 요약 (3문장) | 공통점·차이점 비교 |
| 🔑 키워드 분석 | 핵심 키워드 3~5개 설명 | 전체 키워드 종합 |
| ✏️ 직접 입력 | 자유 자연어 질의 (파일별) | 전체 종합 |

### 동작 방식

- **파일 1개 선택**: 단일 분석 결과 카드 출력
- **파일 복수 선택**: ① 파일별 결과를 테이블로 표시 → ② 하단에 전체 종합 분석 카드 출력
- 파일명 클릭 시 해당 PPTX 파일을 바로 열 수 있습니다
- 모델명 입력란에서 phi4-mini 외 다른 로컬 모델로 변경 가능

---

## 설치 방법

### 1. 필수 프로그램 설치

- [Node.js](https://nodejs.org) LTS 버전
- [Python](https://python.org) — 설치 시 **"Add Python to PATH"** 반드시 체크
- [Java](https://www.java.com/ko/download) — 형태소 분석기(rhinoMorph) 실행에 필요
- [Ollama](https://ollama.com) — AI 분석 탭 사용 시 필요

### 2. 파일 다운로드

GitHub 페이지에서 **Code → Download ZIP** 을 클릭하여 압축을 풀거나, 아래 명령어로 클론합니다.

```bash
git clone https://github.com/SukjaeChoi/pptx-platform.git
cd pptx-platform
```

### 3. 환경 체크 및 설정

다운받은 폴더에서 **`환경체크.bat`** 을 실행합니다.

- Node.js, Python, Java, 필수 패키지 설치 여부를 자동으로 확인합니다.
- **Ollama CLI, phi4-mini 모델, ollama Python 패키지** 설치 여부를 확인하고 미설치 시 방법을 안내합니다.
- 현재 사용자 계정을 감지하여 `public/index.html` 의 기본 탐색 경로를 자동으로 설정합니다.
- 미설치 항목이 있으면 설치 명령어를 안내합니다.

### 4. Python 패키지 설치

```bash
pip install python-pptx openpyxl JPype1 rhinoMorph ollama
```

> `rhinoMorph`와 `JPype1`은 텍스트 분석(3번 탭)에만 필요합니다.  
> `ollama`는 AI 분석(4번 탭)에만 필요합니다.

### 5. Node.js 패키지 설치

```bash
npm install
```

### 6. phi4-mini 모델 설치 (AI 분석 탭 사용 시)

Ollama 설치 후 터미널(CMD 또는 PowerShell)에서 아래 명령어를 실행합니다.  
최초 실행 시 모델을 자동으로 다운로드합니다 (약 2.5GB).

```bash
ollama run phi4-mini
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
├── server.js              # Node.js 서버 (API 엔드포인트 포함)
├── pptx_scanner.py        # PPTX 목록 스캐너 (탭 1)
├── pptx_search.py         # 키워드 검색 (탭 2)
├── text_analyzer.py       # 텍스트 분석 — rhinoMorph (탭 3)
├── ai_analyzer.py         # AI 분석 — ollama/phi4-mini (탭 4)
├── 환경체크.bat            # 환경 확인 및 기본 경로 자동 설정
├── PPT_플랫폼_시작.bat     # 서버 시작
├── package.json
└── public/
    └── index.html         # 웹 프론트엔드 (4탭 SPA)
```

---

## 사용 환경

- Windows 10 / 11
- Node.js 18 이상
- Python 3.9 이상
- Java 8 이상 (rhinoMorph 실행에 필요)
- Ollama (AI 분석 탭 사용 시)
- phi4-mini 모델 또는 다른 Ollama 호환 모델
