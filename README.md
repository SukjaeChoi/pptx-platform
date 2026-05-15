# PPTX 플랫폼

PPTX 파일을 탐색하고 키워드 검색, 텍스트 분석, 로컬 AI 분석을 수행하는 로컬 웹 플랫폼입니다.

---

## 주요 기능

| 탭 | 배지 | 기능 |
|---|---|---|
| 📋 슬라이드 수 확인 | SCANNER | 지정 폴더의 PPTX 파일 목록과 슬라이드 수를 엑셀로 저장 |
| 🔍 키워드 검색 | SEARCH | PPTX 파일 내 키워드를 검색하여 결과를 엑셀로 저장 |
| 📊 텍스트 분석 | ANALYZE | 형태소 분석 기반 단어 빈도·품사 분포·공출현 네트워크·워드 클라우드 시각화 |
| 🤖 AI 분석 | AI | Ollama 로컬 LLM을 활용한 문서 요약·주제 추출·Q&A·비교 분석·키워드 분석. 설치된 모델을 드롭다운에서 선택 가능 |

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
- 누락된 Python 패키지가 있으면 **자동 설치 여부를 물어봅니다** (Y 입력 시 즉시 설치).
- **Ollama CLI, phi4-mini / gemma4 모델, ollama Python 패키지** 설치 여부를 확인하고 미설치 시 설치 명령어를 안내합니다.
- 현재 사용자 계정을 감지하여 `public/index.html` 의 기본 탐색 경로를 자동으로 설정합니다.

### 4. Python 패키지 설치

`환경체크.bat` 실행 시 자동 설치를 선택하면 아래 과정이 자동으로 진행됩니다. 수동으로 설치하려면:

```bash
pip install python-pptx openpyxl JPype1 rhinoMorph ollama
```

> `rhinoMorph`와 `JPype1`은 텍스트 분석(3번 탭)에만 필요합니다.  
> `ollama`는 AI 분석(4번 탭)에만 필요합니다.

### 5. Node.js 패키지 설치

```bash
npm install
```

### 6. Ollama 모델 설치 (AI 분석 탭 사용 시)

Ollama 설치 후 터미널(CMD 또는 PowerShell)에서 원하는 모델을 설치합니다.  
**두 모델 중 하나 이상**을 설치하면 AI 분석 탭의 드롭다운에서 선택할 수 있습니다.

| 모델 | 크기 | 설치 명령 | 특징 |
|------|------|-----------|------|
| phi4-mini | 약 2.5GB | `ollama run phi4-mini` | 가볍고 빠름 |
| gemma4 | 약 8.1GB | `ollama run gemma4` | Google 최신 모델, 높은 정확도 |

> 모델을 여러 개 설치하면 AI 분석 탭 화면에서 드롭다운으로 즉시 전환할 수 있습니다.


---

## 실행

**`PPT_플랫폼_시작.bat`** 을 더블클릭하면 서버가 시작되고 브라우저가 자동으로 열립니다.

또는 터미널에서 직접 실행할 수 있습니다.

```bash
node server.js
```

브라우저에서 `http://localhost:3100` 으로 접속합니다.

| 서비스 | 주소 |
|--------|------|
| PPTX 플랫폼 | http://localhost:3100 |

---

## 파일 구조

```
pptx-platform/
├── server.js              # Node.js 서버 (API 엔드포인트 포함)
├── pptx_scanner.py        # PPTX 목록 스캐너 (탭 1)
├── pptx_search.py         # 키워드 검색 (탭 2)
├── text_analyzer.py       # 텍스트 분석 — rhinoMorph (탭 3)
├── ai_analyzer.py         # AI 분석 — Ollama 로컬 LLM (탭 4)
├── _env_check.ps1         # 환경 체크 및 자동 설치 스크립트
├── 환경체크.bat            # 환경 확인 실행 (더블클릭)
├── PPT_플랫폼_시작.bat     # 서버 시작 (더블클릭)
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
- Ollama (AI 분석 탭 사용 시) — phi4-mini, gemma4 등 호환 모델
