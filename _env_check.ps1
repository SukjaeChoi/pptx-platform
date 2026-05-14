# PPTX 플랫폼 환경 체크 스크립트
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$ok  = "[O]"
$ng  = "[X]"

Write-Host ""
Write-Host " === PPTX 플랫폼 환경 체크 ===" -ForegroundColor Cyan
Write-Host ""

# ── Node.js ────────────────────────────────────────────────────────────────
if (Get-Command node -ErrorAction SilentlyContinue) {
    $v = (node -v 2>&1)
    Write-Host "$ok Node.js $v" -ForegroundColor Green
} else {
    Write-Host "$ng Node.js 없음 - https://nodejs.org 에서 LTS 설치 필요" -ForegroundColor Red
}

# ── Python ─────────────────────────────────────────────────────────────────
if (Get-Command python -ErrorAction SilentlyContinue) {
    $v = (python --version 2>&1)
    Write-Host "$ok $v" -ForegroundColor Green
} else {
    Write-Host "$ng Python 없음 - https://python.org 에서 설치 필요" -ForegroundColor Red
    Write-Host '     설치 시 "Add Python to PATH" 반드시 체크' -ForegroundColor Yellow
}

# ── npm 패키지 ─────────────────────────────────────────────────────────────
$nodeModules = Join-Path $PSScriptRoot "node_modules"
if (Test-Path $nodeModules) {
    Write-Host "$ok npm 패키지 확인" -ForegroundColor Green
} else {
    Write-Host "$ng npm 패키지 없음" -ForegroundColor Red
    Write-Host "     설치하려면 이 폴더에서 아래 명령어 입력:" -ForegroundColor Yellow
    Write-Host "     npm install" -ForegroundColor Yellow
}

# ── python-pptx / openpyxl ────────────────────────────────────────────────
$r = python -c "import pptx, openpyxl" 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-Host "$ok python-pptx / openpyxl 확인" -ForegroundColor Green
} else {
    Write-Host "$ng python-pptx / openpyxl 없음" -ForegroundColor Red
    Write-Host "     pip install python-pptx openpyxl" -ForegroundColor Yellow
}

# ── Java ───────────────────────────────────────────────────────────────────
if (Get-Command java -ErrorAction SilentlyContinue) {
    $v = (java -version 2>&1 | Select-Object -First 1)
    Write-Host "$ok Java: $v" -ForegroundColor Green
} else {
    Write-Host "$ng Java 없음 - rhinoMorph 사용에 필요합니다" -ForegroundColor Red
    Write-Host "     https://www.java.com/ko/download 에서 설치 후 재시도하세요" -ForegroundColor Yellow
}

# ── JPype1 ─────────────────────────────────────────────────────────────────
$r = python -c "import jpype" 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-Host "$ok JPype1 확인" -ForegroundColor Green
} else {
    Write-Host "$ng JPype1 없음" -ForegroundColor Red
    Write-Host "     pip install JPype1" -ForegroundColor Yellow
}

# ── rhinoMorph ─────────────────────────────────────────────────────────────
$r = python -c "import rhinoMorph" 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-Host "$ok rhinoMorph 확인" -ForegroundColor Green
} else {
    Write-Host "$ng rhinoMorph 없음" -ForegroundColor Red
    Write-Host "     pip install rhinoMorph" -ForegroundColor Yellow
}

Write-Host ""
Write-Host " -- AI 분석 환경 --" -ForegroundColor Cyan

# ── Ollama CLI ─────────────────────────────────────────────────────────────
if (Get-Command ollama -ErrorAction SilentlyContinue) {
    $v = (ollama --version 2>&1 | Select-Object -First 1)
    Write-Host "$ok Ollama: $v" -ForegroundColor Green
} else {
    Write-Host "$ng Ollama 미설치" -ForegroundColor Red
    Write-Host "     1. https://ollama.com 에서 설치 파일 다운로드 후 설치" -ForegroundColor Yellow
    Write-Host "     2. 설치 완료 후 이 체크를 다시 실행하세요" -ForegroundColor Yellow
}

# ── phi4-mini 모델 ─────────────────────────────────────────────────────────
if (Get-Command ollama -ErrorAction SilentlyContinue) {
    $list = ollama list 2>$null
    if ($list -match "phi4-mini") {
        Write-Host "$ok phi4-mini 모델 확인" -ForegroundColor Green
    } else {
        Write-Host "$ng phi4-mini 모델 없음" -ForegroundColor Red
        Write-Host "     터미널에서 아래 명령어 실행 (최초 실행시 약 2.5GB 다운로드):" -ForegroundColor Yellow
        Write-Host "     ollama run phi4-mini" -ForegroundColor Yellow
    }
}

# ── ollama Python 패키지 ───────────────────────────────────────────────────
$r = python -c "import ollama" 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-Host "$ok ollama Python 패키지 확인" -ForegroundColor Green
} else {
    Write-Host "$ng ollama Python 패키지 없음" -ForegroundColor Red
    Write-Host "     pip install ollama" -ForegroundColor Yellow
}

# ── 기본 경로 탐색 및 자동 설정 ───────────────────────────────────────────
Write-Host ""
Write-Host " [기본 경로] 사용자 이름: $env:USERNAME" -ForegroundColor Cyan

$docs = [Environment]::GetFolderPath('MyDocuments')
$browsePath = $docs   # MyDocuments가 OneDrive\문서 또는 문서를 자동 반환

$indexFile = Join-Path $PSScriptRoot "public\index.html"
if (Test-Path $indexFile) {
    $escaped = $browsePath -replace '\\', '\\\\'
    $content = [IO.File]::ReadAllText($indexFile, [Text.Encoding]::UTF8)
    $content = $content -replace "const BROWSE_DEFAULT = '[^']*';", "const BROWSE_DEFAULT = '$escaped';"
    [IO.File]::WriteAllText($indexFile, $content, [System.Text.UTF8Encoding]::new($false))
    Write-Host "$ok 기본 경로 설정 완료: $browsePath" -ForegroundColor Green
} else {
    Write-Host "$ng public\index.html 파일 없음" -ForegroundColor Red
}

Write-Host ""
