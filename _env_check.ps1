# PPTX 플랫폼 환경 체크 스크립트
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$ok  = "[O]"
$ng  = "[X]"
$missingPips = @()

# -Command 방식 실행 시 $scriptDir 가 비어있으므로 환경변수로 보완
$scriptDir = if ($PSScriptRoot -ne '') { $PSScriptRoot } else { (Get-Location).Path }

Write-Host ""
Write-Host " === PPTX 플랫폼 환경 체크 ===" -ForegroundColor Gray
Write-Host ""
Write-Host " -- 기본 분석 환경 --" -ForegroundColor Gray

# Node.js
if (Get-Command node -ErrorAction SilentlyContinue) {
    $v = (node -v 2>&1)
    Write-Host "$ok Node.js $v" -ForegroundColor Green
} else {
    Write-Host "$ng Node.js 없음" -ForegroundColor Red
    Write-Host "     https://nodejs.org 에서 LTS 버전 설치 후 재시도하세요" -ForegroundColor Yellow
}

# Python
if (Get-Command python -ErrorAction SilentlyContinue) {
    $v = (python --version 2>&1)
    Write-Host "$ok $v" -ForegroundColor Green
} else {
    Write-Host "$ng Python 없음" -ForegroundColor Red
    Write-Host "     https://python.org 에서 설치 필요" -ForegroundColor Yellow
    Write-Host '     설치 시 "Add Python to PATH" 반드시 체크' -ForegroundColor Yellow
}

# npm 패키지
$nodeModules = Join-Path $scriptDir "node_modules"
if (Test-Path $nodeModules) {
    Write-Host "$ok npm 패키지 확인" -ForegroundColor Green
} else {
    Write-Host "$ng npm 패키지 없음" -ForegroundColor Red
    Write-Host "     이 폴더에서 아래 명령어 실행:" -ForegroundColor Yellow
    Write-Host "     npm install" -ForegroundColor Yellow
}

# python-pptx / openpyxl
python -c "import pptx, openpyxl" 2>$null
if ($LASTEXITCODE -eq 0) {
    Write-Host "$ok python-pptx / openpyxl 확인" -ForegroundColor Green
} else {
    Write-Host "$ng python-pptx / openpyxl 없음" -ForegroundColor Red
    $missingPips += "python-pptx", "openpyxl"
}

# Java
if (Get-Command java -ErrorAction SilentlyContinue) {
    $v = (java -version 2>&1 | Select-Object -First 1)
    Write-Host "$ok Java: $v" -ForegroundColor Green
} else {
    Write-Host "$ng Java 없음 (rhinoMorph 텍스트 분석 탭에 필요)" -ForegroundColor Red
    Write-Host "     https://www.java.com/ko/download 에서 설치 후 재시도하세요" -ForegroundColor Yellow
}

# JPype1
python -c "import jpype" 2>$null
if ($LASTEXITCODE -eq 0) {
    Write-Host "$ok JPype1 확인" -ForegroundColor Green
} else {
    Write-Host "$ng JPype1 없음" -ForegroundColor Red
    $missingPips += "JPype1"
}

# rhinoMorph
python -c "import rhinoMorph" 2>$null
if ($LASTEXITCODE -eq 0) {
    Write-Host "$ok rhinoMorph 확인" -ForegroundColor Green
} else {
    Write-Host "$ng rhinoMorph 없음" -ForegroundColor Red
    $missingPips += "rhinoMorph"
}

Write-Host ""
Write-Host " -- AI 분석 환경 --" -ForegroundColor Gray

# Ollama CLI
$ollamaOk = $false
if (Get-Command ollama -ErrorAction SilentlyContinue) {
    $v = (ollama --version 2>&1 | Select-Object -First 1)
    Write-Host "$ok Ollama: $v" -ForegroundColor Green
    $ollamaOk = $true
} else {
    Write-Host "$ng Ollama 미설치" -ForegroundColor Red
    Write-Host "     1. https://ollama.com 에서 설치 파일 다운로드 후 설치" -ForegroundColor Yellow
    Write-Host "     2. 설치 완료 후 이 체크를 다시 실행하세요" -ForegroundColor Yellow
}

# 설치된 Ollama 모델 확인
if ($ollamaOk) {
    $list = (ollama list 2>$null | Out-String)
    $found = @()
    if ($list -match "phi4-mini") { $found += "phi4-mini" }
    if ($list -match "gemma4")    { $found += "gemma4" }

    $allLines = ($list -split "`n") | Select-Object -Skip 1 |
        Where-Object { $_.Trim() -ne "" } |
        ForEach-Object { ($_.Trim() -split "\s+")[0] }
    $others = $allLines | Where-Object { $_ -and $_ -ne "phi4-mini" -and $_ -notmatch "^gemma4" }
    if ($others) { $found += $others }

    if ($found.Count -gt 0) {
        Write-Host "$ok 설치된 모델: $($found -join ', ')" -ForegroundColor Green
    } else {
        Write-Host "$ng 권장 모델 미설치 - 다음 중 하나 이상을 설치하세요:" -ForegroundColor Red
        Write-Host "     phi4-mini (약 2.5GB) : ollama run phi4-mini" -ForegroundColor Yellow
        Write-Host "     gemma4    (약 8.1GB) : ollama run gemma4" -ForegroundColor Yellow
    }

    if ($list -notmatch "phi4-mini") {
        Write-Host "   . phi4-mini 미설치 - 설치 명령: ollama run phi4-mini" -ForegroundColor DarkYellow
    }
    if ($list -notmatch "gemma4") {
        Write-Host "   . gemma4    미설치 - 설치 명령: ollama run gemma4" -ForegroundColor DarkYellow
    }
}

# ollama Python 패키지
python -c "import ollama" 2>$null
if ($LASTEXITCODE -eq 0) {
    Write-Host "$ok ollama Python 패키지 확인" -ForegroundColor Green
} else {
    Write-Host "$ng ollama Python 패키지 없음" -ForegroundColor Red
    $missingPips += "ollama"
}

# 누락된 pip 패키지 자동 설치 제안
if ($missingPips.Count -gt 0) {
    Write-Host ""
    Write-Host " 누락된 Python 패키지: $($missingPips -join ', ')" -ForegroundColor Yellow
    $ans = Read-Host " 지금 자동으로 설치할까요? (Y/n)"
    if ($ans -ne 'n' -and $ans -ne 'N') {
        Write-Host ""
        Write-Host " pip install 실행 중..." -ForegroundColor Cyan
        pip install @missingPips
        Write-Host ""
        if ($LASTEXITCODE -eq 0) {
            Write-Host " 설치 완료!" -ForegroundColor Green
        } else {
            Write-Host " 일부 패키지 설치 실패. 위 오류 메시지를 확인하세요." -ForegroundColor Red
        }
    }
}

# 기본 경로 탐색 및 자동 설정
Write-Host ""
Write-Host " [기본 경로] 사용자 이름: $env:USERNAME" -ForegroundColor Gray

$docs = [Environment]::GetFolderPath('MyDocuments')
$browsePath = $docs
$escaped = $browsePath -replace '\\', '\\\\'

$indexFile = Join-Path $scriptDir "public\index.html"
if (Test-Path $indexFile) {
    $content = [IO.File]::ReadAllText($indexFile, [Text.Encoding]::UTF8)
    $m = [regex]::Match($content, "const BROWSE_DEFAULT = '[^']*';")
    if ($m.Success) {
        $newVal  = "const BROWSE_DEFAULT = '" + $escaped + "';"
        $content = $content.Replace($m.Value, $newVal)
    }
    [IO.File]::WriteAllText($indexFile, $content, [System.Text.UTF8Encoding]::new($false))
    Write-Host "$ok 기본 경로 설정 완료: $browsePath" -ForegroundColor Green
} else {
    Write-Host "$ng public\index.html 파일 없음" -ForegroundColor Red
}

Write-Host ""
