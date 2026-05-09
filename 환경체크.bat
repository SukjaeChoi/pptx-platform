@echo off
cd /d "%~dp0"
echo.
echo  === PPT 플랫폼 환경 체크 ===
echo.

:: Node.js 확인
where node > nul 2>&1
if errorlevel 1 (
  echo [X] Node.js 없음 - https://nodejs.org 에서 LTS 설치 필요
) else (
  for /f "tokens=*" %%v in ('node -v') do echo [O] Node.js %%v
)

:: Python 확인
where python > nul 2>&1
if errorlevel 1 (
  echo [X] Python 없음 - https://python.org 에서 설치 필요
  echo     설치 시 "Add Python to PATH" 반드시 체크
) else (
  for /f "tokens=*" %%v in ('python --version') do echo [O] %%v
)

:: npm 패키지 확인
if not exist "node_modules\" (
  echo [X] npm 패키지 없음
  echo     설치하려면 이 창에서 아래 명령어 실행:
  echo     npm install
) else (
  echo [O] npm 패키지 확인
)

:: Python 패키지 확인
python -c "import pptx, openpyxl" > nul 2>&1
if errorlevel 1 (
  echo [X] Python 패키지 없음
  echo     설치하려면 이 창에서 아래 명령어 실행:
  echo     pip install python-pptx openpyxl
) else (
  echo [O] python-pptx / openpyxl 확인
)


:: Java 확인
where java > nul 2>&1
if errorlevel 1 (
  echo [X] Java 없음 - rhinoMorph 실행에 필요합니다 [cite: 4]
  echo     https://www.java.com/ko/download 에서 설치 후 재시도하세요 [cite: 4]
) else (
  :: for문 안에서 직접 버전을 출력하고 루프를 마칩니다.
  for /f "tokens=*" %%v in ('java -version 2^>^&1') do (
    echo [O] %%v
    goto :next_step
  )
)

:next_step
:: Python 패키지 확인 - JPype1
python -c "import jpype" > nul 2>&1
if errorlevel 1 (
  echo [X] JPype1 없음
  echo     설치하려면 이 창에서 아래 명령어 실행:
  echo     pip install JPype1
) else (
  echo [O] JPype1 확인
)

:: Python 패키지 확인 - rhinoMorph
python -c "import rhinoMorph" > nul 2>&1
if errorlevel 1 (
  echo [X] rhinoMorph 없음
  echo     설치하려면 이 창에서 아래 명령어 실행:
  echo     pip install rhinoMorph
) else (
  echo [O] rhinoMorph 확인
)


:: ── 기본 탐색 경로 자동 설정 ─────────────────────────────────────────────
echo.
echo  [경로 설정] 사용자 계정: %USERNAME%

set "BROWSE_PATH="
if exist "%USERPROFILE%\OneDrive\문서\" (
  set "BROWSE_PATH=%USERPROFILE%\OneDrive\문서"
)
if not defined BROWSE_PATH (
  if exist "%USERPROFILE%\OneDrive\Documents\" (
    set "BROWSE_PATH=%USERPROFILE%\OneDrive\Documents"
  )
)
if not defined BROWSE_PATH (
  if exist "%USERPROFILE%\문서\" (
    set "BROWSE_PATH=%USERPROFILE%\문서"
  )
)
if not defined BROWSE_PATH (
  if exist "%USERPROFILE%\Documents\" (
    set "BROWSE_PATH=%USERPROFILE%\Documents"
  )
)
if not defined BROWSE_PATH set "BROWSE_PATH=%USERPROFILE%"

set "BP=%BROWSE_PATH%"
set "IDX=%~dp0public\index.html"
powershell -NoProfile -Command "$p=$env:BP -replace '\\','\\\\';$f=$env:IDX;if(Test-Path $f){$c=[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8);$c=$c -replace 'const BROWSE_DEFAULT = ''[^'']*'';',('const BROWSE_DEFAULT = '''+$p+''';');[IO.File]::WriteAllText($f,$c,[System.Text.UTF8Encoding]::new($false));Write-Host '[O] 기본 경로 설정 완료:' $env:BP}else{Write-Host '[X] public\index.html 파일 없음'}"

echo.
pause
