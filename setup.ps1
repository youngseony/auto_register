# AI디지털배움터 등록 업무 자동화 도구(auto_register) - 원클릭 설치 스크립트
# Python 확인/설치 -> 가상환경(venv) 생성 -> 라이브러리 설치 -> Chrome 확인
# 직접 실행하려면:  powershell -ExecutionPolicy Bypass -File .\setup.ps1

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

function Refresh-Path {
    $env:Path = [Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [Environment]::GetEnvironmentVariable("Path", "User")
}

function Test-PythonExe($exe) {
    # 3.9 이상인 진짜 Python인지 확인 (Microsoft Store 안내용 가짜 python.exe 제외)
    try {
        & $exe -c "import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)" 2>$null
        return ($LASTEXITCODE -eq 0)
    } catch {
        return $false
    }
}

function Find-Python {
    # 1) py 런처
    if (Get-Command py -ErrorAction SilentlyContinue) {
        try {
            $exe = (& py -3 -c "import sys; print(sys.executable)" 2>$null | Select-Object -First 1)
            if ($exe -and (Test-PythonExe $exe)) { return $exe }
        } catch {}
    }
    # 2) PATH 의 python
    $cmd = Get-Command python -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Source -notlike "*\WindowsApps\*" -and (Test-PythonExe $cmd.Source)) { return $cmd.Source }
    # 3) 사용자 설치 위치
    $found = Get-ChildItem "$env:LOCALAPPDATA\Programs\Python\Python3*\python.exe" -ErrorAction SilentlyContinue |
        Sort-Object FullName -Descending
    foreach ($f in $found) { if (Test-PythonExe $f.FullName) { return $f.FullName } }
    return $null
}

function Fail($message) {
    Write-Host ""
    Write-Host "[오류] $message" -ForegroundColor Red
    exit 1
}

Write-Host "=== 1/4 Python 확인 ===" -ForegroundColor Cyan
$py = Find-Python
if (-not $py) {
    Write-Host "Python 3.9 이상이 없어 설치합니다. (winget 사용, 1~3분 소요)"
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Start-Process "https://www.python.org/downloads/"
        Fail "winget 을 사용할 수 없습니다. 열린 페이지에서 Python 을 직접 설치(Add Python to PATH 체크)한 뒤 setup.bat 를 다시 실행하세요."
    }
    winget install -e --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements
    Refresh-Path
    $py = Find-Python
    if (-not $py) { Fail "Python 설치 후에도 찾을 수 없습니다. 창을 닫고 setup.bat 를 다시 실행해 보세요." }
}
Write-Host "사용할 Python: $py"

Write-Host ""
Write-Host "=== 2/4 가상환경 준비 ===" -ForegroundColor Cyan
$venvPy = Join-Path $PSScriptRoot "venv\Scripts\python.exe"
if ((Test-Path $venvPy) -and -not (Test-PythonExe $venvPy)) {
    Write-Host "기존 venv 가 손상되어 다시 만듭니다."
    Remove-Item -Recurse -Force (Join-Path $PSScriptRoot "venv")
}
if (-not (Test-Path $venvPy)) {
    & $py -m venv venv
    if ($LASTEXITCODE -ne 0) { Fail "가상환경 생성에 실패했습니다." }
}

Write-Host ""
Write-Host "=== 3/4 라이브러리 설치 ===" -ForegroundColor Cyan
& $venvPy -m pip install --disable-pip-version-check -r requirements.txt
if ($LASTEXITCODE -ne 0) { Fail "라이브러리 설치에 실패했습니다. 인터넷 연결을 확인하고 다시 실행하세요." }

Write-Host ""
Write-Host "=== 4/4 Chrome 확인 ===" -ForegroundColor Cyan
$chromePaths = @(
    "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
    "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
)
if ($chromePaths | Where-Object { Test-Path $_ }) {
    Write-Host "Chrome 이 설치되어 있습니다."
} else {
    Write-Host "Chrome 브라우저를 찾지 못했습니다. 자동화에는 Chrome 이 필요합니다." -ForegroundColor Yellow
    $answer = Read-Host "지금 winget 으로 설치할까요? (Y/N)"
    if ($answer -match "^[Yy]") {
        if (Get-Command winget -ErrorAction SilentlyContinue) {
            winget install -e --id Google.Chrome --silent --accept-package-agreements --accept-source-agreements
        } else {
            Start-Process "https://www.google.com/chrome/"
            Write-Host "열린 페이지에서 Chrome 을 설치하세요." -ForegroundColor Yellow
        }
    } else {
        Write-Host "https://www.google.com/chrome/ 에서 Chrome 을 설치한 뒤 사용하세요." -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "설치가 끝났습니다. 이제 run.bat 를 실행하세요." -ForegroundColor Green
