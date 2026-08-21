# 총무팀 자동화 허브 빌드 스크립트
#
# 빌드.bat 이 이 파일을 호출한다.
# cmd 배치파일에 한글을 넣으면 파일 인코딩과 콘솔 코드페이지가 어긋나면서
# 줄 앞 글자가 잘려나간다(실행 방식마다 코드페이지가 달라 재현이 불규칙하다).
# 그래서 배치파일은 ASCII 만 담은 실행기로 두고, 실제 작업은 여기서 한다.
#
# 이 파일은 UTF-8 with BOM 으로 저장해야 한다.
# Windows PowerShell 5.1 은 BOM 이 없으면 ANSI 로 읽어 한글 파일명이 깨진다.

$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot

$buildDate = Get-Date -Format 'yyyyMMdd'
$appName = "총무팀_자동화_허브_$buildDate"

Write-Host '===================================================='
Write-Host ' 총무팀 자동화 허브 빌드'
Write-Host '===================================================='
Write-Host " 파일명 : $appName.exe"
Write-Host " 위치   : $PSScriptRoot\dist"
Write-Host ''

# ── 허브에 포함할 모듈 ────────────────────────────────────────────────
$dataFiles = @(
    'mod_야근수당.py'
    'mod_야근현황.py'
    'mod_현장경비.py'
    'mod_법률검토.py'
    'law_api.py'
    'law_config.py'
    '야근수당_자동입력.py'
    'mod_4대보험.py'
    'mod_4대보험_parsers.py'
    'mod_4대보험_matcher.py'
    'mod_4대보험_report.py'
    'mod_증명서발급.py'
    'mod_계약서색인.py'
    'mod_보증금현황.py'
    'mod_전산기기.py'
    'wehago_client.py'
)

$hiddenImports = @(
    'tkinterdnd2'
    'openpyxl'
    'xlrd'
    'xlwt'
    'xlutils'
    'xlutils.copy'
    'pandas'
    'win32com.client'
    'requests'
    'keyring.backends.Windows'
)

# 빠진 모듈이 있으면 빌드해도 그 기능이 실행되지 않으므로 미리 확인한다.
$missing = $dataFiles | Where-Object { -not (Test-Path -LiteralPath $_) }
if ($missing) {
    Write-Host '[중단] 다음 파일을 찾을 수 없습니다:' -ForegroundColor Red
    $missing | ForEach-Object { Write-Host "   - $_" -ForegroundColor Red }
    exit 1
}

Write-Host '[1/2] 필요 패키지 확인 중...'
python -m pip install pyinstaller tkinterdnd2 pandas openpyxl xlrd xlwt xlutils pywin32 keyring selenium --quiet
Write-Host ''

Write-Host '[2/2] PyInstaller 빌드 중...'

# 인자를 배열로 넘긴다. 문자열로 이어붙이면 따옴표·세미콜론 처리에서 깨진다.
$pyiArgs = @('--onefile', '--windowed', '--name', $appName, '--noconfirm')
foreach ($f in $dataFiles) { $pyiArgs += '--add-data'; $pyiArgs += "$f;." }
foreach ($h in $hiddenImports) { $pyiArgs += '--hidden-import'; $pyiArgs += $h }
$pyiArgs += '--collect-all'; $pyiArgs += 'selenium'
$pyiArgs += 'hub.py'

python -m PyInstaller @pyiArgs

Write-Host ''
$exe = Join-Path $PSScriptRoot "dist\$appName.exe"
if (Test-Path -LiteralPath $exe) {
    $size = [math]::Round((Get-Item -LiteralPath $exe).Length / 1MB, 1)
    Write-Host '===================================================='
    Write-Host ' 빌드 성공' -ForegroundColor Green
    Write-Host " 파일 : dist\$appName.exe  ($size MB)"
    Write-Host '===================================================='

    # 이전 버전 안내 — 같은 폴더에 날짜별로 쌓이므로 정리 시점을 알려준다.
    $others = Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'dist') -Filter '총무팀_자동화_허브_*.exe' |
              Where-Object { $_.Name -ne "$appName.exe" } | Sort-Object Name -Descending
    if ($others) {
        Write-Host ''
        Write-Host " 이전 버전 $($others.Count)개가 dist 폴더에 남아 있습니다:"
        $others | Select-Object -First 5 | ForEach-Object { Write-Host "   $($_.Name)" }
        if ($others.Count -gt 5) { Write-Host "   ... 외 $($others.Count - 5)개" }
    }
    explorer.exe (Join-Path $PSScriptRoot 'dist')
} else {
    Write-Host '===================================================='
    Write-Host ' 빌드 실패 - 위 오류 메시지를 확인하세요' -ForegroundColor Red
    Write-Host '===================================================='
    exit 1
}
