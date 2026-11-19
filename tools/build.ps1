# 构建脚本：wheel / sdist / exe
#
# 用法：
#   .\tools\build.ps1                 # wheel + sdist
#   .\tools\build.ps1 -Target wheel   # 只构建 wheel + sdist
#   .\tools\build.ps1 -Target exe     # 只构建单文件 exe
#   .\tools\build.ps1 -Target all     # 两者都构建
#   .\tools\build.ps1 -Target test    # 只跑测试

param(
    [ValidateSet('wheel', 'exe', 'all', 'test')]
    [string]$Target = 'wheel',
    [switch]$Clean
)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

$Python = Join-Path $Root 'venv\Scripts\python.exe'
if (-not (Test-Path $Python)) {
    # 兜底：用当前 PATH 里的 python
    $Python = 'python'
}

function Invoke-Tests {
    Write-Host '== 运行测试 ==' -ForegroundColor Cyan
    & $Python -m unittest discover -s tests
    if ($LASTEXITCODE -ne 0) { throw '测试失败，构建中止' }
}

function Build-Wheel {
    Write-Host '== 构建 wheel + sdist ==' -ForegroundColor Cyan
    & $Python -m build
    if ($LASTEXITCODE -ne 0) { throw '构建失败' }
    Get-ChildItem (Join-Path $Root 'dist') | Select-Object Name, Length
}

function Build-Exe {
    Write-Host '== 构建 openvbs.exe ==' -ForegroundColor Cyan
    & $Python -c "import PyInstaller" 2>$null
    if ($LASTEXITCODE -ne 0) {
        Write-Host '安装 PyInstaller...' -ForegroundColor Yellow
        & $Python -m pip install -q "pyinstaller>=6.0"
    }
    & $Python -m PyInstaller (Join-Path $Root 'packaging\openvbs.spec') `
        --noconfirm --distpath (Join-Path $Root 'dist') `
        --workpath (Join-Path $Root 'build\pyinstaller')
    if ($LASTEXITCODE -ne 0) { throw 'PyInstaller 失败' }
    Get-ChildItem (Join-Path $Root 'dist') -Filter *.exe | Select-Object Name, Length
    Write-Host '试试：dist\openvbs.exe examples\control1.vbs' -ForegroundColor Green
}

if ($Clean) {
    Remove-Item -Recurse -Force (Join-Path $Root 'dist') -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force (Join-Path $Root 'build') -ErrorAction SilentlyContinue
}

switch ($Target) {
    'test'  { Invoke-Tests }
    'wheel' { Invoke-Tests; Build-Wheel }
    'exe'   { Invoke-Tests; Build-Exe }
    'all'   { Invoke-Tests; Build-Wheel; Build-Exe }
}
