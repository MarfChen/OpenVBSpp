<#
Run the OpenVBS++ demo, one section per headline feature.

Usage (Windows PowerShell 5.1 or PowerShell 7):
    powershell -NoProfile -ExecutionPolicy Bypass -File examples\demo\run-demo.ps1

Every section prints its title, the output, and the wall-clock time.
README.md in this folder explains how to record the GIF with ScreenToGif.
#>

$ErrorActionPreference = 'Stop'

$demoDir = $PSScriptRoot
$repoRoot = Split-Path -Parent (Split-Path -Parent $demoDir)

$python = Join-Path $repoRoot 'venv\Scripts\python.exe'
if (-not (Test-Path $python)) { $python = 'python' }

# File = script passed to the runner; Args = raw arguments to `python -m OpenVBSPP`.
$steps = @(
    @{ Title = '1/6  legacy VBScript, unchanged';      File = '01-legacy.vbs'  },
    @{ Title = '2/6  classes, inheritance, Try/Catch'; File = '02-modern.vbs'  },
    @{ Title = '3/6  Python interop (#python, COP)';   File = '03-python.vbs'  },
    @{ Title = '4/6  real COM (WScript.Shell, WMI)';   File = '04-com.vbs'     },
    @{ Title = '5/6  multi-file namespaces';           File = '05-modules.vbs' },
    @{ Title = '6/6  package manager (bundled packs)'; Args = @('over', 'list') }
)

$total = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($step in $steps) {
    Write-Host ''
    Write-Host ('-- ' + $step.Title + ' ' + ('-' * [Math]::Max(0, 46 - $step.Title.Length))) -ForegroundColor Cyan

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    if ($step.Args) {
        Write-Host ('   openvbs ' + ($step.Args -join ' '))
        Push-Location $demoDir
        try { & $python -m OpenVBSPP @($step.Args) }
        finally { Pop-Location }
    }
    else {
        Write-Host ('   ' + $step.File)
        & $python -m OpenVBSPP (Join-Path $demoDir $step.File)
    }
    $sw.Stop()
    $exit = $LASTEXITCODE

    Write-Host ('   -> ' + $sw.Elapsed.TotalMilliseconds.ToString('0') + ' ms, exit ' + $exit) -ForegroundColor DarkGray
}

$total.Stop()
Write-Host ''
Write-Host ('-- done in ' + $total.Elapsed.TotalSeconds.ToString('0.0') + ' s') -ForegroundColor Green
Write-Host '   next: openvbs ide     (web IDE in your browser)'
