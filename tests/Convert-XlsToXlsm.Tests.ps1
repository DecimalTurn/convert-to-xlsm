$ErrorActionPreference = 'Stop'
$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("convert-tests-" + [guid]::NewGuid())
$scriptDir = Join-Path $testRoot 'scripts'
$inputDir = Join-Path $testRoot 'input'
$outputDir = Join-Path $testRoot 'output'
$logPath = Join-Path $testRoot 'conversions.log'
$pwsh = (Get-Process -Id $PID).Path

try {
    New-Item -ItemType Directory -Path "$scriptDir/utils", $inputDir, $outputDir -Force | Out-Null
    Copy-Item "$PSScriptRoot/../scripts/Convert-XlsToXlsm.ps1" $scriptDir
    Set-Content -Path "$scriptDir/utils/Screenshot.ps1" -Value 'function Take-Screenshot { param($OutputPath) }'
    @'
param($ScriptPath, $InputDir, $OutputDir, $LogPath)
function Get-Process { param($Name, $ErrorAction) }
function Start-Job {
    param($ScriptBlock, $ArgumentList)
    Add-Content -Path $LogPath -Value ($ArgumentList[0] + '|' + $ArgumentList[1])
    [pscustomobject]@{ State = 'Completed' }
}
function Wait-Job { param($Job, $Timeout) $Job }
function Receive-Job { param($Job) }
function Remove-Job { param($Job, $ErrorAction) }
& $ScriptPath -InputDir $InputDir -OutputDir $OutputDir
if ($null -ne $LASTEXITCODE) { exit $LASTEXITCODE }
'@ | Set-Content -Path "$testRoot/run.ps1"

    function Invoke-Conversion {
        param($Destination = $outputDir)
        $output = & $pwsh -NoProfile -File "$testRoot/run.ps1" `
            -ScriptPath "$scriptDir/Convert-XlsToXlsm.ps1" `
            -InputDir $inputDir -OutputDir $Destination -LogPath $logPath 2>&1
        [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = $output }
    }

    function Assert-True {
        param($Condition, $Message)
        if (-not $Condition) { throw $Message }
    }

    New-Item -ItemType File -Path "$inputDir/legacy.xls", "$inputDir/older.XLSM", "$inputDir/ignore.xlsx" | Out-Null
    $result = Invoke-Conversion
    Assert-True ($result.ExitCode -eq 0) "Mixed conversion failed: $($result.Output)"
    $calls = @(Get-Content $logPath)
    Assert-True ($calls.Count -eq 2) "Expected exactly two conversion jobs, got $($calls.Count)."
    $legacyCall = (Join-Path $inputDir 'legacy.xls') + '|' + (Join-Path $outputDir 'legacy.xlsm')
    $olderCall = (Join-Path $inputDir 'older.XLSM') + '|' + (Join-Path $outputDir 'older.xlsm')
    Assert-True (($calls | Where-Object { $_ -eq $legacyCall }).Count -eq 1) 'The .xls input was not converted.'
    Assert-True (($calls | Where-Object { $_ -eq $olderCall }).Count -eq 1) 'The .xlsm input was not converted.'

    Remove-Item "$inputDir/legacy.xls", "$inputDir/older.XLSM", $logPath
    New-Item -ItemType File -Path "$inputDir/book.xls", "$inputDir/book.xlsm" | Out-Null
    $result = Invoke-Conversion
    Assert-True ($result.ExitCode -ne 0) 'Matching output names should fail.'
    Assert-True (-not (Test-Path $logPath)) 'Collision should be detected before starting any jobs.'

    Remove-Item "$inputDir/book.xls"
    $result = Invoke-Conversion -Destination $inputDir
    Assert-True ($result.ExitCode -ne 0) 'An output that overwrites its input should fail.'
    Assert-True (-not (Test-Path $logPath)) 'Input overwrite should be detected before starting any jobs.'
    Assert-True (Test-Path "$inputDir/book.xlsm") 'The input workbook should remain intact.'

    Remove-Item "$inputDir/book.xlsm"
    $result = Invoke-Conversion
    Assert-True ($result.ExitCode -eq 0) "Empty input failed: $($result.Output)"
    Assert-True (-not (Test-Path $logPath)) 'An unsupported input should not start a job.'
    Write-Host 'Conversion script tests passed.'
} finally {
    Remove-Item -LiteralPath $testRoot -Recurse -Force
}
