$ErrorActionPreference = 'Stop'

$ahk = 'C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe'
if (!(Test-Path $ahk)) {
    throw "AutoHotkey runtime not found at $ahk"
}

$repoRoot = Split-Path -Parent $PSScriptRoot

$tests = @(
    @{ Path = (Join-Path $repoRoot 'Tests\TransferBridgeProtocol.Tests.ahk'); Args = @(); PassToken = 'PASS TransferBridgeProtocol.Tests' },
    @{ Path = (Join-Path $repoRoot 'Tests\TransferShared.Tests.ahk'); Args = @('alpha', 'beta'); PassToken = 'PASS TransferShared.Tests' },
    @{ Path = (Join-Path $repoRoot 'Tests\CitrixTransferLib.Tests.ahk'); Args = @(); PassToken = 'PASS CitrixTransferLib.Tests' }
)

foreach ($test in $tests) {
    Write-Output "Running $($test.Path)"
    $output = & $ahk /ErrorStdOut $test.Path @($test.Args) 2>&1 | Out-String
    Write-Output $output.Trim()

    if ($output -match 'FAIL ') {
        throw "Test failed: $($test.Path)"
    }
    if ($output -notmatch [regex]::Escape($test.PassToken)) {
        throw "Missing PASS token: $($test.Path)"
    }
}

Write-Output 'All AHK tests passed.'
