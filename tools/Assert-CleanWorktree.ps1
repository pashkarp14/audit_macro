$ErrorActionPreference = "Stop"

$workspaceRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$privateRepoRoot = Join-Path $workspaceRoot "_publish\audit-macros-private"
$privateGuard = Join-Path $privateRepoRoot "tools\Assert-CleanWorktree.ps1"
$currencyWorkbook = Join-Path $workspaceRoot "cbr_currency_2020_2026.xlsx"

Push-Location $workspaceRoot
try {
    $status = @(git status --porcelain=v1 --untracked-files=all)
    if ($LASTEXITCODE -ne 0) { throw "Could not read outer workspace Git status." }
    if ($status.Count -gt 0) {
        throw "Outer workspace is not clean:`n$($status -join [Environment]::NewLine)"
    }

    $diffCheck = @(git diff --check 2>&1)
    if ($LASTEXITCODE -ne 0) { throw "Outer workspace git diff --check failed:`n$($diffCheck -join [Environment]::NewLine)" }

    if (-not (Test-Path -LiteralPath $currencyWorkbook -PathType Leaf)) {
        throw "Required local currency workbook is missing: $currencyWorkbook"
    }
    if (-not (Test-Path -LiteralPath (Join-Path $privateRepoRoot ".git") -PathType Container)) {
        throw "Canonical nested private Git repository is missing: $privateRepoRoot"
    }
    if (-not (Test-Path -LiteralPath $privateGuard -PathType Leaf)) {
        throw "Nested private worktree guard is missing: $privateGuard"
    }

    & powershell.exe -NoLogo -NoProfile -NonInteractive -File $privateGuard
    if ($LASTEXITCODE -ne 0) { throw "Nested private repository is not clean." }

    Write-Output "CLEAN_WORKSPACE_OK"
}
finally {
    Pop-Location
}
