<#
.SYNOPSIS
  ミラー（複製）同期を行う PowerShell ラッパー（Robocopy 使用）。

.DESCRIPTION
  指定した `Source` フォルダから `Destination` フォルダへローカルミラー同期します。
  `/MIR` を使うため、ソースで削除したファイルは宛先からも削除されます。

.EXAMPLE
  # 実行（実際にコピー・削除を適用）
  .\sync-mirror.ps1 -Source "C:\path\to\source" -Destination "D:\backup\dest"

  # ドライラン（何が起こるか表示するだけ）
  .\sync-mirror.ps1 -Source "C:\path\to\source" -Destination "D:\backup\dest" -WhatIf
#>

param(
    [Parameter(Mandatory=$true)][string]$Source,
    [Parameter(Mandatory=$true)][string]$Destination,
    [switch]$WhatIf,
    [string]$LogFile = "$PSScriptRoot\sync-mirror.log"
)

if (-not (Test-Path $Source)) {
    Write-Error "Source path does not exist: $Source"
    exit 2
}

# Ensure destination exists (robocopy will create it but make explicit for clarity)
if (-not (Test-Path $Destination)) {
    New-Item -ItemType Directory -Path $Destination -Force | Out-Null
}

$args = @($Source, $Destination, "/MIR", "/Z", "/XA:SH", "/W:5", "/R:2", "/NP", "/TEE", "/LOG+:$LogFile")
if ($WhatIf) { $args += "/L" }

Write-Output "Starting mirror sync"
Write-Output "Source: $Source"
Write-Output "Destination: $Destination"
Write-Output "Log: $LogFile"
if ($WhatIf) { Write-Output "Mode: Dry-run (no changes)" }

Write-Output "Running: robocopy $($args -join ' ')"
& robocopy @args

# Robocopy の終了コードを解釈する
# 0 = no files copied; 1 = files copied; 2..7 = minor issues; >=8 = failure
if ($LASTEXITCODE -ge 8) {
    Write-Error "Robocopy failed with exit code $LASTEXITCODE. See log: $LogFile"
    exit $LASTEXITCODE
} else {
    Write-Output "Robocopy finished with exit code $LASTEXITCODE. See log: $LogFile"
    exit 0
}
