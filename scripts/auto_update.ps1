<#
.SYNOPSIS
Safely auto-updates the local repository from a configured git release branch before ETL runs.

.DESCRIPTION
Opt-in via config.local.yaml:
  auto_update:
    enabled: true
    branch: "release"
    remote: "origin"
    fetch_timeout_seconds: 20

Exit codes:
  0  = No update needed, update succeeded (batch unchanged), or update skipped/offline safely.
  10 = Update succeeded AND scripts\run_daily_etl.bat changed (caller should re-exec once).
#>
param(
    [string]$ProjectDir = "",
    [string]$PythonExe = "",
    [switch]$ForceCheck
)

$ErrorActionPreference = "Stop"

if (-not $ProjectDir) {
    $ProjectDir = Split-Path -Parent $PSScriptRoot
}
$ProjectDir = (Resolve-Path -LiteralPath $ProjectDir).Path
if (-not $PythonExe) {
    $PythonExe = Join-Path $ProjectDir ".venv\Scripts\python.exe"
}

function Write-UpdateLog([string]$Message) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Output "[$ts] [auto_update] $Message"
}

function Get-FileSha256([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return "" }
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Read-AutoUpdateSettings([string]$RootDir) {
    $settings = @{
        enabled = $false
        branch  = "release"
        remote  = "origin"
        timeout = 20
    }
    $localCfg = Join-Path $RootDir "config.local.yaml"
    if (-not (Test-Path -LiteralPath $localCfg)) {
        return $settings
    }

    # Prefer PyYAML via project venv if available; fall back to lightweight parser.
    if (Test-Path -LiteralPath $PythonExe) {
        try {
            $pyCode = @"
import json, yaml, sys
data = yaml.safe_load(open(sys.argv[1], encoding='utf-8')) or {}
au = data.get('auto_update') or {}
print(json.dumps({
    'enabled': bool(au.get('enabled', False)),
    'branch': str(au.get('branch', 'release')),
    'remote': str(au.get('remote', 'origin')),
    'timeout': int(au.get('fetch_timeout_seconds', 20)),
}))
"@
            $jsonOut = & $PythonExe -c $pyCode $localCfg 2>$null
            if ($LASTEXITCODE -eq 0 -and $jsonOut) {
                $parsed = $jsonOut | ConvertFrom-Json
                $settings.enabled = [bool]$parsed.enabled
                $settings.branch  = [string]$parsed.branch
                $settings.remote  = [string]$parsed.remote
                $settings.timeout = [int]$parsed.timeout
                return $settings
            }
        } catch {
            # Fall through to regex parser
        }
    }

    $inSection = $false
    foreach ($line in Get-Content -LiteralPath $localCfg) {
        if ($line -match '^\s*#') { continue }
        if ($line -match '^auto_update\s*:') {
            $inSection = $true
            continue
        }
        if ($inSection -and $line -match '^\S') {
            break
        }
        if ($inSection) {
            if ($line -match '^\s+enabled\s*:\s*(true|false)\s*$') {
                $settings.enabled = ($Matches[1].ToLowerInvariant() -eq "true")
            } elseif ($line -match '^\s+branch\s*:\s*["'']?([^"''\s#]+)["'']?') {
                $settings.branch = $Matches[1]
            } elseif ($line -match '^\s+remote\s*:\s*["'']?([^"''\s#]+)["'']?') {
                $settings.remote = $Matches[1]
            } elseif ($line -match '^\s+fetch_timeout_seconds\s*:\s*(\d+)') {
                $settings.timeout = [int]$Matches[1]
            }
        }
    }
    return $settings
}

function Invoke-GitWithTimeout([string]$WorkDir, [string[]]$GitArgs, [int]$TimeoutSeconds = 20) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "git"
    $psi.WorkingDirectory = $WorkDir
    $psi.Arguments = ($GitArgs | ForEach-Object {
        if ($_ -match '\s') { '"' + ($_ -replace '"', '\"') + '"' } else { $_ }
    }) -join " "
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()
    $stdoutTask = $proc.StandardOutput.ReadToEndAsync()
    $stderrTask = $proc.StandardError.ReadToEndAsync()

    if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
        try { $proc.Kill() } catch {}
        return @{ ExitCode = 124; Stdout = ""; Stderr = "Timed out after ${TimeoutSeconds}s" }
    }
    return @{
        ExitCode = $proc.ExitCode
        Stdout   = $stdoutTask.Result.Trim()
        Stderr   = $stderrTask.Result.Trim()
    }
}

try {
    $cfg = Read-AutoUpdateSettings -RootDir $ProjectDir
    if (-not $cfg.enabled -and -not $ForceCheck) {
        exit 0
    }

    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
        Write-UpdateLog "git is not installed on PATH; skipping auto-update."
        exit 0
    }
    if (-not (Test-Path -LiteralPath (Join-Path $ProjectDir ".git"))) {
        Write-UpdateLog "Project folder is not a git repository; skipping auto-update."
        exit 0
    }

    # Refuse to overwrite uncommitted changes to tracked files (protects developer workstations).
    $dirty = Invoke-GitWithTimeout -WorkDir $ProjectDir -GitArgs @("status", "--porcelain", "--untracked-files=no") -TimeoutSeconds 10
    if ($dirty.ExitCode -ne 0) {
        Write-UpdateLog "Could not inspect git status ($($dirty.Stderr)); skipping auto-update."
        exit 0
    }
    if ($dirty.Stdout) {
        Write-UpdateLog "Tracked files have local uncommitted edits; skipping auto-update to avoid data loss."
        exit 0
    }

    $remote = $cfg.remote
    $branch = $cfg.branch
    $timeout = [Math]::Max(5, $cfg.timeout)

    $fetch = Invoke-GitWithTimeout -WorkDir $ProjectDir -GitArgs @("fetch", "--prune", $remote, $branch) -TimeoutSeconds $timeout
    if ($fetch.ExitCode -ne 0) {
        Write-UpdateLog "Could not fetch $remote/$branch ($($fetch.Stderr)); continuing with local version."
        exit 0
    }

    $headRes = Invoke-GitWithTimeout -WorkDir $ProjectDir -GitArgs @("rev-parse", "HEAD") -TimeoutSeconds 10
    $targetRes = Invoke-GitWithTimeout -WorkDir $ProjectDir -GitArgs @("rev-parse", "FETCH_HEAD") -TimeoutSeconds 10
    if ($headRes.ExitCode -ne 0 -or $targetRes.ExitCode -ne 0) {
        Write-UpdateLog "Could not resolve git revisions; continuing with local version."
        exit 0
    }

    $prevSha = $headRes.Stdout
    $targetSha = $targetRes.Stdout
    if ($prevSha -eq $targetSha) {
        Write-UpdateLog "Already up to date ($($prevSha.Substring(0, [Math]::Min(7, $prevSha.Length))))."
        exit 0
    }

    $batPath = Join-Path $ProjectDir "scripts\run_daily_etl.bat"
    $reqPath = Join-Path $ProjectDir "requirements.txt"
    $reqStampPath = Join-Path $ProjectDir ".venv\.requirements.sha256"
    $prevBatHash = Get-FileSha256 $batPath
    $prevPlaywrightLine = ""
    if (Test-Path -LiteralPath $reqPath) {
        $prevPlaywrightLine = (Get-Content -LiteralPath $reqPath | Where-Object { $_ -match '^\s*playwright\b' } | Select-Object -First 1)
    }

    Write-UpdateLog "Updating from $($prevSha.Substring(0, 7)) to $($targetSha.Substring(0, 7)) ($remote/$branch)..."
    $reset = Invoke-GitWithTimeout -WorkDir $ProjectDir -GitArgs @("reset", "--hard", "FETCH_HEAD") -TimeoutSeconds 20
    if ($reset.ExitCode -ne 0) {
        Write-UpdateLog "git reset failed ($($reset.Stderr)); continuing with existing copy."
        exit 0
    }

    # Verify core Python entrypoints compile cleanly; auto-rollback if broken.
    if (Test-Path -LiteralPath $PythonExe) {
        $coreFiles = @(
            @(
                "main.py",
                "pos_pipeline.py",
                "scripts\refresh_backoffice_session.py",
                "extractors\playwright_extractor.py"
            ) | ForEach-Object { Join-Path $ProjectDir $_ } | Where-Object { Test-Path -LiteralPath $_ }
        )

        if ($coreFiles.Count -gt 0) {
            $prevEap = $ErrorActionPreference
            $ErrorActionPreference = "Continue"
            try {
                & $PythonExe -m py_compile @coreFiles 2>&1 | Out-Null
                $compileExit = $LASTEXITCODE
            } finally {
                $ErrorActionPreference = $prevEap
            }
            if ($compileExit -ne 0) {
                Write-UpdateLog "ERROR: Updated code failed Python syntax check! Rolling back to $($prevSha.Substring(0, 7))."
                [void](Invoke-GitWithTimeout -WorkDir $ProjectDir -GitArgs @("reset", "--hard", $prevSha) -TimeoutSeconds 20)
                exit 0
            }
        }

        # Sync Python dependencies only if requirements.txt changed.
        if (Test-Path -LiteralPath $reqPath) {
            $newReqHash = Get-FileSha256 $reqPath
            $stampedHash = if (Test-Path -LiteralPath $reqStampPath) {
                (Get-Content -LiteralPath $reqStampPath -Raw).Trim().ToLowerInvariant()
            } else { "" }

            if ($newReqHash -ne $stampedHash) {
                Write-UpdateLog "requirements.txt changed; syncing virtual environment packages..."
                $prevEap = $ErrorActionPreference
                $ErrorActionPreference = "Continue"
                try {
                    & $PythonExe -m pip install -r $reqPath 2>&1 | ForEach-Object { Write-Output $_ }
                    $pipExit = $LASTEXITCODE
                } finally {
                    $ErrorActionPreference = $prevEap
                }
                if ($pipExit -eq 0) {
                    $stampDir = Split-Path -Parent $reqStampPath
                    if (Test-Path -LiteralPath $stampDir) {
                        Set-Content -LiteralPath $reqStampPath -Value $newReqHash -Encoding ASCII
                    }
                    $newPlaywrightLine = (Get-Content -LiteralPath $reqPath | Where-Object { $_ -match '^\s*playwright\b' } | Select-Object -First 1)
                    if ($newPlaywrightLine -and ($newPlaywrightLine -ne $prevPlaywrightLine)) {
                        Write-UpdateLog "Playwright requirement changed; ensuring Chromium browser binary is installed..."
                        $env:NODE_USE_SYSTEM_CA = "1"
                        $env:PLAYWRIGHT_DOWNLOAD_CONNECTION_TIMEOUT = "120000"
                        $prevEap = $ErrorActionPreference
                        $ErrorActionPreference = "Continue"
                        try {
                            & $PythonExe -m playwright install chromium 2>&1 | ForEach-Object { Write-Output $_ }
                        } finally {
                            $ErrorActionPreference = $prevEap
                        }
                    }
                } else {
                    Write-UpdateLog "Warning: pip install failed (exit $pipExit); continuing with existing packages."
                }
            }
        }
    }

    Write-UpdateLog "Update complete ($($targetSha.Substring(0, 7)))."
    $newBatHash = Get-FileSha256 $batPath
    if ($prevBatHash -and $newBatHash -and ($prevBatHash -ne $newBatHash)) {
        Write-UpdateLog "run_daily_etl.bat changed; signaling caller to re-launch updated script."
        exit 10
    }
    exit 0
} catch {
    Write-UpdateLog "Warning: auto-update encountered an unexpected error ($($_.Exception.Message)); continuing with local version."
    exit 0
}
