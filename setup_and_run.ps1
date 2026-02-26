param(
    [switch]$NoBrowser,
    [switch]$AutoInstallPython,
    [switch]$StopOldServer = $true
)

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot
$setupStatusPath = Join-Path $PSScriptRoot "output\setup-status.txt"

function Set-SetupStatus {
    param(
        [Parameter(Mandatory = $true)]
        [string]$StatusText
    )

    try {
        $statusDir = Split-Path $setupStatusPath -Parent
        if (-not (Test-Path $statusDir)) {
            New-Item -Path $statusDir -ItemType Directory -Force | Out-Null
        }
        Set-Content -Path $setupStatusPath -Value $StatusText -Encoding UTF8 -Force
    }
    catch {
        # Ignore setup status write failures.
    }
}

function Invoke-NativeNoThrow {
    param(
        [string]$Exe,
        [string[]]$CommandArgs,
        [switch]$Quiet
    )

    $oldErrorAction = $ErrorActionPreference
    $hasNativePref = $null -ne (Get-Variable -Name PSNativeCommandUseErrorActionPreference -ErrorAction SilentlyContinue)
    if ($hasNativePref) {
        $oldNativePref = $PSNativeCommandUseErrorActionPreference
        $PSNativeCommandUseErrorActionPreference = $false
    }

    $ErrorActionPreference = "Continue"
    try {
        if ($Quiet) {
            & $Exe @CommandArgs *> $null
        }
        else {
            & $Exe @CommandArgs | Out-Host
        }
        return [int]$LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $oldErrorAction
        if ($hasNativePref) {
            $PSNativeCommandUseErrorActionPreference = $oldNativePref
        }
    }
}

function Test-PythonCandidate {
    param(
        [string]$Exe,
        [string[]]$PrefixArgs
    )

    $isPathCandidate = ($Exe -match "[\\/]" -or $Exe -match "^[A-Za-z]:")
    if ($isPathCandidate) {
        if (-not (Test-Path $Exe)) {
            return $false
        }
    }
    else {
        if (-not (Get-Command $Exe -ErrorAction SilentlyContinue)) {
            return $false
        }
    }

    $exitCode = Invoke-NativeNoThrow -Exe $Exe -CommandArgs (@($PrefixArgs) + @("-c", "import sys; raise SystemExit(0 if sys.version_info >= (3, 11) else 1)")) -Quiet
    return ($exitCode -eq 0)
}

function Resolve-PythonCommand {
    $localPython311 = Join-Path $env:LocalAppData "Programs\Python\Python311\python.exe"
    $localPython312 = Join-Path $env:LocalAppData "Programs\Python\Python312\python.exe"
    $localPython313 = Join-Path $env:LocalAppData "Programs\Python\Python313\python.exe"
    $programFilesPython311 = Join-Path $env:ProgramFiles "Python311\python.exe"
    $programFilesPython312 = Join-Path $env:ProgramFiles "Python312\python.exe"
    $programFilesPython313 = Join-Path $env:ProgramFiles "Python313\python.exe"

    $candidates = @(
        @{ Exe = "python"; Prefix = @() },
        @{ Exe = "py"; Prefix = @("-3.11") },
        @{ Exe = "py"; Prefix = @("-3") },
        @{ Exe = $localPython311; Prefix = @() },
        @{ Exe = $localPython312; Prefix = @() },
        @{ Exe = $localPython313; Prefix = @() },
        @{ Exe = $programFilesPython311; Prefix = @() },
        @{ Exe = $programFilesPython312; Prefix = @() },
        @{ Exe = $programFilesPython313; Prefix = @() }
    )

    foreach ($candidate in $candidates) {
        if (Test-PythonCandidate -Exe $candidate.Exe -PrefixArgs $candidate.Prefix) {
            return $candidate
        }
    }

    return $null
}

function Ensure-Python {
    Set-SetupStatus -StatusText "Checking Python..."
    $pythonCmd = Resolve-PythonCommand
    if ($pythonCmd) {
        Set-SetupStatus -StatusText "Python is ready."
        return $pythonCmd
    }

    if ($AutoInstallPython) {
        if (Get-Command winget -ErrorAction SilentlyContinue) {
            Set-SetupStatus -StatusText "Installing Python..."
            Write-Host "Python 3.11+ not found. Installing via winget..."
            $exitCode = Invoke-NativeNoThrow -Exe "winget" -CommandArgs @(
                "install", "-e", "--id", "Python.Python.3.11",
                "--scope", "user",
                "--silent",
                "--disable-interactivity",
                "--accept-source-agreements",
                "--accept-package-agreements"
            )
            $pythonCmd = Resolve-PythonCommand
            if ($pythonCmd) {
                Set-SetupStatus -StatusText "Python is ready."
                return $pythonCmd
            }
        }

        Set-SetupStatus -StatusText "Installing Python (fallback)..."
        Write-Host "winget install did not provide a usable Python. Trying direct installer..."
        $installerPath = Join-Path $env:TEMP "python-3.11.9-amd64.exe"
        $downloadOk = $false
        try {
            Invoke-WebRequest -Uri "https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe" -OutFile $installerPath -UseBasicParsing
            $downloadOk = $true
        }
        catch {
            $downloadOk = $false
        }

        if ($downloadOk -and (Test-Path $installerPath)) {
            $installExit = Invoke-NativeNoThrow -Exe $installerPath -CommandArgs @(
                "/quiet",
                "InstallAllUsers=0",
                "PrependPath=1",
                "Include_test=0",
                "SimpleInstall=1",
                "Include_launcher=1"
            )

            $pythonCmd = Resolve-PythonCommand
            if ($installExit -eq 0 -and $pythonCmd) {
                Set-SetupStatus -StatusText "Python is ready."
                return $pythonCmd
            }
        }
    }

    $msg = @"
Python 3.11+ is required but was not found.

Install Python, then run run_web.bat again.

Option A (recommended):
- Download installer from: https://www.python.org/downloads/windows/
- During install, check 'Add python.exe to PATH'

Option B (winget):
- Open PowerShell as user and run:
  winget install -e --id Python.Python.3.11
"@
    throw $msg
}

function Stop-ExistingServerOnPort5000 {
    if (-not $StopOldServer) {
        return
    }

    Set-SetupStatus -StatusText "Checking old server process..."
    $pids = @()
    try {
        $connections = Get-NetTCPConnection -LocalPort 5000 -State Listen -ErrorAction SilentlyContinue
        if ($connections) {
            $pids = $connections | Select-Object -ExpandProperty OwningProcess -Unique
        }
    }
    catch {
        $pids = @()
    }

    foreach ($owningPid in $pids) {
        if (-not $owningPid) { continue }
        try {
            $proc = Get-Process -Id $owningPid -ErrorAction Stop
            Write-Host "Stopping old server process on port 5000: PID=$owningPid Name=$($proc.ProcessName)"
            Stop-Process -Id $owningPid -Force -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 350
        }
        catch {
            # Ignore if process already exited
        }
    }
}

function Ensure-VenvPip {
    param(
        [string]$VenvPython
    )

    function Clear-VenvPipArtifacts {
        param(
            [string]$TargetVenvPython
        )

        $venvScriptsDir = Split-Path $TargetVenvPython -Parent
        $venvDir = Split-Path $venvScriptsDir -Parent
        $sitePackages = Join-Path $venvDir "Lib\site-packages"
        if (-not (Test-Path $sitePackages)) {
            return
        }

        $items = Get-ChildItem $sitePackages -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like "pip*" }

        foreach ($item in $items) {
            try {
                Remove-Item $item.FullName -Recurse -Force -ErrorAction Stop
            }
            catch {
                Start-Sleep -Milliseconds 400
                Remove-Item $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    $pipCheckExit = Invoke-NativeNoThrow -Exe $VenvPython -CommandArgs @("-m", "pip", "--version") -Quiet
    if ($pipCheckExit -eq 0) {
        Set-SetupStatus -StatusText "pip is ready."
        return
    }

    Set-SetupStatus -StatusText "Repairing pip..."
    Write-Host "Pip not found in .venv. Repairing with ensurepip..."
    $ok = $false
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        Clear-VenvPipArtifacts -TargetVenvPython $VenvPython
        Start-Sleep -Seconds 1
        $ensureExit = Invoke-NativeNoThrow -Exe $VenvPython -CommandArgs @("-m", "ensurepip", "--upgrade", "--default-pip")
        if ($ensureExit -eq 0) {
            $ok = $true
            break
        }
    }

    if (-not $ok) {
        throw "Failed to prepare pip inside .venv."
    }
}

function Test-VenvState {
    param(
        [string]$VenvPath,
        [string]$VenvPython
    )

    if (-not (Test-Path $VenvPython)) {
        return $false
    }

    $cfgPath = Join-Path $VenvPath "pyvenv.cfg"
    if (-not (Test-Path $cfgPath)) {
        return $false
    }

    $cfgText = ""
    try {
        $cfgText = Get-Content $cfgPath -Raw -ErrorAction Stop
    }
    catch {
        return $false
    }

    if ([string]::IsNullOrWhiteSpace($cfgText)) {
        return $false
    }

    if ($cfgText -notmatch "(?im)^\s*home\s*=") {
        return $false
    }

    $probeExit = Invoke-NativeNoThrow -Exe $VenvPython -CommandArgs @(
        "-c", "import sys; raise SystemExit(0 if getattr(sys, 'prefix', '') else 1)"
    ) -Quiet
    return ($probeExit -eq 0)
}

function Remove-DirectoryRobust {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathToRemove
    )

    if (-not (Test-Path $PathToRemove)) {
        return
    }

    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            Remove-Item $PathToRemove -Recurse -Force -ErrorAction Stop
        }
        catch {
            Start-Sleep -Milliseconds (300 * $attempt)
            $null = Invoke-NativeNoThrow -Exe "cmd" -CommandArgs @("/c", "attrib", "-r", "$PathToRemove\*", "/s", "/d") -Quiet
            $quotedPath = '"' + $PathToRemove + '"'
            $null = Invoke-NativeNoThrow -Exe "cmd" -CommandArgs @("/c", "rmdir", "/s", "/q", $quotedPath) -Quiet
        }

        if (-not (Test-Path $PathToRemove)) {
            return
        }
    }

    throw "Failed to remove $PathToRemove. Please close apps that may lock this folder and run again."
}

$pythonCommand = Ensure-Python
$pythonExe = $pythonCommand.Exe
$pythonPrefix = $pythonCommand.Prefix

Stop-ExistingServerOnPort5000

$venvPath = Join-Path $PSScriptRoot ".venv"
$venvFallbackPath = Join-Path $PSScriptRoot ".venv_repair"
$venvPython = Join-Path $venvPath "Scripts\python.exe"

if (Test-Path $venvPath) {
    Set-SetupStatus -StatusText "Validating virtual environment..."
    $venvOk = Test-VenvState -VenvPath $venvPath -VenvPython $venvPython
    if (-not $venvOk) {
        Set-SetupStatus -StatusText "Repairing virtual environment..."
        Write-Host "Found broken/incomplete .venv. Recreating virtual environment..."
        try {
            Remove-DirectoryRobust -PathToRemove $venvPath
        }
        catch {
            Write-Warning $_.Exception.Message
            Write-Warning "Using fallback virtual environment '.venv_repair' for this run."
            $venvPath = $venvFallbackPath
            $venvPython = Join-Path $venvPath "Scripts\python.exe"

            if (Test-Path $venvPath) {
                $fallbackOk = Test-VenvState -VenvPath $venvPath -VenvPython $venvPython
                if (-not $fallbackOk) {
                    Remove-DirectoryRobust -PathToRemove $venvPath
                }
            }
        }
    }
}

if (-not (Test-Path $venvPath)) {
    Set-SetupStatus -StatusText "Creating virtual environment..."
    Write-Host "Creating virtual environment ($venvPath)..."
    $venvExit = Invoke-NativeNoThrow -Exe $pythonExe -CommandArgs (@($pythonPrefix) + @("-m", "venv", $venvPath))
    if ($venvExit -ne 0) {
        if (Test-Path $venvPython) {
            Write-Host "venv returned an error. Trying to repair..."
            Ensure-VenvPip -VenvPython $venvPython
        }
        else {
            throw "Failed to create virtual environment."
        }
    }
}

if (-not (Test-Path $venvPython)) {
    throw "Cannot find $venvPython after venv creation."
}
Set-SetupStatus -StatusText "Preparing pip..."
Ensure-VenvPip -VenvPython $venvPython

function Repair-VenvPipState {
    param(
        [string]$VenvPython
    )

    Write-Host "Attempting pip self-repair..."
    $repairPipExit = Invoke-NativeNoThrow -Exe $VenvPython -CommandArgs @(
        "-m", "pip", "install", "--force-reinstall", "--no-deps", "pip==24.0"
    )
    if ($repairPipExit -eq 0) {
        return $true
    }

    Write-Host "Force-reinstall failed, trying ensurepip..."
    $ensureExit = Invoke-NativeNoThrow -Exe $VenvPython -CommandArgs @(
        "-m", "ensurepip", "--upgrade", "--default-pip"
    )
    return ($ensureExit -eq 0)
}

function Start-BrowserWhenServerReady {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BaseUrl
    )

    Start-Job -ScriptBlock {
        param(
            [string]$TargetUrl
        )

        $healthUrl = "$TargetUrl/health"
        $opened = $false
        for ($attempt = 1; $attempt -le 100; $attempt++) {
            try {
                $response = Invoke-WebRequest -Uri $healthUrl -UseBasicParsing -TimeoutSec 2
                if ($response.StatusCode -ge 200 -and $response.StatusCode -lt 400) {
                    Start-Process $TargetUrl
                    $opened = $true
                    break
                }
            }
            catch {
                # Keep waiting until the Flask app is ready.
            }

            Start-Sleep -Milliseconds 350
        }

        if (-not $opened) {
            Start-Process $TargetUrl
        }
    } -ArgumentList $BaseUrl | Out-Null
}

function Get-RequirementPackageNames {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RequirementsPath
    )

    if (-not (Test-Path $RequirementsPath)) {
        return @()
    }

    $names = New-Object System.Collections.Generic.List[string]
    $lines = Get-Content -Path $RequirementsPath -ErrorAction SilentlyContinue
    foreach ($line in $lines) {
        $trimmed = [string]$line
        if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
        $trimmed = $trimmed.Trim()
        if ($trimmed.StartsWith("#")) { continue }

        $hashIndex = $trimmed.IndexOf("#")
        if ($hashIndex -ge 0) {
            $trimmed = $trimmed.Substring(0, $hashIndex).Trim()
        }
        if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
        if ($trimmed.StartsWith("-")) { continue }

        $match = [regex]::Match($trimmed, "^[A-Za-z0-9_.-]+")
        if (-not $match.Success) { continue }

        $name = $match.Value
        if (-not $names.Contains($name)) {
            [void]$names.Add($name)
        }
    }

    return @($names.ToArray())
}

function Get-RequirementsHash {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RequirementsPath
    )

    if (-not (Test-Path $RequirementsPath)) {
        return ""
    }

    try {
        return [string](Get-FileHash -Path $RequirementsPath -Algorithm SHA256 -ErrorAction Stop).Hash
    }
    catch {
        return ""
    }
}

function Test-DependenciesInstalled {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VenvPython,
        [string[]]$PackageNames = @()
    )

    foreach ($package in @($PackageNames)) {
        if ([string]::IsNullOrWhiteSpace($package)) { continue }
        $showExit = Invoke-NativeNoThrow -Exe $VenvPython -CommandArgs @("-m", "pip", "show", $package) -Quiet
        if ($showExit -ne 0) {
            return $false
        }
    }

    return $true
}

function Ensure-Dependencies {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VenvPython,
        [Parameter(Mandatory = $true)]
        [string]$RequirementsPath,
        [Parameter(Mandatory = $true)]
        [string]$StampPath
    )

    if (-not (Test-Path $RequirementsPath)) {
        throw "requirements.txt not found at $RequirementsPath"
    }

    $requirementsHash = Get-RequirementsHash -RequirementsPath $RequirementsPath
    $savedHash = ""
    if (Test-Path $StampPath) {
        try {
            $savedHash = ((Get-Content -Path $StampPath -Raw -ErrorAction Stop) | Out-String).Trim()
        }
        catch {
            $savedHash = ""
        }
    }

    $packageNames = Get-RequirementPackageNames -RequirementsPath $RequirementsPath
    $dependenciesOk = Test-DependenciesInstalled -VenvPython $VenvPython -PackageNames $packageNames

    if ($dependenciesOk -and -not [string]::IsNullOrWhiteSpace($requirementsHash) -and $savedHash -eq $requirementsHash) {
        Set-SetupStatus -StatusText "Dependencies are ready."
        Write-Host "Dependencies are up to date. Skipping install."
        return
    }

    Set-SetupStatus -StatusText "Installing dependencies..."
    Write-Host "Installing dependencies..."
    $pipInstallExit = Invoke-NativeNoThrow -Exe $VenvPython -CommandArgs @("-m", "pip", "install", "--disable-pip-version-check", "-r", $RequirementsPath)
    if ($pipInstallExit -ne 0) {
        Set-SetupStatus -StatusText "Retrying dependency install..."
        Write-Host "Dependency install failed. Trying pip self-repair and retrying once..."
        $repaired = Repair-VenvPipState -VenvPython $VenvPython
        if ($repaired) {
            $pipInstallExit = Invoke-NativeNoThrow -Exe $VenvPython -CommandArgs @("-m", "pip", "install", "--disable-pip-version-check", "-r", $RequirementsPath)
        }
    }

    if ($pipInstallExit -ne 0) {
        throw "Failed to install dependencies from requirements.txt."
    }

    if (-not [string]::IsNullOrWhiteSpace($requirementsHash)) {
        try {
            Set-Content -Path $StampPath -Value $requirementsHash -Encoding ASCII -NoNewline
        }
        catch {
            # Ignore stamp write failures.
        }
    }

    Set-SetupStatus -StatusText "Dependencies are ready."
}

$requirementsPath = Join-Path $PSScriptRoot "requirements.txt"
$requirementsStampPath = Join-Path $venvPath ".requirements.sha256"
Ensure-Dependencies -VenvPython $venvPython -RequirementsPath $requirementsPath -StampPath $requirementsStampPath

if (-not $NoBrowser) {
    Start-BrowserWhenServerReady -BaseUrl "http://127.0.0.1:5000"
}

Set-SetupStatus -StatusText "Starting web server..."
Write-Host "Starting web app on http://127.0.0.1:5000"
& $venvPython app.py
