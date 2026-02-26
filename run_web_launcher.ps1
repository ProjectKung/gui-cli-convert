param(
    [switch]$AutoInstallPython = $true
)

$ErrorActionPreference = "SilentlyContinue"
Set-Location -Path $PSScriptRoot

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class NativeMethods {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern bool DestroyIcon(IntPtr handle);
}
"@
[System.Windows.Forms.Application]::EnableVisualStyles()

if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne [System.Threading.ApartmentState]::STA) {
    $restartArgs = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-Sta", "-WindowStyle", "Hidden", "-File", "`"$PSCommandPath`"")
    if ($AutoInstallPython) {
        $restartArgs += "-AutoInstallPython"
    }
    Start-Process -FilePath "powershell" -ArgumentList $restartArgs -WindowStyle Hidden | Out-Null
    return
}

$baseUrl = "http://127.0.0.1:5000"
$healthUrl = "$baseUrl/health"
$launcherLogPath = Join-Path $PSScriptRoot "output\launcher-error.log"
$setupStatusPath = Join-Path $PSScriptRoot "output\setup-status.txt"

function Write-LauncherLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    try {
        $logDir = Split-Path $launcherLogPath -Parent
        if (-not (Test-Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        Add-Content -Path $launcherLogPath -Value "[$timestamp] $Message"
    }
    catch {
        # Ignore logging failures.
    }
}

function Ensure-ConvertShortcut {
    try {
        $shortcutPath = Join-Path $PSScriptRoot "Convert CLI-GUI.lnk"
        $vbsPath = Join-Path $PSScriptRoot "run_web.vbs"
        $iconPath = Join-Path $PSScriptRoot "static\convert_cli_gui_v2.ico"
        if (-not (Test-Path $shortcutPath) -or -not (Test-Path $vbsPath)) {
            return
        }

        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($shortcutPath)

        $expectedTarget = Join-Path $env:WINDIR "System32\wscript.exe"
        $expectedArguments = '"' + $vbsPath + '"'
        $expectedWorkingDir = $PSScriptRoot
        $expectedIcon = if (Test-Path $iconPath) { $iconPath + ",0" } else { $null }

        $changed = $false
        if ($shortcut.TargetPath -ne $expectedTarget) {
            $shortcut.TargetPath = $expectedTarget
            $changed = $true
        }
        if ($shortcut.Arguments -ne $expectedArguments) {
            $shortcut.Arguments = $expectedArguments
            $changed = $true
        }
        if ($shortcut.WorkingDirectory -ne $expectedWorkingDir) {
            $shortcut.WorkingDirectory = $expectedWorkingDir
            $changed = $true
        }
        if ($expectedIcon -and $shortcut.IconLocation -ne $expectedIcon) {
            $shortcut.IconLocation = $expectedIcon
            $changed = $true
        }

        if ($changed) {
            $shortcut.Save()
            Write-LauncherLog -Message "Shortcut repaired: Convert CLI-GUI.lnk"
        }
    }
    catch {
        Write-LauncherLog -Message ("Shortcut repair failed: {0}" -f $_.Exception.Message)
    }
}

function Get-SetupStatusText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$StatusPath
    )

    if ([string]::IsNullOrWhiteSpace($StatusPath)) {
        return ""
    }

    try {
        if (-not (Test-Path $StatusPath)) {
            return ""
        }
        return [string]((Get-Content -Path $StatusPath -Raw -ErrorAction Stop) | Out-String).Trim()
    }
    catch {
        return ""
    }
}

function New-RoundedRectanglePath {
    param(
        [int]$X,
        [int]$Y,
        [int]$Width,
        [int]$Height,
        [int]$Radius
    )

    $safeRadius = [Math]::Max(1, [Math]::Min([Math]::Min($Width, $Height) / 2, $Radius))
    $diameter = [int]($safeRadius * 2)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc($X, $Y, $diameter, $diameter, 180, 90)
    $path.AddArc($X + $Width - $diameter, $Y, $diameter, $diameter, 270, 90)
    $path.AddArc($X + $Width - $diameter, $Y + $Height - $diameter, $diameter, $diameter, 0, 90)
    $path.AddArc($X, $Y + $Height - $diameter, $diameter, $diameter, 90, 90)
    $path.CloseFigure()
    return $path
}

function New-BrandIcon {
    $bitmap = New-Object System.Drawing.Bitmap(64, 64, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $graphics.Clear([System.Drawing.Color]::Transparent)

    $bgPath = New-RoundedRectanglePath -X 4 -Y 4 -Width 56 -Height 56 -Radius 16
    $bgBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.Rectangle(4, 4, 56, 56)),
        [System.Drawing.Color]::FromArgb(12, 110, 214),
        [System.Drawing.Color]::FromArgb(30, 162, 255),
        45.0
    )
    $graphics.FillPath($bgBrush, $bgPath)

    $docPath = New-RoundedRectanglePath -X 17 -Y 13 -Width 26 -Height 34 -Radius 6
    $docBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(250, 253, 255))
    $graphics.FillPath($docBrush, $docPath)

    $foldBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(210, 233, 255))
    $graphics.FillPolygon($foldBrush, [System.Drawing.Point[]]@(
        (New-Object System.Drawing.Point(34, 13)),
        (New-Object System.Drawing.Point(43, 22)),
        (New-Object System.Drawing.Point(43, 13))
    ))

    $linePen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(182, 210, 241), 2.6)
    $linePen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $linePen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $graphics.DrawLine($linePen, 22, 27, 38, 27)
    $graphics.DrawLine($linePen, 22, 33, 38, 33)
    $graphics.DrawLine($linePen, 22, 39, 33, 39)

    $badgeBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.Rectangle(34, 32, 24, 24)),
        [System.Drawing.Color]::FromArgb(31, 146, 84),
        [System.Drawing.Color]::FromArgb(40, 182, 99),
        45.0
    )
    $graphics.FillEllipse($badgeBrush, (New-Object System.Drawing.Rectangle(34, 32, 24, 24)))

    $checkPen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 4)
    $checkPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $checkPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $checkPen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
    $graphics.DrawLines($checkPen, [System.Drawing.Point[]]@(
        (New-Object System.Drawing.Point(40, 44)),
        (New-Object System.Drawing.Point(45, 49)),
        (New-Object System.Drawing.Point(53, 40))
    ))

    $hIcon = $bitmap.GetHicon()
    $icon = [System.Drawing.Icon]::FromHandle($hIcon).Clone()
    [void][NativeMethods]::DestroyIcon($hIcon)

    $checkPen.Dispose()
    $badgeBrush.Dispose()
    $linePen.Dispose()
    $foldBrush.Dispose()
    $docBrush.Dispose()
    $docPath.Dispose()
    $bgBrush.Dispose()
    $bgPath.Dispose()
    $graphics.Dispose()
    $bitmap.Dispose()

    return $icon
}

Ensure-ConvertShortcut
$appIcon = New-BrandIcon

function Test-ServerReady {
    param(
        [Parameter(Mandatory = $true)]
        [string]$HealthCheckUrl
    )

    try {
        $response = Invoke-WebRequest -Uri $HealthCheckUrl -UseBasicParsing -TimeoutSec 1
        return ($response.StatusCode -ge 200 -and $response.StatusCode -lt 400)
    }
    catch {
        return $false
    }
}

function Stop-WebServer {
    param(
        [int]$BootstrapPid = 0,
        [Parameter(Mandatory = $true)]
        [string]$HealthCheckUrl
    )

    $candidatePids = New-Object "System.Collections.Generic.HashSet[int]"
    if ($BootstrapPid -gt 0) {
        [void]$candidatePids.Add($BootstrapPid)
    }

    try {
        $serverPids = Get-NetTCPConnection -LocalPort 5000 -State Listen -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty OwningProcess -Unique
        foreach ($serverPid in $serverPids) {
            if ($serverPid) {
                [void]$candidatePids.Add([int]$serverPid)
            }
        }
    }
    catch {
        # Ignore and try fallback method below.
    }

    if ($candidatePids.Count -eq 0) {
        try {
            $netstatOutput = & netstat -ano -p tcp 2>$null
            foreach ($line in $netstatOutput) {
                $match = [regex]::Match($line, "^\s*TCP\s+\S+:5000\s+\S+\s+LISTENING\s+(\d+)\s*$")
                if ($match.Success) {
                    [void]$candidatePids.Add([int]$match.Groups[1].Value)
                }
            }
        }
        catch {
            # Ignore fallback parsing failures.
        }
    }

    try {
        $venvPythonPath = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
        $venvPythonPids = Get-Process -Name python -ErrorAction SilentlyContinue |
            Where-Object { $_.Path -and $_.Path -ieq $venvPythonPath } |
            Select-Object -ExpandProperty Id
        foreach ($venvPid in $venvPythonPids) {
            if ($venvPid) {
                [void]$candidatePids.Add([int]$venvPid)
            }
        }
    }
    catch {
        # Ignore process enumeration failures.
    }

    for ($attempt = 1; $attempt -le 4; $attempt++) {
        foreach ($candidatePid in @($candidatePids)) {
            if ($candidatePid -and $candidatePid -ne $PID) {
                try {
                    Stop-Process -Id $candidatePid -Force -ErrorAction SilentlyContinue
                }
                catch {
                    # Process may already be stopped.
                }
            }
        }

        Start-Sleep -Milliseconds (220 * $attempt)
        if (-not (Test-ServerReady -HealthCheckUrl $HealthCheckUrl)) {
            return $true
        }

        try {
            $retryPids = Get-NetTCPConnection -LocalPort 5000 -State Listen -ErrorAction SilentlyContinue |
                Select-Object -ExpandProperty OwningProcess -Unique
            foreach ($retryPid in $retryPids) {
                if ($retryPid) {
                    [void]$candidatePids.Add([int]$retryPid)
                }
            }
        }
        catch {
            # Ignore failures while retrying.
        }
    }

    return -not (Test-ServerReady -HealthCheckUrl $HealthCheckUrl)
}

function Set-BottomRightLocation {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Form]$Window
    )

    $activeScreen = [System.Windows.Forms.Screen]::FromPoint([System.Windows.Forms.Cursor]::Position)
    $workArea = $activeScreen.WorkingArea
    $margin = 12
    $Window.Location = New-Object System.Drawing.Point(
        ($workArea.Right - $Window.Width - $margin),
        ($workArea.Bottom - $Window.Height - $margin)
    )
}

function Resolve-ManagedBrowserExecutable {
    $candidates = @(
        "msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "$env:LocalAppData\Microsoft\Edge\Application\msedge.exe",
        "chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "$env:LocalAppData\Google\Chrome\Application\chrome.exe"
    )

    foreach ($candidate in $candidates) {
        if (-not $candidate) { continue }
        if ($candidate -match "^[A-Za-z]:\\") {
            if (Test-Path $candidate) {
                return $candidate
            }
            continue
        }

        $cmd = Get-Command $candidate -ErrorAction SilentlyContinue
        if ($cmd) {
            return $cmd.Source
        }
    }

    return $null
}

function Start-ManagedBrowserWindow {
    param(
        [string]$BrowserExecutable,
        [Parameter(Mandatory = $true)]
        [string]$Url,
        [Parameter(Mandatory = $true)]
        [string]$ProfileDir,
        [int]$DebugPort = 0,
        [System.Collections.ArrayList]$TrackedProcessIds
    )

    if ([string]::IsNullOrWhiteSpace($BrowserExecutable)) {
        Write-LauncherLog -Message "Start-ManagedBrowserWindow: browser executable not found."
        return $false
    }

    try {
        if (-not (Test-Path $ProfileDir)) {
            New-Item -ItemType Directory -Path $ProfileDir -Force | Out-Null
        }

        $args = @(
            "--user-data-dir=$ProfileDir",
            "--remote-debugging-port=$DebugPort",
            "--no-first-run",
            "--disable-session-crashed-bubble",
            "--hide-crash-restore-bubble",
            "--new-window",
            $Url
        )
        $proc = Start-Process -FilePath $BrowserExecutable -ArgumentList $args -PassThru
        if ($proc -and $TrackedProcessIds) {
            [void]$TrackedProcessIds.Add($proc.Id)
        }
        return $true
    }
    catch {
        Write-LauncherLog -Message ("Start-ManagedBrowserWindow error: {0}" -f $_.Exception.Message)
        return $false
    }
}

function Resolve-ManagedDebugPort {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProfileDir,
        [int]$FallbackPort = 0
    )

    if ([string]::IsNullOrWhiteSpace($ProfileDir)) {
        return $FallbackPort
    }

    $activePortFile = Join-Path $ProfileDir "DevToolsActivePort"
    for ($attempt = 1; $attempt -le 25; $attempt++) {
        try {
            if (Test-Path $activePortFile) {
                $lines = Get-Content -Path $activePortFile -ErrorAction Stop
                if ($lines -and $lines.Count -gt 0) {
                    $portText = [string]$lines[0]
                    $portNumber = 0
                    if ([int]::TryParse($portText, [ref]$portNumber) -and $portNumber -gt 0) {
                        return $portNumber
                    }
                }
            }
        }
        catch {
            # Keep retrying until timeout.
        }

        Start-Sleep -Milliseconds 120
    }

    return $FallbackPort
}

function Reset-ManagedBrowserProfile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProfileDir
    )

    if ([string]::IsNullOrWhiteSpace($ProfileDir)) {
        return $false
    }

    if (-not (Test-Path $ProfileDir)) {
        return $true
    }

    for ($attempt = 1; $attempt -le 4; $attempt++) {
        try {
            if (Test-Path $ProfileDir) {
                Remove-Item -Path $ProfileDir -Recurse -Force -ErrorAction Stop
            }
            return $true
        }
        catch {
            Start-Sleep -Milliseconds (260 * $attempt)
        }
    }

    return (-not (Test-Path $ProfileDir))
}

function Show-LoadingWindow {
    param(
        [Parameter(Mandatory = $true)]
        [System.Diagnostics.Process]$SetupProcess,
        [Parameter(Mandatory = $true)]
        [string]$HealthCheckUrl,
        [Parameter(Mandatory = $true)]
        [System.Drawing.Icon]$AppIcon
    )

    $launchResult = [PSCustomObject]@{
        Ready  = $false
        Failed = $false
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Switch Converter"
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.ControlBox = $false
    $form.TopMost = $true
    $form.ClientSize = New-Object System.Drawing.Size(420, 170)
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 250, 255)
    $form.Icon = $AppIcon

    $title = New-Object System.Windows.Forms.Label
    $title.Text = "Switch Converter"
    $title.AutoSize = $false
    $title.Size = New-Object System.Drawing.Size(380, 34)
    $title.Location = New-Object System.Drawing.Point(20, 18)
    $title.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $title.ForeColor = [System.Drawing.Color]::FromArgb(21, 54, 87)
    $form.Controls.Add($title)

    $status = New-Object System.Windows.Forms.Label
    $status.Text = "Starting web service..."
    $status.AutoSize = $false
    $status.Size = New-Object System.Drawing.Size(380, 24)
    $status.Location = New-Object System.Drawing.Point(20, 62)
    $status.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
    $status.ForeColor = [System.Drawing.Color]::FromArgb(56, 89, 122)
    $form.Controls.Add($status)

    $progress = New-Object System.Windows.Forms.ProgressBar
    $progress.Style = "Marquee"
    $progress.MarqueeAnimationSpeed = 35
    $progress.Size = New-Object System.Drawing.Size(380, 16)
    $progress.Location = New-Object System.Drawing.Point(20, 96)
    $form.Controls.Add($progress)

    $hint = New-Object System.Windows.Forms.Label
    $hint.Text = "Please wait. The browser will open automatically."
    $hint.AutoSize = $false
    $hint.Size = New-Object System.Drawing.Size(380, 24)
    $hint.Location = New-Object System.Drawing.Point(20, 124)
    $hint.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $hint.ForeColor = [System.Drawing.Color]::FromArgb(80, 107, 133)
    $form.Controls.Add($hint)

    $spinnerFrames = @("|", "/", "-", "\")
    $spinnerIndex = 0
    $startTime = Get-Date

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 450
    $timer.Add_Tick({
        $elapsed = [int]((Get-Date) - $startTime).TotalSeconds
        $frame = $spinnerFrames[$spinnerIndex % $spinnerFrames.Length]
        $status.Text = "Starting web service... $frame  (${elapsed}s)"
        $spinnerIndex += 1

        if (Test-ServerReady -HealthCheckUrl $HealthCheckUrl) {
            $launchResult.Ready = $true
            $timer.Stop()
            $form.Close()
            return
        }

        if ($SetupProcess.HasExited) {
            $timer.Stop()
            if ($SetupProcess.ExitCode -ne 0) {
                Write-LauncherLog -Message ("Show-LoadingWindow startup failed: exitCode={0}" -f $SetupProcess.ExitCode)
                $launchResult.Failed = $true
            }
            $form.Close()
        }
    })

    $form.Add_Shown({ $timer.Start() })
    $form.Add_FormClosing({ $timer.Stop() })

    [void]$form.ShowDialog()
    return $launchResult
}

function Start-TrayController {
    param(
        [Parameter(Mandatory = $true)]
        [System.Diagnostics.Process]$SetupProcess,
        [Parameter(Mandatory = $true)]
        [string]$Url,
        [Parameter(Mandatory = $true)]
        [string]$HealthCheckUrl,
        [Parameter(Mandatory = $true)]
        [System.Drawing.Icon]$AppIcon,
        [Parameter(Mandatory = $true)]
        [string]$ShowPanelEventName,
        [Parameter(Mandatory = $true)]
        [string]$SetupStatusPath
    )

    $script:allowWidgetClose = $false
    $script:isExiting = $false
    $script:isDraggingWidget = $false
    $script:widgetDragOffset = New-Object System.Drawing.Point(0, 0)
    $managedBrowserExecutable = Resolve-ManagedBrowserExecutable
    $managedBrowserProfileDir = Join-Path $env:LocalAppData "SwitchConverter\managed-browser-profile"
    $managedBrowserDebugPort = 9334
    $managedBrowserProcessIds = New-Object System.Collections.ArrayList
    $showPanelEvent = New-Object System.Threading.EventWaitHandle($false, [System.Threading.EventResetMode]::AutoReset, $ShowPanelEventName)

    $appContext = New-Object System.Windows.Forms.ApplicationContext

    $widget = New-Object System.Windows.Forms.Form
    $widget.Text = "Switch Converter"
    $widget.StartPosition = "Manual"
    $widget.FormBorderStyle = "None"
    $widget.ShowInTaskbar = $false
    $widget.TopMost = $true
    $widget.MaximizeBox = $false
    $widget.MinimizeBox = $false
    $widget.ClientSize = New-Object System.Drawing.Size(372, 162)
    $widget.BackColor = [System.Drawing.Color]::FromArgb(17, 44, 70)
    $widget.Icon = $AppIcon

    $topBar = New-Object System.Windows.Forms.Panel
    $topBar.Location = New-Object System.Drawing.Point(0, 0)
    $topBar.Size = New-Object System.Drawing.Size(372, 30)
    $topBar.BackColor = [System.Drawing.Color]::FromArgb(21, 61, 95)
    $widget.Controls.Add($topBar)

    $topBarTitle = New-Object System.Windows.Forms.Label
    $topBarTitle.Text = "Switch Converter"
    $topBarTitle.AutoSize = $false
    $topBarTitle.Location = New-Object System.Drawing.Point(10, 6)
    $topBarTitle.Size = New-Object System.Drawing.Size(250, 18)
    $topBarTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $topBarTitle.ForeColor = [System.Drawing.Color]::FromArgb(229, 241, 251)
    $topBar.Controls.Add($topBarTitle)

    $topBarHide = New-Object System.Windows.Forms.Button
    $topBarHide.Text = "×"
    $topBarHide.Size = New-Object System.Drawing.Size(30, 24)
    $topBarHide.Location = New-Object System.Drawing.Point(336, 3)
    $topBarHide.FlatStyle = "Flat"
    $topBarHide.BackColor = [System.Drawing.Color]::FromArgb(32, 80, 121)
    $topBarHide.ForeColor = [System.Drawing.Color]::White
    $topBarHide.FlatAppearance.BorderSize = 0
    $topBarHide.Add_Click({
        $widget.Hide()
    })
    $topBar.Controls.Add($topBarHide)

    $topBarHide.Add_MouseEnter({
        $topBarHide.BackColor = [System.Drawing.Color]::FromArgb(45, 103, 150)
    })
    $topBarHide.Add_MouseLeave({
        $topBarHide.BackColor = [System.Drawing.Color]::FromArgb(32, 80, 121)
    })

    $dragStart = {
        param($sender, $e)
        if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
        $script:isDraggingWidget = $true
        $screenPoint = $sender.PointToScreen([System.Drawing.Point]::new($e.X, $e.Y))
        $script:widgetDragOffset = New-Object System.Drawing.Point(
            ($screenPoint.X - $widget.Left),
            ($screenPoint.Y - $widget.Top)
        )
    }
    $dragMove = {
        param($sender, $e)
        if (-not $script:isDraggingWidget) { return }
        $screenPoint = $sender.PointToScreen([System.Drawing.Point]::new($e.X, $e.Y))
        $widget.Location = New-Object System.Drawing.Point(
            ($screenPoint.X - $script:widgetDragOffset.X),
            ($screenPoint.Y - $script:widgetDragOffset.Y)
        )
    }
    $dragEnd = {
        $script:isDraggingWidget = $false
    }

    $topBar.Add_MouseDown($dragStart)
    $topBar.Add_MouseMove($dragMove)
    $topBar.Add_MouseUp($dragEnd)
    $topBarTitle.Add_MouseDown($dragStart)
    $topBarTitle.Add_MouseMove($dragMove)
    $topBarTitle.Add_MouseUp($dragEnd)

    $title = New-Object System.Windows.Forms.Label
    $title.Text = "Switch Converter is running"
    $title.AutoSize = $false
    $title.Size = New-Object System.Drawing.Size(332, 24)
    $title.Location = New-Object System.Drawing.Point(14, 38)
    $title.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $title.ForeColor = [System.Drawing.Color]::White
    $widget.Controls.Add($title)

    $status = New-Object System.Windows.Forms.Label
    $status.Text = "Status: Starting..."
    $status.AutoSize = $false
    $status.Size = New-Object System.Drawing.Size(332, 20)
    $status.Location = New-Object System.Drawing.Point(14, 62)
    $status.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $status.ForeColor = [System.Drawing.Color]::FromArgb(213, 232, 247)
    $widget.Controls.Add($status)

    $urlLink = New-Object System.Windows.Forms.LinkLabel
    $urlLink.Text = $Url
    $urlLink.AutoSize = $false
    $urlLink.Size = New-Object System.Drawing.Size(332, 20)
    $urlLink.Location = New-Object System.Drawing.Point(14, 83)
    $urlLink.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Underline)
    $urlLink.LinkColor = [System.Drawing.Color]::FromArgb(137, 201, 255)
    $urlLink.ActiveLinkColor = [System.Drawing.Color]::White
    $urlLink.VisitedLinkColor = [System.Drawing.Color]::FromArgb(137, 201, 255)
    $openWebAction = {
        $opened = $false
        try {
            Start-Process -FilePath $Url -ErrorAction Stop | Out-Null
            $opened = $true
        }
        catch {
            $opened = Start-ManagedBrowserWindow `
                -BrowserExecutable $managedBrowserExecutable `
                -Url $Url `
                -ProfileDir $managedBrowserProfileDir `
                -DebugPort $managedBrowserDebugPort `
                -TrackedProcessIds $managedBrowserProcessIds
        }

        if ($opened) {
            $script:openedWebAtLeastOnce = $true
        }
        Write-LauncherLog -Message ("openWebAction: opened={0}, url='{1}', browserFallback='{2}'" -f $opened, $Url, $managedBrowserExecutable)
    }

    $urlLink.Add_LinkClicked({ & $openWebAction })
    $widget.Controls.Add($urlLink)

    $openButton = New-Object System.Windows.Forms.Button
    $openButton.Text = "Open Web"
    $openButton.Size = New-Object System.Drawing.Size(104, 30)
    $openButton.Location = New-Object System.Drawing.Point(14, 116)
    $openButton.FlatStyle = "Flat"
    $openButton.BackColor = [System.Drawing.Color]::FromArgb(38, 124, 199)
    $openButton.ForeColor = [System.Drawing.Color]::White
    $openButton.FlatAppearance.BorderSize = 0
    $openButton.Enabled = $false
    $openButton.Add_Click({ & $openWebAction })
    $widget.Controls.Add($openButton)

    $stopButton = New-Object System.Windows.Forms.Button
    $stopButton.Text = "Stop Server"
    $stopButton.Size = New-Object System.Drawing.Size(116, 30)
    $stopButton.Location = New-Object System.Drawing.Point(126, 116)
    $stopButton.FlatStyle = "Flat"
    $stopButton.BackColor = [System.Drawing.Color]::FromArgb(173, 74, 68)
    $stopButton.ForeColor = [System.Drawing.Color]::White
    $stopButton.FlatAppearance.BorderSize = 0
    $widget.Controls.Add($stopButton)

    $hideButton = New-Object System.Windows.Forms.Button
    $hideButton.Text = "Hide"
    $hideButton.Size = New-Object System.Drawing.Size(84, 30)
    $hideButton.Location = New-Object System.Drawing.Point(248, 116)
    $hideButton.FlatStyle = "Flat"
    $hideButton.BackColor = [System.Drawing.Color]::FromArgb(66, 103, 131)
    $hideButton.ForeColor = [System.Drawing.Color]::White
    $hideButton.FlatAppearance.BorderSize = 0
    $hideButton.Add_Click({ $widget.Hide() })
    $widget.Controls.Add($hideButton)

    $showWidget = {
        if ($widget.IsDisposed) { return }
        if (-not $widget.Visible) {
            Set-BottomRightLocation -Window $widget
            $widget.Show()
        }
        $widget.WindowState = [System.Windows.Forms.FormWindowState]::Normal
        $widget.Activate()
    }

    $monitorTimer = New-Object System.Windows.Forms.Timer
    $monitorTimer.Interval = 700

    $trayMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $menuShow = $trayMenu.Items.Add("Show Panel")
    $menuOpen = $trayMenu.Items.Add("Open Web")
    $null = $trayMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    $menuStop = $trayMenu.Items.Add("Stop Server")
    $menuExit = $trayMenu.Items.Add("Exit Controller")

    $trayIcon = New-Object System.Windows.Forms.NotifyIcon
    $trayIcon.Icon = $AppIcon
    $trayIcon.Text = "Switch Converter (Starting)"
    $trayIcon.ContextMenuStrip = $trayMenu
    $trayIcon.Visible = $true

    $startupErrorShown = $false
    $script:openedWebAtLeastOnce = $false
    $autoOpenAttempted = $false
    $lastSetupStatusText = ""

    $cleanupAndExit = {
        if ($script:isExiting) { return }
        $script:isExiting = $true

        $monitorTimer.Stop()

        if ($trayIcon) {
            $trayIcon.Visible = $false
            $trayIcon.Dispose()
        }

        if ($showPanelEvent) {
            $showPanelEvent.Dispose()
        }

        if ($widget -and -not $widget.IsDisposed) {
            $script:allowWidgetClose = $true
            $widget.Close()
            $widget.Dispose()
        }

        $appContext.ExitThread()
    }

    $stopServerAction = {
        $stopButton.Enabled = $false
        $status.Text = "Stopping server..."
        $status.ForeColor = [System.Drawing.Color]::FromArgb(255, 220, 180)

        $stopped = Stop-WebServer -BootstrapPid $SetupProcess.Id -HealthCheckUrl $HealthCheckUrl

        if ($stopped) {
            Write-LauncherLog -Message "stopServer: stopped=True, autoTabCloseDisabled=True"
            $status.Text = "Status: Stopped"
            $status.ForeColor = [System.Drawing.Color]::FromArgb(255, 190, 190)
            & $cleanupAndExit
        }
        else {
            $status.Text = "Stop failed. Try run_web.bat --debug"
            $status.ForeColor = [System.Drawing.Color]::FromArgb(255, 190, 190)
            $stopButton.Enabled = $true
        }
    }

    $stopButton.Add_Click({ & $stopServerAction })
    $menuStop.Add_Click({ & $stopServerAction })
    $menuShow.Add_Click({ & $showWidget })
    $menuOpen.Add_Click({ & $openWebAction })
    $menuExit.Add_Click({ & $cleanupAndExit })

    $trayIcon.Add_DoubleClick({ & $showWidget })
    $trayIcon.Add_MouseClick({
        param($sender, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            & $showWidget
        }
    })

    $widget.Add_FormClosing({
        param($sender, $e)
        if (-not $script:allowWidgetClose -and $e.CloseReason -eq [System.Windows.Forms.CloseReason]::UserClosing) {
            $e.Cancel = $true
            $widget.Hide()
        }
    })

    $monitorTimer.Add_Tick({
        if ($showPanelEvent -and $showPanelEvent.WaitOne(0)) {
            & $showWidget
        }

        $isReady = Test-ServerReady -HealthCheckUrl $HealthCheckUrl

        if ($isReady) {
            $status.Text = "Status: Running"
            $status.ForeColor = [System.Drawing.Color]::FromArgb(161, 236, 182)
            $trayIcon.Text = "Switch Converter (Running)"
            $openButton.Enabled = $true
            $stopButton.Enabled = $true
            if (-not $autoOpenAttempted) {
                $autoOpenAttempted = $true
                if (-not $script:openedWebAtLeastOnce) {
                    & $openWebAction
                }
            }
            return
        }

        if ($SetupProcess.HasExited) {
            $status.Text = "Status: Stopped"
            $status.ForeColor = [System.Drawing.Color]::FromArgb(255, 190, 190)
            $trayIcon.Text = "Switch Converter (Stopped)"
            $openButton.Enabled = $false
            $stopButton.Enabled = $false
            if ($SetupProcess.ExitCode -ne 0 -and -not $startupErrorShown) {
                $startupErrorShown = $true
                Write-LauncherLog -Message ("startup failed: exitCode={0}" -f $SetupProcess.ExitCode)
            }
            & $cleanupAndExit
            return
        }

        $setupStatusText = Get-SetupStatusText -StatusPath $SetupStatusPath
        if (-not [string]::IsNullOrWhiteSpace($setupStatusText)) {
            $lastSetupStatusText = $setupStatusText
        }

        if (-not [string]::IsNullOrWhiteSpace($lastSetupStatusText)) {
            $status.Text = "Status: $lastSetupStatusText"
        }
        else {
            $status.Text = "Status: Starting..."
        }
        $status.ForeColor = [System.Drawing.Color]::FromArgb(213, 232, 247)
        $trayIcon.Text = "Switch Converter (Starting)"
        $openButton.Enabled = $false
    })

    Set-BottomRightLocation -Window $widget
    $widget.Show()

    $monitorTimer.Start()
    [System.Windows.Forms.Application]::Run($appContext)
}

$mutexName = "Local\SwitchConverterLauncherSingleton"
$showPanelEventName = "Local\SwitchConverterShowPanelEvent"
$isMutexOwner = $false
$singleInstanceMutex = New-Object System.Threading.Mutex($true, $mutexName, [ref]$isMutexOwner)

if (-not $isMutexOwner) {
    if (Test-ServerReady -HealthCheckUrl $healthUrl) {
        try {
            $showEvent = [System.Threading.EventWaitHandle]::OpenExisting($showPanelEventName)
            [void]$showEvent.Set()
            $showEvent.Dispose()
        }
        catch {
            # Keep silent if panel signal cannot be sent.
        }
    }

    $singleInstanceMutex.Dispose()
    if ($appIcon) {
        $appIcon.Dispose()
    }
    return
}

$setupScript = Join-Path $PSScriptRoot "setup_and_run.ps1"
$setupArgs = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $setupScript, "-NoBrowser")
if ($AutoInstallPython) {
    $setupArgs += "-AutoInstallPython"
}

try {
    $setupStatusDir = Split-Path $setupStatusPath -Parent
    if (-not (Test-Path $setupStatusDir)) {
        New-Item -Path $setupStatusDir -ItemType Directory -Force | Out-Null
    }
    Set-Content -Path $setupStatusPath -Value "Starting..." -Encoding UTF8 -Force
}
catch {
    # Ignore status initialization failures.
}

$setupProc = Start-Process -FilePath "powershell" -ArgumentList $setupArgs -WindowStyle Hidden -PassThru

try {
    try {
        Start-TrayController -SetupProcess $setupProc -Url $baseUrl -HealthCheckUrl $healthUrl -AppIcon $appIcon -ShowPanelEventName $showPanelEventName -SetupStatusPath $setupStatusPath
    }
    catch {
        Write-LauncherLog -Message $_.Exception.ToString()
    }
}
finally {
    if ($appIcon) {
        $appIcon.Dispose()
    }
    if ($singleInstanceMutex) {
        try {
            if ($isMutexOwner) {
                $singleInstanceMutex.ReleaseMutex()
            }
        }
        catch {
            # Ignore mutex release errors on shutdown.
        }
        $singleInstanceMutex.Dispose()
    }
}
