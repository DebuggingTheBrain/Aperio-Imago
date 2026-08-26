# ImageScope_FullBatch_RECURSIVE_ALL_Q99_v4.ps1
# Fully automated ImageScope batch workflow:
# Open SVS -> Fit -> read the full slide dimensions from ImageScope -> Extract Region Tool ->
# draw a temporary region -> set the region numerically to the full slide ->
# Thumbnail OFF -> TIF:LZW -> Extract -> wait until the TIFF is finished ->
# close the Extract dialog -> close the image -> continue with the next image.
#
# IMPORTANT:
# - ImageScope itself creates every TIFF file.
# - Windows must stay unlocked.
# - Please do not use the mouse or keyboard while the batch is running.
# - Do not manually move or minimize the ImageScope window.
# Abhinav Singh
# 26.08.2026



$ErrorActionPreference = "Stop"

# =========================
# CONFIGURATION
# =========================

# Enter the folder that contains the .svs files here:
$InputFolder = "C:\PATH\TO\YOUR\SVS_FOLDER"





# Also search all subfolders recursively.
$Recurse = $true








# 0 = process ALL .svs files that were found.
$MaxFiles = 0




# Optional: If .svs files are NOT associated with ImageScope in Windows,
# enter the full path to ImageScope.exe here.
# Otherwise, leave this empty.
$ImageScopeExe = ""

$AppName = "ImageScope"

# Maximum time allowed for loading one slide.
$OpenTimeoutSeconds = 180

# Maximum time allowed for each export.
$ExportTimeoutMinutes = 120

# Optional: Before switching to the final TIF:LZW format, the script tries to
# set the visible Quality value to 99. ImageScope normally disables this field
# for TIF:LZW, so Quality 99 is only prepared as a best-effort step before the
# script switches back to TIF:LZW.
$TryVisualQuality99 = $true

# Log file used for resuming the batch and investigating errors.
$LogFile = Join-Path $InputFolder "ImageScope_batch_log.csv"

# =========================
# HELPER FUNCTIONS
# =========================

function Invoke-WinAppRaw {
    param(
        [Parameter(Mandatory=$true)]
        [string[]]$Arguments
    )

    # winapp also writes expected "not found" messages to STDERR.
    # With the global $ErrorActionPreference set to "Stop", these messages could
    # otherwise stop the complete batch, even during optional cleanup steps.
    $oldErrorActionPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        $output = @(& winapp @Arguments 2>&1)
        $exitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $oldErrorActionPreference
    }

    [pscustomobject]@{
        ExitCode = $exitCode
        Lines    = $output
        Text     = ($output | ForEach-Object { "$_" }) -join "`n"
    }
}

function Invoke-WinAppChecked {
    param(
        [Parameter(Mandatory=$true)]
        [string[]]$Arguments,
        [string]$What = "winapp command"
    )

    $r = Invoke-WinAppRaw -Arguments $Arguments
    if ($r.ExitCode -ne 0) {
        throw "$What failed.`n$($r.Text)"
    }
    return $r
}

function Wait-Until {
    param(
        [Parameter(Mandatory=$true)]
        [scriptblock]$Condition,
        [int]$TimeoutSeconds = 30,
        [int]$PollMilliseconds = 500,
        [string]$Description = "condition"
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        try {
            if (& $Condition) {
                return $true
            }
        } catch {
            # The UI may only be temporarily unavailable, so try again.
        }

        Start-Sleep -Milliseconds $PollMilliseconds
    } while ((Get-Date) -lt $deadline)

    throw "Timeout waiting for $Description."
}

function Get-ImageNavBounds {
    $r = Invoke-WinAppChecked `
        -Arguments @("ui","inspect","ImageNav1","-a",$AppName,"--depth","1","--quiet") `
        -What "Inspect ImageNav1"

    # Example: ImageNav1 Pane (2125,74 1713x934)
    if ($r.Text -match 'ImageNav1.*?\((-?\d+),(-?\d+)\s+(\d+)x(\d+)\)') {
        return [pscustomobject]@{
            X      = [int]$matches[1]
            Y      = [int]$matches[2]
            Width  = [int]$matches[3]
            Height = [int]$matches[4]
        }
    }

    throw "Could not parse ImageNav1 bounds.`n$($r.Text)"
}

function Get-SlideDimensions {
    $r = Invoke-WinAppChecked `
        -Arguments @("ui","inspect","staStatusBar","-a",$AppName,"--depth","2","--quiet") `
        -What "Read ImageScope status bar"

    # Example: "57764 x 37348 x 3 = 6.0GB, File = 1.2GB"
    if ($r.Text -match '"(\d+)\s*x\s*(\d+)\s*x\s*(\d+)\s*=') {
        return [pscustomobject]@{
            Width    = [int]$matches[1]
            Height   = [int]$matches[2]
            Channels = [int]$matches[3]
        }
    }

    throw "No full-slide dimensions found in status bar.`n$($r.Text)"
}

function Get-ImageScopeHwnds {
    $r = Invoke-WinAppRaw -Arguments @(
        "ui","list-windows","-a",$AppName,"--show-hidden","--quiet"
    )

    $ids = New-Object System.Collections.Generic.List[long]

    foreach ($line in $r.Lines) {
        $s = "$line"
        if ($s -match 'HWND\s+(\d+)') {
            $id = [long]$matches[1]
            if (-not $ids.Contains($id)) {
                $ids.Add($id)
            }
        }
    }

    return @($ids)
}

function Invoke-ExtractRegionTool {
    # Right-click inside the image area. winapp click also brings ImageScope to the front.
    Invoke-WinAppChecked `
        -Arguments @("ui","click","ImageNav1","-a",$AppName,"--right","--quiet") `
        -What "Right-click image canvas" | Out-Null

    Start-Sleep -Milliseconds 400

    # Context menus and pop-ups may have their own HWNDs.
    # winapp list-windows also returns pop-ups and dialogs.
    $hwnds = Get-ImageScopeHwnds

    foreach ($hwnd in $hwnds) {
        $search = Invoke-WinAppRaw -Arguments @(
            "ui","search","Extract Region Tool","-w","$hwnd","--max","10","--quiet"
        )

        if ($search.ExitCode -eq 0 -and $search.Text -match 'Extract Region') {
            $invoke = Invoke-WinAppRaw -Arguments @(
                "ui","invoke","Extract Region Tool","-w","$hwnd","--quiet"
            )
            if ($invoke.ExitCode -eq 0) {
                return
            }
        }
    }

    # Second attempt using a shorter search term.
    $hwnds = Get-ImageScopeHwnds
    foreach ($hwnd in $hwnds) {
        $search = Invoke-WinAppRaw -Arguments @(
            "ui","search","Extract Region","-w","$hwnd","--max","10","--quiet"
        )

        if ($search.ExitCode -eq 0 -and $search.Text -match 'Extract Region') {
            $invoke = Invoke-WinAppRaw -Arguments @(
                "ui","invoke","Extract Region","-w","$hwnd","--quiet"
            )
            if ($invoke.ExitCode -eq 0) {
                return
            }
        }
    }

    # Close the context menu if it is still open.
    Invoke-WinAppRaw -Arguments @(
        "ui","send-keys","esc","-a",$AppName,"--via","send-input","--quiet"
    ) | Out-Null

    throw 'Could not invoke "Extract Region Tool" from the ImageScope context menu.'
}

function Open-ExtractDialogByDrag {
    $bounds = Get-ImageNavBounds

    # Draw any temporary region near the center of the image.
    # The exact full-slide region is entered numerically afterwards.
    # Using 30%-65% normally avoids zoom controls and the thumbnail overlay.
    $x1 = [int][math]::Round($bounds.X + 0.30 * $bounds.Width)
    $y1 = [int][math]::Round($bounds.Y + 0.30 * $bounds.Height)
    $x2 = [int][math]::Round($bounds.X + 0.65 * $bounds.Width)
    $y2 = [int][math]::Round($bounds.Y + 0.68 * $bounds.Height)

    Invoke-WinAppChecked `
        -Arguments @(
            "ui","drag",
            "$x1,$y1",
            "$x2,$y2",
            "-a",$AppName,
            "--hold-ms","150",
            "--dwell-ms","150",
            "--quiet"
        ) `
        -What "Draw dummy extract region" | Out-Null

    Invoke-WinAppChecked `
        -Arguments @(
            "ui","wait-for","frmExtract",
            "-a",$AppName,
            "--timeout","15000",
            "--quiet"
        ) `
        -What 'Wait for "Extract Image Region" dialog' | Out-Null
}

function Set-FullSlideRegion {
    param(
        [Parameter(Mandatory=$true)][int]$Width,
        [Parameter(Mandatory=$true)][int]$Height
    )

    # Left and Top are intentionally set first.
    # This keeps the full Width and Height within the valid boundaries afterwards.
    Invoke-WinAppChecked `
        -Arguments @("ui","set-value","TextILeft","0","-a",$AppName,"--quiet") `
        -What "Set Left=0" | Out-Null

    Invoke-WinAppChecked `
        -Arguments @("ui","set-value","TextITop","0","-a",$AppName,"--quiet") `
        -What "Set Top=0" | Out-Null

    Invoke-WinAppChecked `
        -Arguments @("ui","set-value","TextIWidth","$Width","-a",$AppName,"--quiet") `
        -What "Set full width" | Out-Null

    Invoke-WinAppChecked `
        -Arguments @("ui","set-value","TextIHeight","$Height","-a",$AppName,"--quiet") `
        -What "Set full height" | Out-Null

    # Verify that ImageScope accepted all four values.
    $left   = (Invoke-WinAppChecked -Arguments @("ui","get-value","TextILeft","-a",$AppName,"--quiet") -What "Verify Left").Text.Trim()
    $top    = (Invoke-WinAppChecked -Arguments @("ui","get-value","TextITop","-a",$AppName,"--quiet") -What "Verify Top").Text.Trim()
    $width2 = (Invoke-WinAppChecked -Arguments @("ui","get-value","TextIWidth","-a",$AppName,"--quiet") -What "Verify Width").Text.Trim()
    $height2= (Invoke-WinAppChecked -Arguments @("ui","get-value","TextIHeight","-a",$AppName,"--quiet") -What "Verify Height").Text.Trim()

    if ($left -notmatch '(^|\D)0(\D|$)' -or
        $top -notmatch '(^|\D)0(\D|$)' -or
        $width2 -notmatch [regex]::Escape("$Width") -or
        $height2 -notmatch [regex]::Escape("$Height")) {
        throw "Full-slide region verification failed: Left=$left Top=$top Width=$width2 Height=$height2"
    }
}

function Try-SetVisualQuality99 {
    if (-not $TryVisualQuality99) {
        return
    }

    function Get-QualityNumericSelector {
        $inspect = Invoke-WinAppRaw -Arguments @(
            "ui","inspect","frmExtract","-a",$AppName,"--depth","5","--quiet"
        )

        foreach ($line in $inspect.Lines) {
            $s = "$line"
            if ($s -match '^\s*([^\s]+)\s+Edit\s+"RGB"\s+value="\d+"') {
                return $matches[1]
            }
        }

        return $null
    }

    $candidateFormats = @(
        "OptionJP2",        # JPEG 2000
        "OptionJPG_JPEG",   # JPEG
        "OptionSVS_JPEG",
        "OptionCWS_JPEG"
    )

    foreach ($formatSelector in $candidateFormats) {
        $sel = $null
        try {
            Invoke-WinAppChecked `
                -Arguments @("ui","invoke",$formatSelector,"-a",$AppName,"--quiet") `
                -What "Temporarily enable quality control via $formatSelector" | Out-Null

            Start-Sleep -Milliseconds 300

            $sel = Get-QualityNumericSelector

            if ($sel) {
                $set = Invoke-WinAppRaw -Arguments @(
                    "ui","set-value",$sel,"99","-a",$AppName,"--quiet"
                )

                if ($set.ExitCode -ne 0) {
                    Invoke-WinAppRaw -Arguments @(
                        "ui","send-keys","ctrl+a",
                        "--target",$sel,
                        "--via","send-input",
                        "-a",$AppName,"--quiet"
                    ) | Out-Null

                    Start-Sleep -Milliseconds 120

                    Invoke-WinAppRaw -Arguments @(
                        "ui","send-keys","99",
                        "--verbatim",
                        "--target",$sel,
                        "--via","send-input",
                        "-a",$AppName,"--quiet"
                    ) | Out-Null
                }

                Start-Sleep -Milliseconds 120

                $verify = Invoke-WinAppRaw -Arguments @(
                    "ui","get-value",$sel,"-a",$AppName,"--quiet"
                )

                if ($verify.ExitCode -eq 0 -and $verify.Text -match '(^|\D)99(\D|$)') {
                    return
                }
            }
        } catch {
            # This option did not work, so try the next candidate.
        }
    }

    Write-Warning "Quality 99 could not be set visually before switching back to TIF:LZW. Continuing with TIF:LZW export."
}

function Ensure-ThumbnailOff {
    # Read the live state. This is more reliable than using an old inspect dump.
    $thumb = Invoke-WinAppChecked `
        -Arguments @("ui","get-property","CheckThumb","-p","ToggleState","-a",$AppName,"--quiet") `
        -What "Read Thumbnail state"

    if ($thumb.Text -match 'ToggleState:\s*On') {
        Invoke-WinAppChecked `
            -Arguments @("ui","invoke","CheckThumb","-a",$AppName,"--quiet") `
            -What "Turn Thumbnail OFF" | Out-Null

        Start-Sleep -Milliseconds 250
    }

    # Check the live state again. If ImageScope ignored the first toggle,
    # make exactly one more attempt.
    $thumb2 = Invoke-WinAppChecked `
        -Arguments @("ui","get-property","CheckThumb","-p","ToggleState","-a",$AppName,"--quiet") `
        -What "Verify Thumbnail state"

    if ($thumb2.Text -match 'ToggleState:\s*On') {
        Invoke-WinAppChecked `
            -Arguments @("ui","invoke","CheckThumb","-a",$AppName,"--quiet") `
            -What "Retry turning Thumbnail OFF" | Out-Null

        Start-Sleep -Milliseconds 250

        $thumb3 = Invoke-WinAppChecked `
            -Arguments @("ui","get-property","CheckThumb","-p","ToggleState","-a",$AppName,"--quiet") `
            -What "Final Thumbnail verification"

        if ($thumb3.Text -match 'ToggleState:\s*On') {
            throw "Thumbnail is still ON after two toggle attempts. Export aborted to avoid saving a thumbnail."
        }
    }

    Write-Host "  Thumbnail: OFF"
}

function Set-ExportOptions {
    # 1) Best effort: set the visible Quality value to 99.
    #    ImageScope may briefly switch to a JPEG or JPEG 2000 format for this.
    Try-SetVisualQuality99

    # 2) The FINAL format must be TIF:LZW.
    Invoke-WinAppChecked `
        -Arguments @("ui","invoke","OptionTIF_LZW","-a",$AppName,"--quiet") `
        -What "Select final TIF:LZW" | Out-Null

    Start-Sleep -Milliseconds 250

    # 3) Finally, make sure Thumbnail is OFF.
    #    Important: Changing the format may change this setting again in ImageScope.
    Ensure-ThumbnailOff
}

function Get-OutputFile {
    # Use JSON so long paths are not wrapped in the middle of the filename by
    # the console width. get-value --json returns a "text" field.
    $r = Invoke-WinAppChecked `
        -Arguments @("ui","get-value","TextOutputFile","-a",$AppName,"--json") `
        -What "Read ImageScope output file"

    try {
        $json = $r.Text | ConvertFrom-Json
    }
    catch {
        throw "Could not parse winapp JSON for TextOutputFile.`n$($r.Text)"
    }

    $candidate = $null

    if ($json.PSObject.Properties.Name -contains "text") {
        $candidate = [string]$json.text
    }
    elseif ($json.PSObject.Properties.Name -contains "value") {
        $candidate = [string]$json.value
    }

    if ($candidate) {
        $candidate = $candidate.Trim()

        if ($candidate -match '^[A-Za-z]:\\.*\.(tif|tiff)$') {
            return $candidate
        }
    }

    throw "Could not read TIFF output path from TextOutputFile.`nJSON:`n$($r.Text)"
}

function Test-FileExclusiveReadable {
    param([Parameter(Mandatory=$true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return $false
    }

    try {
        $stream = [System.IO.File]::Open(
            $Path,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::None
        )
        $stream.Close()
        return $true
    } catch {
        return $false
    }
}

function Wait-ForExportComplete {
    param(
        [Parameter(Mandatory=$true)][string]$OutputPath,
        [Parameter(Mandatory=$true)][datetime]$ExportStarted
    )

    $deadline = (Get-Date).AddMinutes($ExportTimeoutMinutes)
    $previousLength = -1L
    $exclusiveSuccesses = 0

    while ((Get-Date) -lt $deadline) {
        if (Test-Path -LiteralPath $OutputPath) {
            try {
                $item = Get-Item -LiteralPath $OutputPath
                $length = [long]$item.Length

                # Make sure this file belongs to the current export.
                $isCurrent = ($item.LastWriteTime -ge $ExportStarted.AddSeconds(-2))

                if ($length -gt 0 -and $isCurrent) {
                    $exclusive = Test-FileExclusiveReadable -Path $OutputPath

                    if ($exclusive -and $length -eq $previousLength) {
                        $exclusiveSuccesses++
                    } else {
                        $exclusiveSuccesses = 0
                    }

                    # Two consecutive exclusive opens with an unchanged file size mean
                    # that ImageScope has most likely closed and finalized the file.
                    if ($exclusiveSuccesses -ge 2) {
                        return
                    }

                    $previousLength = $length
                }
            } catch {
                # The file is still being created, renamed, or finalized. Keep waiting.
            }
        }

        Start-Sleep -Seconds 1
    }

    throw "Export timeout after $ExportTimeoutMinutes minutes: $OutputPath"
}

function Close-ExtractDialog {
    $r = Invoke-WinAppRaw -Arguments @(
        "ui","invoke","CommandCancel","-a",$AppName,"--quiet"
    )

    if ($r.ExitCode -eq 0) {
        Invoke-WinAppRaw -Arguments @(
            "ui","wait-for","frmExtract","-a",$AppName,
            "--gone","--timeout","5000","--quiet"
        ) | Out-Null
    }
}

function Close-CurrentImage {
    $r = Invoke-WinAppRaw -Arguments @(
        "ui","invoke","mnu-closeimage-0011","-a",$AppName,"--quiet"
    )
    Start-Sleep -Milliseconds 800
}

function Reset-ImageScopeUiAfterError {
    # Escape may close an open message box or context menu.
    Invoke-WinAppRaw -Arguments @(
        "ui","send-keys","esc","-a",$AppName,"--via","send-input","--quiet"
    ) | Out-Null

    Start-Sleep -Milliseconds 300
    Close-ExtractDialog
    Close-CurrentImage
}

function Open-Slide {
    param([Parameter(Mandatory=$true)][System.IO.FileInfo]$File)

    if ($ImageScopeExe -and (Test-Path -LiteralPath $ImageScopeExe)) {
        Start-Process -FilePath $ImageScopeExe -ArgumentList "`"$($File.FullName)`""
    } else {
        # Use the Windows file association. .svs files must be associated with ImageScope.
        Start-Process -FilePath $File.FullName
    }

    # Wait until ImageScope reports real slide dimensions in the status bar.
    Wait-Until `
        -TimeoutSeconds $OpenTimeoutSeconds `
        -PollMilliseconds 700 `
        -Description "ImageScope to load $($File.Name)" `
        -Condition {
            $r = Invoke-WinAppRaw -Arguments @(
                "ui","inspect","staStatusBar","-a",$AppName,
                "--depth","2","--quiet"
            )
            return ($r.ExitCode -eq 0 -and
                    $r.Text -match '"\d+\s*x\s*\d+\s*x\s*\d+\s*=')
        } | Out-Null
}

function Write-LogRow {
    param(
        [string]$Source,
        [string]$Output,
        [string]$Width,
        [string]$Height,
        [string]$Status,
        [datetime]$Started,
        [datetime]$Ended,
        [string]$Message
    )

    $row = [pscustomobject]@{
        Source  = $Source
        Output  = $Output
        Width   = $Width
        Height  = $Height
        Status  = $Status
        Started = $Started.ToString("s")
        Ended   = $Ended.ToString("s")
        Message = $Message
    }

    $row | Export-Csv -Path $LogFile -Append -NoTypeInformation -Encoding UTF8
}

# =========================
# KEEP THE COMPUTER AWAKE
# =========================

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class Awake {
    [DllImport("kernel32.dll")]
    public static extern uint SetThreadExecutionState(uint esFlags);
}
"@

# Important: Windows PowerShell 5.1 otherwise interprets 0x80000000 as a
# negative Int32, but the Win32 API expects a UInt32 value here.
[uint32]$ES_CONTINUOUS       = [Convert]::ToUInt32("80000000", 16)
[uint32]$ES_SYSTEM_REQUIRED  = 1
[uint32]$ES_DISPLAY_REQUIRED = 2
[uint32]$WakeFlags = $ES_CONTINUOUS -bor $ES_SYSTEM_REQUIRED -bor $ES_DISPLAY_REQUIRED

[Awake]::SetThreadExecutionState($WakeFlags) | Out-Null

# =========================
# FILE DISCOVERY AND RESUME SUPPORT
# =========================

if (-not (Test-Path -LiteralPath $InputFolder)) {
    throw "InputFolder does not exist: $InputFolder"
}

$getChildArgs = @{
    Path   = $InputFolder
    File   = $true
    Filter = "*.svs"
}
if ($Recurse) {
    $getChildArgs.Recurse = $true
}

$files = @(Get-ChildItem @getChildArgs | Sort-Object FullName)

if ($MaxFiles -gt 0) {
    $files = @($files | Select-Object -First $MaxFiles)
}

if ($files.Count -eq 0) {
    throw "No .svs files found in $InputFolder"
}

$completed = @{}
if (Test-Path -LiteralPath $LogFile) {
    try {
        foreach ($row in (Import-Csv -LiteralPath $LogFile)) {
            if ($row.Status -eq "SUCCESS") {
                $completed[$row.Source] = $true
            }
        }
    } catch {
        Write-Warning "Could not read old log; continuing without resume data."
    }
}

Write-Host ""
Write-Host "ImageScope full batch"
Write-Host "Input: $InputFolder"
Write-Host "Files found: $($files.Count)"
Write-Host "Log: $LogFile"
Write-Host ""
Write-Host "IMPORTANT: Keep Windows unlocked and do not use mouse/keyboard during the batch."
Write-Host ""

try {
    # If ImageScope is already open, try to close old dialogs first.
    # These steps are optional and must never stop the batch.
    Invoke-WinAppRaw -Arguments @(
        "ui","send-keys","esc","-a",$AppName,"--via","send-input","--quiet"
    ) | Out-Null

    $closeAllProbe = Invoke-WinAppRaw -Arguments @(
        "ui","search","Close All Images","-a",$AppName,"--max","5","--quiet"
    )
    if ($closeAllProbe.ExitCode -eq 0 -and $closeAllProbe.Text -match 'Close All Images') {
        Invoke-WinAppRaw -Arguments @(
            "ui","invoke","Close All Images","-a",$AppName,"--quiet"
        ) | Out-Null
        Start-Sleep -Milliseconds 500
    }

    $index = 0

    foreach ($file in $files) {
        $index++

        if ($completed.ContainsKey($file.FullName)) {
            Write-Host "[$index/$($files.Count)] SKIP already completed: $($file.Name)"
            continue
        }

        $started = Get-Date
        $outputPath = ""
        $dims = $null

        Write-Host ""
        Write-Host "[$index/$($files.Count)] START $($file.Name)"

        try {
            # 1) Open the slide.
            Open-Slide -File $file

            # 2) Fit the complete slide into the window.
            Invoke-WinAppChecked `
                -Arguments @("ui","invoke","cmdFit","-a",$AppName,"--quiet") `
                -What "Fit slide" | Out-Null

            Start-Sleep -Milliseconds 500

            # 3) Read the real full-slide dimensions directly from ImageScope.
            $dims = Get-SlideDimensions
            Write-Host "  Dimensions: $($dims.Width) x $($dims.Height) x $($dims.Channels)"

            # 4) Activate Extract Region Tool from the ImageScope context menu.
            Invoke-ExtractRegionTool

            Start-Sleep -Milliseconds 300

            # 5) Draw a temporary region automatically to open the dialog.
            Open-ExtractDialogByDrag

            # 6) Replace the temporary region with the exact full-slide dimensions.
            Set-FullSlideRegion -Width $dims.Width -Height $dims.Height

            # 7) Set Thumbnail OFF and select TIF:LZW.
            Set-ExportOptions

            # 8) Read the output path generated by ImageScope.
            $outputPath = Get-OutputFile
            Write-Host "  Output: $outputPath"

            # If the target already exists, ImageScope normally chooses a numbered filename.
            # If the exact path still exists, do not delete or overwrite it here.
            if (Test-Path -LiteralPath $outputPath) {
                $old = Get-Item -LiteralPath $outputPath
                Write-Host "  Note: target path already exists; waiting for ImageScope's current write/overwrite behavior."
            }

            # 9) Start the export.
            $exportStarted = Get-Date

            Invoke-WinAppChecked `
                -Arguments @("ui","invoke","CommandExtract","-a",$AppName,"--quiet") `
                -What "Start ImageScope Extract" | Out-Null

            # 10) Continue only after the TIFF is finished and released by ImageScope.
            Wait-ForExportComplete -OutputPath $outputPath -ExportStarted $exportStarted

            Write-Host "  Export complete."

            # 11) Close the Extract dialog.
            Close-ExtractDialog

            # 12) Close the current slide.
            Close-CurrentImage

            $ended = Get-Date

            Write-LogRow `
                -Source $file.FullName `
                -Output $outputPath `
                -Width "$($dims.Width)" `
                -Height "$($dims.Height)" `
                -Status "SUCCESS" `
                -Started $started `
                -Ended $ended `
                -Message ""

            Write-Host "[$index/$($files.Count)] SUCCESS $($file.Name)"
        }
        catch {
            $ended = Get-Date
            $message = $_.Exception.Message

            Write-Warning "FAILED: $($file.Name)"
            Write-Warning $message

            Write-LogRow `
                -Source $file.FullName `
                -Output $outputPath `
                -Width $(if ($dims) { "$($dims.Width)" } else { "" }) `
                -Height $(if ($dims) { "$($dims.Height)" } else { "" }) `
                -Status "FAILED" `
                -Started $started `
                -Ended $ended `
                -Message $message

            Reset-ImageScopeUiAfterError

            # Do not stop a large batch because of one failed slide.
            # Log the error and continue with the next file.
            continue
        }
    }
}
finally {
    # Restore the normal Windows power-saving behavior.
    [Awake]::SetThreadExecutionState([uint32]$ES_CONTINUOUS) | Out-Null
}

Write-Host ""
Write-Host "DONE."
Write-Host "Log: $LogFile"
