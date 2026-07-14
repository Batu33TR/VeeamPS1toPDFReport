<#
.SYNOPSIS
    Generates a N-day Veeam Backup & Replication HTML/PDF report (success/failed, speeds,
    backed-up machines).

.DESCRIPTION
    Run this script LOCALLY on the Veeam Backup & Replication server (Windows Server
    2019+). It uses the Veeam PowerShell module to pull job/session history for the
    last N days, builds a clean HTML report (summary cards + job table + per-machine
    table), saves it to disk, and automatically converts it to PDF.

.NOTES
    Requires: Veeam Backup & Replication PowerShell module (installed automatically
    with the Veeam B&R console). Tested against Veeam v11/v12 cmdlets.
    PDF output uses headless Microsoft Edge or Google Chrome (Edge is recommended
    on Windows Server 2019/2022).

.PARAMETER ReportDays
    How many days back to include. Default 7.

.PARAMETER CustomerName
    Friendly name shown in the report header (e.g. "CompanyName Ltd.").

.PARAMETER OutputFolder
    Where to save the generated HTML and PDF files.

.PARAMETER SkipPdf
    Switch. If present, only the HTML file is saved (no PDF conversion).

.EXAMPLE
    .\VeeamReport30Days.ps1 -CustomerName "CompanyName Ltd." -OutputFolder "D:\Reports" -ReportDays 30

.EXAMPLE
    # HTML only, skip PDF
    .\VeeamReport30Days.ps1 -CustomerName "CompanyName Ltd." -SkipPdf
#>

[CmdletBinding()]
param(
    [int]    $ReportDays     = 30,
    [string] $CustomerName   = "Customer",
    [string] $OutputFolder   = "C:\VeeamReports",
    [switch] $SkipPdf
)

$ErrorActionPreference = "Stop"

# ------------------------------------------------------------------
# 0. Helpers
# ------------------------------------------------------------------

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$ts] [$Level] $Message"
}

function Format-Bytes {
    param([double]$Bytes)
    if ($Bytes -ge 1TB) { return "{0:N2} TB" -f ($Bytes / 1TB) }
    elseif ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    elseif ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    elseif ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    else { return "$Bytes B" }
}

function Format-Speed {
    param([double]$BytesPerSec)
    if (-not $BytesPerSec -or $BytesPerSec -le 0) { return "-" }
    return (Format-Bytes -Bytes $BytesPerSec) + "/s"
}

function Format-Duration {
    param([TimeSpan]$Duration)
    if (-not $Duration) { return "-" }
    "{0:00}h {1:00}m {2:00}s" -f $Duration.Hours, $Duration.Minutes, $Duration.Seconds
}

function Format-DT {
    # Renders dates/times as DD/MM/YYYY HH:mm:ss regardless of server locale
    param($DateTimeValue)
    if (-not $DateTimeValue) { return "-" }
    try { return ([datetime]$DateTimeValue).ToString("dd/MM/yyyy HH:mm:ss") }
    catch { return $DateTimeValue }
}

function HtmlEncode {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    [System.Net.WebUtility]::HtmlEncode($Text)
}

function Get-UsageColor {
    # Used-space thresholds: high usage = bad (opposite sense of success-rate color)
    param([double]$Pct)
    if ($Pct -ge 90) { return "#c9432f" }
    elseif ($Pct -ge 75) { return "#c98a1f" }
    else { return "#1e8e5a" }
}

function Get-BytesValue {
    param($MemorySizeObj)
    if ($null -eq $MemorySizeObj) { return 0 }
    if ($MemorySizeObj -is [long] -or $MemorySizeObj -is [int] -or $MemorySizeObj -is [double]) {
        return [long]$MemorySizeObj
    }
    foreach ($propName in @('InBytes', 'InBytesAsInt64', 'Bytes', 'TotalBytes')) {
        if ($MemorySizeObj.PSObject.Properties.Name -contains $propName) {
            return [long]$MemorySizeObj.$propName
        }
    }
    $s = $MemorySizeObj.ToString()
    if ($s -match '\((\d+)\)') { return [long]$Matches[1] }
    return 0
}

function Get-TaskStartTime {
    param($Task, $FallbackSession)
    foreach ($getter in @(
            { $Task.Progress.StartTime },
            { $Task.Progress.StartTimeLocal },
            { $Task.StartTime },
            { $Task.StartTimeLocal },
            { $Task.CreationTime }
        )) {
        try {
            $val = & $getter
            if ($val -and $val -ne [datetime]::MinValue) { return $val }
        }
        catch { }
    }
    return $FallbackSession.CreationTime
}

function Convert-HtmlFileToPdf {
    param(
        [Parameter(Mandatory)]
        [string]$HtmlPath,
        [Parameter(Mandatory)]
        [string]$PdfPath
    )

    if (-not (Test-Path -LiteralPath $HtmlPath)) {
        throw "HTML file not found: $HtmlPath"
    }

    $browser = @(
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1

    if (-not $browser) {
        Write-Log "Skipping PDF: Microsoft Edge or Google Chrome is required for HTML-to-PDF conversion." "WARN"
        return $false
    }

    $htmlUri = ([Uri][System.IO.Path]::GetFullPath($HtmlPath)).AbsoluteUri
    $pdfFullPath = [System.IO.Path]::GetFullPath($PdfPath)
    $pdfDir = Split-Path $pdfFullPath -Parent
    if (-not (Test-Path $pdfDir)) {
        New-Item -Path $pdfDir -ItemType Directory -Force | Out-Null
    }
    if (Test-Path -LiteralPath $pdfFullPath) {
        Remove-Item -LiteralPath $pdfFullPath -Force
    }

    $argSets = @(
        @(
            '--headless=new',
            '--disable-gpu',
            '--run-all-compositor-stages-before-draw',
            '--virtual-time-budget=15000',
            "--print-to-pdf=$pdfFullPath",
            '--no-pdf-header-footer',
            $htmlUri
        ),
        @(
            '--headless',
            '--disable-gpu',
            '--run-all-compositor-stages-before-draw',
            '--virtual-time-budget=15000',
            "--print-to-pdf=$pdfFullPath",
            '--no-pdf-header-footer',
            $htmlUri
        )
    )

    foreach ($browserArgs in $argSets) {
        try {
            Start-Process -FilePath $browser -ArgumentList $browserArgs -WindowStyle Hidden -Wait -ErrorAction Stop | Out-Null
        }
        catch {
            continue
        }

        $deadline = (Get-Date).AddSeconds(30)
        while (-not (Test-Path -LiteralPath $pdfFullPath) -and (Get-Date) -lt $deadline) {
            Start-Sleep -Milliseconds 250
        }
        if (Test-Path -LiteralPath $pdfFullPath) {
            return $true
        }
    }

    Write-Log "PDF conversion failed for $HtmlPath (browser: $browser)." "WARN"
    return $false
}

# ------------------------------------------------------------------
# 1. Connect to Veeam
# ------------------------------------------------------------------

Write-Log "Loading Veeam PowerShell module..."
try {
    if (Get-Module -ListAvailable -Name Veeam.Backup.PowerShell) {
        Import-Module Veeam.Backup.PowerShell -ErrorAction Stop
    }
    else {
        Add-PSSnapin VeeamPSSnapin -ErrorAction Stop
    }
}
catch {
    Write-Log "Could not load Veeam PowerShell module/snapin. Is this running on the Veeam B&R server? $_" "ERROR"
    throw
}

try {
    if (-not (Get-VBRServerSession)) {
        Connect-VBRServer -Server "localhost"
    }
}
catch {
    Write-Log "Failed to connect to the local Veeam B&R server: $_" "ERROR"
    throw
}

# ------------------------------------------------------------------
# 2. Pull session history
# ------------------------------------------------------------------

$startDate = (Get-Date).AddDays(-$ReportDays)
Write-Log "Collecting backup sessions since $startDate ..."

$allSessions = Get-VBRBackupSession | Where-Object { $_.CreationTime -ge $startDate }

if (-not $allSessions -or $allSessions.Count -eq 0) {
    Write-Log "No backup sessions found in the last $ReportDays days." "WARN"
}

# ------------------------------------------------------------------
# 2b. Repository storage usage
# ------------------------------------------------------------------

Write-Log "Collecting repository storage usage..."

$repoRows = @()

Get-VBRBackupRepository -WarningAction SilentlyContinue | ForEach-Object {
    $r = $_
    $total = 0; $free = 0
    try {
        $c = $r.GetContainer()
        $total = Get-BytesValue $c.CachedTotalSpace
        $free  = Get-BytesValue $c.CachedFreeSpace
    }
    catch {
        Write-Log "Could not read capacity for repository '$($r.Name)': $_" "WARN"
    }
    $used = [math]::Max($total - $free, 0)
    $repoRows += [PSCustomObject]@{
        Name   = $r.Name
        Type   = "Repository"
        TotalB = $total
        UsedB  = $used
        FreeB  = $free
    }
}

if (Get-Command Get-VBRScaleOutBackupRepository -ErrorAction SilentlyContinue) {
    Get-VBRScaleOutBackupRepository -WarningAction SilentlyContinue | ForEach-Object {
        $sobr = $_
        $total = 0; $free = 0
        foreach ($extent in $sobr.Extent) {
            try {
                $c = $extent.Repository.GetContainer()
                $total += Get-BytesValue $c.CachedTotalSpace
                $free  += Get-BytesValue $c.CachedFreeSpace
            }
            catch {
                Write-Log "Could not read capacity for SOBR extent on '$($sobr.Name)': $_" "WARN"
            }
        }
        $used = [math]::Max($total - $free, 0)
        $repoRows += [PSCustomObject]@{
            Name   = $sobr.Name
            Type   = "Scale-Out ($($sobr.Extent.Count) extents)"
            TotalB = $total
            UsedB  = $used
            FreeB  = $free
        }
    }
}
else {
    Write-Log "Get-VBRScaleOutBackupRepository not available on this Veeam install - skipping SOBR." "INFO"
}

$repoRows = $repoRows | ForEach-Object {
    $pct = if ($_.TotalB -gt 0) { [math]::Round(($_.UsedB / $_.TotalB) * 100, 1) } else { 0 }
    $_ | Add-Member -NotePropertyName UsedPct -NotePropertyValue $pct -PassThru
} | Sort-Object Name

$totalCapacityB = ($repoRows | Measure-Object -Property TotalB -Sum).Sum
$totalUsedB     = ($repoRows | Measure-Object -Property UsedB -Sum).Sum
$totalFreeB     = ($repoRows | Measure-Object -Property FreeB -Sum).Sum
$overallUsedPct = if ($totalCapacityB -gt 0) { [math]::Round(($totalUsedB / $totalCapacityB) * 100, 1) } else { 0 }

# ------------------------------------------------------------------
# 2c. All configured jobs
# ------------------------------------------------------------------

Write-Log "Collecting configured jobs list..."

$configuredJobs = Get-VBRJob -WarningAction SilentlyContinue | Sort-Object Name | ForEach-Object {
    $j = $_
    $lastStatus = $null
    try {
        if ($j.Info -and $j.Info.LatestStatus -and $j.Info.LatestStatus.ToString() -ne "None") {
            $lastStatus = $j.Info.LatestStatus.ToString()
        }
    }
    catch { }

    $enabled = $true
    try { $enabled = [bool]$j.IsScheduleEnabled } catch { }

    [PSCustomObject]@{
        Name       = $j.Name
        Type       = $j.JobType.ToString()
        Enabled    = $enabled
        LastStatus = $lastStatus
    }
}

# ------------------------------------------------------------------
# 3. Build summary stats
# ------------------------------------------------------------------

$totalSessions   = $allSessions.Count
$successSessions = ($allSessions | Where-Object { $_.Result -eq "Success" }).Count
$warningSessions = ($allSessions | Where-Object { $_.Result -eq "Warning" }).Count
$failedSessions  = ($allSessions | Where-Object { $_.Result -eq "Failed"  }).Count

$successRate = if ($totalSessions -gt 0) { [math]::Round(($successSessions / $totalSessions) * 100, 1) } else { 0 }

$totalTransferred = ($allSessions | ForEach-Object { $_.Progress.TransferedSize } | Measure-Object -Sum).Sum
$avgSpeedBps       = ($allSessions | Where-Object { $_.Progress.AvgSpeed -gt 0 } | ForEach-Object { $_.Progress.AvgSpeed } | Measure-Object -Average).Average

# ------------------------------------------------------------------
# 4. Per-job (session) rows
# ------------------------------------------------------------------

$jobRows = $allSessions | Sort-Object CreationTime -Descending | ForEach-Object {
    $s = $_
    $duration = $null
    if ($s.EndTime -and $s.CreationTime -and $s.EndTime -gt $s.CreationTime) {
        $duration = $s.EndTime - $s.CreationTime
    }

    [PSCustomObject]@{
        JobName     = $s.JobName
        JobType     = $s.JobTypeString
        Start       = $s.CreationTime
        End         = $s.EndTime
        Duration    = $duration
        Result      = $s.Result.ToString()
        Transferred = $s.Progress.TransferedSize
        Speed       = $s.Progress.AvgSpeed
    }
}

# ------------------------------------------------------------------
# 5. Per-machine rows
# ------------------------------------------------------------------

Write-Log "Collecting per-machine task details..."

$machineRows = foreach ($s in $allSessions) {
    $tasks = Get-VBRTaskSession -Session $s -ErrorAction SilentlyContinue
    foreach ($t in $tasks) {
        [PSCustomObject]@{
            Machine     = $t.Name
            JobName     = $s.JobName
            Start       = Get-TaskStartTime -Task $t -FallbackSession $s
            End         = $t.Progress.StopTime
            Status      = $t.Status.ToString()
            ProcessedGB = if ($t.Progress.ProcessedSize) { $t.Progress.ProcessedSize } else { 0 }
            Transferred = if ($t.Progress.TransferedSize) { $t.Progress.TransferedSize } else { 0 }
            Speed       = $t.Progress.AvgSpeed
        }
    }
}

$machineSummary = $machineRows | Group-Object Machine | ForEach-Object {
    $rows   = $_.Group | Sort-Object Start -Descending
    $latest = $rows | Select-Object -First 1
    $fails  = ($rows | Where-Object { $_.Status -match "Fail" }).Count
    $warns  = ($rows | Where-Object { $_.Status -match "Warning" }).Count

    [PSCustomObject]@{
        Machine       = $_.Name
        LastJob       = $latest.JobName
        LastBackup    = $latest.Start
        LastStatus    = $latest.Status
        LastSize      = $latest.Transferred
        LastSpeed     = $latest.Speed
        RunsInWindow  = $rows.Count
        FailuresCount = $fails
        WarningsCount = $warns
    }
} | Sort-Object Machine

# ------------------------------------------------------------------
# 5b. Daily trend + failed-job callout list
# ------------------------------------------------------------------

$dayBuckets = for ($i = $ReportDays - 1; $i -ge 0; $i--) {
    $day = (Get-Date).AddDays(-$i).Date
    $dayCandidates = $allSessions | Where-Object { $_.CreationTime.Date -eq $day }
    [PSCustomObject]@{
        Date    = $day
        Success = ($dayCandidates | Where-Object { $_.Result -eq "Success" }).Count
        Warning = ($dayCandidates | Where-Object { $_.Result -eq "Warning" }).Count
        Failed  = ($dayCandidates | Where-Object { $_.Result -eq "Failed"  }).Count
    }
}
$dayMax = [math]::Max((($dayBuckets | ForEach-Object { $_.Success + $_.Warning + $_.Failed } | Measure-Object -Maximum).Maximum), 1)

$failedJobsList = $jobRows | Where-Object { $_.Result -eq "Failed" } | Select-Object -First 8

# --- Trigger CSS Grid class if report days are greater than 15 ---
$trendLayoutClass = if ($ReportDays -gt 15) { "trend-grid" } else { "" }

# ------------------------------------------------------------------
# 6. Build HTML
# ------------------------------------------------------------------

Write-Log "Building HTML report..."

function Get-StatusBadge {
    param([string]$Status)
    switch -Regex ($Status) {
        "Success" { return '<span class="badge badge-success">&#10003; Success</span>' }
        "Warning" { return '<span class="badge badge-warning">&#9888; Warning</span>' }
        "Fail"    { return '<span class="badge badge-failed">&#10007; Failed</span>' }
        default   { return "<span class=`"badge`">$(HtmlEncode $Status)</span>" }
    }
}

$reportGenerated = Format-DT (Get-Date)
$periodText = "{0:dd/MM/yyyy} to {1:dd/MM/yyyy}" -f $startDate, (Get-Date)

$jobRowsHtml = ($jobRows | ForEach-Object {
    @"
        <tr>
          <td>$(HtmlEncode $_.JobName)</td>
          <td>$(HtmlEncode $_.JobType)</td>
          <td>$(Format-DT $_.Start)</td>
          <td>$(Format-Duration $_.Duration)</td>
          <td>$(Get-StatusBadge $_.Result)</td>
          <td class="num">$(Format-Bytes $_.Transferred)</td>
          <td class="num">$(Format-Speed $_.Speed)</td>
        </tr>
"@
}) -join "`n"

if (-not $jobRowsHtml) {
    $jobRowsHtml = '<tr><td colspan="7" class="empty">No backup sessions in this period.</td></tr>'
}

$machineRowsHtml = ($machineSummary | ForEach-Object {
    @"
        <tr>
          <td>$(HtmlEncode $_.Machine)</td>
          <td>$(HtmlEncode $_.LastJob)</td>
          <td>$(Format-DT $_.LastBackup)</td>
          <td>$(Get-StatusBadge $_.LastStatus)</td>
          <td class="num">$(Format-Bytes $_.LastSize)</td>
          <td class="num">$(Format-Speed $_.LastSpeed)</td>
          <td class="num">$($_.RunsInWindow)</td>
          <td class="num">$($_.FailuresCount)</td>
        </tr>
"@
}) -join "`n"

if (-not $machineRowsHtml) {
    $machineRowsHtml = '<tr><td colspan="8" class="empty">No machine data in this period.</td></tr>'
}

# --- Donut chart angle calculations (Fixed inline spacing) ---
$donutDeg = [math]::Round(3.6 * $successRate, 1)
$donutColor = if ($successRate -ge 95) { "#1e8e5a" } elseif ($successRate -ge 80) { "#c98a1f" } else { "#c9432f" }

# --- Daily trend mini bar chart -------------------------------------
$dayBarsHtml = ($dayBuckets | ForEach-Object {
    $total = $_.Success + $_.Warning + $_.Failed
    $hS = if ($total -gt 0) { [math]::Round(($_.Success / $dayMax) * 100, 1) } else { 0 }
    $hW = if ($total -gt 0) { [math]::Round(($_.Warning / $dayMax) * 100, 1) } else { 0 }
    $hF = if ($total -gt 0) { [math]::Round(($_.Failed  / $dayMax) * 100, 1) } else { 0 }
    $dLabel = $_.Date.ToString("ddd d")
    @"
      <div class="daybar" title="$($_.Date.ToString('dd/MM/yyyy')): $($_.Success) success, $($_.Warning) warning, $($_.Failed) failed">
        <div class="daybar-stack">
          <div class="seg seg-failed"  style="height:$($hF)%"></div>
          <div class="seg seg-warning" style="height:$($hW)%"></div>
          <div class="seg seg-success" style="height:$($hS)%"></div>
        </div>
        <div class="daybar-label">$dLabel</div>
      </div>
"@
}) -join "`n"

# --- Storage overview donut ------------------------------------------
$storageDonutDeg   = [math]::Round(3.6 * $overallUsedPct, 1)
$storageDonutColor = Get-UsageColor -Pct $overallUsedPct

# --- Per-repository usage bars -----------------------------------------
$repoBarsHtml = ($repoRows | ForEach-Object {
    $color = Get-UsageColor -Pct $_.UsedPct
    $totalText = if ($_.TotalB -gt 0) { Format-Bytes $_.TotalB } else { "unknown" }
    @"
    <div class="repo-row">
      <div class="repo-head">
        <span class="repo-name">$(HtmlEncode $_.Name)</span>
        <span class="repo-type">$(HtmlEncode $_.Type)</span>
        <span class="repo-stats">$(Format-Bytes $_.UsedB) / $totalText &nbsp;&bull;&nbsp; $($_.UsedPct)% used</span>
      </div>
      <div class="repo-bar"><div class="repo-bar-fill" style="width:$($_.UsedPct)%;background:$color;"></div></div>
    </div>
"@
}) -join "`n"

if (-not $repoBarsHtml) {
    $repoBarsHtml = '<div class="empty">No backup repositories found.</div>'
}

# --- Configured jobs list -------------------------------------------
$jobListHtml = ($configuredJobs | ForEach-Object {
    $stateTag = if ($_.Enabled) { '<span class="job-type job-enabled">Enabled</span>' } else { '<span class="job-type job-disabled">Disabled</span>' }
    @"
      <div class="job-item">
        <span class="job-name">$(HtmlEncode $_.Name)</span>
        <span class="job-type">$(HtmlEncode $_.Type)</span>
        $stateTag
      </div>
"@
}) -join "`n"

if (-not $jobListHtml) {
    $jobListHtml = '<div class="empty">No jobs configured on this server.</div>'
}

# --- Failed-job callout ----------------------------------------------
$failedCalloutHtml = ""
if ($failedJobsList -and $failedJobsList.Count -gt 0) {
    $failedItems = ($failedJobsList | ForEach-Object {
        "<li><strong>$(HtmlEncode $_.JobName)</strong> &mdash; $(Format-DT $_.Start) <span class='muted'>($(HtmlEncode $_.JobType))</span></li>"
    }) -join "`n"

    $failedCalloutHtml = @"
  <div class="callout callout-failed">
    <div class="callout-title">&#9888; Attention needed &ndash; $($failedJobsList.Count) failed job(s) in this period</div>
    <ul>
      $failedItems
    </ul>
  </div>
"@
}

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$(HtmlEncode $CustomerName) - Backup Report</title>
<style>
  * { box-sizing: border-box; }
  body { font-family: 'Segoe UI', Arial, sans-serif; background:#eef1f5; margin:0; padding:0; color:#26313c; }
  .wrapper { max-width:1040px; margin:0 auto; padding:24px; }

  .header { background:linear-gradient(135deg,#16324f 0%,#1c4a75 100%); color:#fff; padding:28px 32px; border-radius:10px; position:relative; overflow:hidden; }
  .header h1 { margin:0 0 6px 0; font-size:24px; font-weight:600; }
  .header p { margin:0; color:#cfe0ee; font-size:13px; }
  .header .tag { display:inline-block; margin-top:12px; background:rgba(255,255,255,0.14); border-radius:20px; padding:4px 12px; font-size:12px; }

  .overview { display:flex; gap:16px; margin:20px 0; flex-wrap:wrap; align-items:stretch; }
  .donut-card { background:#fff; border-radius:10px; box-shadow:0 1px 4px rgba(20,30,40,0.08); padding:18px 22px; display:flex; align-items:center; gap:18px; min-width:250px; }
  
  /* Fixed Donut CSS Bugs (ensured continuous $($variable)deg string format) */
  .donut { width:92px; height:92px; border-radius:50%; flex-shrink:0; background:conic-gradient($donutColor 0deg $($donutDeg)deg, #e7ebee $($donutDeg)deg 360deg); display:flex; align-items:center; justify-content:center; }
  .donut-inner { width:66px; height:66px; border-radius:50%; background:#fff; display:flex; align-items:center; justify-content:center; flex-direction:column; }
  .donut-inner .pct { font-size:18px; font-weight:700; color:$donutColor; }
  .donut-inner .lbl { font-size:9px; color:#8a97a1; text-transform:uppercase; }
  .donut-text .big { font-size:14px; font-weight:600; }
  .donut-text .small { font-size:12px; color:#6b7a86; margin-top:2px; }

  .trend-card { background:#fff; border-radius:10px; box-shadow:0 1px 4px rgba(20,30,40,0.08); padding:18px 22px; flex:1; min-width:280px; }
  .trend-card h3 { margin:0 0 12px 0; font-size:13px; color:#3c4b57; text-transform:uppercase; letter-spacing:.03em; }
  
  /* Daily Trend Layout Area */
  .daychart { display:flex; align-items:flex-end; gap:8px; height:80px; }
  
  /* CSS Grid Layout triggered when days exceed 15 */
  .daychart.trend-grid { display: grid; grid-template-columns: repeat(15, minmax(0, 1fr)); gap: 16px 8px; height: auto; }
  .daychart.trend-grid .daybar { height: auto; }

  .daybar { display:flex; flex-direction:column; align-items:center; flex:1; height:100%; justify-content:flex-end; }
  .daybar-stack { width:100%; max-width:22px; height:64px; display:flex; flex-direction:column-reverse; border-radius:3px; overflow:hidden; background:#f0f3f5; }
  .seg { width:100%; }
  .seg-success { background:#2fa870; }
  .seg-warning { background:#e0a63e; }
  .seg-failed  { background:#dd5847; }
  .daybar-label { font-size:10px; color:#8a97a1; margin-top:6px; white-space: nowrap; }

  .card-row { display:flex; gap:12px; margin:0 0 20px 0; flex-wrap:wrap; }
  .card { background:#fff; border-radius:10px; box-shadow:0 1px 4px rgba(20,30,40,0.08); padding:16px 20px; flex:1; min-width:150px; text-align:center; }
  .card .value { font-size:24px; font-weight:700; }
  .card .label { font-size:11px; color:#6b7a86; text-transform:uppercase; letter-spacing:.03em; margin-top:4px; }
  .card.success .value { color:#1e8e5a; }
  .card.warning .value { color:#c98a1f; }
  .card.failed .value  { color:#c9432f; }

  .callout { border-radius:10px; padding:16px 20px; margin-bottom:20px; }
  .callout-failed { background:#fdf1ef; border:1px solid #f3c8c0; }
  .callout-title { font-weight:700; color:#a83226; margin-bottom:8px; font-size:13px; }
  .callout ul { margin:0; padding-left:20px; font-size:13px; color:#5c3a34; }
  .callout li { margin-bottom:4px; }
  .callout .muted { color:#8a736d; }

  .section { background:#fff; border-radius:10px; box-shadow:0 1px 4px rgba(20,30,40,0.08); padding:20px; margin-bottom:20px; }
  .section h2 { font-size:15px; margin:0 0 14px 0; color:#1c2b3a; display:flex; align-items:center; gap:8px; }
  .section h2::before { content:""; width:4px; height:16px; background:#1c4a75; border-radius:2px; display:inline-block; }

  table { width:100%; border-collapse:collapse; font-size:13px; }
  th { text-align:left; background:#f4f7f9; padding:9px 10px; font-weight:600; color:#3c4b57; font-size:11px; text-transform:uppercase; letter-spacing:.02em; }
  td { padding:9px 10px; border-bottom:1px solid #eef1f3; }
  td.num { text-align:right; font-variant-numeric:tabular-nums; }
  tbody tr:nth-child(even) td { background:#fafbfc; }
  tbody tr:hover td { background:#f0f6fa; }

  .badge { display:inline-block; padding:3px 10px; border-radius:12px; font-size:11px; font-weight:600; white-space:nowrap; }
  .badge-success { background:#e3f5ea; color:#1e8e5a; }
  .badge-warning { background:#fdf2df; color:#c98a1f; }
  .badge-failed  { background:#fbe6e3; color:#c9432f; }

  .job-list { display:flex; flex-direction:column; gap:2px; }
  .job-item { display:flex; align-items:center; gap:10px; padding:10px 12px; border-radius:6px; }
  .job-item:nth-child(even) { background:#fafbfc; }
  .job-item:hover { background:#f0f6fa; }
  .job-name { font-weight:600; font-size:13px; color:#1c2b3a; }
  .job-type { font-size:11px; color:#8a97a1; background:#f0f3f5; border-radius:8px; padding:1px 8px; white-space:nowrap; }
  .job-enabled { color:#1e8e5a; background:#e3f5ea; }
  .job-disabled { color:#8a97a1; background:#f0f3f5; }

  .repo-row { margin-bottom:14px; }
  .repo-row:last-child { margin-bottom:0; }
  .repo-head { display:flex; flex-wrap:wrap; align-items:baseline; gap:8px; margin-bottom:5px; font-size:13px; }
  .repo-name { font-weight:600; color:#1c2b3a; }
  .repo-type { font-size:11px; color:#8a97a1; background:#f0f3f5; border-radius:8px; padding:1px 8px; }
  .repo-stats { margin-left:auto; font-size:12px; color:#6b7a86; font-variant-numeric:tabular-nums; }
  .repo-bar { width:100%; height:8px; border-radius:4px; background:#eef1f3; overflow:hidden; }
  .repo-bar-fill { height:100%; border-radius:4px; }

  .storage-layout { display:flex; gap:24px; align-items:center; }
  .storage-donut-card { flex:0 0 300px; align-self:flex-start; }
  .repo-list { flex:1; min-width:0; background:#fafbfc; border-radius:10px; padding:16px 18px; }

  .empty { text-align:center; color:#8a97a1; padding:16px; }
  .footer { text-align:center; font-size:12px; color:#8a97a1; padding:16px 0 30px; }

  @media (max-width:600px) {
    .overview { flex-direction:column; }
    .card-row { flex-direction:column; }
    .storage-layout { flex-direction:column; }
    .storage-donut-card { flex:1 1 auto; width:100%; }
  }
</style>
</head>
<body>
<div class="wrapper">

  <div class="header">
    <h1>$(HtmlEncode $CustomerName) &ndash; Backup Report</h1>
    <p>Period: $periodText &nbsp;|&nbsp; Generated: $reportGenerated</p>
    <span class="tag">Last $ReportDays days &nbsp;&bull;&nbsp; $totalSessions job runs</span>
  </div>

  <div class="overview">
    <div class="donut-card">
      <div class="donut"><div class="donut-inner"><div class="pct">$successRate%</div><div class="lbl">Success</div></div></div>
      <div class="donut-text">
        <div class="big">$successSessions of $totalSessions jobs succeeded</div>
        <div class="small">$warningSessions warning &nbsp;&bull;&nbsp; $failedSessions failed</div>
      </div>
    </div>
    <div class="trend-card">
      <h3>Daily result trend</h3>
      <div class="daychart $trendLayoutClass">
        $dayBarsHtml
      </div>
    </div>
  </div>

  <div class="card-row">
    <div class="card success"><div class="value">$successSessions</div><div class="label">Successful</div></div>
    <div class="card warning"><div class="value">$warningSessions</div><div class="label">Warnings</div></div>
    <div class="card failed"><div class="value">$failedSessions</div><div class="label">Failed</div></div>
    <div class="card"><div class="value">$(Format-Bytes $totalTransferred)</div><div class="label">Data Transferred</div></div>
    <div class="card"><div class="value">$(Format-Speed $avgSpeedBps)</div><div class="label">Average Speed</div></div>
  </div>

  <div class="section">
    <h2>Backup Repository Storage</h2>
    <div class="storage-layout">
      <div class="donut-card storage-donut-card">
        <div class="donut" style="background:conic-gradient($storageDonutColor 0deg $($storageDonutDeg)deg, #e7ebee $($storageDonutDeg)deg 360deg);">
          <div class="donut-inner"><div class="pct" style="color:$storageDonutColor;">$overallUsedPct%</div><div class="lbl">Used</div></div>
        </div>
        <div class="donut-text">
          <div class="big">$(Format-Bytes $totalUsedB) used of $(Format-Bytes $totalCapacityB)</div>
          <div class="small">$(Format-Bytes $totalFreeB) free across $($repoRows.Count) repositor$(if ($repoRows.Count -eq 1) { 'y' } else { 'ies' })</div>
        </div>
      </div>
      <div class="repo-list">
        $repoBarsHtml
      </div>
    </div>
  </div>

$failedCalloutHtml

  <div class="section">
    <h2>Configured Jobs ($($configuredJobs.Count))</h2>
    <div class="job-list">
      $jobListHtml
    </div>
  </div>

  <div class="section">
    <h2>Backup Job Sessions (last $ReportDays days)</h2>
    <table>
      <thead>
      <tr>
        <th>Job Name</th><th>Type</th><th>Start Time</th><th>Duration</th>
        <th>Result</th><th class="num">Transferred</th><th class="num">Avg Speed</th>
      </tr>
      </thead>
      <tbody>
      $jobRowsHtml
      </tbody>
    </table>
  </div>

  <div class="section">
    <h2>Protected Machines Summary</h2>
    <table>
      <thead>
      <tr>
        <th>Machine</th><th>Last Job</th><th>Last Backup</th><th>Last Status</th>
        <th class="num">Last Size</th><th class="num">Last Speed</th>
        <th class="num">Runs</th><th class="num">Failures</th>
      </tr>
      </thead>
      <tbody>
      $machineRowsHtml
      </tbody>
    </table>
  </div>

  <div class="footer">
    This is an automated report generated from Veeam Backup &amp; Replication.
  </div>

</div>
</body>
</html>
"@

# ------------------------------------------------------------------
# 7. Save to disk
# ------------------------------------------------------------------

if (-not (Test-Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}

$fileName = "BackupReport_{0}_{1}.html" -f ($CustomerName -replace '[^\w-]', '_'), (Get-Date -Format "yyyyMMdd_HHmmss")
$outputPath = Join-Path $OutputFolder $fileName
$html | Out-File -FilePath $outputPath -Encoding UTF8

Write-Log "Report saved to $outputPath"

# ------------------------------------------------------------------
# 8. Convert HTML to PDF
# ------------------------------------------------------------------

if (-not $SkipPdf) {
    $pdfPath = [System.IO.Path]::ChangeExtension($outputPath, '.pdf')
    Write-Log "Converting report to PDF..."
    if (Convert-HtmlFileToPdf -HtmlPath $outputPath -PdfPath $pdfPath) {
        Write-Log "PDF saved to $pdfPath"
    }
}

Write-Log "Done."
