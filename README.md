# Veeam Backup & Replication HTML&PDF Backup Reporter

A PowerShell script that runs locally on your Veeam Backup & Replication server and generates a polished, self-contained HTML report covering the last N days (default 7) of backup activity — job results, speeds, protected machines, and repository storage usage. It can save the report to disk.

---

## What's in the report

- **Success-rate donut** — overall % of jobs that succeeded in the report window
- **7-day trend chart** — stacked bar per day showing success / warning / failed counts
- **Attention-needed callout** — auto-appears only when there are failed jobs, listing them at the top
- **Configured Jobs** — every job defined on the server (name, type, Enabled/Disabled), independent of the report window
- **Backup Job Sessions** — full table of every session that ran in the window (start time, duration, result, data transferred, average speed)
- **Protected Machines Summary** — per-machine rollup (last job, last backup time, last status, last size/speed, run count, failure count)
- **Backup Repository Storage** — used/free/total capacity per repository (including Scale-Out repositories where available), plus an overall usage donut

All dates/times are rendered as `DD/MM/YYYY HH:mm:ss` regardless of the server's regional settings.

---

## Requirements

- Windows Server 2019 or newer
- Must run **locally on the Veeam Backup & Replication server** (it connects to `localhost`)
- Veeam Backup & Replication PowerShell module (installed automatically with the Veeam console — no separate install needed)

---

## Basic usage

Generate the report only, no email — good for a scheduled task:

```powershell
.\VeeamReport7Days.ps1 -CustomerName "CustomerName Ltd." -OutputFolder "D:\Reports"
```

---

## Parameters

| Parameter | Default | Description |
|---|---|---|
| `-ReportDays` | `7` | How many days back to include |
| `-CustomerName` | `"Customer"` | Friendly name shown in the report header and used in the output filename |
| `-OutputFolder` | `C:\VeeamReports` | Where the generated `.html` file is saved (created automatically if missing) |

---

## Output

Each run creates a uniquely named file so old reports are never overwritten:

```
<OutputFolder>\BackupReport_<CustomerName>_<yyyyMMdd_HHmmss>.html
```

Example: `C:\VeeamReports\BackupReport_CustomerName_Ltd_20260714_091238.html`

---

## Scheduling it

Run it automatically with Windows Task Scheduler:

1. Create a new task, trigger it daily/weekly as needed.
2. Action: `powershell.exe`
3. Arguments:
   ```
   -NoProfile -ExecutionPolicy Bypass -File "C:\Path\To\VeeamReport7Days.ps1" -CustomerName "CustomerName Ltd."
   ```
4. Run as a service account with local admin rights on the Veeam server (needed for the Veeam PowerShell module).

For multiple customers on one Veeam server, run the script once per customer with different `-CustomerName` values.

---

## Known quirks / compatibility notes

These came up during testing and are already handled in the script, documented here so you know what's going on if you see similar log lines:

- **`Get-VBRScaleOutBackupRepository not available`** — this cmdlet only exists if your Veeam edition/version supports Scale-Out Backup Repositories. The script checks for it first and simply skips SOBR collection if it's not present. Your regular repositories are unaffected.
- **Repository capacity numbers** — Veeam returns `CachedTotalSpace`/`CachedFreeSpace` as a `VMemorySize` object, not a plain number. The script unwraps this safely (`Get-BytesValue`); if you ever see `0 B` for a repository, check the console output for a `WARN` line — some repository types (e.g. cloud/object storage tiers) don't report a fixed total size.
- **Per-machine "Last Backup" time** — some Veeam versions/job types don't populate the task-level start time consistently. The script tries several known property paths and falls back to the parent job session's start time if all else fails, so this column should never be blank.

---

## Customizing

The report's colors, fonts, and layout are all in the `<style>` block inside the script (search for `$html = @"`). Common tweaks:

- **Company logo**: add an `<img>` tag inside the `.header` div, pointing to a hosted image URL (email clients generally won't render locally embedded images without extra work).
- **Color thresholds**: `Get-UsageColor` (storage) and the inline success-rate logic near `$donutColor` control the green/amber/red cutoffs.
- **Which jobs appear**: currently all jobs/sessions are included. To scope by customer, filter `$allSessions` and `$configuredJobs` by job name pattern (e.g. `Where-Object { $_.JobName -like "CustomerName*" }`).

---

## Troubleshooting

If the script errors out immediately on `Import-Module` or `Connect-VBRServer`, double check:
- It's running **on the Veeam B&R server itself**, not a remote machine
- The account running it has permissions in Veeam (local admin or a Veeam-assigned role)
- The Veeam Backup & Replication console/PowerShell module is actually installed on that machine
