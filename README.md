# Internship Scanner

This scanner reads company names from column A of `JOB APPLICATION TEMPLATE.xlsx`, checks live official ATS feeds where available, and writes a timestamped workbook with active internship results.

## Run It Manually

From PowerShell:

```powershell
& "C:\Users\chris\Documents\Codex\2026-06-23\files-mentioned-by-the-user-job\outputs\summer_2027_scanner\run_daily_scan.ps1"
```

Outputs are saved here:

```text
C:\Users\chris\Documents\Codex\2026-06-23\files-mentioned-by-the-user-job\outputs\daily_scans
```

Each run creates:

- A timestamped `.xlsx` tracker copy
- A timestamped `.json` file with raw scan details

## What It Checks

The primary scanner checks official job-board APIs for:

- Greenhouse
- Lever
- Ashby

It marks a company as `Yes` when a live posting title matches an internship pattern, such as:

- `Intern`
- `Internship`
- `Co-op`
- `Summer Analyst`
- `Summer Associate`

Rows without a verified match are marked `No confirmed active posting found`.

The role title and source URL still show whether the internship is specifically Summer 2027. This is intentionally broader than filtering only on `2027`, because some companies post active internship roles without putting the year in the title.

## Schedule It Once Per Day

Open PowerShell and run:

```powershell
$Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument '-NoProfile -ExecutionPolicy Bypass -File "C:\Users\chris\Documents\Codex\2026-06-23\files-mentioned-by-the-user-job\outputs\summer_2027_scanner\run_daily_scan.ps1"'
$Trigger = New-ScheduledTaskTrigger -Daily -At 9am
Register-ScheduledTask -TaskName "Internship Scan" -Action $Action -Trigger $Trigger -Description "Scan tracked companies for active internship postings."
```

To run at a different time, change `9am`.

## Update The Company List

Edit `C:\Users\chris\Downloads\JOB APPLICATION TEMPLATE.xlsx`. The scanner reads the current list from column A each time it runs.

## Known Limits

Some companies use Workday, custom careers sites, or blocked search pages that do not expose a simple public API. The scanner flags those as unverified unless an official posting URL is added to the manual official postings list in `summer_2027_scanner.py`.
