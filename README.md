# HPE Support CaseBot (Playwright)

Automates an overview of your **HPE Support Center** cases and exports:
- `cases_overview.json`
- `cases_overview.csv`
- Per-case **Communications** export (redacted for obvious tokens/passwords)

Supports **headless** operation and can run **fully hidden** via **Windows Scheduled Task** (VBS wrapper).

> **Disclaimer:** Not affiliated with HPE. Use at your own risk and always respect your organization’s security policy and the applicable HPE terms.

---

## Table of contents
- [English](#english)
- [Nederlands](#nederlands)
- [Security notes (important)](#security-notes-important)
- [Repository hygiene (.gitignore)](#repository-hygiene-gitignore)

---

## English

### What it does
This bot logs into the HPE Support portal and generates a clean overview of open cases, including:
- case metadata (status, severity, product, etc.)
- action-plan related signals (e.g., “Awaiting Customer Action”)
- a per-case communications export (with redaction applied)

### Example run

```powershell
c:\HPESUPBOT> Run-HPECaseBot.cmd
Output (example):

Open cases: https://support.hpe.com/connect/s/?tab=cases

Found cases: 2 -> 540XXXXXXX, 540XXXXXXX (MASKED FOR PRIVACY)

=== [1/2] Case 540XXXXXXX ===
OK: 540XXXXXXX | CZ29XXXXGT | Awaiting Customer Action - Complete action plan

=== [2/2] Case 540XXXXXXX ===
OK: 540XXXXXXX | CZ21XXXXGG | Awaiting Customer Action - Approve case closure

CSV:  C:\HPESUPBOT\out_hpe\cases_overview.csv
JSON: C:\HPESUPBOT\out_hpe\cases_overview.json

Done. Cases: 2
Example JSON (trimmed)
{
  "generated_at": "2026-02-16T15:05:46Z",
  "cases": [
    {
      "case_no": "CASE-001",
      "serial": "SERIAL-001",
      "status": "Awaiting Customer Action - Complete action plan",
      "severity": "3-No Business Impact",
      "product": "Server (model redacted)",
      "product_no": "PRODUCTNO-001",
      "group": "GROUP-001",
      "hpe_request_category": "ACTION_PLAN",
      "hpe_request_summary": "HPE is waiting for completion of the action plan.",
      "hpe_requested_actions": "Complete the action plan and confirm in the HPE portal.",
      "comms_file": "REDACTED_PATH/cases/CASE-001_communications_redacted.txt",
      "generated_at": "2026-02-16T15:05:42Z"
    }
  ],
  "errors": []
}
---


### Requirements

- Windows 10/11 or Windows Server
- PowerShell **5.1+**
- Python **3.12** (installer uses `py` launcher or `python`)
- Internet access to `support.hpe.com`
- You must be able to open a browser once to complete MFA (for `hpe_state.json`)

### Quick start (recommended)

1) Unzip the package somewhere (example `C:\HPE_CaseBot`)

2) Run the interactive installer:
cd /d C:\HPE_CaseBot
install_me.cmd

The installer will:
- unblock files
- create `.venv` and install dependencies
- install Playwright Chromium
- prompt you to login and create/update `hpe_state.json`
- optionally create a **10-minute hidden** Scheduled Task
- optionally run a quick headless test run

---

## Usage

### Manual usage

#### 1) Setup venv + Playwright
```powershell
powershell -ExecutionPolicy Bypass -File .\00_Setup.ps1
```

#### 2) Login once (interactive, MFA)
```powershell
powershell -ExecutionPolicy Bypass -File .\01_Login.ps1
```

#### 3) Run once (headless)
```powershell
powershell -ExecutionPolicy Bypass -File .\Run-HPECaseBot.ps1 -Headless
```

### Parameters (important)

Most people should run the **PowerShell wrapper**: `Run-HPECaseBot.ps1`  
It forwards parameters to Python and also handles logging, Nagios status output, and archiving.

#### Run-HPECaseBot.ps1 parameters

- `-Headless`  
  Run Chromium headless (no browser window). **Recommended for Scheduled Task.**

- `-Max <int>`  
  Limit how many cases are processed in this run.  
  Example: `-Max 10` (0 = all cases found)

- `-NoArchive`  
  Disable copying `out_hpe\` into `out_hpe_archive\YYYY-MM-DD\`

- `-Format csv|json|both`  
  Output format for the overview export (default: `both`)

- `-OutDir <path>`  
  Output directory (default: `.\out_hpe`)

**Examples**
```powershell
# headless, process max 10 cases, do NOT archive
powershell -ExecutionPolicy Bypass -File .\Run-HPECaseBot.ps1 -Headless -Max 10 -NoArchive

# JSON only
powershell -ExecutionPolicy Bypass -File .\Run-HPECaseBot.ps1 -Headless -Format json
```

#### Python CLI parameters (advanced)

If you want to run Python directly:
```powershell
.\.venv\Scripts\python.exe .\hpe_cases_overview.py --headless --max 10 --format both
```

Supported args:
- `--state hpe_state.json`
- `--selectors hpe_selectors.json`
- `--outdir out_hpe`
- `--max 10` (0 = all)
- `--headless`
- `--format csv|json|both`
- `--alarm-file ALERT_SESSION_EXPIRED.txt`
- `--alarm-cmd "<command to execute on session expiry>"`

---

## Scheduled Task (hidden + headless)

Create/update the task that runs every 10 minutes:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\03_CreateTask_10min.ps1 -RootPath (Get-Location) -Minutes 10
```

Remove it:
```powershell
schtasks /Delete /TN "HPE CaseBot (10min)" /F
```

**How it stays hidden**
- `03_CreateTask_10min.ps1` writes `Run-HPECaseBot_hidden.vbs`
- The task runs `wscript.exe //B //NoLogo Run-HPECaseBot_hidden.vbs`
- VBS starts PowerShell with `-WindowStyle Hidden` and calls `Run-HPECaseBot.ps1 -Headless`

> Note: Scheduled Tasks created by this package run with an **Interactive token**.  
> That means: it runs for the current user and typically requires the user to be logged on.

---

## Outputs

Main outputs in `out_hpe\`:

- `cases_overview.json` – full structured output
- `cases_overview.csv` – same overview as CSV
- `cases\<CASE>_communications_redacted.txt` – communications text (with obvious `Password/Token` values redacted)
- `nagios\hpe_casebot.status` – small status JSON (OK/CRITICAL + timestamp)
- `run.log` – wrapper log

Archives (unless `-NoArchive`): `out_hpe_archive\YYYY-MM-DD\...`

---

## How filtering works

The bot does **not** hardcode a customer list. It processes **the cases that are visible in your HPE portal** for the logged-in account.

- If your portal shows only “My Group (Private)”, only those cases are collected.
- If you switch groups in the HPE UI, your visible cases change, and so will the bot output.

**Extraction strategy (high-level):**
- Case numbers are collected from the Cases view using a regex like `Case 5401234567`
- Details/Communications are scraped using selectors in `hpe_selectors.json`
- Fields are normalized using label mappings in `hpe_cases_overview.py` (`FIELD_LABELS`)
- Actions/category are inferred from:
  - the case status text
  - keywords in the communications thread

---

## Monitoring in Nagios

You have two practical monitoring options:

### Option A — “Is it still running?” (file age)
Monitor the age of:
- `out_hpe\nagios\hpe_casebot.status` **or**
- `out_hpe\cases_overview.json`

**Recommended thresholds**
- WARNING if older than **30 minutes**
- CRITICAL if older than **2 hours**

### Option B — “Is it healthy + how many cases need action?”
Parse `cases_overview.json`:
- CRITICAL if `errors[]` is not empty
- WARNING/CRITICAL based on number of cases in `status` “Awaiting Customer Action”
- optionally: alert on certain inferred categories (`LOG_REQUEST`, `CLOSE_APPROVAL`)

#### Example PowerShell Nagios plugin (runs on Windows)
Save as `C:\HPE_CaseBot\check_hpe_casebot.ps1`:

```powershell
param(
  [string]$Root = "C:\HPE_CaseBot",
  [int]$WarnAgeMinutes = 30,
  [int]$CritAgeMinutes = 120,
  [int]$WarnCases = 1,
  [int]$CritCases = 5
)

$ErrorActionPreference = "Stop"
$jsonPath = Join-Path $Root "out_hpe\cases_overview.json"
if (!(Test-Path $jsonPath)) { Write-Host "CRITICAL - missing $jsonPath"; exit 2 }

$data = Get-Content $jsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
$gen  = [datetime]::Parse($data.generated_at)
$ageM = [int]((Get-Date).ToUniversalTime() - $gen.ToUniversalTime()).TotalMinutes

$errCount = @($data.errors).Count
$caseCount = @($data.cases).Count
$awaiting  = @($data.cases | Where-Object { $_.status -match "Awaiting Customer Action" }).Count

if ($ageM -ge $CritAgeMinutes) { Write-Host "CRITICAL - stale ($ageM min) | age_min=$ageM errors=$errCount cases=$caseCount awaiting=$awaiting"; exit 2 }
if ($ageM -ge $WarnAgeMinutes) { Write-Host "WARNING  - stale ($ageM min) | age_min=$ageM errors=$errCount cases=$caseCount awaiting=$awaiting"; exit 1 }

if ($errCount -gt 0) { Write-Host "CRITICAL - bot errors=$errCount | age_min=$ageM errors=$errCount cases=$caseCount awaiting=$awaiting"; exit 2 }

if ($awaiting -ge $CritCases) { Write-Host "CRITICAL - awaiting=$awaiting | age_min=$ageM errors=$errCount cases=$caseCount awaiting=$awaiting"; exit 2 }
if ($awaiting -ge $WarnCases) { Write-Host "WARNING  - awaiting=$awaiting | age_min=$ageM errors=$errCount cases=$caseCount awaiting=$awaiting"; exit 1 }

Write-Host "OK - cases=$caseCount awaiting=$awaiting | age_min=$ageM errors=$errCount cases=$caseCount awaiting=$awaiting"
exit 0
```

**Nagios integration notes**
- **NSClient++** / **NRPE** / **check_by_ssh** are common patterns
- Always return standard exit codes: `0 OK`, `1 WARNING`, `2 CRITICAL`, `3 UNKNOWN`

---

## Session expiry / re-login

If HPE expires your session, runs will start failing until you re-create `hpe_state.json`:

```powershell
powershell -ExecutionPolicy Bypass -File .\01_Login.ps1
```

---

## Making it work in another language (international / universal)

The easiest approach: set your HPE portal language to **English**.  
If you must use another UI language, these parts may need adjustments:

1) **Page readiness / session expired detection**  
Edit `hpe_selectors.json`:
- `session_expired_text_any`
- `ready_text_any`

2) **Field label mapping**  
Edit `hpe_cases_overview.py` → `FIELD_LABELS`:
- add translated labels for `status`, `severity`, `product`, `serial`, etc.

3) **Action inference (category + suggested actions)**  
Edit `hpe_cases_overview.py`:
- update keywords used to infer categories and requested actions

> Tip: if you plan many languages, move keywords + output strings into a JSON config.

---

## Nederlands

### Wat doet dit?

1. Je logt **1 keer interactief** in en bewaart je sessie in `hpe_state.json`.
2. De bot opent de **Cases** pagina en neemt de cases die **zichtbaar zijn voor jouw account** (en jouw actieve “Group” in de portal).
3. Per case:
   - leest **Details**
   - leest **Communications**
   - normaliseert velden (status, severity, product, serial, …)
   - bepaalt een **categorie** + “wat moet je doen” (bv. `CLOSE_APPROVAL`, `LOG_REQUEST`)
4. Schrijft alles weg naar `out_hpe\` + een statusfile voor monitoring.

### Vereisten

- Windows 10/11 of Windows Server
- PowerShell **5.1+**
- Python **3.12**
- Internet naar `support.hpe.com`
- Browser login 1 keer (MFA) om `hpe_state.json` te maken

### Snel starten

```bat
cd /d C:\HPE_CaseBot
install_me.cmd
```

---

## Gebruik

### Handmatig

```powershell
powershell -ExecutionPolicy Bypass -File .\00_Setup.ps1
powershell -ExecutionPolicy Bypass -File .\01_Login.ps1
powershell -ExecutionPolicy Bypass -File .\Run-HPECaseBot.ps1 -Headless
```

### Parameters (belangrijk)

#### Run-HPECaseBot.ps1 parameters

- `-Headless`  
  Headless draaien (geen browservenster). **Aanrader voor Scheduled Task.**

- `-Max <int>`  
  Max. aantal cases verwerken in deze run.  
  Voorbeeld: `-Max 10` (0 = alles)

- `-NoArchive`  
  Geen archief kopie maken naar `out_hpe_archive\YYYY-MM-DD\`

- `-Format csv|json|both`  
  Exportformaat (default: `both`)

- `-OutDir <pad>`  
  Output directory (default: `.\out_hpe`)

**Voorbeelden**
```powershell
# headless, max 10 cases, GEEN archief
powershell -ExecutionPolicy Bypass -File .\Run-HPECaseBot.ps1 -Headless -Max 10 -NoArchive

# enkel JSON
powershell -ExecutionPolicy Bypass -File .\Run-HPECaseBot.ps1 -Headless -Format json
```

---

## Scheduled Task (hidden + headless)

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\03_CreateTask_10min.ps1 -RootPath (Get-Location) -Minutes 10
```

Verwijderen:
```powershell
schtasks /Delete /TN "HPE CaseBot (10min)" /F
```

---

## Security notes (important)

- **Treat `hpe_state.json` like a password.** It contains session cookies.
- Commit **never**: `hpe_state.json`, `out_hpe\`, `out_hpe_archive\` (case data/serials/addresses).
- Communications are “redacted”, but **not guaranteed** to remove all sensitive data.

---

## Repository hygiene (.gitignore)

Recommended `.gitignore`:

```gitignore
# Virtual env / dependencies
.venv/

# Runtime output
out_hpe/
out_hpe_archive/
out_hpe/**

# Session cookies
hpe_state.json

# Debug dumps / alarms
**/debug/
ALERT_SESSION_EXPIRED.txt

# OS
.DS_Store
Thumbs.db
```
