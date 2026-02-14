# HPE Support CaseBot (Playwright)

Automates an overview of your **HPE Support Center** cases and exports:
- `cases_overview.json`
- `cases_overview.csv`
- per-case **Communications** export (redacted for obvious tokens/passwords)

It can run **headless + fully hidden** via Windows Scheduled Task (VBS wrapper).

> Not affiliated with HPE. Use at your own risk and respect your organization’s security policy and HPE terms.

---

## Table of contents

- [English](#english)
- [Nederlands](#nederlands)
- [Security notes (important)](#security-notes-important)
- [Repository hygiene (.gitignore)](#repository-hygiene-gitignore)

---

## English

### What it does

1. Uses an **interactive login once** to save your authenticated session into `hpe_state.json`.
2. Opens your HPE **Cases** page and collects the case numbers that are **visible for your account** (and your current “Group” scope in the portal).
3. For each case:
   - reads **Details**
   - reads **Communications**
   - extracts/normalizes fields (status, severity, product, serial, etc.)
   - infers an **action category** + suggested actions (example: `CLOSE_APPROVAL`, `LOG_REQUEST`)
4. Writes exports to `out_hpe\` and a lightweight status file for monitoring.

### Requirements

- Windows 10/11 or Windows Server
- PowerShell **5.1+**
- Python **3.12** (the installer can use `py` launcher or `python`)
- Internet access to `support.hpe.com`
- You must be able to open a browser once to complete MFA (for `hpe_state.json`)

### Quick start (recommended)

1) Unzip the package somewhere (example `C:\HPE_CaseBot`)

2) Run the interactive installer:
```bat
cd /d C:\HPE_CaseBot
install_me.cmd
```

The installer will:
- unblock files
- create `.venv` and install dependencies
- install Playwright Chromium
- prompt you to login and create/update `hpe_state.json`
- optionally create a **10-minute hidden** Scheduled Task
- optionally run a quick headless test run

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

### Scheduled Task (hidden + headless)

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

### Outputs

Main outputs in `out_hpe\`:

- `cases_overview.json` – full structured output
- `cases_overview.csv` – same overview as CSV
- `cases\<CASE>_communications_redacted.txt` – communications text (with obvious `Password/Token` values redacted)
- `nagios\hpe_casebot.status` – small status JSON (OK/CRITICAL + timestamp)
- `run.log` – wrapper log

Archives (unless `-NoArchive`): `out_hpe_archive\YYYY-MM-DD\...`

### How filtering works

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

### Monitoring in Nagios

You have two practical monitoring options:

#### Option A — “Is it still running?” (file age)
Monitor the age of:
- `out_hpe\nagios\hpe_casebot.status` **or**
- `out_hpe\cases_overview.json`

**Recommended thresholds**
- WARNING if older than **30 minutes**
- CRITICAL if older than **2 hours**

#### Option B — “Is it healthy + how many cases need action?”
Parse `cases_overview.json`:
- CRITICAL if `errors[]` is not empty
- WARNING/CRITICAL based on number of cases in `status` “Awaiting Customer Action”
- optionally: alert on certain inferred categories (`LOG_REQUEST`, `CLOSE_APPROVAL`)

##### Example PowerShell Nagios plugin (runs on Windows)
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

# Status by age first
if ($ageM -ge $CritAgeMinutes) { Write-Host "CRITICAL - stale ($ageM min) | age_min=$ageM errors=$errCount cases=$caseCount awaiting=$awaiting"; exit 2 }
if ($ageM -ge $WarnAgeMinutes) { Write-Host "WARNING  - stale ($ageM min) | age_min=$ageM errors=$errCount cases=$caseCount awaiting=$awaiting"; exit 1 }

# Errors from bot
if ($errCount -gt 0) { Write-Host "CRITICAL - bot errors=$errCount | age_min=$ageM errors=$errCount cases=$caseCount awaiting=$awaiting"; exit 2 }

# Case thresholds
if ($awaiting -ge $CritCases) { Write-Host "CRITICAL - awaiting=$awaiting | age_min=$ageM errors=$errCount cases=$caseCount awaiting=$awaiting"; exit 2 }
if ($awaiting -ge $WarnCases) { Write-Host "WARNING  - awaiting=$awaiting | age_min=$ageM errors=$errCount cases=$caseCount awaiting=$awaiting"; exit 1 }

Write-Host "OK - cases=$caseCount awaiting=$awaiting | age_min=$ageM errors=$errCount cases=$caseCount awaiting=$awaiting"
exit 0
```

##### Hooking into Nagios (common patterns)
- **NSClient++**: execute the PowerShell script remotely from Nagios
- **NRPE on Windows**: run a command that executes PowerShell
- **check_by_ssh**: if OpenSSH is enabled on the Windows host

Because every Nagios environment differs, the universal rule is:
- execute the script on the Windows host
- return standard exit codes: `0 OK`, `1 WARNING`, `2 CRITICAL`, `3 UNKNOWN`

### Session expiry / re-login

If HPE expires your session, runs will start failing until you re-create `hpe_state.json`:

```powershell
powershell -ExecutionPolicy Bypass -File .\01_Login.ps1
```

You can also run Python directly and use its built-in alarm file:
```powershell
.\.venv\Scripts\python.exe .\hpe_cases_overview.py --headless --alarm-file ALERT_SESSION_EXPIRED.txt
```

### Making it work in another language (international / universal)

The easiest approach: set your HPE portal language to **English**.  
If you must use another UI language, these parts may need adjustments:

1) **Page readiness / session expired detection**  
Edit `hpe_selectors.json`:
- `session_expired_text_any`
- `ready_text_any`
Add the translated strings you see in your portal.

2) **Field label mapping**  
Edit `hpe_cases_overview.py` → `FIELD_LABELS`:
- add translated labels for `status`, `severity`, `product`, `serial`, etc.

3) **Action inference (category + suggested actions)**  
Edit `hpe_cases_overview.py` → `infer_requested_actions()`:
- it checks status text (“Approve case closure”, “Complete action plan”, …)
- it checks communications keywords (“log file request”, “AHS log”, …)
Add keywords in your language and update the output strings.

> Tip for maintainability: move the keywords and output strings into a JSON config if you plan to support many languages.

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

### Handmatig

```powershell
powershell -ExecutionPolicy Bypass -File .\00_Setup.ps1
powershell -ExecutionPolicy Bypass -File .\01_Login.ps1
powershell -ExecutionPolicy Bypass -File .\Run-HPECaseBot.ps1 -Headless
```

### Scheduled Task (hidden + headless)

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\03_CreateTask_10min.ps1 -RootPath (Get-Location) -Minutes 10
```

Verwijderen:
```powershell
schtasks /Delete /TN "HPE CaseBot (10min)" /F
```

### Output bestanden (kort)

- `out_hpe\cases_overview.json` (alles)
- `out_hpe\cases_overview.csv` (overzicht)
- `out_hpe\cases\<CASE>_communications_redacted.txt` (communicatie, tokens/passwords geredact)
- `out_hpe\nagios\hpe_casebot.status` (OK/CRITICAL + timestamp)
- `out_hpe\run.log` (wrapper log)
- `out_hpe_archive\YYYY-MM-DD\...` (archief)

### Nagios monitoring

**Minimum**: check “file age” van `cases_overview.json` of `nagios\hpe_casebot.status`.

**Beter**: parse `cases_overview.json` en alert op:
- `errors[]` niet leeg → CRITICAL
- aantal “Awaiting Customer Action” → WARNING/CRITICAL

Zie het PowerShell voorbeeld hierboven (kan 1-op-1 gebruikt worden).

### Andere taal gebruiken (universeel maken)

- Voeg vertalingen toe in `hpe_selectors.json` (ready/session expired)
- Voeg vertaalde labels toe in `FIELD_LABELS` (Python)
- Voeg vertaalde keywords toe in `infer_requested_actions()` (Python)

---

## Security notes (important)

- **Treat `hpe_state.json` like a password.** It contains session cookies.
- Do **not** commit anything from `out_hpe\` or `out_hpe_archive\` to GitHub (can contain serials, addresses, case data).
- `cases\*_communications_redacted.txt` redacts obvious tokens/passwords, but **does not remove all sensitive information**.

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

# Debug dumps
**/debug/
ALERT_SESSION_EXPIRED.txt

# OS
.DS_Store
Thumbs.db
```
