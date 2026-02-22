# HPE Support CaseBot (Playwright)

Automates an overview export of your HPE Support Center cases and writes:

- `out_hpe/cases_overview.json`
- `out_hpe/cases_overview.csv`

It supports headless execution and can run fully hidden via Windows Task Scheduler.

> Disclaimer  
> This project is not affiliated with HPE. Use at your own risk and in accordance with your organization's security policy and HPE terms.

---

## Requirements

- Windows 10/11 or Windows Server
- PowerShell 5.1+
- Python 3.12+
- Internet access to `support.hpe.com`

---

## Install

```bat
cd /d C:\HPESUPBOT
install_me.cmd
```

---

## Authentication (choose one)

### Option A - Session state (recommended)
Creates/updates `hpe_state.json` via an interactive login:

```bat
01_Login.cmd
```

### Option B - Auto-login (username/password, no MFA)
Stores HPE credentials encrypted with DPAPI for the current Windows user:

```bat
02_SetHPECredentials.cmd
```

Creates `hpe_credential.xml`.

DPAPI note: Scheduled Task must run under the same Windows user that created `hpe_credential.xml`.

---

## Run

```bat
Run-HPECaseBot.cmd
```

Outputs:
- `out_hpe/cases_overview.json`
- `out_hpe/cases_overview.csv`
- `out_hpe/run.log`

---

## How login works (state vs auto-login)

1) If `hpe_state.json` exists and is valid:
- Uses it (fast)
- Scrapes cases
- Refreshes `hpe_state.json` after success

2) If `hpe_state.json` is missing/invalid/expired:
- If DPAPI creds exist (`hpe_credential.xml`) or `HPE_USERNAME/HPE_PASSWORD` are provided:
  - Auto-login (user/pass)
  - Rebuilds `hpe_state.json`
  - Continues scraping
- Otherwise: interactive login required (`01_Login.cmd`)

---

## Scheduling (Windows Task Scheduler)

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\03_CreateTask_10min.ps1
```

Recommended:
- Run as same Windows user used for `02_SetHPECredentials.cmd`
- Do not start a new instance
- Start in: `C:\HPESUPBOT`

---

## Security
Do not commit:
- `hpe_state.json`
- `hpe_credential.xml`
- `out_hpe/`, `out_hpe_archive/`
