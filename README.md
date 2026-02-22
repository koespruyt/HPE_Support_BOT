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
```bat

c:\testhpebot\HPE_Support_BOT>install_me.cmd

[HPE CaseBot] Package location: c:\testhpebot\HPE_Support_BOT

Install folder (default = package folder) >
Using in-place install.
[HPE CaseBot] Package location: C:\testhpebot\HPE_Support_BOT
[HPE CaseBot] Install root:      C:\testhpebot\HPE_Support_BOT
[HPE CaseBot] Unblocking scripts...
[HPE CaseBot] Checking Python...
[HPE CaseBot] Python OK: 3.12
[HPE CaseBot] Running setup (00_Setup.ps1)...
[Setup] Using system python: python
[Setup] Creating venv .venv ...
[Setup] Upgrading pip...
Requirement already satisfied: pip in c:\testhpebot\hpe_support_bot\.venv\lib\site-packages (25.0.1)
Collecting pip
  Using cached pip-26.0.1-py3-none-any.whl.metadata (4.7 kB)
Using cached pip-26.0.1-py3-none-any.whl (1.8 MB)
Installing collected packages: pip
  Attempting uninstall: pip
    Found existing installation: pip 25.0.1
    Uninstalling pip-25.0.1:
      Successfully uninstalled pip-25.0.1
Successfully installed pip-26.0.1
[Setup] Installing requirements...
Collecting playwright>=1.58.0 (from -r C:\testhpebot\HPE_Support_BOT\requirements.txt (line 1))
  Using cached playwright-1.58.0-py3-none-win_amd64.whl.metadata (3.5 kB)
Collecting pyee<14,>=13 (from playwright>=1.58.0->-r C:\testhpebot\HPE_Support_BOT\requirements.txt (line 1))
  Using cached pyee-13.0.1-py3-none-any.whl.metadata (3.0 kB)
Collecting greenlet<4.0.0,>=3.1.1 (from playwright>=1.58.0->-r C:\testhpebot\HPE_Support_BOT\requirements.txt (line 1))
  Using cached greenlet-3.3.2-cp312-cp312-win_amd64.whl.metadata (3.8 kB)
Collecting typing-extensions (from pyee<14,>=13->playwright>=1.58.0->-r C:\testhpebot\HPE_Support_BOT\requirements.txt (line 1))
  Using cached typing_extensions-4.15.0-py3-none-any.whl.metadata (3.3 kB)
Using cached playwright-1.58.0-py3-none-win_amd64.whl (36.8 MB)
Using cached greenlet-3.3.2-cp312-cp312-win_amd64.whl (231 kB)
Using cached pyee-13.0.1-py3-none-any.whl (15 kB)
Using cached typing_extensions-4.15.0-py3-none-any.whl (44 kB)
Installing collected packages: typing-extensions, greenlet, pyee, playwright
Successfully installed greenlet-3.3.2 playwright-1.58.0 pyee-13.0.1 typing-extensions-4.15.0
[Setup] Installing Playwright Chromium...
[Setup] OK (marker: .setup.ok)
Run login now to create/update hpe_state.json? (Y/N): n
Create/update Scheduled Task to run every 10 minutes? (Y/N): y
Minutes interval (default 10): 30
Scheduled Task created/updated: HPE CaseBot (10min)
  Root:     C:\testhpebot\HPE_Support_BOT
  Interval: every 30 minutes (starts ~1 minute from now)
To remove: schtasks /Delete /TN "HPE CaseBot (10min)" /F
Run bot once now (quick test, shows output)? (Y/N): : n
[HPE CaseBot] Install complete.

Next:
  1) Login (interactive):     .\01_Login.cmd   (CMD)  OR  powershell -ExecutionPolicy Bypass -File .\01_Login.ps1
  2) Run bot once:            .\Run-HPECaseBot.cmd   (CMD)  OR  powershell -ExecutionPolicy Bypass -File .\Run-HPECaseBot.ps1
  3) Output folder:           .\out_hpe

```
---

## Run Install: 
Choose NO on "Run login now to create/update hpe_state.json? (Y/N):N"
Choose NO on "Run bot once now (quick test, shows output)? (Y/N):N"

Make sure there is no MFA active on your login account HPE support.
RUN:
```bat
cd /d C:\HPESUPBOT
02_SetHPECredentials.cmd
```
Fill in User / Pass of the HPE account.

Now you may run the BOT

```bat
cd /d C:\HPESUPBOT
Run-HPECaseBot.cmd
```

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

EXAMPLE
```json
{
  "generated_at": "2026-02-22T14:40:XXZ",
  "cases": [
    {
      "case_no": "CASE-001",
      "serial": "SERIAL-001",
      "host_name": "",
      "contact_name": "",
      "addr_street": "",
      "addr_city": "",
      "addr_state": "",
      "addr_postal_code": "",
      "addr_country": "",
      "status": "In Progress",
      "severity": "2-Limited Business Impact",
      "product": "HPE ProLiant DL3xx Gen10 Server",
      "product_no": "PXXXXX-XXX",
      "group": "My Group (Redacted)",
      "action_plan": "",
      "hpe_last_update": "",
      "hpe_last_subject": "",
      "hpe_request_category": "ONSITE_SERVICE",
      "hpe_request_summary": "HPE onsite interventie/dispatch loopt (technieker gepland/onderweg).",
      "hpe_requested_actions": "Check Onsite Service tab voor planning/status (task ID, scheduling status, latest service start). | Zorg dat toegang/site contact klopt; bereid interventie/onderdelen/remote access voor.",
      "hpe_key_links": "",
      "event_ids": "",
      "problem_descriptions": "",
      "ahs_links": "",
      "dropbox_hosts": "",
      "dropbox_logins": "",
      "onsite_detected": "1",
      "onsite_task_ref": "",
      "onsite_task_id": "TASK-001",
      "onsite_scheduling_status": "HPE Hold",
      "onsite_latest_service_start": "Feb XX, 2026, XX:XX XX",
      "onsite_engineer": "",
      "comms_file": "C:\\REDACTED\\out_hpe\\cases\\CASE-001_communications_redacted.txt",
      "generated_at": "2026-02-22T14:39:XXZ"
    },
    {
      "case_no": "CASE-002",
      "serial": "SERIAL-002",
      "host_name": "",
      "contact_name": "",
      "addr_street": "",
      "addr_city": "",
      "addr_state": "",
      "addr_postal_code": "",
      "addr_country": "",
      "status": "In Progress",
      "severity": "3-No Business Impact",
      "product": "HPE ProLiant DL3xx Gen10 Server",
      "product_no": "PXXXXX-XXX",
      "group": "My Group (Redacted)",
      "action_plan": "",
      "hpe_last_update": "",
      "hpe_last_subject": "",
      "hpe_request_category": "ONSITE_SERVICE",
      "hpe_request_summary": "HPE onsite interventie/dispatch loopt (technieker gepland/onderweg).",
      "hpe_requested_actions": "Check Onsite Service tab voor planning/status (task ID, scheduling status, latest service start). | Zorg dat toegang/site contact klopt; bereid interventie/onderdelen/remote access voor.",
      "hpe_key_links": "",
      "event_ids": "",
      "problem_descriptions": "",
      "ahs_links": "",
      "dropbox_hosts": "",
      "dropbox_logins": "",
      "onsite_detected": "1",
      "onsite_task_ref": "",
      "onsite_task_id": "TASK-002",
      "onsite_scheduling_status": "Closed",
      "onsite_latest_service_start": "Feb XX, 2026, XX:XX XX",
      "onsite_engineer": "",
      "comms_file": "C:\\REDACTED\\out_hpe\\cases\\CASE-002_communications_redacted.txt",
      "generated_at": "2026-02-22T14:40:XXZ"
    }
  ],
  "errors": []
}
```
---
## Security
Do not commit:
- `hpe_state.json`
- `hpe_credential.xml`
- `out_hpe/`, `out_hpe_archive/`
