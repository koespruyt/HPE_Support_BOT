HPE CaseBot — Universal README (NL + EN)
========================================

NL — Overzicht
--------------
HPE CaseBot logt (1x interactief) in op HPE Support Center, leest je "Cases"-overzicht en haalt per case de belangrijkste info uit "Details" en "Communications".
Daarna schrijft hij een overzicht (CSV/JSON) + per case een redacted communications-dump, en een kleine status-file voor monitoring (Nagios).

Nieuw in deze build
- Verwerkt ALLE zichtbare cases (niet enkel "Awaiting Customer Action"), incl. "In Progress"
- Detecteert "Onsite Service" (als tab aanwezig) en zet extra JSON velden:
  onsite_detected, onsite_task_id, onsite_scheduling_status, onsite_latest_service_start, onsite_task_ref, onsite_engineer
- Refresh van hpe_state.json na elke run (rolling cookies/tokens) -> meestal geen dagelijkse herlogin meer
- Console: "LOGIN OK" vóór navigatie naar /cases
- Robuuster in headless + workaround voor sporadische SPA fout: net::ERR_ABORTED

------------------------------------------------------------
1) Quick start (1x)
------------------------------------------------------------
1. Pak de ZIP uit (bv. C:\HPESUPBOT)
2. Start als normale user (Admin is niet nodig):
   install_me.cmd

Tijdens install_me.cmd:
- Unblock: haalt Windows "blocked" vlag weg van scripts
- Setup (00_Setup.ps1): maakt .venv, installeert requirements, installeert Playwright Chromium
- Login (01_Login.ps1): opent browser om MFA/login te doen en bewaart hpe_state.json
- Optional: Scheduled Task (10min of daily)
- Optional: run once test

------------------------------------------------------------
2) Login (wanneer sessie écht verlopen is)
------------------------------------------------------------
Als hpe_state.json ontbreekt of je krijgt SESSION_EXPIRED:

  powershell -NoProfile -ExecutionPolicy Bypass -File .\01_Login.ps1

Security
- hpe_state.json bevat cookies/session. Behandel dit als een wachtwoord (niet mailen/committen).

------------------------------------------------------------
3) Runnen (debug / test)
------------------------------------------------------------
Vanuit CMD:
  .\Run-HPECaseBot.cmd

Vanuit PowerShell (aanrader):
  powershell -NoProfile -ExecutionPolicy Bypass -File .\Run-HPECaseBot.ps1 -Headless

Opties (PowerShell wrapper)
- -Headless         : geen browservenster
- -Format both|csv|json
- -Max 10           : max. aantal cases verwerken
- -TimeoutSeconds 90
- -NoArchive        : geen archief-copy naar out_hpe_archive

------------------------------------------------------------
4) Output / bestanden
------------------------------------------------------------
Folder: .\out_hpe\
- cases_overview.csv
- cases_overview.json   (bevat ook extra onsite_* velden indien beschikbaar)
- cases\<case_no>_communications_redacted.txt
- nagios\hpe_casebot.status
- run.log

------------------------------------------------------------
5) Scheduled Task (hidden + headless)
------------------------------------------------------------
10-min (klassiek):
  powershell -NoProfile -ExecutionPolicy Bypass -File .\03_CreateTask_10min.ps1 -RootPath (Get-Location) -Minutes 10

Daily (aanrader):
  powershell -NoProfile -ExecutionPolicy Bypass -File .\Install-HPECaseBotScheduledTask.ps1 -Time "07:00"

Let op (schtasks fallback)
- Als je ziet: "WARNING: ScheduledTasks module not usable. Falling back to schtasks.exe"
  dan is dat ok. In deze build wordt /TR correct gequotet (alles in één string).

------------------------------------------------------------
EN — Summary
-----------
Exports a CSV/JSON overview of all visible HPE cases + per-case communications dump (redacted).
Adds JSON-only enrichment for Onsite Service cases (if the tab exists), refreshes hpe_state.json after each run, and improves headless reliability.

