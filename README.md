# HPE_Support_BOT
HPE Support Web Page crawler export cases / serials / adres / current state / hpe required info into csv / json

Readme is in NL / EN
EN version below 

HPE CaseBot (v7) — Universal README (NL + EN)
================================================

NL — Overzicht
--------------
HPE CaseBot logt (1x interactief) in op HPE Support Center, leest je "Cases"-overzicht en haalt per case de belangrijkste info uit "Details" en "Communications".
Daarna schrijft hij een overzicht (CSV/JSON) + per case een redacted communications-dump, en een kleine status-file voor monitoring (Nagios).

Belangrijkste eigenschappen
- Windows + Python (via eigen venv .\.venv)
- Playwright Chromium wordt lokaal geïnstalleerd (1x)
- Scheduled Task draait volledig hidden (geen PowerShell-venster) en headless (geen browser window)
- Fail-safe outputs + debug artifacts (screenshots/html bij fouten)

------------------------------------------------------------
1) Quick start (1x)
------------------------------------------------------------
1. Pak de ZIP uit (bv. C:\hpe-casebot)
2. Start als normale user (Admin is niet nodig):
   install_me.cmd

Tijdens install_me.cmd:
- Unblock: haalt Windows "blocked" vlag weg van scripts
- Python check: gebruikt bij voorkeur 'py -3.12' of anders 'python'
- Setup (00_Setup.ps1): maakt .venv, installeert requirements, installeert Playwright Chromium
- Login (01_Login.ps1): opent browser om MFA/login te doen en bewaart hpe_state.json
- Optional: Scheduled Task elke X minuten (03_CreateTask_10min.ps1)
- Optional: run once test

------------------------------------------------------------
2) Login (wanneer sessie verlopen is)
------------------------------------------------------------
Als hpe_state.json ontbreekt of je krijgt SESSION_EXPIRED:

  powershell -NoProfile -ExecutionPolicy Bypass -File .\01_Login.ps1

- Er opent een Chromium window
- Log in (MFA)
- Ga tot je de HPE portal/cases ziet
- Druk ENTER in de console om de session state op te slaan

Security
- hpe_state.json bevat cookies/session. Behandel dit als een wachtwoord (niet mailen/committen).

------------------------------------------------------------
3) Handmatig runnen (debug / test)
------------------------------------------------------------
Vanuit CMD:
  .\Run-HPECaseBot.cmd

Vanuit PowerShell:
  powershell -NoProfile -ExecutionPolicy Bypass -File .\Run-HPECaseBot.ps1 -Headless

Opties (PowerShell wrapper)
- -Headless        : geen browservenster
- -Format both|csv|json
- -Max 10           : max. aantal cases verwerken
- -TimeoutSeconds 90
- -NoArchive        : geen archief-copy naar out_hpe_archive

------------------------------------------------------------
4) Output / bestanden (wat wordt waar geschreven)
------------------------------------------------------------
Belangrijkste output (folder: .\out_hpe\)
- cases_overview.csv
- cases_overview.json
- cases\<case_no>_communications_redacted.txt
- nagios\hpe_casebot.status          (JSON health/status voor monitoring)
- debug\                            (screenshots/html/errors bij failures)

Archief
- .\out_hpe_archive\YYYY-MM-DD\      (dagelijkse kopie van out_hpe)

Wat zit er in cases_overview.json?
- generated_at (UTC)
- cases[]: per case onder andere:
  - case_no, serial, product, severity, status, group
  - host_name (heuristisch uit Communications)
  - contact_name (uit "Dear <name>,")
  - hpe_last_update + hpe_last_subject (laatste HPE-message)
  - hpe_request_category / summary / requested_actions (heuristieken)
  - key links (HPRC, AHS links, support links)
  - comms_file: pad naar redacted communications dump
- errors[]: per case een fout + debug artifacts

------------------------------------------------------------
5) Scheduled Task (hidden + headless)
------------------------------------------------------------
10-min (of elk interval) task aanmaken/updaten:

  powershell -NoProfile -ExecutionPolicy Bypass -File .\03_CreateTask_10min.ps1 -RootPath "C:\hpe-casebot\FINAL" -Minutes 10

Wat doet dit?
- Schrijft (of overschrijft) Run-HPECaseBot_hidden.vbs
- Maakt een Task Scheduler entry "HPE CaseBot (10min)"
- Task start via: wscript.exe //B //NoLogo Run-HPECaseBot_hidden.vbs
- VBS roept Run-HPECaseBot.ps1 aan met -Headless

Verwijderen:
  schtasks /Delete /TN "HPE CaseBot (10min)" /F

Let op
- De task gebruikt LogonType=InteractiveToken: hij draait alleen als die user ingelogd is.
  (Dit is bewust: Playwright/Chromium + cookies/session zijn user-context gevoelig.)

------------------------------------------------------------
6) Hoe werkt het — filtering / scraping logica
------------------------------------------------------------
Welke cases worden verwerkt?
- De bot gaat naar: https://support.hpe.com/connect/s/?tab=cases
- Hij zoekt teksten die matchen met "Case <nummer>" (7-12 digits), scrollt door de lijst en verzamelt case-nummers.
- De bot filtert NIET op "open/closed" in code: hij verwerkt wat jij in de portal ziet.

Hoe filter ik wél cases?
- Zet filters in de HPE portal (My Group, status, enz.) en sla daarna de state opnieuw op (01_Login.ps1).
- Alternatief: pas cases_url aan in hpe_selectors.json.

Wat wordt per case gedaan?
1) Open case via search/click
2) Tab "Details" lezen → label/value parsing (Serial, Status, Severity, Product, ...)
3) Tab "Communications" lezen
   - Expand all communications + klik "Read more" waar nodig
   - Redaction: lines met Password/Token/Wrap token worden gemaskeerd
   - Extracties:
     - host_name ("System Name/Host Name" heuristiek)
     - contact_name ("Dear <...>,")
     - address (Street/City/State/Postal/Country block)
     - key links (HPRC / AHS / support URLs)
     - event IDs + problem descriptions
4) Heuristiek "requested actions" op basis van status + keywords

GUI wijzigingen (HPE verandert selectors)
- Alles dat UI-afhankelijk is staat in: hpe_selectors.json
- Bij errors: kijk in .\out_hpe\debug\ (screenshot + html)

------------------------------------------------------------
7) Monitoring in Nagios (aanrader)
------------------------------------------------------------
De wrapper schrijft altijd:
  .\out_hpe\nagios\hpe_casebot.status

Voorbeeld inhoud:
{
  "generated_at": "2026-02-14T17:26:54.1234567Z",
  "state": "OK" | "CRITICAL",
  "message": "OK" | "..."
}

A) Simpele monitoring (file age)
- ALERT als de status file te oud is (bot loopt niet / task faalt).
- Op Windows met NSClient++ (voorbeeldconcept): check bestandsleeftijd op hpe_casebot.status

B) Betere monitoring (age + state)
Maak op de bot-machine bv. C:\hpe-casebot\check_hpe_casebot_status.ps1 met:

  param(
    [string]$Path = "C:\\hpe-casebot\\FINAL\\out_hpe\\nagios\\hpe_casebot.status",
    [int]$WarnMinutes = 20,
    [int]$CritMinutes = 40
  )

  if (!(Test-Path -LiteralPath $Path)) { Write-Host "CRITICAL - status file missing"; exit 2 }

  $j = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json
  $ts = [datetime]::Parse($j.generated_at)
  $ageMin = [int]((Get-Date).ToUniversalTime() - $ts.ToUniversalTime()).TotalMinutes

  if ($ageMin -ge $CritMinutes) { Write-Host "CRITICAL - stale ($ageMin min) - $($j.message)"; exit 2 }
  if ($ageMin -ge $WarnMinutes) { Write-Host "WARNING  - stale ($ageMin min) - $($j.message)"; exit 1 }

  if ($j.state -ne "OK") { Write-Host "CRITICAL - $($j.state) - $($j.message)"; exit 2 }

  Write-Host "OK - updated $ageMin min ago"; exit 0

Nagios command (vanop NagiosXI) via check_nrpe (concept):
- Installeer NRPE/NSClient++ op de Windows host
- Expose een command dat bovenstaande PS script runt
- In Nagios:
  check_nrpe -H <windows-host> -c check_hpe_casebot

Extra alarm bij sessie-expiry
- Python schrijft: .\out_hpe\ALERT_SESSION_EXPIRED.txt bij SESSION_EXPIRED
- Je kan dit ook laten checken (bestandsleeftijd of aanwezigheid)

------------------------------------------------------------
8) Andere taal / universal usage (IMPORTANT)
------------------------------------------------------------
Standaard is deze bot gebouwd rond ENGELSE HPE UI-teksten (tabs, labels, hints), omdat:
- hpe_selectors.json zoekt naar teksten zoals "Details", "Communications", "Expand all communications", "Read more", ...
- infer_requested_actions() zoekt keywords in status/communicatie zoals "complete action plan" en "approve case closure".

Aanpak 1 (aanrader): Forceer HPE portal UI op Engels
- Zet je HPE profiel/portal taal op Engels
- Run opnieuw 01_Login.ps1 zodat de opgeslagen state bij Engels past

Aanpak 2: Localiseren
1) Update hpe_selectors.json
   - ready_text_any
   - tab_details_any / tab_communications_any
   - expand_all_any / read_more_any
   - (en evt. contract banner teksten)

2) Update hpe_cases_overview.py
   - FIELD_LABELS: voeg vertaalde labelvarianten toe (bv. "Numéro de série", "Statut", ...)
   - infer_requested_actions(): voeg vertaalde keywords toe voor status en communications

3) Output taal
- Samenvattingen/actions worden nu in het NL geschreven.
- Wil je EN/FR output: pas de tekststrings in infer_requested_actions() aan of maak een kleine translation-dict.

------------------------------------------------------------
9) Troubleshooting (snelle fixes)
------------------------------------------------------------
1) "Executable doesn't exist ... ms-playwright ... chromium"
   → Setup opnieuw:
     powershell -NoProfile -ExecutionPolicy Bypass -File .\00_Setup.ps1

2) SESSION_EXPIRED / redirected to login
   → Login opnieuw:
     powershell -NoProfile -ExecutionPolicy Bypass -File .\01_Login.ps1

3) GUI gewijzigd / selectors fail
   → Check .\out_hpe\debug\ (screenshot/html)
   → Pas hpe_selectors.json aan

4) Task draait maar geen output
   → Check .\out_hpe\run.log en .\out_hpe\debug\wrapper_exception_*.txt


======================================================================
EN — Overview
----------------------------------------------------------------------
HPE CaseBot logs in to the HPE Support Center (one interactive login), reads your "Cases" list and extracts key fields from "Details" and "Communications" for each case.
It exports a consolidated CSV/JSON overview, one redacted communications dump per case, and a small JSON status file for monitoring (Nagios).

Key features
- Windows + Python (isolated venv at .\.venv)
- Playwright Chromium installed locally (one-time)
- Scheduled Task runs fully hidden (no PowerShell console) and headless (no browser window)
- Fail-safe outputs + debug artifacts (screenshots/html on errors)

------------------------------------------------------------
1) Quick start (one-time)
------------------------------------------------------------
1. Extract the ZIP (example: C:\hpe-casebot\FINAL)
2. Run as a normal user (no admin required):
   install_me.cmd

What install_me.cmd does:
- Unblocks scripts (removes Windows "blocked" flag)
- Detects Python (prefers 'py -3.12', otherwise 'python')
- Setup (00_Setup.ps1): creates .venv, installs requirements, installs Playwright Chromium
- Login (01_Login.ps1): opens a browser for MFA/login and stores hpe_state.json
- Optional: create a Scheduled Task every X minutes (03_CreateTask_10min.ps1)
- Optional: quick test run

------------------------------------------------------------
2) Login (when the session expires)
------------------------------------------------------------
If hpe_state.json is missing or you get SESSION_EXPIRED:

  powershell -NoProfile -ExecutionPolicy Bypass -File .\01_Login.ps1

- Chromium opens
- Sign in (MFA)
- Once your portal/cases are loaded, press ENTER in the console

Security
- hpe_state.json contains cookies/session data. Treat it like a password.

------------------------------------------------------------
3) Manual run (debug / testing)
------------------------------------------------------------
From CMD:
  .\Run-HPECaseBot.cmd

From PowerShell:
  powershell -NoProfile -ExecutionPolicy Bypass -File .\Run-HPECaseBot.ps1 -Headless

Wrapper options
- -Headless
- -Format both|csv|json
- -Max 10
- -TimeoutSeconds 90
- -NoArchive

------------------------------------------------------------
4) Output / files
------------------------------------------------------------
Main output (folder: .\out_hpe\)
- cases_overview.csv
- cases_overview.json
- cases\<case_no>_communications_redacted.txt
- nagios\hpe_casebot.status          (JSON health/status for monitoring)
- debug\                            (screenshots/html/errors)

Archive
- .\out_hpe_archive\YYYY-MM-DD\

cases_overview.json contains:
- generated_at (UTC)
- cases[]: case_no, serial, product, severity, status, group, host_name, contact_name,
  last HPE subject/timestamp, inferred requested actions, key links, and more.
- errors[]: per-case errors and debug artifacts.

------------------------------------------------------------
5) Scheduled Task (hidden + headless)
------------------------------------------------------------
Create/update an interval task:

  powershell -NoProfile -ExecutionPolicy Bypass -File .\03_CreateTask_10min.ps1 -RootPath "C:\\hpe-casebot\\FINAL" -Minutes 10

What it does
- Writes/overwrites Run-HPECaseBot_hidden.vbs
- Creates a task "HPE CaseBot (10min)"
- Task action: wscript.exe //B //NoLogo Run-HPECaseBot_hidden.vbs
- VBS runs Run-HPECaseBot.ps1 with -Headless

Remove:
  schtasks /Delete /TN "HPE CaseBot (10min)" /F

Note
- LogonType=InteractiveToken: the task only runs while that user is logged on.

------------------------------------------------------------
6) How it works — filtering / scraping logic
------------------------------------------------------------
Which cases are processed?
- The bot opens: https://support.hpe.com/connect/s/?tab=cases
- It collects case numbers by scanning for "Case <number>" (7-12 digits) while scrolling.
- It does NOT filter by open/closed in code; it processes whatever is visible in your portal view.

How to apply filters
- Configure filters in the HPE portal (group, status, etc.) and re-run 01_Login.ps1 to save a new state.
- Or change cases_url in hpe_selectors.json.

Per-case flow
1) Open the case
2) Read "Details" tab and parse label/value pairs
3) Read "Communications" tab
   - expand all communications + click "Read more" where needed
   - redact obvious secrets (password/token lines)
   - extract hostname/contact name/address/links/event IDs/problem descriptions
4) Infer requested actions based on status + keyword heuristics

If HPE changes the UI
- Update hpe_selectors.json
- Check .\out_hpe\debug\ for screenshots/html

------------------------------------------------------------
7) Nagios monitoring
------------------------------------------------------------
The wrapper always writes:
  .\out_hpe\nagios\hpe_casebot.status

Recommended checks
A) File age (stale = task not running)
B) JSON state + age (better)

Example PowerShell checker is included in the NL section above.
Use NRPE/NSClient++ (or any Windows agent) to execute it and return Nagios exit codes.

Also note
- SESSION_EXPIRED creates .\out_hpe\ALERT_SESSION_EXPIRED.txt

------------------------------------------------------------
8) Other language support (IMPORTANT)
------------------------------------------------------------
This version is built around ENGLISH HPE UI texts because:
- hpe_selectors.json relies on "Details", "Communications", "Expand all communications", "Read more", ...
- infer_requested_actions() uses keywords like "complete action plan" and "approve case closure".

Option 1 (recommended): force the HPE portal UI to English and re-run 01_Login.ps1.
Option 2: localize
- Update hpe_selectors.json texts
- Add translated labels to FIELD_LABELS in hpe_cases_overview.py
- Add translated keywords in infer_requested_actions()
- Adjust output language strings as needed

------------------------------------------------------------
9) Troubleshooting
------------------------------------------------------------
1) "Executable doesn't exist ... chromium"
   → rerun setup:
     powershell -NoProfile -ExecutionPolicy Bypass -File .\00_Setup.ps1

2) SESSION_EXPIRED
   → rerun login:
     powershell -NoProfile -ExecutionPolicy Bypass -File .\01_Login.ps1

3) UI changed
   → check .\out_hpe\debug\ and adjust hpe_selectors.json

4) Task runs but no output
   → check .\out_hpe\run.log and .\out_hpe\debug\wrapper_exception_*.txt
