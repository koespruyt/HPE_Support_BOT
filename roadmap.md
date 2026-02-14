# HPE Support BOT — Roadmap v2.0

# 1. Architectuurverbeteringen

## 🔧 1.1 Modulair framework
- [ ] Opsplitsen van scraping, parsing, export, logging en monitoring
- [ ] Duidelijke interfaces tussen modules
- [ ] Betere testbaarheid en onderhoudbaarheid

## 🧱 1.2 Selector‑plugins
- [ ] Selectors in aparte JSON/YAML‑bestanden
- [ ] Fallback‑selectors bij UI‑wijzigingen
- [ ] Ondersteuning voor meerdere HPE‑portals (Aruba, Nimble, OneView, …)

---

# 2. Betrouwbaarheid & stabiliteit

## 🔄 2.1 Resilient scraping
- [ ] Retry‑mechanisme per stap
- [ ] Detectie van login‑timeouts
- [ ] Automatische herauthenticatie

## 🧪 2.2 Testframework
- [ ] Unit tests voor parsing
- [ ] Mocked Playwright‑tests
- [ ] GitHub Actions CI‑pipeline voor linting en tests

---

# 3. Security‑verbeteringen

## 🔐 3.1 Secure credential storage
- [ ] Windows Credential Manager integratie
- [ ] Linux Secret Service ondersteuning
- [ ] Optionele encrypted config met master password

## 🕵️ 3.2 Secure session handling
- [ ] Encryptie van `hpe_state.json`
- [ ] Automatische invalidatie na X dagen
- [ ] Optionele MFA‑flow

---

# 4. Nieuwe functionaliteit

## 📬 4.1 Case‑change notifications
- [ ] E‑mailnotificaties
- [ ] Teams‑webhook
- [ ] Slack‑notificaties
- [ ] Digest‑mode (1x per dag)

## 📊 4.2 Dashboard‑export
- [ ] HTML‑dashboard
- [ ] Grafieken (open cases, aging, SLA‑risico’s)
- [ ] Optionele export naar Prometheus / Loki

## 🧾 4.3 Case‑history tracking
- [ ] Lokale SQLite‑database
- [ ] Vergelijking met vorige runs
- [ ] Detectie van nieuwe comments, statuswijzigingen, assigned engineers

---

# 5. Deployment & beheer

## 🖥 5.1 Cross‑platform support
- [ ] Windows Scheduled Task
- [ ] Linux systemd timer
- [ ] Docker‑container met Playwright

## 📦 5.2 Installer / Setup script
- [ ] Automatische Playwright‑installatie
- [ ] Config‑wizard
- [ ] Logging‑directory setup

---

# 6. Monitoring & observability

## 📈 6.1 Verbeterde Nagios‑integratie
- [ ] Statuscodes per component (login, scrape, export)
- [ ] Duidelijkere foutmeldingen

## 🪵 6.2 Structured logging
- [ ] JSON‑logging
- [ ] Loglevels (DEBUG/INFO/WARN/ERROR)
- [ ] Logrotatie

---

# 7. Documentatie & community

## 📘 7.1 Documentatie‑structuur
- [ ] `/docs` folder
- [ ] How‑to’s, troubleshooting, architecture overview

## 🌐 7.2 GitHub Pages website
- [ ] Automatische documentatie‑site
- [ ] Screenshots, flowcharts, voorbeelden

---

# 🎯 Milestones

## **Milestone: v2.0 — Architectuur & Security**
- Modulair framework
- Secure credential storage
- Resilient scraping
- Basis CI‑pipeline

## **Milestone: v2.1 — Functionaliteit & Monitoring**
- Notificaties
- Dashboard
- History tracking
- Verbeterde Nagios‑integratie

## **Milestone: v2.2 — Deployment & Community**
- Docker‑support
- Setup‑wizard
- Documentatie‑site
- Uitgebreide voorbeelden

---

# 📄 Status
Deze roadmap wordt bijgewerkt naarmate features worden geïmplementeerd of herprioritiseerd.
