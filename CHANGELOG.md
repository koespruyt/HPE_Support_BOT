
## v9-autologin-watchdog
- Robust auto-login: always start from signin_direct_url; supports 1-step/2-step forms and iframe IdPs.
- Visibility-aware login state detection (avoids hidden DOM false positives).
- Watchdog: optional --timeout / -TimeoutSeconds; scheduled task defaults to 900s; task ExecutionTimeLimit 15m.
- Debug artifacts on login failures: out_hpe/debug/login_*.png + .html (via HPE_OUTDIR).
- Safer cleanup: always close browser/context; handles Ctrl+C gracefully.

# Changelog

## [0.3.0] - 2026-02-20

### Added
- Track all visible cases (including **In Progress**), not only Awaiting Customer Action.
- Onsite Service enrichment (JSON-only): `onsite_task_id`, `onsite_scheduling_status`, `onsite_latest_service_start`, etc.
- Automatic refresh + atomic save of `hpe_state.json` after successful runs (reduces daily re-login).

### Changed
- Console now prints `LOGIN OK` before navigating to `/cases`.
- Headless reliability improvements (viewport, tab-click reliability).

### Fixed
- Scheduled Task creation via `schtasks.exe` uses correct `/TR` quoting.
- Robust navigation to cases page with a workaround for sporadic `net::ERR_ABORTED` on SPA routing.

## [0.1.0] - 2026-02-14
- Initial release.

