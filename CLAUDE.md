# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

Desktop GUI app (single file, [cribis_export.py](cribis_export.py)) that logs in to the CRIBIS X portal with Selenium, searches each term from a dropped input file (VAT/fiscal codes or company names), opens the first result and exports company registry data to an `.xls` file. UI strings and output column names are in Italian. User-facing docs are in [README.md](README.md). No tests or linter exist.

## Commands

```powershell
python -m venv .venv; .\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe cribis_export.py                              # run the GUI
.\.venv\Scripts\python.exe -m PyInstaller cribis_export.spec --noconfirm # build dist\cribis_export.exe (windowed, one-file)
& "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe" installer\cribis_export.iss  # installer -> installer\Output\CRIBIS_Export_Setup_<ver>.exe
gh release create vX.Y.Z dist\cribis_export.exe installer\Output\CRIBIS_Export_Setup_X.Y.Z.exe --title "..." --notes "..."
```

Release flow: bump the default `AppVersion` in [installer/cribis_export.iss](installer/cribis_export.iss), rebuild the exe, compile the installer, commit/push, then `gh release create`. Distribution is only through GitHub Releases of the private repo `PshowGit/CRIBIS` (installer recommended, portable exe as alternative); `build/`, `dist/` and `installer/Output/` are git-ignored.

There are no tests: verify GUI changes by instantiating `App()` from the venv, and verify frozen-only behaviour by running the built exe (a plain launch does not prove Selenium works in the bundle).

## Architecture

- **Paths**: `BASE_DIR` is the exe's folder when frozen, else the script's folder. `config.json`, `cribis_export.log`, `user_agents.txt` and the output `.xls` are all resolved against it (never the working directory). `OUTPUT_FILE` is forced to the `.xls` extension.
- **Threading**: `Scraper.run()` executes in a daemon thread started by `App._start`. The worker never touches Tk; it sends `("log"|"progress"|"done", …)` tuples through `App.messages`, drained by `App._poll_messages` every 100 ms. Stop (button or window close) calls `Scraper.request_stop()`, which sets `stop_event` and quits the driver from a helper thread so blocked Selenium waits abort. All waits go through `stop_event.wait`.
- **Result**: `Scraper.summary` (`status` = completed/stopped/error) is shown in the in-app status label; end of run is intentionally not a dialog.
- **Output**: `XlsWriter` (xlwt) re-saves the whole workbook after every row so interrupted runs keep partial data; `.xls` is limited to 65,535 rows and fails to save if the file is open in Excel. First column is "Termine di Ricerca" so `N/D` rows stay aligned with inputs.
- **Input**: `read_search_terms` takes `.txt` lines, or the first non-empty column of `.csv/.xls/.xlsx` via pandas, dropping a header row that contains no digits. `describe_terms` labels terms as IVA (9–11 digits or 16-char fiscal code) vs Ragione sociale.
- **Scraping**: `extract_data()` uses positional XPaths (Nth `<p>` after `span.upcas.link-silver-small`); a site layout change yields `N/D` silently rather than errors. Anti-bot behaviour (random user-agent, CDP masking, random delays, long pauses, driver restart every `DRIVER_RESTART_EVERY` terms) is tuned through `config.json`. `SHOW_BROWSER=false` runs Chrome headless. ChromeDriver comes from `webdriver-manager` (matches the installed Chrome automatically, needs Internet), falling back to Selenium Manager; `CHROME_DRIVER_PATH` in config is no longer used.
- **Config**: `config.json` is created on first save and is git-ignored (it holds portal credentials; they remain in the first commit's history). The GUI only edits user/password/`SHOW_BROWSER`; other keys (`LOGIN_URL`, delays, `PROXY`…) are preserved and edited by hand.
- **PyInstaller gotcha**: Selenium lazily imports its browser modules and ships JS/selenium-manager data files, so [cribis_export.spec](cribis_export.spec) uses `collect_submodules('selenium')` and `collect_data_files('selenium')` (plus customtkinter/tkinterdnd2 data). Removing them produces `No module named 'selenium.webdriver.chrome.options'` only in the frozen exe, not from source.
- **Install location**: the Inno Setup installer installs per user (`PrivilegesRequired=lowest`, under `%LOCALAPPDATA%\Programs`) because the app writes config/log/xls next to the exe; installing under Program Files would break that. Uninstall removes `config.json` and the log but keeps the `.xls`.
