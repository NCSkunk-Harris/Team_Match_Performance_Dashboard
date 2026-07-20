# North Carolina Courage — Team Performance Dashboard

Weekly performance dashboard for the NC Courage (NWSL). A Python script reads
match data from a spreadsheet and renders a self-contained HTML dashboard.

**Live dashboard:** _(GitHub Pages URL appears here once Pages is enabled — see Setup)_

---

## Repository structure

```
Team Performance Dashboard/
├── build_dashboard_from_csv.py              # ENTRY POINT — this is what you run
├── build_dashboard.py                       # Shared module: MANUAL_OVERRIDES +
│                                            #   render/read helpers. Imported by
│                                            #   the entry point — do NOT delete.
├── validate_event_aggregation.py            # QC check: CSV totals vs spreadsheet
├── template_dashboard.html                  # HTML template (do NOT edit)
├── Courage_Team_Performance_Dashboard.html  # Generated output (overwritten each run)
├── Archive/                                 # Retired one-off artifacts (audits etc.)
├── .gitignore                               # (must stay at root)
├── .github/workflows/build-and-deploy.yml   # Deploy committed HTML (must stay at root)
└── GitHub Automation/                       # Automation docs + config
    ├── README.md                            # This file
    ├── update_and_push.py                   # Rebuild → commit → push, in one command
    ├── CONTRIBUTING.md
    ├── SETUP_GITHUB.md                      # One-time push/Pages setup
    └── requirements.txt                     # Python dependencies (openpyxl)
```

**Two build files, one entry point.** `build_dashboard_from_csv.py` is the
production build (v2, CSV-first). It imports `MANUAL_OVERRIDES` and most render
helpers *from* `build_dashboard.py`, so both files must stay. Running
`build_dashboard.py` on its own uses the old spreadsheet-only path and will
produce a stale dashboard — don't.

Match data lives OUTSIDE this repo, in the shared data folder used by all
projects: `../Data Organization And Cleaning/` — the three provider spreadsheets
(NWSL StatsBomb / Impect / InHouse Data.xlsx, each with a "Team Event Data"
sheet) plus the raw event CSVs the v2 build reads. Because the data isn't in the
repo, CI does not rebuild the dashboard; it deploys the committed HTML. Rebuild
locally, then commit/push.

---

## Weekly workflow

1. Drop the new match's event CSV into the shared data folder, named
   `<HomeTeam>_<AwayTeam>_<match_id>.csv`.
2. Add the match's fixture row (Date / Game Week / Match / Team) to
   **NWSL StatsBomb Data.xlsx**, sheet "Team Event Data". Metric columns can stay
   blank — the v2 build computes them from the CSV. Add Impect and InHouse rows
   whenever those providers deliver (InHouse: Courage row only).
3. If the match needs manual values (Shots on Target, Press Regains, narrative
   notes), add an entry to the `MANUAL_OVERRIDES` dict in `build_dashboard.py`,
   keyed by game number (`"M12": {...}`). Use `None` for anything not yet available.
4. Rebuild and publish:

   ```bash
   python3 "GitHub Automation/update_and_push.py"
   ```

   This rebuilds, and commits + pushes **only if the build succeeds**. GitHub
   Actions then republishes the live page from the committed HTML.

### Rebuilding without pushing

```bash
pip install -r "GitHub Automation/requirements.txt"
python3 build_dashboard_from_csv.py
```

Then open `Courage_Team_Performance_Dashboard.html` in a browser. Confirm there
are no warnings in the script output. `python3 validate_event_aggregation.py`
cross-checks the CSV-derived totals against the spreadsheet.

---

## Data sources

- **Raw event CSVs (automatic):** all StatsBomb metrics — xG, SPxG, PSxG, shots,
  Shots on Target, pressing, box metrics, OBV, deep progressions, press regains,
  and match results.
- **NWSL StatsBomb Data.xlsx:** fixture info (Date, Game Week, Match, Team) and
  Possession % — possession comes from the sheet only, with no CSV fallback.
- **NWSL Impect Data.xlsx:** packing / line-break metrics.
- **NWSL InHouse Data.xlsx:** zone metrics (SOP), Courage rows only.
- **Spreadsheet (video-coded):** QC and Big Chances.
- **MANUAL_OVERRIDES (in `build_dashboard.py`):** narrative fields — top shooters,
  key passers, notes — that have no spreadsheet column.

Anything a provider hasn't delivered renders as pending/blank rather than failing
the build.

---

## Important

- **Do not edit `template_dashboard.html`** — it's the unchanging template.
- **Run `build_dashboard_from_csv.py`, not `build_dashboard.py`.** See "Two build
  files, one entry point" above.
- Both scripts resolve paths relative to their own location, so a rebuild works
  the same from any working directory.
