# North Carolina Courage — Team Performance Dashboard

Weekly performance dashboard for the NC Courage (NWSL). A Python script reads
match data from a spreadsheet and renders a self-contained HTML dashboard.

**Live dashboard:** _(GitHub Pages URL appears here once Pages is enabled — see Setup)_

---

## Repository structure

```
Team Performance Dashboard/
├── build_dashboard.py                       # Build script — run weekly
├── template_dashboard.html                  # HTML template (do NOT edit)
├── Courage_Team_Performance_Dashboard.html  # Generated output (overwritten each run)
├── .gitignore                               # (must stay at root)
├── .github/workflows/build-and-deploy.yml   # Auto-rebuild + deploy (must stay at root)
└── GitHub Automation/                       # Automation docs + config
    ├── README.md                            # This file
    ├── CONTRIBUTING.md
    ├── SETUP_GITHUB.md                      # One-time push/Pages setup
    └── requirements.txt                     # Python dependencies (openpyxl)
```

Match data lives OUTSIDE this repo, in the shared data folder used by all
projects: `../Data Organization And Cleaning/` (NWSL StatsBomb Data.xlsx,
NWSL Impect Data.xlsx, NWSL InHouse Data.xlsx — each with a "Team Event Data"
sheet). Because the data isn't in the repo, CI does not rebuild the dashboard;
it deploys the committed HTML. Rebuild locally, then commit/push.

---

## Weekly workflow

1. Add the new match's rows to each provider file in
   `../Data Organization And Cleaning/` (sheet **"Team Event Data"**): one
   North Carolina Courage row + one opponent row (InHouse: Courage row only).
2. If the match needs manual values (Shots on Target, Press Regains, narrative
   notes), add a new entry to the `MANUAL_OVERRIDES` dict in `build_dashboard.py`,
   keyed by game number (`"M12": {...}`). Use `None` for anything not yet available.
3. Run `python build_dashboard.py` (or use the watcher script), then commit and
   push. GitHub Actions republishes the live page from the committed HTML.

### Running it locally (optional)

```bash
pip install -r "GitHub Automation/requirements.txt"
python build_dashboard.py
```

Then open `Courage_Team_Performance_Dashboard.html` in a browser. Confirm there
are no warnings in the script output.

---

## Data sources

- **Spreadsheet (automatic):** all metrics, match metadata, result (computed
  from Goals), Tackles Won, Opp Possession %, and Post-Shot xG (from the
  "Post Shot xG" column).
- **MANUAL_OVERRIDES (in `build_dashboard.py`):** Shots on Target, Press Regains,
  and narrative fields (top shooters, key passers, notes) that have no
  spreadsheet column.

---

## Important

- **Do not edit `template_dashboard.html`** — it's the unchanging template.
- The build script resolves all paths relative to its own location, so it works
  the same locally and in CI.
