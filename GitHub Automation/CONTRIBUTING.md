# Contributing

Quick guide for updating the dashboard collaboratively.

## Making a change

1. Create a branch off `main`:
   ```bash
   git checkout -b update/match-week-12
   ```
2. Make your edits (usually: update the spreadsheet and/or `MANUAL_OVERRIDES`).
3. Run the build locally to confirm there are no warnings:
   ```bash
   python build_dashboard.py
   ```
4. Commit and push your branch, then open a Pull Request on GitHub.
5. After review and merge to `main`, CI rebuilds and republishes automatically.

## Ground rules

- **Never edit `template_dashboard.html`.**
- Don't hand-edit the generated `Courage_Team_Performance_Dashboard.html` —
  change the data or the script and let the build regenerate it.
- Keep one PR focused on one update (e.g. a single match week) where possible.
