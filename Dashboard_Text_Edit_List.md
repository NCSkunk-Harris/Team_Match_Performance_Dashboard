# Dashboard Text — Mark-up Edit List

**Target file(s):** `template_dashboard.html` **and** `Courage_Team_Performance_Dashboard.html` — both carry identical text. Make every change in BOTH, or make it in the build source (`build_dashboard.py`) and regenerate. Line numbers below are from `Courage_Team_Performance_Dashboard.html`.

**⚠ One decision needed from you before applying:** the true match count. The dashboard currently says 4, 7, and 14 in different places. Everywhere below I've written `[N]` — replace with the correct number. If the season figures (goals, xG) should be live rather than typed, see Edit 3.

---

## PRIORITY 1 — Contradictions

### Edit 1a — Header match count (line 1879)
**REMOVE:**
```html
      <div>7 matches played</div>
```
**REPLACE WITH:**
```html
      <div>[N] matches played</div>
```

### Edit 1b — Header update date (line 1880)
**REMOVE:**
```html
      <div style="margin-top:2px;">Updated May 4, 2026</div>
```
**REPLACE WITH:** (use the real last-update date)
```html
      <div style="margin-top:2px;">Updated [Month DD, 2026]</div>
```

### Edit 1c — Defensive Intensity source note (line 2024)
**REMOVE:**
```
Source: Match CSVs (raw StatsBomb events), all 14 matches. Tackles + Interceptions = ...
```
**REPLACE WITH:**
```
Source: Match CSVs (raw StatsBomb events), all [N] matches. Tackles + Interceptions = ...
```

### Edit 1d — Context "Important Limitations" card (line 4038)
**REMOVE:**
```
<strong>Small sample size.</strong> With 7 matches, the season range is building ...
```
**REPLACE WITH:**
```
<strong>Small sample size.</strong> With [N] matches, the season range is building ...
```

### Edit 1e — Data Sources table, three "Coverage" cells (lines ~4145, ~4151, ~4157)
**REMOVE (×3):**
```html
<td>All 14 matches — complete</td>
```
**REPLACE WITH (×3):**
```html
<td>All [N] matches — complete</td>
```

---

### Edit 2 — Grading card: delete the contradictory z-score paragraph (line 3998)
The card describes Match Grade as a **z-score**, then the next paragraph says z-score was **rejected** in favour of min-max, and the formula + grade scale below are min-max. Paragraph 1 is wrong. Replace it with a min-max description.

**REMOVE (entire paragraph 1):**
```html
<p class="ctx-body-text">Each metric uses <strong>two separate grading functions</strong> that answer different questions, and both are <strong>fixed once established</strong> — a grade is computed only from matches played up to and including the graded match, so later results never rewrite it. <strong>Match Grade</strong> is a z-score: the match value versus the mean and standard deviation of the team's performances to that point (z ≥ +1.25 Excellent · +0.5 Good · ±0.5 Typical · −1.25 Poor · below Abysmal, direction-aware). <strong>3-Match Form</strong> applies the same z-bands to the rolling average of the last three consecutive matches, against the same to-date baseline — because rolling averages are smoother, the identical bands are effectively a stricter test of sustained form. Both grades require four matches of history; earlier matches show as pending. A match can still grade Poor on Match Grade yet Good on 3-Match Form — a single-game dip inside a form window that tracks above baseline.</p>
```
**REPLACE WITH:**
```html
<p class="ctx-body-text">Each metric uses <strong>two separate grading functions</strong> that answer different questions, and both are <strong>fixed once established</strong> — a grade is computed only from matches played up to and including the graded match, so later results never rewrite it. <strong>Match Grade</strong> places the match value on the season's min–max range for that metric to date (see Formula below): higher position = better, direction-aware for negative metrics. <strong>3-Match Form</strong> applies the grade bands to the rolling average of the last three consecutive matches, measured against the <em>season mean</em> rather than the min/max range, with wider deviation bands (±10% Typical, ±35% Excellent/Abysmal) so rolling averages aren't over-graded. Both grades require four matches of history; earlier matches show as pending. A match can still grade Poor on Match Grade yet Good on 3-Match Form — a single-game dip inside a form window that tracks above baseline.</p>
```
*Note:* this makes paragraph 1 agree with paragraph 2 ("Min-max will remain the method"), the Formula line, the grade scale, the main-page legend, and the "3-Match Form" description in the Limitations card.

---

### Edit 3 — Season-bar hardcoded figures (lines 1913–1934) — VERIFY / make live
These numbers are typed into the HTML, not computed. Confirm each is current for `[N]` matches, or wire them to the data like `stat-goals` already is.

- `<div class="season-stat-value">4</div>` → should equal `[N]`
- Goals Conceded `<div class="season-stat-value" id="stat-conceded">6</div>` and sub `1.50 per match` → verify
- Avg xG `<div class="season-stat-value" id="stat-xg">1.35</div>` → verify

---

## PRIORITY 2 — Remove text describing things not on the page

### Edit 4 — "Box Resilience" context card (lines ~4087–4096)
There is no Box Resilience tab in the nav. **Either** delete the whole `ctx-card` block (from its `<div class="ctx-card">` through its closing `</div>`), **or**, if the tab is planned, prefix the title:
**REMOVE:**
```html
<div class="ctx-card-title">Box Resilience</div>
```
**REPLACE WITH:**
```html
<div class="ctx-card-title">Box Resilience <span style="font-weight:400;opacity:0.6;">(planned)</span></div>
```
Same check applies to the "Box Resilience" print page (line ~4188) and the `'box-res'` label (line ~4259).

### Edit 5 — Trim Box Domination / Defensive Intensity context to match the live tabs
The live Box Domination tab shows only a single "Box Battle" chart, but its context card describes "Penetration" and "Finishing" tables. The live Defensive Intensity tab shows only "Pressing Battle," but its context describes three areas with tables. Edit these two cards so they describe the head-to-head Battle view that actually renders, or add "(detailed tables planned)". No exact string given — depends on your decision on scope.

---

## PRIORITY 3 — Redundant text (say it once)

### Edit 6 — Post-shot xG / chance-quality definitions appear 3×
Keep the full definition in the Context page (line 4081). Shorten the two in-tab notes to pointers.

**Line 2006 — Box Battle note. REMOVE the definitions tail:**
```
Source: Match CSVs (StatsBomb events). xG = shot xG; set-piece xG = corner/free-kick/throw-in shots; post-shot xG = keeper save-difficulty xG; Quality Chances = shots xG 0.15–0.30; Big Chances = shots xG ≥ 0.30.
```
**REPLACE WITH:**
```
Source: Match CSVs (StatsBomb events). Metric definitions in Context &amp; Logic → Box Domination.
```
Leave line 4081 (Context page) as the single source of truth for these definitions.

### Edit 7 — "Completion % = Completions ÷ Attempts" stated twice
Appears in the SOP context card and the Data Sources table. Keep it in the SOP context card; delete the trailing "Completion % calculated as Completions ÷ Attempts." from the Data Sources table "Notes" cell for the InHouse row.

---

## PRIORITY 4 — Naming / minor

### Edit 8 — Disambiguate the two "Insights" entry points
A disabled subtab reads "Match Insights" while the top button "Insights" opens **Season** Insights. Rename the disabled subtab so nothing else is called Insights:
**REMOVE:**
```html
<div class="subtab disabled" aria-disabled="true" title="Coming soon">Match Insights</div>
```
**REPLACE WITH:**
```html
<div class="subtab disabled" aria-disabled="true" title="Coming soon">Per-Match Report</div>
```

### Edit 9 — Strengths card wording (build_dashboard JS, ~line 3835)
`top ${100-pct+1}%` reads awkwardly. Consider `top ${pct}% of the season range` for plain-language clarity. Cosmetic, not a contradiction.

---

## KEEP — do not touch
- In-tab descriptor lines ("Reads down the pitch toward goal", "bar reaches whoever pressed harder", "negative = value lost") — each is specific and non-redundant.
- Grade legend caption — dense but earns its space.
- Dynamically generated Insights cards — clean and already use consistent min-max language.

---

## Durable fix (recommended)
Contradictions 1 and 2 exist because prose is hand-typed into a file that is also machine-generated. Have `build_dashboard.py` inject match count, update date, and season totals as variables so the narrative can never disagree with the data again. Then this edit list only needs applying once, in the template/source.
