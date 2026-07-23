"""
Courage Team Performance Dashboard — CSV-first build (v2)
==========================================================

NEW PROCESS (replaces spreadsheet-only build for StatsBomb metrics):

  StatsBomb metrics   → computed from raw event CSVs in
                        ../Data Organization And Cleaning/Data/Statsbomb Match CSVs/
                        (xG, SPxG, PSxG, shots, SOT, pressing, box metrics,
                        OBV, deep progressions, press regains, results)
  Fixtures (date/GW)  → union of Impect + InHouse "Team Event Data"
  Packing / LB        → NWSL Impect Data.xlsx (unchanged)
  Zone metrics (SOP)  → NWSL InHouse Data.xlsx (unchanged)
  QC / Big Chances    → computed from shot xG bands in the CSVs
                        (QC 0.15-0.30, BC >0.30). See QC_LO/QC_HI for accuracy.
  Possession % (PA)   → NWSL Impect Data.xlsx "Team Event Data" (no CSV fallback).
                        Validated 26/26 exact vs the retired StatsBomb workbook.
  Narrative text      → MANUAL_OVERRIDES in build_dashboard.py (unchanged)

Anything missing from a provider simply renders as pending/blank.

GAME WEEKS ARE NOT UNIQUE: a game week can contain two fixtures (GW6 holds
both 2026-04-30 and 2026-05-03). Match identity is game_number (M1, M2, ...),
derived from chronological date order — never Game Week. Labels render as
"GW6 · M6" and "GW6 · M7" so a doubled week reads clearly.

WEEKLY WORKFLOW:
  1. Drop the new match CSV into Data/Statsbomb Match CSVs/
     (filename: "<HomeTeam>_<AwayTeam>_<match_id>.csv")
  2. Add the match row (Date/Game Week/Venue/Match/Team) to Impect and/or
     InHouse — metric columns may stay blank; only fixture info is needed.
  3. Add Impect/InHouse rows whenever those providers deliver.
  4. Run this script.

OUTPUT: Courage_Team_Performance_Dashboard.html (production file)
"""

import csv
import re
import sys
from collections import defaultdict
from pathlib import Path

from build_dashboard import (
    MANUAL_OVERRIDES, read_spreadsheet, get_opp_possession_pct, pa_adjust,
    ZONE_POINT_WEIGHTS,
    build_bypass_segments, render_bypass_segments_block, render_matches_block,
    render_match_info_block, render_data_block, render_match_report_data_block,
    replace_block, safe_round, safe_div, info, warn,
)

SCRIPT_DIR = Path(__file__).resolve().parent
# Raw event CSVs live in the shared data folder (sibling of this project),
# alongside the three provider spreadsheets — not beside the script.
CSV_DIR = SCRIPT_DIR.parent / "Data Organization And Cleaning" / "Data" / "Statsbomb Match CSVs"
TEMPLATE = SCRIPT_DIR / "template_dashboard.html"
OUTPUT = SCRIPT_DIR / "Courage_Team_Performance_Dashboard.html"

NCC = "North Carolina Courage"
BOX_X, BOX_Y_LO, BOX_Y_HI = 102.0, 18.0, 62.0
FINAL_THIRD_X = 80.0
HALFWAY_X = 60.0
ON_TARGET = {"Goal", "Saved", "Saved to Post", "Saved To Post"}
PRESS_REGAIN_WINDOW = 5.0

# ── QUALITY / BIG CHANCE THRESHOLDS (xG bands) ───────────────────────────────
# Per Tom's definition: a Quality Chance is a shot worth 0.15–0.30 xG; a Big
# Chance is a shot worth more than 0.30. Bands are mutually exclusive, so a
# shot is counted once: QC if QC_LO <= xg <= QC_HI, BC if xg > QC_HI.
#
# VALIDATED against the 26 team-rows in the retired NWSL StatsBomb workbook
# (13 matches x 2 sides), which carried hand/video-coded QC and BC values:
#
#   Big Chances     21/26 exact (81%). Totals 24 computed vs 25 recorded;
#                   mean signed error -0.04 — effectively unbiased. A grid
#                   search over every cut from 0.15 to 0.70 peaked at 22/26
#                   (0.29), so 0.30 is at the optimum and is trustworthy.
#
#   Quality Chances 15/26 exact (58%). Totals 32 computed vs 40 recorded;
#                   mean signed error -0.31 — a systematic ~20% UNDERCOUNT
#                   (8 rows under, 3 over). A grid search over every band
#                   between 0.05 and 0.60 peaked at only 17/26, so the gap is
#                   NOT a threshold-tuning problem: the original QC metric
#                   counted chances that never became shots (a chance created
#                   but not struck has no shot event and therefore no xG).
#                   Treat the QC row as a shot-based proxy, not a like-for-like
#                   replacement for the old video-coded column.
QC_LO, QC_HI = 0.15, 0.30

# Play patterns that are NOT open play (see open_play() in compute_team_metrics).
SET_PIECE_PATTERNS = {"From Corner", "From Free Kick"}


# ─────────────────────────────────────────────────────────────────────────────
# CSV EVENT PARSING / METRICS  (definitions validated in
# CSV_Rebuild_Feasibility_Report.md; * = approximation of IQ definition)
# ─────────────────────────────────────────────────────────────────────────────

def _f(row, key):
    try:
        return float(row.get(key, ""))
    except (TypeError, ValueError):
        return None


def read_match_csv(path):
    """Read one event CSV, deduping on event id (freeze frames repeat rows)."""
    events, seen = [], set()
    with open(path, newline="", encoding="utf-8") as fh:
        for row in csv.DictReader(fh):
            eid = row.get("id")
            if eid in seen:
                continue
            seen.add(eid)
            events.append(row)
    return events


def in_box(x, y):
    return x is not None and y is not None and x >= BOX_X and BOX_Y_LO <= y <= BOX_Y_HI


def compute_team_metrics(events, team):
    te = [e for e in events if e.get("team_name") == team]
    m = {}

    shots = [e for e in te if e["event_type_name"] == "Shot"
             and int(e["period"] or 0) <= 4]
    m["shots"] = len(shots)
    m["goals"] = sum(1 for s in shots if s.get("outcome_name") == "Goal") \
        + sum(1 for e in te if e["event_type_name"] == "Own Goal For")
    m["xg"] = sum(_f(s, "statsbomb_xg") or 0 for s in shots)
    def is_set_piece(s):
        # Set-piece shot: any shot from a corner, free-kick or throw-in play
        # pattern. This matches the dashboard's Set Piece xG column exactly. A
        # broader rule (type_name Free Kick/Corner, or a set_piece_phase flag)
        # was tested but over-counted vs the on-screen values, so we keep it
        # to the three play patterns. Penalties are excluded by convention.
        if s.get("type_name") == "Penalty":
            return False
        return s.get("play_pattern_name") in {"From Corner", "From Free Kick",
                                              "From Throw In"}
    m["spxg"] = sum(_f(s, "statsbomb_xg") or 0 for s in shots if is_set_piece(s))
    m["xgps"] = safe_div(m["xg"], m["shots"])
    on_t = [s for s in shots if s.get("outcome_name") in ON_TARGET]
    m["sot"] = len(on_t)
    m["sot_pct"] = safe_div(m["sot"], m["shots"])
    # Finishing vs chance quality: goals scored minus xG created. Positive =
    # outscored the chances created, negative = wasteful. Uses the same shot
    # set as m["xg"], and m["goals"] includes own goals scored FOR this team,
    # which xG cannot model — so a match with an own goal will read high.
    m["xg_perf"] = m["goals"] - m["xg"]
    m["psxg"] = sum(_f(s, "gk_save_difficulty_xg") or 0 for s in on_t)

    # Quality / Big Chances from shot xG bands (see QC_LO/QC_HI notes above).
    shot_xgs = [x for x in (_f(s, "statsbomb_xg") for s in shots) if x is not None]
    # Boundary at exactly QC_HI (0.30) counts as a BIG chance, matching Tom's
    # definition: "0.15 to 0.30 = quality, 0.30 and larger = big".
    m["qc"] = sum(1 for x in shot_xgs if QC_LO <= x < QC_HI)
    m["bc"] = sum(1 for x in shot_xgs if x >= QC_HI)

    pressures = [e for e in te if e["event_type_name"] == "Pressure"]
    m["pressures_raw"] = len(pressures)
    m["press_oh_raw"] = sum(1 for e in pressures
                            if (_f(e, "location_x") or 0) > HALFWAY_X)
    m["press_oh_pct"] = safe_div(m["press_oh_raw"], m["pressures_raw"])
    cps = [e for e in te if (e.get("counterpress") or "").lower() == "true"]
    m["cp_raw"] = len(cps)
    m["cp_oh_raw"] = sum(1 for e in cps if (_f(e, "location_x") or 0) > HALFWAY_X)

    def tsec(e):
        return int(e["period"] or 0) * 10000 + int(e["minute"] or 0) * 60 \
            + int(e["second"] or 0)

    import bisect
    press_times = sorted(tsec(e) for e in pressures)
    regains = 0
    for e in te:
        is_regain = e["event_type_name"] in {"Ball Recovery", "Interception"} \
            or (e["event_type_name"] == "Duel" and e.get("type_name") == "Tackle"
                and e.get("outcome_name") in {"Won", "Success",
                                              "Success In Play", "Success Out"})
        if not is_regain:
            continue
        t = tsec(e)
        i = bisect.bisect_right(press_times, t)
        if i and t - press_times[i - 1] <= PRESS_REGAIN_WINDOW:
            regains += 1
    m["press_regains"] = regains

    tackles_won = sum(
        1 for e in te if e["event_type_name"] == "Duel"
        and e.get("type_name") == "Tackle"
        and e.get("outcome_name") in {"Won", "Success", "Success In Play",
                                      "Success Out"})
    interceptions = sum(1 for e in te if e["event_type_name"] == "Interception")
    m["tack_int_raw"] = tackles_won + interceptions              # *
    m["agg_raw"] = m["pressures_raw"] + tackles_won + sum(       # *
        1 for e in te if e["event_type_name"] == "Foul Committed")

    # Battle-view raw counts (NOT possession-adjusted). Tackles+Interceptions
    # here counts ALL tackle-type duels (any outcome) plus interceptions, and
    # ball recoveries are a standalone count. These feed defint.battle.
    all_tackles = sum(1 for e in te if e["event_type_name"] == "Duel"
                      and e.get("type_name") == "Tackle")
    m["tack_int_battle"] = interceptions + all_tackles
    m["ball_recoveries"] = sum(
        1 for e in te if e["event_type_name"] == "Ball Recovery")

    passes = [e for e in te if e["event_type_name"] == "Pass"]
    comp = [p for p in passes if not (p.get("outcome_name") or "").strip()]
    m["pass_pct"] = safe_div(len(comp), len(passes))
    carries = [e for e in te if e["event_type_name"] == "Carries"]

    def entered_box(ev):
        """Event that ends inside the box having started outside it.
        Box geometry (x >= 102, 18 <= y <= 62) is EXACT, not an approximation:
        counting shots outside this area reproduced the retired workbook's
        "Shots Outside Box" column 68/68 times with zero error."""
        return in_box(_f(ev, "end_location_x"), _f(ev, "end_location_y")) \
            and not in_box(_f(ev, "location_x"), _f(ev, "location_y"))

    box_pass = [p for p in passes if entered_box(p)]
    box_pass_c = [p for p in box_pass if not (p.get("outcome_name") or "").strip()]
    box_carry = [c for c in carries if entered_box(c)]
    m["bea_cp"] = len(box_pass)
    m["bea_carry"] = len(box_carry)
    m["sbe"] = len(box_pass_c) + len(box_carry)                  # *

    def open_play(ev):
        return ev.get("play_pattern_name") not in SET_PIECE_PATTERNS

    def is_cross(ev):
        return (ev.get("pass_cross") or "").strip().lower() == "true"

    # ── OPEN PLAY BOX ENTRIES (pass / cross / carry) ─────────────────────────
    # "Open play" = every play pattern EXCEPT corners and free kicks. This is
    # StatsBomb's own definition, reverse-engineered rather than assumed: their
    # "Open Play Shots Outside Box" column was reproduced most accurately by
    # this filter (51/68 exact, MAE 0.31). Throw-ins, goal kicks, keeper
    # distribution, kick-offs and counters all COUNT as open play — restricting
    # to play_pattern "Regular Play" alone scored far worse (26/68) and
    # undercounted by a full shot per team per match.
    op_pass = [p for p in box_pass if open_play(p)]
    m["op_bea_cross"] = sum(1 for p in op_pass if is_cross(p))
    m["op_bea_pass"] = sum(1 for p in op_pass if not is_cross(p))
    m["op_bea_carry"] = sum(1 for c in box_carry if open_play(c))
    m["op_bea_total"] = m["op_bea_cross"] + m["op_bea_pass"] + m["op_bea_carry"]
    m["op_sbe"] = sum(1 for p in op_pass
                      if not (p.get("outcome_name") or "").strip()) \
        + m["op_bea_carry"]

    # ── BOX METRICS VALIDATED EXACTLY AGAINST THE RETIRED WORKBOOK ───────────
    # Each of the following reproduced its column in NWSL StatsBomb Data.xlsx
    # 68/68 times with zero error, so these are definitions, not estimates:
    #   passes_into_box, succ_passes_into_box, passes_inside_box,
    #   succ_box_crosses, shots_outside_box
    # "Inside box" means the pass STARTS AND ENDS in the box — origin-only
    # overcounted by +2.18/match. "Into box" means it starts outside, ends in.
    m["passes_into_box"] = len(box_pass)
    m["succ_passes_into_box"] = len(box_pass_c)
    inside = [p for p in passes
              if in_box(_f(p, "location_x"), _f(p, "location_y"))
              and in_box(_f(p, "end_location_x"), _f(p, "end_location_y"))]
    m["passes_inside_box"] = len(inside)
    m["succ_passes_inside_box"] = sum(
        1 for p in inside if not (p.get("outcome_name") or "").strip())
    m["box_crosses"] = sum(1 for p in box_pass if is_cross(p))
    m["succ_box_crosses"] = sum(
        1 for p in box_pass
        if is_cross(p) and not (p.get("outcome_name") or "").strip())
    m["box_cross_pct"] = safe_div(m["succ_box_crosses"], m["box_crosses"])

    outside_shots = [sh for sh in shots
                     if not in_box(_f(sh, "location_x"), _f(sh, "location_y"))]
    m["shots_outside_box"] = len(outside_shots)
    m["op_shots_outside_box"] = sum(1 for sh in outside_shots
                                    if open_play(sh))

    def f3(x):
        return x is not None and x >= FINAL_THIRD_X
    m["deep_prog"] = sum(1 for p in comp if f3(_f(p, "end_location_x"))
                         and not f3(_f(p, "location_x"))) \
        + sum(1 for c in carries if f3(_f(c, "end_location_x"))
              and not f3(_f(c, "location_x")))

    touch_types = {"Pass", "Shot", "Carries", "Dribble"}
    m["tib"] = sum(1 for e in te if e["event_type_name"] in touch_types
                   and in_box(_f(e, "location_x"), _f(e, "location_y")))  # *

    def obv(types):
        return sum(_f(e, "obv_total_net") or 0 for e in te
                   if e["event_type_name"] in types)
    m["obv_total"] = sum(_f(e, "obv_total_net") or 0 for e in te)
    m["obv_pass"] = obv({"Pass"})
    m["obv_shot"] = obv({"Shot"})
    m["obv_dribble"] = obv({"Carries", "Dribble"})
    m["obv_def"] = obv({"Ball Recovery", "Interception", "Block",
                        "Clearance", "Duel"})
    return m


# ─────────────────────────────────────────────────────────────────────────────
# CSV ↔ FIXTURE MATCHING
# ─────────────────────────────────────────────────────────────────────────────

def _norm(name):
    """Normalize a team name for matching (filenames sanitize '/' etc.)."""
    n = re.sub(r"[^a-z0-9]", "", name.lower())
    # Provider quirks: 'Boston Legacy W', 'Denver Summit W' vs plain names
    if n.endswith("w") and len(n) > 6:
        n_stripped = n[:-1]
    else:
        n_stripped = n
    return n_stripped


def _same_team(a, b):
    na, nb = _norm(a), _norm(b)
    return na == nb or na in nb or nb in na


def load_csv_matches():
    """Parse every CSV → {(norm_opp, loc): computed metrics bundle}."""
    out = {}
    for path in sorted(CSV_DIR.glob("*.csv")):
        mm = re.match(r"(.+)_(.+)_(\d+)\.csv$", path.name)
        if not mm:
            warn(f"skipping unrecognized filename: {path.name}")
            continue
        home, away, mid = mm.group(1), mm.group(2), int(mm.group(3))
        if not (_same_team(home, NCC) or _same_team(away, NCC)):
            warn(f"skipping non-Courage file: {path.name}")
            continue
        loc = "Home" if _same_team(home, NCC) else "Away"
        opp_file = away if loc == "Home" else home
        events = read_match_csv(path)
        # Resolve exact team names as they appear inside the CSV
        team_names = {e.get("team_name") for e in events if e.get("team_name")}
        opp_name = next((t for t in team_names if not _same_team(t, NCC)), opp_file)
        dur = defaultdict(float)
        for e in events:
            t, d = e.get("possession_team_name"), _f(e, "duration")
            if t and d:
                dur[t] += d
        total = sum(dur.values()) or 1.0
        bundle = {
            "match_id": mid, "file": path.name, "loc": loc,
            "ncc": compute_team_metrics(events, NCC),
            "opp": compute_team_metrics(events, opp_name),
            "ncc_poss_csv": dur.get(NCC, 0) / total,
        }
        out[(_norm(opp_file), loc)] = bundle
    return out


def attach_csv_to_matches(matches, csv_matches):
    """Attach each fixture's CSV bundle (matched on opponent + venue)."""
    used = set()
    for m in matches:
        key = None
        for (nopp, loc), bundle in csv_matches.items():
            if loc == m["loc"] and _same_team(nopp, m["opp"]) \
                    and (nopp, loc) not in used:
                key = (nopp, loc)
                break
        if key is None:
            warn(f"M{m['game_number']} vs {m['opp']} ({m['loc']}): "
                 f"no CSV found — StatsBomb metrics will be blank")
            m["csv"] = None
        else:
            used.add(key)
            m["csv"] = csv_matches[key]
            info(f"M{m['game_number']} vs {m['opp']:<22} ← {csv_matches[key]['file']}")
    leftovers = set(csv_matches) - used
    for k in leftovers:
        warn(f"CSV not matched to any fixture: {csv_matches[k]['file']}")


# ─────────────────────────────────────────────────────────────────────────────
# BUILD DATA (CSV for StatsBomb, spreadsheets for Impect/InHouse/QC-BC)
# ─────────────────────────────────────────────────────────────────────────────

def _net(matches, key, dec=3):
    """Courage value minus opponent value for a CSV-computed metric."""
    out = []
    for m in matches:
        b = m.get("csv")
        if b is None:
            out.append(None); continue
        a, o = b["ncc"].get(key), b["opp"].get(key)
        out.append(None if a is None or o is None else safe_round(a - o, dec))
    return out


def build_data(matches, sheet_data):
    n = len(matches)
    # Possession comes exclusively from NWSL Impect Data.xlsx ("Possession %"
    # on the OPP row). No CSV fallback — a missing sheet value warns + defaults
    # to 0.50 inside get_opp_possession_pct.
    opp_poss = get_opp_possession_pct(matches, sheet_data)

    def c(side, key, dec=None, adj=False):
        out = []
        for i, m in enumerate(matches):
            b = m.get("csv")
            v = None if b is None else b[side].get(key)
            if adj and v is not None:
                p = opp_poss[i] if side == "ncc" else 1 - opp_poss[i]
                v = pa_adjust(v, p)
            out.append(safe_round(v, dec) if dec is not None else v)
        return out

    def s(side, header):
        return [sheet_data.get((m["game_number"], side), {}).get(header)
                for m in matches]

    def spct(side, header):
        return [safe_round(v, 4) if v is not None else None
                for v in s(side, header)]

    def row(id_, label, vals, hb=True, **kw):
        d = {"id": id_, "label": label, "vals": vals, "higherBetter": hb}
        d.update(kw)
        return d

    def bt(label, ncc_vals, opp_vals, hb=True):
        """A head-to-head 'battle' row: NCC vs opponent arrays on one metric."""
        return {"label": label, "ncc": ncc_vals, "opp": opp_vals,
                "higherBetter": hb}

    def zpts(header):
        """Zone points from the sheet, falling back to completions x weight
        when the Excel formula cache is empty (see ZONE_POINT_WEIGHTS)."""
        comp_hdr, w = ZONE_POINT_WEIGHTS[header]
        sheet_v, comp_v = s("NCC", header), s("NCC", comp_hdr)
        out, derived = [], 0
        for sv, cv in zip(sheet_v, comp_v):
            if sv is not None:
                out.append(sv)
            elif cv is not None:
                out.append(cv * w); derived += 1
            else:
                out.append(None)
        if derived:
            warn(f"{header}: {derived}/{len(out)} values derived from "
                 f"{comp_hdr} x{w} (sheet formula cache empty)")
        return out

    z2a, z2c = s("NCC", "Z2 Attempts"), s("NCC", "Z2 Completions")
    z3a, z3c = s("NCC", "Z3 Attempts"), s("NCC", "Z3 Completions")
    pza, pzc = s("NCC", "PZ Attempts"), s("NCC", "PZ Completions")
    sza, szc = s("NCC", "SZ Attempts"), s("NCC", "SZ Completions")
    pct = lambda cs, as_: [safe_round(safe_div(x, y), 4)
                           for x, y in zip(cs, as_)]
    bw_def = [pa_adjust(v, p) if v is not None else None
              for v, p in zip(s("NCC", "Ball Win Removed Opponents Defenders"),
                              opp_poss)]

    DATA = {
        "sop": {
            "zones": [
                row("Z2_att", "Zone 2 Attempts", z2a),
                row("Z2_comp", "Zone 2 Completions", z2c),
                row("Z2_pct", "Zone 2 Completion %", pct(z2c, z2a), pct=True, dec=3),
                row("Z3_att", "Zone 3 Attempts", z3a),
                row("Z3_comp", "Zone 3 Completions", z3c),
                row("Z3_pct", "Zone 3 Completion %", pct(z3c, z3a), pct=True, dec=3),
                row("Z2_pts", "Zone 2 Points (x1)", zpts("Z2 Points")),
                row("Z3_pts", "Zone 3 Points (x2)", zpts("Z3 Points")),
                row("PZ_att", "Pass Zone Attempts", pza),
                row("PZ_comp", "Pass Zone Completions", pzc),
                row("PZ_pct", "Pass Zone Completion %", pct(pzc, pza), pct=True, dec=3),
                row("SZ_att", "Shoot Zone Attempts", sza),
                row("SZ_comp", "Shoot Zone Completions", szc),
                row("SZ_pct", "Shoot Zone Completion %", pct(szc, sza), pct=True, dec=3),
                row("PZ_pts", "Pass Zone Points (x3)", zpts("PZ Points")),
                row("SZ_pts", "Shoot Zone Points (x5)", zpts("SZ Points")),
                row("zone_pts_total", "Total Zone Points (weighted)",
                    [None if any(v is None for v in vs) else sum(vs)
                     for vs in zip(zpts("Z2 Points"), zpts("Z3 Points"),
                                   zpts("PZ Points"), zpts("SZ Points"))]),
            ],
        },
        "packing": {
            "volume": [
                row("byp_opp", "Bypassed Opponents", s("NCC", "Bypassed Opponents")),
                row("byp_def", "Bypassed Defenders", s("NCC", "Bypassed Defenders")),
                row("bw_opp", "Ball Win Removed Opponents", s("NCC", "Ball Win Removed Opponents")),
                row("bw_def", "Ball Win Removed Defenders (PA)",
                    [safe_round(v, 1) for v in bw_def], dec=1),
                row("crit_loss", "Critical Ball Loss", s("NCC", "Critical Ball Loss Number"), hb=False, context="neg"),
                row("bl_team", "Ball Loss Removed Teammates", s("NCC", "Ball Loss Removed Teammates"), hb=False, context="neg"),
                row("deep_prog", "Final 3rd Entries", c("ncc", "deep_prog")),
            ],
            # NCC vs Opponent packing (Impect, both rows). Raw per-match counts.
            # bw_def here is the RAW "Ball Win Removed Opponents Defenders", not
            # the PA version used in the volume table above.
            "battle": [
                bt("Bypassed Opponents", s("NCC", "Bypassed Opponents"), s("OPP", "Bypassed Opponents")),
                bt("Bypassed Defenders", s("NCC", "Bypassed Defenders"), s("OPP", "Bypassed Defenders")),
                bt("Ball Wins (opp removed)", s("NCC", "Ball Win Removed Opponents"), s("OPP", "Ball Win Removed Opponents")),
                bt("Ball Wins (def removed)", s("NCC", "Ball Win Removed Opponents Defenders"), s("OPP", "Ball Win Removed Opponents Defenders")),
                bt("Critical Ball Loss", s("NCC", "Critical Ball Loss Number"), s("OPP", "Critical Ball Loss Number"), hb=False),
                bt("Players lost to turnovers", s("NCC", "Ball Loss Removed Teammates"), s("OPP", "Ball Loss Removed Teammates"), hb=False),
            ],
        },
        "boxdom": {
            # Penetration metrics removed per request (box-entry / passes-into-box
            # / crosses / touches-in-box). Empty list keeps the key so dependent
            # JS (guarded by matchVal/null checks) and the panel-hide logic work.
            "penetration": [],
            "finishing": [
                row("xg", "Total xG", c("ncc", "xg", 3), dec=3),
                row("spxg", "Set Piece xG", c("ncc", "spxg", 3), dec=3),
                row("shots", "Shots", c("ncc", "shots")),
                row("sot_pct", "Shots on Target %", c("ncc", "sot_pct", 4), pct=True, dec=3),
                row("xgps", "xG/Shot", c("ncc", "xgps", 3), dec=3),
                row("psxg", "Post-Shot xG", c("ncc", "psxg", 2), dec=2),
                row("qc", "Quality Chances", c("ncc", "qc")),
                row("bc", "Big Chances", c("ncc", "bc")),
            ],
        },
        "boxres": {
            # Opponent penetration (Denying) metrics removed per request.
            "denying": [],
            "nullifying": [
                row("opp_xg", "Opp. Total xG", c("opp", "xg", 3), hb=False, dec=3, context="opp"),
                row("opp_spxg", "Opp. Set Piece xG", c("opp", "spxg", 3), hb=False, dec=3, context="opp"),
                row("opp_shots", "Opp. Shots", c("opp", "shots"), hb=False, context="opp"),
                row("opp_sot_pct", "Opp. Shots on Target %", c("opp", "sot_pct", 4), hb=False, pct=True, dec=3, context="opp"),
                row("opp_xgps", "Opp. xG/Shot", c("opp", "xgps", 3), hb=False, dec=3, context="opp"),
                row("opp_psxg", "Opp. Post-Shot xG", c("opp", "psxg", 2), hb=False, dec=2, context="opp"),
                row("opp_qc", "Opp. Quality Chances", c("opp", "qc"), hb=False, context="opp"),
                row("opp_bc", "Opp. Big Chances", c("opp", "bc"), hb=False, context="opp"),
            ],
        },
        "defint": {
            "volume": [
                row("pressures", "Pressures (PA)", c("ncc", "pressures_raw", 1, adj=True), dec=1),
                row("cp", "Counterpressures (PA)", c("ncc", "cp_raw", 1, adj=True), dec=1),
                row("agg_actions", "Aggressive Actions (PA)", c("ncc", "agg_raw", 1, adj=True), dec=1),
                row("tack_int", "Tackles + Interceptions (PA)", c("ncc", "tack_int_raw", 1, adj=True), dec=1),
                row("press_regains", "Press Regains (PA)", c("ncc", "press_regains", 1, adj=True), dec=1),
            ],
            "forward": [
                row("press_oh", "Pressures in Opp. Half (PA)", c("ncc", "press_oh_raw", 1, adj=True), dec=1),
                row("press_oh_pct", "Pressures in Opp. Half %", c("ncc", "press_oh_pct", 4), pct=True, dec=3),
                row("cp_oh", "Counterpressures in Opp. Half (PA)", c("ncc", "cp_oh_raw", 1, adj=True), dec=1),
            ],
            # "Progression Allowed" and "Opposition Intensity" sections removed
            # per request. Empty lists keep the keys so the render loop hides
            # both panels cleanly.
            "progallowed": [],
            "oppintensity": [],
            # NCC vs Opponent pressing (from the Match CSVs, both teams). Raw
            # per-match counts — the Pressing Battle deliberately does NOT
            # possession-adjust, so a team on the back foot reads as pressing
            # more, which is the intended head-to-head story.
            "battle": [
                bt("Pressures", c("ncc", "pressures_raw"), c("opp", "pressures_raw")),
                bt("Counterpressures", c("ncc", "cp_raw"), c("opp", "cp_raw")),
                bt("Tackles + Interceptions", c("ncc", "tack_int_battle"), c("opp", "tack_int_battle")),
                bt("Ball Recoveries", c("ncc", "ball_recoveries"), c("opp", "ball_recoveries")),
                bt("Pressures in Opp. Half", c("ncc", "press_oh_raw"), c("opp", "press_oh_raw")),
                bt("Counterpressures in Opp. Half", c("ncc", "cp_oh_raw"), c("opp", "cp_oh_raw")),
            ],
        },
        "obv": {
            "ncc": [
                row("obv_total", "Total OBV", c("ncc", "obv_total", 3), dec=3),
                row("obv_pass", "Pass OBV", c("ncc", "obv_pass", 3), dec=3),
                row("obv_shot", "Shot OBV", c("ncc", "obv_shot", 3), dec=3),
                row("obv_dribble", "Dribble & Carry OBV", c("ncc", "obv_dribble", 3), dec=3),
                row("obv_def", "Defensive Action OBV", c("ncc", "obv_def", 3), dec=3),
            ],
            "opp": [
                row("opp_obv_total", "Opp. Total OBV", c("opp", "obv_total", 3), hb=False, dec=3, context="opp"),
                row("opp_obv_pass", "Opp. Pass OBV", c("opp", "obv_pass", 3), hb=False, dec=3, context="opp"),
                row("opp_obv_shot", "Opp. Shot OBV", c("opp", "obv_shot", 3), hb=False, dec=3, context="opp"),
                row("opp_obv_dribble", "Opp. Dribble & Carry OBV", c("opp", "obv_dribble", 3), hb=False, dec=3, context="opp"),
                row("opp_obv_def", "Opp. Defensive Action OBV", c("opp", "obv_def", 3), hb=False, dec=3, context="opp"),
            ],
            "net": [
                # Net = Courage minus opponent, per match. OBV is already a
                # possession-neutral value model, so the difference is the
                # cleanest single read of who controlled the match.
                row("net_obv_total", "Net Total OBV", _net(matches, "obv_total"), dec=3),
                row("net_obv_pass", "Net Pass OBV", _net(matches, "obv_pass"), dec=3),
                row("net_obv_shot", "Net Shot OBV", _net(matches, "obv_shot"), dec=3),
                row("net_obv_dribble", "Net Dribble & Carry OBV", _net(matches, "obv_dribble"), dec=3),
                row("net_obv_def", "Net Defensive Action OBV", _net(matches, "obv_def"), dec=3),
            ],
        },
    }
    return DATA, opp_poss


def build_match_arrays(matches):
    SM = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    marr, miarr = [], []
    for m in matches:
        d = m["date"]
        label = f"GW{m['game_week']} · M{m['game_number']}"
        result = m["result"] or "—"
        marr.append({"label": label, "date": f"{SM[d.month-1]} {d.day}",
                     "opp": m["opp"], "loc": m["loc"], "result": result})
        miarr.append({"label": label, "date": f"{SM[d.month-1]} {d.day}, {d.year}",
                      "opp": m["opp"], "loc": m["loc"],
                      "result": result.replace("-", "–") if result != "—" else result})
    return marr, miarr


def build_match_report(matches, sheet_data, opp_poss):
    entries = []
    for i, m in enumerate(matches):
        gn = m["game_number"]
        ov = MANUAL_OVERRIDES.get(f"M{gn}", {})
        b = m.get("csv")
        ncc = b["ncc"] if b else {}
        opp = b["opp"] if b else {}
        ng, og = ncc.get("goals"), opp.get("goals")
        if ng is not None and og is not None:
            letter = "W" if ng > og else ("L" if ng < og else "D")
        else:
            letter = (m["result"] or "?")[0]
        ncc_poss = round((1 - opp_poss[i]) * 100)
        entries.append({
            "ncc_goals": ng if ng is not None else 0,
            "opp_goals": og if og is not None else 0,
            "result": letter,
            "ncc_xg": safe_round(ncc.get("xg"), 3),
            "opp_xg": safe_round(opp.get("xg"), 3),
            "ncc_xg_sp": safe_round(ncc.get("spxg"), 3),
            "opp_xg_sp": safe_round(opp.get("spxg"), 3),
            "ncc_shots": ncc.get("shots"), "opp_shots": opp.get("shots"),
            "ncc_sot": ncc.get("sot", ov.get("ncc_sot")),
            "opp_sot": opp.get("sot", ov.get("opp_sot")),
            "ncc_poss": ncc_poss, "opp_poss": 100 - ncc_poss,
            "ncc_pass_pct": round(ncc["pass_pct"] * 100) if ncc.get("pass_pct") else None,
            "opp_pass_pct": round(opp["pass_pct"] * 100) if opp.get("pass_pct") else None,
            "ncc_pressures": ncc.get("pressures_raw"),
            "opp_pressures": opp.get("pressures_raw"),
            "ncc_press_regains": ncc.get("press_regains", ov.get("ncc_press_regains")),
            "opp_press_regains": opp.get("press_regains", ov.get("opp_press_regains")),
            "top_shooters": ov.get("top_shooters"),
            "top_kp": ov.get("top_kp"),
            "note": ov.get("note"),
        })
        info(f"M{gn} insights: {letter} {ng}-{og} xG {entries[-1]['ncc_xg']}"
             f"-{entries[-1]['opp_xg']} SOT {entries[-1]['ncc_sot']}"
             f"-{entries[-1]['opp_sot']}")
    return entries


def main():
    print("=" * 70)
    print("Courage Team Performance Dashboard — CSV-first Build")
    print("=" * 70)
    matches, sheet_data = read_spreadsheet()
    csv_matches = load_csv_matches()
    print(f"\n[CSV] {len(csv_matches)} match CSVs found in {CSV_DIR.name}/")
    attach_csv_to_matches(matches, csv_matches)

    # Fill result from CSV goals when the sheet's Goals column is blank
    for m in matches:
        if m["result"] is None and m.get("csv"):
            ng = m["csv"]["ncc"]["goals"]
            og = m["csv"]["opp"]["goals"]
            letter = "W" if ng > og else ("L" if ng < og else "D")
            m["result"] = f"{letter} {ng}-{og}"

    DATA, opp_poss = build_data(matches, sheet_data)
    marr, miarr = build_match_arrays(matches)
    mrd = build_match_report(matches, sheet_data, opp_poss)
    segments = build_bypass_segments(matches, sheet_data)

    html = TEMPLATE.read_text(encoding="utf-8")
    # Fill header tokens from the data so the count/date never go stale.
    # marr is sorted by game number, so the last entry is the latest match.
    html = html.replace("{{MATCH_COUNT}}", str(len(marr)))
    html = html.replace("{{UPDATED_DATE}}", miarr[-1]["date"] if miarr else "—")
    html = replace_block(html, "const MATCHES = [", "];", render_matches_block(marr))
    html = replace_block(html, "const DATA = {", "};", render_data_block(DATA))
    html = replace_block(html, "const MATCH_INFO = [", "];", render_match_info_block(miarr))
    html = replace_block(html, "const MATCH_REPORT_DATA = [", "];", render_match_report_data_block(mrd))
    html = replace_block(html, "const BYPASS_SEGMENTS = [", "];", render_bypass_segments_block(segments))
    OUTPUT.write_text(html, encoding="utf-8")
    print(f"\n✅ Wrote {OUTPUT.name} ({len(html):,} chars)")


if __name__ == "__main__":
    main()
