"""
Courage Team Performance Dashboard — CSV-first build (v2)
==========================================================

NEW PROCESS (replaces spreadsheet-only build for StatsBomb metrics):

  StatsBomb metrics   → computed from raw event CSVs in ./Match CSV Files/
                        (xG, SPxG, PSxG, shots, SOT, pressing, box metrics,
                        OBV, deep progressions, press regains, results)
  Fixtures (date/GW)  → NWSL StatsBomb Data.xlsx "Team Event Data"
  Packing / LB        → NWSL Impect Data.xlsx (unchanged)
  Zone metrics (SOP)  → NWSL InHouse Data.xlsx (unchanged)
  QC / Big Chances    → spreadsheet columns (video-coded, unchanged)
  Possession % (PA)   → NWSL StatsBomb Data.xlsx only (no CSV fallback)
  Narrative text      → MANUAL_OVERRIDES in build_dashboard.py (unchanged)

Anything missing from a provider simply renders as pending/blank.

WEEKLY WORKFLOW:
  1. Drop the new match CSV into ./Match CSV Files/
     (filename: "<HomeTeam>_<AwayTeam>_<match_id>.csv")
  2. Add the match row (Date/Game Week/Match/Team) to the StatsBomb xlsx —
     metric columns may stay blank; only fixture info is needed now.
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
    build_bypass_segments, render_bypass_segments_block, render_matches_block,
    render_match_info_block, render_data_block, render_match_report_data_block,
    replace_block, safe_round, safe_div, info, warn,
)

SCRIPT_DIR = Path(__file__).resolve().parent
# Raw event CSVs live in the shared data folder (sibling of this project),
# alongside the three provider spreadsheets — not beside the script.
CSV_DIR = SCRIPT_DIR.parent / "Data Organization And Cleaning" / "Statsbomb Match CSV Files"
TEMPLATE = SCRIPT_DIR / "template_dashboard.html"
OUTPUT = SCRIPT_DIR / "Courage_Team_Performance_Dashboard.html"

NCC = "North Carolina Courage"
BOX_X, BOX_Y_LO, BOX_Y_HI = 102.0, 18.0, 62.0
FINAL_THIRD_X = 80.0
HALFWAY_X = 60.0
ON_TARGET = {"Goal", "Saved", "Saved to Post", "Saved To Post"}
PRESS_REGAIN_WINDOW = 5.0


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
        # Set-piece shot: direct free kick / corner, any shot in a set-piece
        # play pattern (corner, free kick, throw-in), or flagged SP phase.
        # Penalties are excluded (reported separately by convention).
        if s.get("type_name") == "Penalty":
            return False
        return s.get("type_name") in {"Free Kick", "Corner"} \
            or s.get("play_pattern_name") in {"From Corner", "From Free Kick",
                                              "From Throw In"} \
            or (s.get("set_piece_phase") or "").strip() != ""
    m["spxg"] = sum(_f(s, "statsbomb_xg") or 0 for s in shots if is_set_piece(s))
    m["xgps"] = safe_div(m["xg"], m["shots"])
    on_t = [s for s in shots if s.get("outcome_name") in ON_TARGET]
    m["sot"] = len(on_t)
    m["psxg"] = sum(_f(s, "gk_save_difficulty_xg") or 0 for s in on_t)

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

    passes = [e for e in te if e["event_type_name"] == "Pass"]
    comp = [p for p in passes if not (p.get("outcome_name") or "").strip()]
    m["pass_pct"] = safe_div(len(comp), len(passes))
    carries = [e for e in te if e["event_type_name"] == "Carries"]

    def entered_box(ev):
        return in_box(_f(ev, "end_location_x"), _f(ev, "end_location_y")) \
            and not in_box(_f(ev, "location_x"), _f(ev, "location_y"))

    box_pass = [p for p in passes if entered_box(p)]
    box_pass_c = [p for p in box_pass if not (p.get("outcome_name") or "").strip()]
    box_carry = [c for c in carries if entered_box(c)]
    m["bea_cp"] = len(box_pass)
    m["bea_carry"] = len(box_carry)
    m["sbe"] = len(box_pass_c) + len(box_carry)                  # *

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
    m["obv_def"] = obv({"Duel", "Block", "Interception", "Foul Committed"})
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

def build_data(matches, sheet_data):
    n = len(matches)
    # Possession comes exclusively from NWSL StatsBomb Data.xlsx ("Possession %"
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
            "progression": [
                row("Z2_att", "Zone 2 Attempts", z2a),
                row("Z2_comp", "Zone 2 Completions", z2c),
                row("Z2_pct", "Zone 2 Completion %", pct(z2c, z2a), pct=True, dec=3),
                row("Z3_att", "Zone 3 Attempts", z3a),
                row("Z3_comp", "Zone 3 Completions", z3c),
                row("Z3_pct", "Zone 3 Completion %", pct(z3c, z3a), pct=True, dec=3),
            ],
            "finishing": [
                row("PZ_att", "Pass Zone Attempts", pza),
                row("PZ_comp", "Pass Zone Completions", pzc),
                row("PZ_pct", "Pass Zone Completion %", pct(pzc, pza), pct=True, dec=3),
                row("SZ_att", "Shoot Zone Attempts", sza),
                row("SZ_comp", "Shoot Zone Completions", szc),
                row("SZ_pct", "Shoot Zone Completion %", pct(szc, sza), pct=True, dec=3),
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
            "linebreak": [
                row("lb_comp_f3", "LB Passes Completed in Final Third", s("NCC", "Line Breaking Passes Completed in Final Third")),
                row("lb_pct_f3", "LB Completion % (Final Third)", spct("NCC", "Line Breaking Passes Completed in Final Third%"), pct=True, dec=3),
                row("lb_obv_f3", "LB Passes OBV (Final Third)",
                    [safe_round(v, 3) if v is not None else None
                     for v in s("NCC", "Line Breaking Passes On Ball Value in Final Third")], dec=3),
            ],
        },
        "boxdom": {
            "penetration": [
                row("bea_cp", "Box Entries Attempted (Cross&Pass)", c("ncc", "bea_cp")),
                row("bea_carry", "Box Entries Attempted (Carry)", c("ncc", "bea_carry")),
                row("sbe", "Successful Box Entries", c("ncc", "sbe")),
                row("tib", "Touches in Box", c("ncc", "tib")),
            ],
            "finishing": [
                row("xg", "Total xG", c("ncc", "xg", 3), dec=3),
                row("spxg", "Set Piece xG", c("ncc", "spxg", 3), dec=3),
                row("shots", "Shots", c("ncc", "shots")),
                row("xgps", "xG/Shot", c("ncc", "xgps", 3), dec=3),
                row("psxg", "Post-Shot xG", c("ncc", "psxg", 2), dec=2),
                row("qc", "Quality Chances", s("NCC", "Quality Chances")),
                row("bc", "Big Chances", s("NCC", "Big Chances")),
            ],
        },
        "boxres": {
            "denying": [
                row("opp_bea_cp", "Opp. Box Entries Attempted (Cross&Pass)", c("opp", "bea_cp"), hb=False, context="opp"),
                row("opp_bea_carry", "Opp. Box Entries Attempted (Carry)", c("opp", "bea_carry"), hb=False, context="opp"),
                row("opp_sbe", "Opp. Successful Box Entries", c("opp", "sbe"), hb=False, context="opp"),
                row("opp_tib", "Opp. Touches in Box", c("opp", "tib"), hb=False, context="opp"),
            ],
            "nullifying": [
                row("opp_xg", "Opp. Total xG", c("opp", "xg", 3), hb=False, dec=3, context="opp"),
                row("opp_spxg", "Opp. Set Piece xG", c("opp", "spxg", 3), hb=False, dec=3, context="opp"),
                row("opp_shots", "Opp. Shots", c("opp", "shots"), hb=False, context="opp"),
                row("opp_xgps", "Opp. xG/Shot", c("opp", "xgps", 3), hb=False, dec=3, context="opp"),
                row("opp_psxg", "Opp. Post-Shot xG", c("opp", "psxg", 2), hb=False, dec=2, context="opp"),
                row("opp_qc", "Opp. Quality Chances", s("OPP", "Quality Chances"), hb=False, context="opp"),
                row("opp_bc", "Opp. Big Chances", s("OPP", "Big Chances"), hb=False, context="opp"),
            ],
        },
        "defint": {
            "volume": [
                row("pressures", "Pressures (PA)", c("ncc", "pressures_raw", 1, adj=True), dec=1),
                row("cp", "Counterpressures (PA)", c("ncc", "cp_raw", 1, adj=True), dec=1),
                row("agg_actions", "Aggressive Actions (PA)", c("ncc", "agg_raw", 1, adj=True), dec=1),
                row("tack_int", "Tackles + Interceptions (PA)", c("ncc", "tack_int_raw", 1, adj=True), dec=1),
            ],
            "forward": [
                row("press_oh", "Pressures in Opp. Half (PA)", c("ncc", "press_oh_raw", 1, adj=True), dec=1),
                row("press_oh_pct", "Pressures in Opp. Half %", c("ncc", "press_oh_pct", 4), pct=True, dec=3),
                row("cp_oh", "Counterpressures in Opp. Half (PA)", c("ncc", "cp_oh_raw", 1, adj=True), dec=1),
            ],
            "progallowed": [
                row("opp_dp", "Opp. Final 3rd Entries", c("opp", "deep_prog"), hb=False, context="opp"),
                row("opp_lb_comp_f3", "Opp. LB Completed (Final Third)", s("OPP", "Line Breaking Passes Completed in Final Third"), hb=False, context="opp"),
                row("opp_lb_pct", "Opp. LB Completion % (F3)", spct("OPP", "Line Breaking Passes Completed in Final Third%"), hb=False, pct=True, dec=3, context="opp"),
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
    html = replace_block(html, "const MATCHES = [", "];", render_matches_block(marr))
    html = replace_block(html, "const DATA = {", "};", render_data_block(DATA))
    html = replace_block(html, "const MATCH_INFO = [", "];", render_match_info_block(miarr))
    html = replace_block(html, "const MATCH_REPORT_DATA = [", "];", render_match_report_data_block(mrd))
    html = replace_block(html, "const BYPASS_SEGMENTS = [", "];", render_bypass_segments_block(segments))
    OUTPUT.write_text(html, encoding="utf-8")
    print(f"\n✅ Wrote {OUTPUT.name} ({len(html):,} chars)")


if __name__ == "__main__":
    main()
