"""
Validation probe: can the raw StatsBomb Match CSV (event-level) reproduce the
aggregated StatsBomb IQ metrics already sitting in NWSL StatsBomb Data.xlsx?

Method: pick a match that exists BOTH as a Match CSV and as an IQ row in the
workbook, compute a core metric set from events, and diff against the IQ values.
Nothing is written to the workbook. This only measures agreement.
"""

import csv
import math
from collections import defaultdict
from pathlib import Path

import openpyxl

DATA_DIR = Path("/Users/tomharris/Desktop/Claude/Projects/Data Organization And Cleaning")
CSV_PATH = DATA_DIR / "Statsbomb Match CSV Files/North Carolina Courage/North Carolina Courage_Washington Spirit_4047603.csv"
WB_PATH = DATA_DIR / "NWSL StatsBomb Data.xlsx"

GOAL = (120.0, 40.0)


def truthy(v):
    return str(v).strip().lower() in ("true", "1", "yes", "t")


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def compute(csv_path):
    raw = list(csv.DictReader(open(csv_path, encoding="utf-8-sig")))
    # CRITICAL: the export is exploded by freeze-frame / formation rows — one CSV
    # line per tracked player per event. Collapse to one line per event id.
    seen = set()
    rows = []
    for r in raw:
        eid = r["id"]
        if eid in seen:
            continue
        seen.add(eid)
        rows.append(r)
    print(f"  [rows: {len(raw)} raw -> {len(rows)} unique events]\n")
    teams = sorted({r["team_name"] for r in rows if r["team_name"]})
    out = {t: defaultdict(float) for t in teams}

    shot_dists = defaultdict(list)

    for r in rows:
        t = r["team_name"]
        if not t:
            continue
        m = out[t]
        et = r["event_type_name"]
        tn = (r.get("type_name") or "").strip()
        oc = (r.get("outcome_name") or "").strip()

        if et == "Shot":
            m["Shots"] += 1
            xg = num(r.get("statsbomb_xg"))
            if xg is not None:
                m["xG"] += xg
            if oc == "Goal":
                m["Goals"] += 1
            if tn == "Penalty":
                m["Penalties"] += 1
                if oc == "Goal":
                    m["Penalty Goals"] += 1
            else:
                m["Non Penalty Shots"] += 1
            if tn == "Open Play":
                m["Open Play Shots"] += 1
                x = num(r.get("location_x"))
                if x is not None and x < 102:
                    m["Open Play Shots Outside Box"] += 1
            x, y = num(r.get("location_x")), num(r.get("location_y"))
            if x is not None and y is not None:
                shot_dists[t].append(math.hypot(GOAL[0] - x, GOAL[1] - y))
                if x < 102:
                    m["Shots Outside Box"] += 1

        elif et == "Pass":
            m["Passes"] += 1
            if oc == "":
                m["Successful Passes"] += 1
            if truthy(r.get("pass_cross")):
                m["Crosses"] += 1
            if tn == "Throw-in":
                m["Throw-ins"] += 1
            if tn == "Corner":
                m["Corners"] += 1
            if tn == "Free Kick":
                m["Free Kicks"] += 1
            if tn not in ("Throw-in", "Corner", "Free Kick", "Goal Kick", "Kick Off"):
                m["Open Play Passes"] += 1
            ex, ey = num(r.get("end_location_x")), num(r.get("end_location_y"))
            if ex is not None and ey is not None and ex >= 102 and 18 <= ey <= 62:
                m["Passes Into Box"] += 1
                if oc == "":
                    m["Successful Passes Into Box"] += 1

        elif et == "Pressure":
            m["Pressures"] += 1
            x = num(r.get("location_x"))
            if x is not None and x >= 60:
                m["Pressures in Opposing Half"] += 1

        elif et == "Interception":
            m["Interceptions"] += 1
        elif et == "Clearance":
            m["Clearances"] += 1
        elif et == "Ball Recovery":
            m["Ball Recoveries"] += 1
        elif et == "Foul Committed":
            m["Fouls"] += 1
        elif et == "Foul Won":
            m["Fouls Won"] += 1
        elif et == "Dribble":
            m["Dribbles"] += 1
            if oc == "Complete":
                m["Successful Dribbles"] += 1
            else:
                m["Failed Dribbles"] += 1
        elif et == "Dribbled Past":
            m["Dribbled Past"] += 1

        if truthy(r.get("counterpress")):
            m["Counterpressures"] += 1

        obv = num(r.get("obv_total_net"))
        if obv is not None:
            m["OBV"] += obv

        # touches in box (any on-ball event located in the box)
        x, y = num(r.get("location_x")), num(r.get("location_y"))
        if x is not None and y is not None and x >= 102 and 18 <= y <= 62:
            if et in ("Pass", "Shot", "Carry", "Ball Receipt*", "Ball Receipt", "Dribble"):
                m["Touches in box"] += 1

    for t in teams:
        m = out[t]
        if m["Passes"]:
            m["Passing%"] = m["Successful Passes"] / m["Passes"]
        if m["Dribbles"]:
            m["Dribble%"] = m["Successful Dribbles"] / m["Dribbles"]
        if m["Shots"]:
            m["xG/Shot"] = m["xG"] / m["Shots"]
        if shot_dists[t]:
            m["Shot Distance"] = sum(shot_dists[t]) / len(shot_dists[t])
        if m["Pressures"]:
            m["Pressures in Opposing Half%"] = m["Pressures in Opposing Half"] / m["Pressures"]
    return out


def load_iq(wb_path, sbd_id):
    wb = openpyxl.load_workbook(wb_path, read_only=True, data_only=True)
    ws = wb["Team Event Data"]
    it = ws.iter_rows(values_only=True)
    hdr = [str(c).strip() if c is not None else "" for c in next(it)]
    idx = {h: i for i, h in enumerate(hdr)}
    found = {}
    for row in it:
        gid = row[idx["Game SBD ID"]]
        if gid is not None and str(gid).strip() == str(sbd_id):
            found[row[idx["Team"]]] = {h: row[i] for h, i in idx.items()}
    return found


def main():
    sbd_id = CSV_PATH.stem.split("_")[-1]
    print(f"Match: {CSV_PATH.stem}  (SBD ID {sbd_id})\n")

    computed = compute(CSV_PATH)
    iq = load_iq(WB_PATH, sbd_id)

    if not iq:
        print("No IQ rows found for that Game SBD ID — cannot validate.")
        return

    for team in sorted(iq):
        if team not in computed:
            print(f"!! {team}: no event rows under this name in CSV")
            continue
        print("=" * 74)
        print(team)
        print("=" * 74)
        print(f"{'Metric':<34}{'IQ':>12}{'Computed':>12}{'Diff':>12}")
        print("-" * 74)
        c = computed[team]
        exact = close = off = missing = 0
        for metric in sorted(c):
            iqv = num(iq[team].get(metric))
            if iqv is None:
                missing += 1
                continue
            cv = c[metric]
            d = cv - iqv
            tol = max(0.02 * abs(iqv), 0.01)
            flag = "OK" if abs(d) <= 0.001 else ("~" if abs(d) <= tol else "XX")
            if flag == "OK":
                exact += 1
            elif flag == "~":
                close += 1
            else:
                off += 1
            print(f"{metric:<34}{iqv:>12.3f}{cv:>12.3f}{d:>+12.3f}  {flag}")
        print("-" * 74)
        print(f"exact={exact}  within2%={close}  mismatched={off}  "
              f"(computed metrics with no IQ counterpart: {missing})\n")


if __name__ == "__main__":
    main()
