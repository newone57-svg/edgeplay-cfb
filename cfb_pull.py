#!/usr/bin/env python3
"""
EdgePlay Analytics - NCAA CFB Nightly Box Score Pull
====================================================
Pulls completed FBS + FCS box scores for a given date from ESPN's public
JSON feed (no key needed), optionally enriches with CollegeFootballData
advanced stats, and writes three things:

    data/team_games.csv    - one row per team per game  (model feed)
    data/player_games.csv  - one row per player per game (props feed)
    reports/report_YYYY-MM-DD.html - the morning breakdown

Run it:
    python cfb_pull.py                    # yesterday's games
    python cfb_pull.py --date 2025-11-15  # a specific date (for testing)

Everything appends to the season master CSVs and de-dupes, so re-running
the same date twice is harmless.
"""

import argparse
import datetime as dt
import json
import os
import sys
import time

import pandas as pd
import requests

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

ESPN_BASE = "https://site.api.espn.com/apis/site/v2/sports/football/college-football"
CFBD_BASE = "https://api.collegefootballdata.com"

# ESPN division groups: 80 = FBS, 81 = FCS.
# FBS only for now. To add FCS back later, just add "FCS": 81 to this line.
GROUPS = {"FBS": 80}

DATA_DIR = "data"
REPORT_DIR = "reports"

VERSION = "2.2"   # printed in the log so you can always tell which
                  # version actually ran. If the log does not say 2.1,
                  # GitHub is running an older copy of this file.

HEADERS = {"User-Agent": "EdgePlayAnalytics/1.0"}

# Column order for the model feed. Do NOT reorder without updating the
# Excel Power Query - it maps by position as a fallback.
TEAM_COLS = [
    "game_id", "date", "division", "team", "opponent", "home_away",
    "points_for", "points_against", "result", "margin",
    "total_yards", "pass_yards", "rush_yards", "plays", "yards_per_play",
    "first_downs", "third_down_conv", "third_down_att", "third_down_pct",
    "fourth_down_conv", "fourth_down_att",
    "turnovers", "fumbles_lost", "interceptions_thrown",
    "penalties", "penalty_yards", "possession_time", "possession_seconds",
    "completions", "pass_attempts", "sacks_allowed", "sack_yards_lost",
    "rush_attempts", "yards_per_rush",
    "epa_per_play", "success_rate", "explosiveness", "line_yards",
    "pulled_at",
]

PLAYER_COLS = [
    "game_id", "date", "division", "team", "opponent", "home_away", "player", "jersey",
    "roles",
    # passing
    "pass_comp", "pass_att", "pass_yards", "pass_avg", "pass_td", "pass_int", "qbr",
    # rushing
    "rush_att", "rush_yards", "rush_avg", "rush_td", "rush_long",
    # receiving
    "rec", "rec_yards", "rec_avg", "rec_td", "rec_long",
    # fumbles
    "fum", "fum_lost", "fum_rec",
    # defense
    "tackles_tot", "tackles_solo", "sacks", "tfl", "pass_defended", "qb_hits", "def_td",
    # interceptions
    "int_made", "int_yards", "int_td",
    # returns
    "kr_no", "kr_yards", "kr_avg", "kr_long", "kr_td",
    "pr_no", "pr_yards", "pr_avg", "pr_long", "pr_td",
    # kicking / punting
    "fg_made", "fg_att", "fg_pct", "fg_long", "xp_made", "xp_att", "kick_pts",
    "punts", "punt_yards", "punt_avg", "punt_tb", "punt_in20", "punt_long",
    "pulled_at",
]

# Maps ESPN's stat keys -> our column names, per category.
# "PAIR:a|b" means the value looks like "23/37" or "1-2" and splits into two columns.
CATEGORY_MAP = {
    "passing": {"label": "PASS", "keys": {
        "c/att": "PAIR:pass_comp|pass_att", "yds": "pass_yards", "avg": "pass_avg",
        "td": "pass_td", "int": "pass_int", "qbr": "qbr",
        "completions/passingattempts": "PAIR:pass_comp|pass_att",
        "passingyards": "pass_yards", "yardsperpassattempt": "pass_avg",
        "passingtouchdowns": "pass_td", "interceptions": "pass_int",
        "adjqbr": "qbr", "qbrating": "qbr"}},
    "rushing": {"label": "RUSH", "keys": {
        "car": "rush_att", "yds": "rush_yards", "avg": "rush_avg",
        "td": "rush_td", "long": "rush_long",
        "rushingattempts": "rush_att", "rushingyards": "rush_yards",
        "yardsperrushattempt": "rush_avg", "rushingtouchdowns": "rush_td",
        "longrushing": "rush_long"}},
    "receiving": {"label": "REC", "keys": {
        "rec": "rec", "yds": "rec_yards", "avg": "rec_avg",
        "td": "rec_td", "long": "rec_long",
        "receptions": "rec", "receivingyards": "rec_yards",
        "yardsperreception": "rec_avg", "receivingtouchdowns": "rec_td",
        "longreception": "rec_long"}},
    "fumbles": {"label": "FUM", "keys": {
        "fum": "fum", "lost": "fum_lost", "rec": "fum_rec"}},
    "defensive": {"label": "DEF", "keys": {
        "tot": "tackles_tot", "solo": "tackles_solo", "sacks": "sacks", "tfl": "tfl",
        "pd": "pass_defended", "qb hts": "qb_hits", "td": "def_td",
        "totaltackles": "tackles_tot", "solotackles": "tackles_solo",
        "tacklesforloss": "tfl", "passesdefended": "pass_defended",
        "qbhits": "qb_hits", "defensivetouchdowns": "def_td"}},
    "interceptions": {"label": "INT", "keys": {
        "int": "int_made", "yds": "int_yards", "td": "int_td",
        "interceptions": "int_made", "interceptionyards": "int_yards",
        "interceptiontouchdowns": "int_td"}},
    "kickReturns": {"label": "KR", "keys": {
        "no": "kr_no", "yds": "kr_yards", "avg": "kr_avg",
        "long": "kr_long", "td": "kr_td"}},
    "puntReturns": {"label": "PR", "keys": {
        "no": "pr_no", "yds": "pr_yards", "avg": "pr_avg",
        "long": "pr_long", "td": "pr_td"}},
    "kicking": {"label": "K", "keys": {
        "fg": "PAIR:fg_made|fg_att", "pct": "fg_pct", "long": "fg_long",
        "xp": "PAIR:xp_made|xp_att", "pts": "kick_pts"}},
    "punting": {"label": "P", "keys": {
        "no": "punts", "yds": "punt_yards", "avg": "punt_avg",
        "tb": "punt_tb", "in 20": "punt_in20", "long": "punt_long"}},
}

# Display order + column headers for the full box score in the HTML report
REPORT_SECTIONS = [
    ("passing", "Passing", ["pass_comp/pass_att", "pass_yards", "pass_avg", "pass_td", "pass_int", "qbr"],
     ["C/ATT", "YDS", "AVG", "TD", "INT", "QBR"]),
    ("rushing", "Rushing", ["rush_att", "rush_yards", "rush_avg", "rush_td", "rush_long"],
     ["CAR", "YDS", "AVG", "TD", "LONG"]),
    ("receiving", "Receiving", ["rec", "rec_yards", "rec_avg", "rec_td", "rec_long"],
     ["REC", "YDS", "AVG", "TD", "LONG"]),
    ("fumbles", "Fumbles", ["fum", "fum_lost", "fum_rec"], ["FUM", "LOST", "REC"]),
    ("defensive", "Defense", ["tackles_tot", "tackles_solo", "sacks", "tfl", "pass_defended", "def_td"],
     ["TOT", "SOLO", "SACKS", "TFL", "PD", "TD"]),
    ("interceptions", "Interceptions", ["int_made", "int_yards", "int_td"], ["INT", "YDS", "TD"]),
    ("kickReturns", "Kick Returns", ["kr_no", "kr_yards", "kr_avg", "kr_long", "kr_td"],
     ["NO", "YDS", "AVG", "LONG", "TD"]),
    ("puntReturns", "Punt Returns", ["pr_no", "pr_yards", "pr_avg", "pr_long", "pr_td"],
     ["NO", "YDS", "AVG", "LONG", "TD"]),
    ("kicking", "Kicking", ["fg_made/fg_att", "fg_pct", "fg_long", "xp_made/xp_att", "kick_pts"],
     ["FG", "PCT", "LONG", "XP", "PTS"]),
    ("punting", "Punting", ["punts", "punt_yards", "punt_avg", "punt_tb", "punt_in20", "punt_long"],
     ["NO", "YDS", "AVG", "TB", "IN20", "LONG"]),
]


def log(msg):
    print(f"[{dt.datetime.now():%H:%M:%S}] {msg}", flush=True)


def get_json(url, params=None, headers=None, tries=3):
    """GET with retries. Returns None instead of blowing up the whole run."""
    for attempt in range(1, tries + 1):
        try:
            r = requests.get(url, params=params, headers=headers or HEADERS, timeout=30)
            if r.status_code == 200:
                return r.json()
            log(f"  HTTP {r.status_code} from {url} (attempt {attempt})")
        except Exception as e:
            log(f"  Error on {url}: {e} (attempt {attempt})")
        time.sleep(2 * attempt)
    return None


# ---------------------------------------------------------------------------
# ESPN parsing helpers
# ---------------------------------------------------------------------------

def num(val):
    """Turn ESPN's string stats into numbers. Returns None if not a number."""
    if val is None:
        return None
    s = str(val).strip().replace(",", "")
    if s in ("", "-", "--"):
        return None
    try:
        return float(s) if "." in s else int(s)
    except ValueError:
        return None


def split_pair(val):
    """ESPN gives '5-13' for third downs, '3-25' for sacks. Split into two."""
    if not val:
        return None, None
    s = str(val).strip()
    sep = "/" if "/" in s else "-"
    parts = s.split(sep)
    if len(parts) != 2:
        return None, None
    return num(parts[0]), num(parts[1])


def mmss_to_seconds(val):
    if not val or ":" not in str(val):
        return None
    m, s = str(val).split(":")[:2]
    try:
        return int(m) * 60 + int(s)
    except ValueError:
        return None


def stat_map(team_block):
    """Flatten ESPN's [{name, displayValue}, ...] into a dict."""
    out = {}
    for s in team_block.get("statistics", []) or []:
        key = s.get("name") or s.get("label")
        if key:
            out[key] = s.get("displayValue")
    return out


def get_completed_games(date_str):
    """Return list of (game_id, division) for finished games on a date."""
    yyyymmdd = date_str.replace("-", "")
    games = []
    for div, group in GROUPS.items():
        data = get_json(
            f"{ESPN_BASE}/scoreboard",
            params={"dates": yyyymmdd, "groups": group, "limit": 400},
        )
        if not data:
            log(f"  Could not load {div} scoreboard for {date_str}")
            continue
        events = data.get("events", []) or []
        done = 0
        for ev in events:
            comps = ev.get("competitions", [])
            if not comps:
                continue
            status = comps[0].get("status", {}).get("type", {})
            if status.get("completed") is True:
                games.append((ev["id"], div))
                done += 1
        log(f"  {div}: {done} completed of {len(events)} scheduled")
    return games


def parse_game(game_id, division, date_str):
    """Pull one game's summary and return (team_rows, player_rows)."""
    data = get_json(f"{ESPN_BASE}/summary", params={"event": game_id})
    if not data:
        return [], []

    box = data.get("boxscore", {})
    teams = box.get("teams", []) or []
    if len(teams) != 2:
        return [], []

    # Scores live in the header, not the boxscore
    scores = {}
    homeaway = {}
    for comp in data.get("header", {}).get("competitions", []):
        for c in comp.get("competitors", []):
            tid = str(c.get("id"))
            scores[tid] = num(c.get("score"))
            homeaway[tid] = c.get("homeAway")

    pulled = dt.datetime.utcnow().isoformat(timespec="seconds")
    names = {}
    for t in teams:
        tid = str(t.get("team", {}).get("id"))
        names[tid] = t.get("team", {}).get("displayName")

    team_rows = []
    for t in teams:
        tid = str(t.get("team", {}).get("id"))
        other = [x for x in names if x != tid]
        oid = other[0] if other else None
        st = stat_map(t)

        pf = scores.get(tid)
        pa = scores.get(oid)
        td_conv, td_att = split_pair(st.get("thirdDownEff"))
        fd_conv, fd_att = split_pair(st.get("fourthDownEff"))
        comp_att = st.get("completionAttempts")
        comps_, atts_ = split_pair(comp_att)
        sacks_n, sack_yds = split_pair(st.get("sacksYardsLost"))
        top = st.get("possessionTime")

        row = {
            "game_id": game_id,
            "date": date_str,
            "division": division,
            "team": names.get(tid),
            "opponent": names.get(oid),
            "home_away": homeaway.get(tid),
            "points_for": pf,
            "points_against": pa,
            "result": ("W" if (pf or 0) > (pa or 0) else "L" if (pf or 0) < (pa or 0) else "T")
                      if pf is not None and pa is not None else None,
            "margin": (pf - pa) if (pf is not None and pa is not None) else None,
            "total_yards": num(st.get("totalYards")),
            "pass_yards": num(st.get("netPassingYards")),
            "rush_yards": num(st.get("rushingYards")),
            "plays": None,           # computed below
            "yards_per_play": None,  # computed below
            "first_downs": num(st.get("firstDowns")),
            "third_down_conv": td_conv,
            "third_down_att": td_att,
            "third_down_pct": round(td_conv / td_att * 100, 1) if td_conv is not None and td_att else None,
            "fourth_down_conv": fd_conv,
            "fourth_down_att": fd_att,
            "turnovers": num(st.get("turnovers")),
            "fumbles_lost": num(st.get("fumblesLost")),
            "interceptions_thrown": num(st.get("interceptions")),
            "penalties": split_pair(st.get("totalPenaltiesYards"))[0],
            "penalty_yards": split_pair(st.get("totalPenaltiesYards"))[1],
            "possession_time": top,
            "possession_seconds": mmss_to_seconds(top),
            "completions": comps_,
            "pass_attempts": atts_,
            "sacks_allowed": sacks_n,
            "sack_yards_lost": sack_yds,
            "rush_attempts": num(st.get("rushingAttempts")),
            "yards_per_rush": num(st.get("yardsPerRushAttempt")),
            "epa_per_play": None,      # filled by CFBD if available
            "success_rate": None,
            "explosiveness": None,
            "line_yards": None,
            "pulled_at": pulled,
        }
        # ESPN college doesn't supply these two -- derive them.
        if row["plays"] is None:
            pa_, ra_ = row["pass_attempts"], row["rush_attempts"]
            if pa_ is not None and ra_ is not None:
                row["plays"] = pa_ + ra_
        if row["yards_per_play"] is None and row["plays"] and row["total_yards"] is not None:
            row["yards_per_play"] = round(row["total_yards"] / row["plays"], 2)

        team_rows.append(row)

    # ---- players: ALL ESPN categories -------------------------------------
    player_rows = []
    acc = {}
    unresolved = set()
    for tblock in box.get("players", []) or []:
        tid = str(tblock.get("team", {}).get("id"))
        other = [x for x in names if x != tid]
        oid = other[0] if other else None
        for group in tblock.get("statistics", []) or []:
            gname = group.get("name") or ""
            cat = CATEGORY_MAP.get(gname)
            if not cat:
                continue                      # unknown category, skip quietly
            labels = [str(k).strip().lower() for k in (group.get("labels") or [])]
            keys = [str(k).strip().lower() for k in (group.get("keys") or [])]
            for ath in group.get("athletes", []) or []:
                athlete = ath.get("athlete", {}) or {}
                pname = athlete.get("displayName")
                if not pname:
                    continue
                vals = ath.get("stats", []) or []
                stats = {}
                for namelist in (labels, keys):
                    for k, v in zip(namelist, vals):
                        if k and k not in stats:
                            stats[k] = v

                pid = (tid, pname)
                rec = acc.setdefault(pid, {
                    "game_id": game_id, "date": date_str, "division": division,
                    "team": names.get(tid), "opponent": names.get(oid),
                    "home_away": homeaway.get(tid), "player": pname,
                    "jersey": athlete.get("jersey"), "roles": "",
                    "pulled_at": pulled,
                })
                seen = set(filter(None, rec["roles"].split("/")))
                seen.add(cat["label"])
                rec["roles"] = "/".join(sorted(seen))

                if not any(k in stats for k in cat["keys"]):
                    unresolved.add(f"{gname}:{labels or keys}")
                for espn_key, target in cat["keys"].items():
                    if espn_key not in stats:
                        continue
                    raw = stats[espn_key]
                    if target.startswith("PAIR:"):
                        c1, c2 = target[5:].split("|")
                        a, b = split_pair(raw)
                        rec[c1], rec[c2] = a, b
                    else:
                        rec[target] = num(raw)

    if unresolved:
        log(f"  !! unmapped stat headers in game {game_id}: {sorted(unresolved)}")

    player_rows = list(acc.values())
    return team_rows, player_rows


# ---------------------------------------------------------------------------
# CFBD enrichment (optional - only runs if an API key is present)
# ---------------------------------------------------------------------------

def _norm(name):
    """ESPN says 'Louisville Cardinals', CFBD says 'Louisville'. Normalize both."""
    s = str(name or "").lower().strip()
    for junk in ("(", ")", ".", "'", "-"):
        s = s.replace(junk, " ")
    return " ".join(s.split())


def _team_key(espn_name, cfbd_names):
    """Find which CFBD team name matches an ESPN display name."""
    e = _norm(espn_name)
    best = None
    for c in cfbd_names:
        cn = _norm(c)
        if e == cn or e.startswith(cn + " ") or e == cn:
            # prefer the longest match so "Miami" doesn't beat "Miami OH"
            if best is None or len(cn) > len(_norm(best)):
                best = c
    return best


def enrich_with_cfbd(team_df, date_str):
    """Add per-game EPA / success rate / explosiveness from CollegeFootballData.

    Optional. No key -> skipped entirely, ESPN data is unaffected.
    """
    key = os.environ.get("CFBD_API_KEY", "").strip()
    if not key:
        log("No CFBD_API_KEY set - skipping advanced stats (ESPN data still complete).")
        return team_df

    year = int(date_str[:4])
    if int(date_str[5:7]) <= 2:      # Jan/Feb bowls belong to the prior season
        year -= 1

    log("Fetching CFBD per-game advanced stats...")
    data = get_json(
        f"{CFBD_BASE}/stats/game/advanced",
        params={"year": year, "excludeGarbageTime": "false"},
        headers={**HEADERS, "Authorization": f"Bearer {key}"},
    )
    if not data:
        log("  CFBD returned nothing (bad key, quota, or outage). Continuing without it.")
        return team_df
    if not isinstance(data, list):
        log(f"  CFBD sent an unexpected shape: {type(data).__name__}. Skipping.")
        return team_df

    log(f"  CFBD returned {len(data)} team-game records for {year}.")

    # Index by (team, opponent) so we attach the right GAME, not a season average.
    idx, cfbd_names = {}, set()
    for rec in data:
        t, o = rec.get("team"), rec.get("opponent")
        if not t:
            continue
        cfbd_names.add(t)
        if o:
            cfbd_names.add(o)
            idx[(_norm(t), _norm(o))] = rec

    hits = misses = 0
    unmatched = []
    for i, row in team_df.iterrows():
        ct = _team_key(row["team"], cfbd_names)
        co = _team_key(row["opponent"], cfbd_names)
        rec = idx.get((_norm(ct), _norm(co))) if (ct and co) else None
        if not rec:
            misses += 1
            if len(unmatched) < 5:
                unmatched.append(f"{row['team']} vs {row['opponent']}")
            continue
        off = rec.get("offense") or {}
        hits += 1
        team_df.at[i, "epa_per_play"] = off.get("ppa")
        team_df.at[i, "success_rate"] = off.get("successRate")
        team_df.at[i, "explosiveness"] = off.get("explosiveness")
        team_df.at[i, "line_yards"] = off.get("lineYards")

    log(f"  ADVANCED STATS MATCHED: {hits} of {len(team_df)} team rows.")
    if misses:
        log(f"  {misses} unmatched (name mismatch between ESPN and CFBD). Examples: {unmatched}")
        log("  Unmatched rows keep every ESPN stat -- only the 4 advanced columns stay blank.")
    return team_df


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

def append_dedupe(df, path, key_cols):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    if os.path.exists(path):
        old = pd.read_csv(path)
        df = pd.concat([old, df], ignore_index=True)
    df = df.drop_duplicates(subset=key_cols, keep="last")
    df.to_csv(path, index=False)
    return df


def _fmt(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _combo(row, spec):
    """spec may be 'a/b' meaning show two columns as one cell."""
    if "/" in spec:
        a, b = spec.split("/")
        av, bv = _fmt(row.get(a)), _fmt(row.get(b))
        return f"{av}/{bv}" if (av or bv) else ""
    return _fmt(row.get(spec))


NO_SUM = {"pass_avg", "rush_avg", "rec_avg", "kr_avg", "pr_avg",
          "punt_avg", "fg_pct", "qbr", "rush_long", "rec_long", "fg_long",
          "kr_long", "pr_long", "punt_long"}
MAXES = {"rush_long", "rec_long", "fg_long", "kr_long", "pr_long", "punt_long"}


def _section_table(players, section):
    """One category table for one team, ESPN style. Returns '' if nobody qualifies."""
    gname, title, fields, headers = section
    label = CATEGORY_MAP[gname]["label"]
    rows = [p for p in players if label in str(p.get("roles", "")).split("/")]
    if not rows:
        return ""

    # Sort by the first numeric stat, biggest first
    key0 = fields[0].split("/")[0]
    rows.sort(key=lambda r: (r.get(key0) if isinstance(r.get(key0), (int, float)) else -1), reverse=True)

    body, totals = "", [0] * len(fields)
    can_total = [True] * len(fields)
    for p in rows:
        cells = ""
        for i, f in enumerate(fields):
            cells += f"<td>{_combo(p, f)}</td>"
            base = f.split("/")[0]
            v = p.get(base)
            if base in NO_SUM and base not in MAXES:
                can_total[i] = False        # averages/percentages don't add up
            elif isinstance(v, (int, float)) and not pd.isna(v):
                totals[i] = max(totals[i], v) if base in MAXES else totals[i] + v
            else:
                can_total[i] = False
        jersey = f' <span class="jsy">#{_fmt(p.get("jersey"))}</span>' if p.get("jersey") else ""
        body += f'<tr><td class="pname">{p["player"]}{jersey}</td>{cells}</tr>'

    tcells = "".join(
        f"<td>{_fmt(round(t,1)) if can_total[i] and '/' not in fields[i] else ''}</td>"
        for i, t in enumerate(totals))
    head = "".join(f"<th>{h}</th>" for h in headers)
    return f"""
      <div class="cat">
        <div class="cat-title">{title}</div>
        <table class="box">
          <tr><th class="pcol"></th>{head}</tr>
          {body}
          <tr class="tot"><td class="pname">TEAM</td>{tcells}</tr>
        </table>
      </div>"""


def _game_boxscore(gteam, gplayers):
    """Full two-column box score for one game."""
    a, b = gteam.iloc[0], gteam.iloc[1]
    pa = [p for p in gplayers if p["team"] == a["team"]]
    pb = [p for p in gplayers if p["team"] == b["team"]]

    cols = ""
    for team_row, plist in ((a, pa), (b, pb)):
        secs = "".join(_section_table(plist, s) for s in REPORT_SECTIONS)
        cols += f"""
        <td class="side">
          <div class="side-head">{team_row['team']}
            <span class="score">{_fmt(team_row['points_for'])}</span>
          </div>
          {secs or '<div class="muted">No player stats available.</div>'}
        </td>"""

    def teamline(r):
        return (f"<tr><td class='tl'>{r['team']}</td>"
                f"<td>{_fmt(r['total_yards'])}</td><td>{_fmt(r['pass_yards'])}</td>"
                f"<td>{_fmt(r['rush_yards'])}</td><td>{_fmt(r['plays'])}</td>"
                f"<td>{_fmt(r['yards_per_play'])}</td><td>{_fmt(r['first_downs'])}</td>"
                f"<td>{_fmt(r['third_down_conv'])}-{_fmt(r['third_down_att'])}</td>"
                f"<td>{_fmt(r['turnovers'])}</td>"
                f"<td>{_fmt(r['penalties'])}-{_fmt(r['penalty_yards'])}</td>"
                f"<td>{_fmt(r['possession_time'])}</td></tr>")

    return f"""
    <div class="game">
      <div class="game-head">
        {a['team']} {_fmt(a['points_for'])} &nbsp;—&nbsp; {b['team']} {_fmt(b['points_for'])}
        
      </div>
      <table class="teamstats">
        <tr><th></th><th>TOT YDS</th><th>PASS</th><th>RUSH</th><th>PLAYS</th><th>Y/P</th>
            <th>1ST</th><th>3RD DN</th><th>TO</th><th>PEN</th><th>TOP</th></tr>
        {teamline(a)}{teamline(b)}
      </table>
      <table class="cols"><tr>{cols}</tr></table>
    </div>"""


def build_report(team_df, player_df, date_str):
    """Navy + gold report: summary page, then a full box score for every game."""
    os.makedirs(REPORT_DIR, exist_ok=True)
    path = f"{REPORT_DIR}/report_{date_str}.html"

    precs = player_df.to_dict("records") if len(player_df) else []
    by_game = {}
    for p in precs:
        by_game.setdefault(p["game_id"], []).append(p)

    games, summary_rows = [], ""
    for gid, g in team_df.groupby("game_id"):
        if len(g) != 2:
            continue
        g = g.sort_values("points_for", ascending=False)
        w, l = g.iloc[0], g.iloc[1]
        total = (w["points_for"] or 0) + (l["points_for"] or 0)
        ymarg = (w["total_yards"] or 0) - (l["total_yards"] or 0)
        pmarg = (w["points_for"] or 0) - (l["points_for"] or 0)
        flags = []
        if pmarg >= 21 and ymarg < 75:
            flags.append("SCORE &gt; YARDS")
        if ymarg >= 150 and pmarg <= 7:
            flags.append("YARDS &gt; SCORE")
        if (w["turnovers"] or 0) + (l["turnovers"] or 0) >= 5:
            flags.append("TURNOVER CHAOS")
        if total >= 70:
            flags.append("SHOOTOUT")
        if total <= 27:
            flags.append("ROCK FIGHT")
        games.append({"gid": gid, "div": w["division"], "g": g, "w": w, "l": l,
                      "total": total, "flags": flags})

    games.sort(key=lambda x: -x["total"])

    for i, gm in enumerate(games, 1):
        w, l = gm["w"], gm["l"]
        fl = "".join(f'<span class="flag">{f}</span>' for f in gm["flags"]) or '<span class="muted">—</span>'
        summary_rows += f"""
        <tr><td class="gnum">{i}</td>
        <td><a href="#g{i}"><b>{w['team']}</b> {_fmt(w['points_for'])},
            {l['team']} {_fmt(l['points_for'])}</a></td>
        <td>{gm['total']:.0f}</td>
        <td>{_fmt(w['total_yards'])} / {_fmt(l['total_yards'])}</td>
        <td>{_fmt(w['yards_per_play'])} / {_fmt(l['yards_per_play'])}</td>
        <td>{_fmt(w['turnovers'])} / {_fmt(l['turnovers'])}</td>
        <td>{fl}</td></tr>"""

    boxes = ""
    for i, gm in enumerate(games, 1):
        boxes += f'<a id="g{i}"></a>' + _game_boxscore(gm["g"], by_game.get(gm["gid"], []))

    n = len(games)
    avg_total = (sum(g["total"] for g in games) / n) if n else 0
    avg_margin = (sum((g["w"]["points_for"] or 0) - (g["l"]["points_for"] or 0)
                      for g in games) / n) if n else 0
    tot_to = int(team_df["turnovers"].fillna(0).sum()) if len(team_df) else 0

    html = f"""<!doctype html>
<html><head><meta charset="utf-8">
<title>EdgePlay CFB Breakdown — {date_str}</title>
<style>
  @page {{ size: landscape; margin: .35in; }}
  body {{ font-family: -apple-system, "Segoe UI", Roboto, Helvetica, sans-serif;
         margin:0; padding:20px; color:#10203a; background:#fff; font-size:12px; }}
  a {{ color:inherit; text-decoration:none; }}
  .head {{ background:#0d1f3c; color:#fff; padding:16px 20px; border-radius:5px;
           border-left:6px solid #c8a24a; }}
  .head h1 {{ margin:0; font-size:19px; letter-spacing:.5px; }}
  .head .sub {{ color:#c8a24a; font-size:12px; margin-top:3px; }}
  .kpis {{ display:flex; gap:12px; margin:16px 0; }}
  .kpi {{ flex:1; border:1px solid #dde3ec; border-top:3px solid #c8a24a;
          padding:8px 12px; border-radius:4px; }}
  .kpi .n {{ font-size:20px; font-weight:700; color:#0d1f3c; }}
  .kpi .l {{ font-size:10px; text-transform:uppercase; color:#6b7789; letter-spacing:.6px; }}
  h2 {{ font-size:13px; text-transform:uppercase; letter-spacing:1px; color:#0d1f3c;
        border-bottom:2px solid #c8a24a; padding-bottom:5px; margin:26px 0 10px; }}
  table {{ width:100%; border-collapse:collapse; }}
  .summary th {{ background:#0d1f3c; color:#fff; text-align:left; padding:7px 8px;
                 font-size:10px; text-transform:uppercase; letter-spacing:.5px; }}
  .summary td {{ padding:6px 8px; border-bottom:1px solid #e7ecf3; }}
  .summary tr:nth-child(even) td {{ background:#f7f9fc; }}
  .gnum {{ font-weight:700; color:#c8a24a; font-size:10px; }}
  .flag {{ display:inline-block; background:#0d1f3c; color:#c8a24a; padding:2px 6px;
           border-radius:3px; font-size:9px; margin:0 3px 2px 0; white-space:nowrap; }}
  .muted {{ color:#a8b2c1; }}
  .game {{ page-break-before:always; margin-top:26px; }}
  .game-head {{ background:#0d1f3c; color:#fff; padding:10px 14px; font-size:14px;
                font-weight:700; border-left:5px solid #c8a24a; border-radius:4px; }}
  .gdiv {{ float:right; color:#c8a24a; font-size:10px; letter-spacing:1px; }}
  .teamstats {{ margin:10px 0 14px; font-size:11px; }}
  .teamstats th {{ background:#eef2f8; color:#41506b; padding:5px 6px; font-size:9px;
                   text-transform:uppercase; border-bottom:1px solid #c8a24a; }}
  .teamstats td {{ padding:5px 6px; border-bottom:1px solid #e7ecf3; text-align:center; }}
  .teamstats .tl {{ text-align:left; font-weight:700; }}
  .cols {{ width:100%; border-collapse:separate; border-spacing:9px 0; }}
  .cols > tr > .side {{ width:50%; vertical-align:top; }}
  .side-head {{ background:#f2f5fa; border-left:4px solid #c8a24a; padding:6px 10px;
                font-weight:700; font-size:12px; margin-bottom:6px; }}
  .score {{ float:right; color:#0d1f3c; }}
  .cat {{ margin-bottom:11px; }}
  .cat-title {{ font-size:10px; font-weight:700; text-transform:uppercase;
                color:#0d1f3c; letter-spacing:.7px; margin-bottom:2px; }}
  .box {{ font-size:10.5px; }}
  .box th {{ background:#fafbfd; color:#6b7789; font-size:8.5px; padding:3px 4px;
             text-align:center; border-bottom:1px solid #dde3ec; }}
  .box .pcol {{ text-align:left; width:44%; }}
  .box td {{ padding:3px 4px; text-align:center; border-bottom:1px solid #f0f3f8; }}
  .box .pname {{ text-align:left; font-weight:600; }}
  .jsy {{ color:#a8b2c1; font-weight:400; }}
  .box .tot td {{ background:#f7f9fc; font-weight:700; border-top:1px solid #c8a24a; }}
  .foot {{ margin-top:16px; font-size:10px; color:#6b7789;
           border-top:1px solid #dde3ec; padding-top:8px; }}
</style></head><body>
<div class="head">
  <h1>EDGEPLAY ANALYTICS — NCAA FOOTBALL BREAKDOWN</h1>
  <div class="sub">Games completed {date_str} &nbsp;·&nbsp; FBS &nbsp;·&nbsp; full box scores</div>
</div>
<div class="kpis">
  <div class="kpi"><div class="n">{n}</div><div class="l">FBS Games</div></div>
  <div class="kpi"><div class="n">{avg_total:.1f}</div><div class="l">Avg Total</div></div>
  <div class="kpi"><div class="n">{avg_margin:.1f}</div><div class="l">Avg Margin</div></div>
  <div class="kpi"><div class="n">{tot_to}</div><div class="l">Turnovers</div></div>
  <div class="kpi"><div class="n">{len(precs)}</div><div class="l">Player Lines</div></div>
</div>
<h2>FBS Slate — tap any game to jump to its box score</h2>
<table class="summary">
  <tr><th>#</th><th>Result</th><th>Total</th><th>Yards W/L</th><th>Y/P W/L</th>
      <th>TO W/L</th><th>Flags</th></tr>
  {summary_rows or '<tr><td colspan="7" class="muted">No completed FBS games on this date.</td></tr>'}
</table>
<div class="foot">
  Flags mark box-score disagreements worth a look, not bets. SCORE &gt; YARDS means the
  scoreboard outran the offense — usually turnovers or special teams, and it tends to regress.
  YARDS &gt; SCORE is the reverse. Check against your own ratings before acting.
</div>
{boxes}
</body></html>"""

    with open(path, "w") as f:
        f.write(html)
    return path, n


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", help="YYYY-MM-DD. Defaults to yesterday.")
    args = ap.parse_args()

    date_str = args.date or (dt.date.today() - dt.timedelta(days=1)).isoformat()
    log(f"=== EdgePlay CFB pull v{VERSION} for {date_str} ===")

    games = get_completed_games(date_str)
    if not games:
        log("No completed games found. Writing an empty report and exiting clean.")
        build_report(pd.DataFrame(columns=TEAM_COLS), pd.DataFrame(columns=PLAYER_COLS), date_str)
        return 0

    log(f"Pulling box scores for {len(games)} games...")
    all_team, all_player = [], []
    failed = []
    for i, (gid, div) in enumerate(games, 1):
        t, p = parse_game(gid, div, date_str)
        if not t:
            failed.append(gid)
        all_team.extend(t)
        all_player.extend(p)
        if i % 10 == 0:
            log(f"  {i}/{len(games)} done")
        time.sleep(0.4)          # be polite to ESPN

    if failed:
        log(f"WARNING: {len(failed)} games failed to parse: {failed}")

    team_df = pd.DataFrame(all_team).reindex(columns=TEAM_COLS)
    player_df = pd.DataFrame(all_player).reindex(columns=PLAYER_COLS)

    team_df = enrich_with_cfbd(team_df, date_str)

    append_dedupe(team_df, f"{DATA_DIR}/team_games.csv", ["game_id", "team"])
    append_dedupe(player_df, f"{DATA_DIR}/player_games.csv", ["game_id", "team", "player"])

    # Also drop a single-night file, handy for spot checks
    team_df.to_csv(f"{DATA_DIR}/latest_night.csv", index=False)

    path, n = build_report(team_df, player_df, date_str)
    # Stable filename so the same link always shows last night
    with open(f"{REPORT_DIR}/latest.html", "w") as f:
        f.write(open(path).read())

    statcols = [c for c in PLAYER_COLS if c not in
                ("game_id","date","division","team","opponent","home_away",
                 "player","jersey","roles","pulled_at")]
    filled = int(player_df[statcols].notna().sum().sum()) if len(player_df) else 0
    log(f"Wrote {len(team_df)} team rows, {len(player_df)} player rows.")
    log(f"STAT CELLS FILLED: {filled}   <-- must be in the thousands. Zero means broken.")
    log(f"Report: {path} ({n} games)")
    log("=== done ===")
    return 0


if __name__ == "__main__":
    sys.exit(main())
