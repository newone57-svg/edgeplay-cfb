# EdgePlay CFB Nightly Agent — Setup

Read this once, do it once, then forget it exists.

**Total time: about 25 minutes.** You do not need to know how to code.
You are copying files and clicking buttons.

---

## What you're building

```
   4:00 AM Central, every night
            |
            v
   GitHub runs the script (free, in the cloud, your computer can be off)
            |
            +--> data/team_games.csv     <- the model feed
            +--> data/player_games.csv   <- the props feed
            +--> reports/latest.html     <- your morning breakdown
            |
            v
   You open your Excel workbook at 7 AM.
   Power Query pulls the new numbers in automatically.
```

Your Excel file stays on your computer. The cloud job publishes the numbers;
Excel goes and gets them. Nothing to copy, nothing to paste.

---

## Part 1 — Put the code on GitHub (10 min)

1. Go to **github.com** and make a free account if you don't have one.
2. Click the **+** in the top right → **New repository**.
3. Name it `edgeplay-cfb`. Set it to **Public**. Click **Create repository**.
   - Public matters: it's what lets Excel read the CSVs without a password.
     No personal information goes in this repo — only public game stats.
4. On the next screen click **uploading an existing file**.
5. Drag in `cfb_pull.py` and `SETUP.md`. Click **Commit changes**.
6. Now the workflow file. It has to sit in a specific folder:
   - Click **Add file** → **Create new file**.
   - In the filename box type exactly: `.github/workflows/nightly.yml`
     (typing the slashes creates the folders automatically)
   - Paste in the contents of `nightly.yml`. Click **Commit changes**.

---

## Part 2 — Add your CFBD key (5 min, optional but do it)

Without this, you still get every ESPN box score number. With it, you also get
EPA, success rate, and explosiveness — the stats that actually move a power rating.

1. Go to **collegefootballdata.com/key**, enter your email, get a free key.
2. In your GitHub repo: **Settings** → **Secrets and variables** → **Actions**.
3. Click **New repository secret**.
   - Name: `CFBD_API_KEY`
   - Secret: paste your key
   - Click **Add secret**.

The free tier is 1,000 calls a month. This script uses about 30. You're fine.

---

## Part 3 — Test it right now (2 min)

It's the offseason, so don't test on last night — there are no games.
Test on a real Saturday from last season.

1. In your repo, click the **Actions** tab.
2. Click **EdgePlay CFB Nightly Pull** on the left.
3. Click **Run workflow** (button on the right).
4. In the date box type: `2025-11-15`
5. Click the green **Run workflow**.

Wait 3–5 minutes. Refresh. You want a **green checkmark**.

Then click the **Code** tab and check:
- `data/team_games.csv` exists and has a few hundred rows
- `reports/report_2025-11-15.html` exists

Click into the report file, then click **Raw**, then save it and open it in
your browser. That's your morning report.

**If you got a red X:** click into the failed run, click the "Run the pull"
step, and send me the red error text. That's all I need to fix it.

---

## Part 4 — Wire it into Excel (8 min)

You need the raw link to your CSV. It looks like this — swap in your username:

```
https://raw.githubusercontent.com/YOURUSERNAME/edgeplay-cfb/main/data/team_games.csv
```

Now in your CFB workbook:

1. **Data** tab → **Get Data** → **From Other Sources** → **From Web**.
2. Paste that link. Click **OK**.
3. Excel shows you a preview of the table. Click **Load To...** →
   **Table** → **New worksheet**. Name that sheet `CFB_FEED`.
4. Right-click the new table → **Table** → **External Data Properties**.
   Check **Refresh data when opening the file**. Click OK.

Repeat for `player_games.csv` into a sheet called `PROPS_FEED`.

Done. From here on, opening the workbook pulls last night's numbers.
Point your existing formulas at `CFB_FEED` instead of wherever they read now.

---

## Part 5 — Get the report emailed (optional, later)

The report already lands in the repo every morning. If you want it pushed to
your inbox instead of going to look for it, that's a 10-line addition to the
workflow using a Gmail app password. Ask me when you want it — I left it out
of v1 because it means storing a credential and I didn't want to hand you
something with an extra failure point before the core is proven.

---

## Column order — don't reorder these

Your Power Query maps by column name, but if a name changes the formula
breaks silently. If you ever want a column added or renamed, tell me and
I'll change it in the script rather than in Excel. That keeps one source
of truth.

**team_games.csv** — one row per team per game:

`game_id, date, division, team, opponent, home_away, points_for,
points_against, result, margin, total_yards, pass_yards, rush_yards, plays,
yards_per_play, first_downs, third_down_conv, third_down_att,
third_down_pct, fourth_down_conv, fourth_down_att, turnovers, fumbles_lost,
interceptions_thrown, penalties, penalty_yards, possession_time,
possession_seconds, completions, pass_attempts, sacks_allowed,
sack_yards_lost, rush_attempts, yards_per_rush, epa_per_play, success_rate,
explosiveness, line_yards, pulled_at`

**player_games.csv** — one row per player per game:

`game_id, date, team, opponent, home_away, player, position_group, pass_comp,
pass_att, pass_yards, pass_td, pass_int, rush_att, rush_yards, rush_td,
rush_long, rec, rec_yards, rec_td, rec_long, pulled_at`

---

## Things that will eventually go wrong

**ESPN changes their feed.** It's an unofficial endpoint. It's been stable for
years, but if the job starts failing mid-season, that's usually why. Fix is
a small edit to the script.

**A game gets stat-corrected after we pull.** Rare, and the script de-dupes on
re-run, so just trigger a manual re-pull for that date and it overwrites.

**Want FCS back later?** One line in `cfb_pull.py` controls it. Find the line
that reads `GROUPS = {"FBS": 80}` and change it to `GROUPS = {"FBS": 80, "FCS": 81}`.
Nothing else needs to change. Ask me and I'll walk you through it.

**Cron drift.** GitHub's scheduler is not exact. It usually fires within a few
minutes of 4 AM but can run up to 15 minutes late on a busy morning. Doesn't
matter for a report you read at 7.
