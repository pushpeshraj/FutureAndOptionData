# NSE F&O Participant-wise Open Interest — Auto-updating Dashboard

A zero-server dashboard that shows NSE's daily **Participant-wise Open Interest**
report. A scheduled GitHub Action fetches the latest published file each weekday
evening and commits it to the repo; GitHub Pages serves a webpage that displays it.

```
fetch_data.py              → downloads Participant-wise OI, writes data.json + raw CSV
fetch_fii_stats.py         → downloads FII Derivative Statistics, writes fii_stats.json + raw .xls
index.html                 → the webpage (two tabs: Participant OI and FII Stats)
data.json / fii_stats.json → the data each tab renders (refreshed by the Action)
requirements.txt           → Python dependency (requests)
.github/workflows/update.yml → scheduled fetch + auto-commit (both reports)
```

## One-time setup

1. **Create a new GitHub repository** and upload all of these files (keep the
   folder structure — the workflow must stay at `.github/workflows/update.yml`).

2. **Enable GitHub Pages**
   Repo → **Settings → Pages** → *Build and deployment* → Source: **Deploy from a
   branch** → Branch: `main` / `(root)` → **Save**.
   After a minute your page is live at:
   `https://<your-username>.github.io/<repo-name>/`

3. **Allow the Action to commit**
   Repo → **Settings → Actions → General** → *Workflow permissions* →
   select **Read and write permissions** → **Save**.

4. **Run it once manually** to populate real data
   Repo → **Actions** tab → *Update NSE Participant OI* → **Run workflow**.
   When it finishes, refresh your Pages URL.

That's it. From then on it updates itself on the schedule.

## When does it update?

The Action runs at **13:00 and 15:00 UTC, Monday–Friday** (~6:30 PM and ~8:30 PM
IST), just after NSE publishes. If the report isn't out yet on the first run, the
second one catches it. `fetch_data.py` always grabs the **most recent available**
report, so on weekends/holidays the page simply keeps showing the last trading
day's data.

## Run locally (optional)

```bash
pip install -r requirements.txt
python fetch_data.py          # writes data.json + participant_oi_DDMMYYYY.csv
python -m http.server         # then open http://localhost:8000
```

## Good-to-know caveats

- GitHub's scheduled Actions can be delayed a few minutes under load, and GitHub
  **disables scheduled workflows after ~60 days of no repo activity** — just visit
  the Actions tab and re-enable if that happens.
- If NSE changes its archive URL or cookie scheme, update `ARCHIVE_URL` /
  warm-up requests in `fetch_data.py`.
- This is an unofficial mirror for personal use. Always verify against the
  [official NSE reports page](https://www.nseindia.com/all-reports-derivatives).
