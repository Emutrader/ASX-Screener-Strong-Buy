# ASX Strong Buy Stock Screener

Generates an Excel report (`ASX 200` tab + `Penny Stocks` tab) with price,
valuation ratios, analyst sentiment, revenue/profit history, and a business
notes field for each ASX-listed stock in your watchlists. Built to run
automatically every weekday morning before 7:00am (Brisbane time).

**Read `screener.py`'s module docstring and the "Read Me - Legend" tab in
every report before relying on this data — they explain exactly which
fields are genuine analyst forecasts vs. approximations, and where free
data sources simply don't have coverage. This tool is not financial advice.**

## What it does

1. Reads two editable ticker lists: `tickers_asx200.csv` and
   `tickers_penny_stocks.csv`.
2. Pulls price, 52-week range, valuation ratios, analyst target price,
   analyst recommendation, and 4-5 years of revenue/net profit for each
   ticker via Yahoo Finance (`yfinance` — free, no API key required).
3. Writes a formatted `.xlsx` with Strong-Buy rows highlighted green, an
   AutoFilter on every column, and a legend tab explaining the data.

## One-time setup (recommended: GitHub Actions — free, cloud-based)

You don't need your own computer running for this option — it runs on
GitHub's servers on schedule.

1. Create a free GitHub account if you don't have one, and create a new
   **private** repository (e.g. `asx-screener`).
2. Upload all the files in this folder to that repo, keeping the folder
   structure (the `.github/workflows/asx_screener.yml` file must stay at
   that exact path).
3. In the repo, go to **Settings → Actions → General → Workflow
   permissions**, and select **"Read and write permissions"**. This lets
   the workflow commit each day's report back into the `/reports` folder
   of your repo.
4. That's it. The workflow will run automatically at 6:30am AEST every
   Monday–Friday and commit that day's `.xlsx` file into `/reports/`.
   You can also trigger a run manually any time from the repo's
   **Actions** tab → "ASX Strong Buy Screener" → **Run workflow**.
5. Each run also uploads the report as a downloadable "artifact" on the
   Actions run page, in case you'd rather download from there than pull
   from `/reports/`.

### Want it emailed to you instead of just committed to the repo?

That needs a small addition (an SMTP step using a GitHub Actions secret
for your email password/app-password) which isn't included by default
since it requires your mail provider's credentials. Let me know if you'd
like this added and I'll build it in.

## Alternative: run it yourself on a schedule (Windows Task Scheduler / cron)

If you'd rather not use GitHub:

1. Install Python 3.11+ if you don't have it.
2. `pip install -r requirements.txt`
3. Test it manually: `python screener.py` — the report lands in
   `output/`.
4. Schedule it:
   - **Windows Task Scheduler**: create a daily task, repeat Mon-Fri,
     trigger time 6:30am, action = run
     `python C:\path\to\screener.py`, and make sure your PC is on and
     online at that time (it won't run if the machine is asleep/off).
   - **macOS/Linux cron**: `crontab -e` and add
     `30 6 * * 1-5 cd /path/to/asx_screener && /usr/bin/python3 screener.py`

## Editing the stock lists

Open `tickers_asx200.csv` or `tickers_penny_stocks.csv` in Excel or a
text editor — one ASX ticker per line, no `.AX` suffix, no exchange
prefix. The bundled `tickers_asx200.csv` is a starting list; the S&P/ASX
200's actual constituents change every quarter (March/June/September/
December rebalance), so review and update it periodically — the ASX 200
list on marketindex.com.au is a good reference. There's no official
"penny stock" index, so `tickers_penny_stocks.csv` is entirely a
manually curated watchlist — add/remove whatever you want tracked.

## Known limitations (see also the in-workbook legend tab)

- **Free data only.** No paid market-data subscription is used. Yahoo
  Finance's analyst coverage of ASX stocks — especially small caps and
  virtually all penny stocks — is patchy; expect "No data" in many
  Strong Buy / forecast growth cells for smaller names.
- **3/6/12-month forecast growth is an approximation**, not three
  independent analyst forecasts — see the script docstring for exactly
  how it's derived from the single 12-month analyst target price.
- **"Strong Buy" is Yahoo's aggregated analyst score**, not a specific
  broker's rating and not a recommendation from this tool.
- **Revenue/Net Profit history is usually 4 years, not 5** — that's the
  limit of what the free data source exposes.
- **The Notes field is a static company description**, not a live
  forward-looking outlook — no free source publishes that.
- Runtime for ~150-200 tickers is roughly 5-12 minutes depending on
  Yahoo Finance's response time; the GitHub Actions job has a 25-minute
  timeout, which should comfortably cover both tabs.
