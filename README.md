# ASX 200 Strong Buy Screener — Daily Automated Email

**Files in this package:**
- `ASX200_Strong_Buy_Screener_MANUAL_TEMPLATE.xlsx` — the standalone manual version (fill in yourself, run whenever you like, no setup required)
- `daily_screener.py`, `tickers.csv`, `requirements.txt`, `.github/workflows/daily-screener.yml` — the automated version below, which emails you a fresh copy every weekday morning

Pulls price/52-week/revenue data for your ticker list from Yahoo Finance every
weekday morning, runs the same screener logic as the manual workbook, and
emails you the resulting spreadsheet. Runs on GitHub's free infrastructure —
your computer does not need to be on.

## What you're getting, honestly

- **Yahoo Finance data is automated.** Reliable free source, no terms-of-use issue.
- **Market Index and TradingView are NOT automated.** Neither offers a free
  API, and scraping either on a schedule would breach their terms of use. The
  spreadsheet has two blank orange columns on the "Data Input" tab so you can
  paste in a Market Index or TradingView price yourself for a quick 3-way
  sanity check on any stock the screener flags — takes a few seconds per stock.
- **"Growth Outlook OK?" uses Yahoo's trailing revenue-growth figure**, not a
  true forward analyst forecast — free automatable forward estimates don't
  really exist for ASX stocks. Treat it as a rough proxy and verify manually
  before acting on it.
- This is a research shortlist tool, not financial advice.

## One-time setup (about 10 minutes)

### 1. Create a GitHub account (skip if you have one)
https://github.com/signup — free.

### 2. Create a new repository
- Click "+" → "New repository" → name it e.g. `asx-screener` → Private → Create.

### 3. Upload these files
Upload all files in this folder, **keeping the folder structure** (the
`.github/workflows/daily-screener.yml` file must stay in that exact path —
GitHub only detects scheduled workflows there). Easiest way: on the repo
page, "Add file" → "Upload files", drag the whole folder in.

### 4. Create a Gmail App Password
Regular Gmail passwords won't work for this — you need an "app password":
1. Turn on 2-Step Verification on your Google account (if not already on):
   https://myaccount.google.com/security
2. Go to https://myaccount.google.com/apppasswords
3. Create a new app password (name it anything, e.g. "ASX Screener") → copy
   the 16-character code it gives you.

### 5. Add three secrets to your repo
In your repo: **Settings → Secrets and variables → Actions → New repository
secret**. Add each of these:

| Secret name | Value |
|---|---|
| `GMAIL_ADDRESS` | your full Gmail address |
| `GMAIL_APP_PASSWORD` | the 16-character app password from step 4 |
| `RECIPIENT_EMAIL` | the email address you want the report sent to |

### 6. Test it
Go to the **Actions** tab → "Daily ASX 200 Screener" → **Run workflow** button
→ Run. Watch it go green (takes ~1-2 minutes). Check your inbox — you should
have the spreadsheet. If it fails, click into the run to see which step
errored (almost always a typo'd secret).

### 7. Let it run
Once the test works, you're done — it'll run automatically every weekday
morning (~7:30am Brisbane time, before the ASX opens) and email you the
report. No further action needed.

## Editing the ticker list

`tickers.csv` currently has 20 well-known ASX 200 large caps as a starter set.
To screen the full ASX 200, get the current constituent list free from
https://www.marketindex.com.au/asx200 (updates quarterly — the index is
rebalanced every March/June/September/December) and paste the Ticker/Company
Name into `tickers.csv`, one row per stock, same two-column format.

## Changing the screener thresholds

Edit the constants at the top of `daily_screener.py`:
```python
NEAR_LOW_THRESHOLD = 0.05        # "near 52-week low" = within 5% of it
MIN_REVENUE_GROWTH = 0.00        # min YoY revenue growth to count as "growing"
MIN_FORECAST_GROWTH = 0.05       # min growth-proxy to count as "growth outlook OK"
MIN_SIGNALS_FOR_STRONG_BUY = 3   # how many of the 3 signals required (1-3)
```
You can also just edit these directly in the emailed spreadsheet each day —
the yellow cells on the Screener tab recalculate live in Excel. Editing the
script only changes what's used for the *next* automated run.
