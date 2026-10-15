# upload_excel_gsheets

Syncs daily **orders** and **revenue** from the CLA Excel workbook (monthly tabs `CLA <MON> 26 NEW`) to four Google Sheets (SE, NO, DK, FI), writing into the `Daily` tab on the row whose date matches.

## How it works
- Reads visible tabs named `CLA <MONTH> <YY> NEW` (old/hidden tabs are ignored).
- Column positions differ between months, so revenue/orders are located from the headers: row 2 = market (Sweden/Norway/Denmark/Finland), row 3 = `Real Revenue SEK|NOK|DKK|€` and `Real Orders`. SE/NO/DK revenue is in local currency, FI in €.
- Only the last `lookback_days` (default 15) up to today are synced. Days with empty cells (future days) are skipped.
- On the Google side only the orders (I) and revenue (J) cells of the matching date row (date in D) are written. Dates not found in the sheet are logged and skipped, never appended.
- Works on a temp copy of the file, so it is safe while Excel has it open. Re-running is idempotent.

## Setup
1. `pip install -r requirements.txt`
2. Google Cloud: create a project, enable the **Google Sheets API**, create a **service account**, download its JSON key as `service_account.json` (keep it private, it is git-ignored).
3. Share each of the 4 Google Sheets with the service account email (`...@...iam.gserviceaccount.com`) as **Editor**. "Public" view access is not enough to write.
4. `cp config.example.json config.json` and set `excel_path`.

## Run
```
python sync_excel_to_gsheets.py --dry-run          # show what would be written
python sync_excel_to_gsheets.py                    # sync last 15 days
python sync_excel_to_gsheets.py --force            # ignore the "unchanged file" shortcut
python sync_excel_to_gsheets.py --from 2026-01-01  # backfill
```
The date column of the `Daily` tab may be a real date or text (`YYYY-MM-DD`, `DD/MM/YYYY`, `DD.MM.YYYY`); check with `--dry-run`.

## Schedule
- Linux cron (daily 07:00): `0 7 * * * cd /opt/sync && python3 sync_excel_to_gsheets.py >> sync.log 2>&1`
- Windows Task Scheduler: daily trigger, action `python.exe C:\sync\sync_excel_to_gsheets.py`, "Start in" `C:\sync`.
