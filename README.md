# upload_excel_gsheets

Syncs daily **orders** and **revenue** from the CLA Excel workbook (monthly tabs `CLA <MON> 26 NEW`) to four Google Sheets (SE, NO, DK, FI), writing into the `Daily` tab on the row whose date matches. You run it when you want to push.

## How it works
- Reads visible tabs named `CLA <MONTH> <YY> NEW` (old/hidden tabs are ignored).
- Column positions differ between months, so revenue/orders are located from the headers: row 2 = market (Sweden/Norway/Denmark/Finland), row 3 = `Real Revenue SEK|NOK|DKK|€` and `Real Orders`. SE/NO/DK revenue is in local currency, FI in €.
- Only the last `lookback_days` (default 15) up to today are synced. Days with empty cells (future days) are skipped.
- On the Google side only the orders (I) and revenue (J) cells of the matching date row (date in D) are written. Dates not found in the sheet are logged and skipped, never appended.
- Works on a temp copy of the file, so it is safe while Excel has it open. If Excel or OneDrive has the file locked, the copy is retried for about 30 seconds before giving up with a clear message. Re-running is idempotent.
- **Confirmation:** before writing, it prints a summary per market (days, date range, cells, skipped dates) and asks `[y/N]`. Nothing is written unless you answer `y`.

## Setup
1. `pip install -r requirements.txt` (Python 3.9+)
2. Google Cloud: create a project, enable the **Google Sheets API**, create a **service account**, download its JSON key as `service_account.json` (keep it private, it is git-ignored). Organizations may block key creation with the `iam.disableServiceAccountKeyCreation` policy; it can be overridden per project by an Organization Policy Administrator.
3. Share each of the 4 Google Sheets with the service account email (`...@...iam.gserviceaccount.com`) as **Editor**. "Public" view access is not enough to write.
4. `cp config.example.json config.json` and set `excel_path`.

### Where the files live
The script, `config.json`, `service_account.json` and the state file can sit in their own folder, separate from the Excel file. Relative paths in `config.json` are resolved against the script's folder, not the current directory. Use forward slashes (`D:/Reports/file.xlsx`) or a relative path such as `../file.xlsx` for `excel_path`. A leading slash on `service_account_file` (e.g. `/service_account.json`) points to the drive root on Windows, so don't use one.

## Run
```
python sync_excel_to_gsheets.py --dry-run          # list every cell that would be written, write nothing
python sync_excel_to_gsheets.py                    # summary, then asks for confirmation
python sync_excel_to_gsheets.py --yes              # skip the question (needed when there is no terminal)
python sync_excel_to_gsheets.py --force            # ignore the "unchanged file" shortcut
python sync_excel_to_gsheets.py --from 2026-01-01  # backfill
```
Without an interactive terminal (for example under cron) the script refuses to write unless `--yes` is given.
The date column of the `Daily` tab may be a real date or text (`YYYY-MM-DD`, `DD/MM/YYYY`, `DD.MM.YYYY`); check with `--dry-run`.

## Optional: schedule it
- Linux cron (daily 07:00): `0 7 * * * cd /opt/sync && python3 sync_excel_to_gsheets.py --yes >> sync.log 2>&1`
- Windows Task Scheduler: daily trigger, action `python.exe C:\sync\sync_excel_to_gsheets.py --yes`, "Start in" `C:\sync`.
