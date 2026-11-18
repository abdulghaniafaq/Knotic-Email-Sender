# Knotic Labs – Email Sender

Sends a personalised outreach email to each company in an Excel sheet, from your own email account (Gmail SMTP by default), CCs your team, and records every send in an Excel tracker. A second script reads your mailbox and marks replies and bounces in that tracker.

## Files in this folder

| File | What it is |
|---|---|
| `send_emails_final_v3.py` | **Main script — use this one.** Sends emails, adds CC, writes the tracker, can send only today's sectors on a weekly schedule (`--schedule`), and reconnects automatically if the connection drops. |
| `check_replies.py` | Reads your mailbox (IMAP) and updates the tracker: marks replies and bounced emails. |
| `send_emails_final.py` | Previous version of the main script. Same as v3 but without `--schedule`, and test sends are **not** logged. Backup only. |
| `send_emails.py` | Oldest version. Logs to `sent_log.csv` only (no tracker), CC is hard-coded, reads `UAE_Email_Outreach_v4.xlsx`. Backup only. |
| `UAE_Email_Outreach_final.xlsx` | All 85 companies (original 30 + 55 new) with ready-to-send emails. |
| `UAE_Email_Outreach_new_only.xlsx` | Only the 55 newly added companies (Banking, Education, Real Estate, Logistics, Insurance, Trading, Accounting). |
| `UAE_Email_Outreach_v4.xlsx` | Only the original 30 companies (Healthcare, Banking, Education). |
| `sent_log.csv` | Created automatically. Remembers who was emailed so nobody gets it twice. **Don't delete.** |
| `Email_Sent_Tracker.xlsx` | Created automatically. One line per email: date, time, company, sector, sent to, CC, subject, status, error, follow-up dates, reply columns. |

## Before you start

- Python 3 with `openpyxl` installed (`python3 -m pip install --user openpyxl`).
- A **Gmail app password** for the sending account (16 letters, no spaces). Create one at myaccount.google.com/apppasswords.
- For `check_replies.py`, IMAP must be enabled on the mailbox (Gmail: Settings → Forwarding and POP/IMAP).

## Settings (top of `send_emails_final_v3.py`)

| Setting | Meaning |
|---|---|
| `SHEET_FILE` | Which sheet to send from (currently `"UAE_Email_Outreach_new_only.xlsx"`) |
| `SENDER_NAME` | Name the recipient sees in "From" |
| `SENDER_EMAIL` | Email account you send from (default `abdulghani@knotic.co`). Can also be set with the `SENDER_EMAIL` environment variable. |
| `SMTP_HOST`, `SMTP_PORT` | Mail server (default `smtp.gmail.com`, `587`). Override with env vars `SMTP_HOST` / `SMTP_PORT`. Outlook: `smtp.office365.com`, Zoho: `smtp.zoho.in`. |
| `CC_EMAILS` | CC addresses, comma-separated inside one pair of quotes. Added to every email. |
| `DELAY_SECONDS` | Wait between emails (90 = looks human, avoids spam flags) |
| `DAILY_LIMIT` | Max emails per run (25). `--limit` overrides it. |
| `SKIP_UNCONFIRMED` | `True` = skip emails marked "unconfirmed" in the sheet |
| `FOLLOWUP_1_DAYS`, `FOLLOWUP_2_DAYS` | Days after sending for the follow-up reminder dates in the tracker (4 and 10) |
| `SCHEDULE` | Which sectors go out on which weekday, and the start time in India time (IST) — see [Weekly schedule](#weekly-schedule) |
| `LATEST_START` | A scheduled run won't start after this time, IST (`13:00`) |
| `LOG_FILE`, `TRACKER_FILE` | Names of the CSV log and the Excel tracker |

The app password is never stored in the code. You're asked for it each run, or you can set it first with `export SMTP_PASSWORD="yourapppassword"`.

Save with **Cmd+S** after any change.

## How to run (VS Code terminal, inside this folder)

1. **Preview** (sends nothing, prints each email):
   ```bash
   python3 send_emails_final_v3.py
   ```
2. **Test** (sends every email to you instead of the company):
   ```bash
   python3 send_emails_final_v3.py --send --test-to ghaniabdul712@gmail.com --limit 1
   ```
3. **Send for real:**
   ```bash
   python3 send_emails_final_v3.py --send
   ```
   Optional: `--limit 10` to send fewer.
4. **Send today's sectors only** (see schedule below):
   ```bash
   python3 send_emails_final_v3.py --send --schedule
   ```
   Add `--schedule` without `--send` to preview what today's run would send.

When asked, paste your **Gmail app password**. Nothing shows while you paste — that's normal.

## Weekly schedule

With `--schedule`, only rows whose **Sector** contains one of the day's keywords are sent. Start times are IST.

| Day | Start | Sectors (keywords matched) |
|---|---|---|
| Tuesday | 10:15 | bank, exchange, freight, logistic |
| Wednesday | 10:30 | health, dental, lab, insurance, account |
| Thursday | 10:15 | educat, school, nurser, real estate, property, trading, fintech |

- Started before the start time → it waits, then sends. Keep the laptop awake.
- Started after `LATEST_START` (13:00 IST) → it refuses and tells you to run on the next planned day.
- Other days (Mon, Fri, Sat, Sun) → nothing is sent.
- The app password is asked **after** the wait. To start it early and leave it, run `export SMTP_PASSWORD="..."` first; otherwise it will sit at the password prompt when the time comes.
- Edit `SCHEDULE` at the top of the script to change days, times or sectors.

## Checking replies and bounces

```bash
python3 check_replies.py
```

Asks for the same app password (or uses `SMTP_PASSWORD`; set `SENDER_EMAIL` / `IMAP_HOST` env vars if not Gmail). It looks at every **SENT** row that has no reply yet and:

- **Reply found** → "Reply received?" = Yes, "Reply date" filled, "Notes" = `Reply: <subject>`. For company domains, a reply from anyone at that domain counts; for free mail (gmail, yahoo, hotmail, outlook, icloud, live) only an exact address match counts.
- **Delivery-failed message found** → Status changes to **BOUNCED** and Notes says to find another address.

Close `Email_Sent_Tracker.xlsx` in Excel/Numbers before running. A bounced address stays in `sent_log.csv`, so it won't be retried — put a new address in the sheet if you find one.

## What the sheet needs

First sheet of the workbook, with these column headings (case doesn't matter, only the start of the heading is matched): **Company name**, **Sector**, **Email id**, **Email type**, **Email template**. Other columns are ignored.

- The first line of the template should be `Subject: your subject`. If it's missing, the subject is "Quick Idea from Knotic Labs". The rest of the template is the email body.
- If the **Email id** cell has more than one address, only the first is used.
- A row is skipped (and the reason printed) when it has no email, its email type says "unconfirmed" (if `SKIP_UNCONFIRMED` is on), that address was already sent to, or the template still contains a `[placeholder]`.

Emails go out as plain text with an HTML version; `https://labs.knotic.co/` becomes a clickable `labs.knotic.co` link.

## The tracker

| Status | Meaning |
|---|---|
| **SENT** | Delivered to the mail server. Follow-up 1 / 2 due dates are filled and "Reply received?" starts as No. |
| **FAILED** | Send failed; the reason is in the Error column. No follow-up dates. It is **not** counted as sent, so the next run tries it again. |
| **TEST** | A `--test-to` send. Logged with the test address, no follow-up dates, and doesn't stop the real email going out later. |
| **BOUNCED** | Set by `check_replies.py` when delivery failed. |

Fill in "Notes" yourself as needed. If the tracker is first created after emails were already sent, the old `sent_log.csv` entries are imported into it.

## Good to know

- Run once a day; each run sends up to 25 and skips anyone already emailed.
- Keep the laptop awake while it runs (25 emails ≈ 40 minutes).
- If the connection drops mid-run, v3 waits 10 seconds, reconnects and retries that email once.
- Close `Email_Sent_Tracker.xlsx` in Excel while sending so it can be updated. If it's open, the script prints a warning and still saves the entry to `sent_log.csv`.
- Rows with no email ("Email - NA") are skipped automatically.

## Troubleshooting

| Message | Fix |
|---|---|
| `Username and Password not accepted` | Create a new app password at myaccount.google.com/apppasswords and paste it without spaces. |
| `No such file or directory` | Run `cd` into the email-sender folder, and check `SHEET_FILE` matches the file name exactly. |
| `No module named openpyxl` | `python3 -m pip install --user openpyxl` |
| `Email_Sent_Tracker.xlsx not found - run this in the email-sender folder.` | `check_replies.py` needs the tracker, which is created by the first real send. Run it from inside this folder. |
| `Close Email_Sent_Tracker.xlsx first, then run again.` | Close the tracker in Excel/Numbers, then re-run `check_replies.py`. |
| `(close Email_Sent_Tracker.xlsx in Excel so it can be updated…)` | Close the tracker. The send itself was logged in `sent_log.csv`. |
| `No sending planned for <day>` / `It's past 13:00 IST` | Using `--schedule` on a non-planned day or too late. Run on the next planned day, or drop `--schedule`. |
| IMAP login error in `check_replies.py` | Turn on IMAP in the mailbox settings and use the app password. |
| Tracker file not showing | Click refresh in the VS Code Explorer. |
