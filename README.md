# Knotic Labs – Email Sender

Sends a personalised outreach email to each company in an Excel sheet, from your own email account (Gmail SMTP), and records every send in an Excel tracker.

## Files in this folder

| File | What it is |
|---|---|
| `send_emails_final.py` | **Main script — use this one.** Sends emails, adds CC, writes the tracker. |
| `send_emails.py` | Older version without the tracker. Keep as a backup or delete. |
| `UAE_Email_Outreach_final.xlsx` | All 85 companies (original 30 + 55 new) with ready-to-send emails. |
| `UAE_Email_Outreach_new_only.xlsx` | Only the 55 newly added companies. |
| `UAE_Email_Outreach_v4.xlsx` | Only the original 30 companies. |
| `sent_log.csv` | Created automatically. Remembers who was emailed so nobody gets it twice. **Don't delete.** |
| `Email_Sent_Tracker.xlsx` | Created automatically. One line per email: date, company, sector, sent to, CC, subject, status, follow-up dates, reply columns. |

## Settings (top of `send_emails_final.py`)

| Setting | Meaning |
|---|---|
| `SHEET_FILE` | Which sheet to send from, e.g. `"UAE_Email_Outreach_final.xlsx"` |
| `SENDER_NAME` | Name the recipient sees in "From" |
| `SENDER_EMAIL` | Email account you send from |
| `CC_EMAILS` | CC addresses, comma-separated inside one pair of quotes |
| `DELAY_SECONDS` | Wait between emails (90 = looks human, avoids spam flags) |
| `DAILY_LIMIT` | Max emails per run (25) |
| `SKIP_UNCONFIRMED` | `True` = skip emails marked "unconfirmed" in the sheet |
| `FOLLOWUP_1_DAYS`, `FOLLOWUP_2_DAYS` | Days after sending for follow-up reminders in the tracker (4 and 10) |

Save with **Cmd+S** after any change.

## How to run (VS Code terminal, inside this folder)

1. **Preview** (sends nothing):
   ```bash
   python3 send_emails_final.py
   ```
2. **Test** (sends to you only):
   ```bash
   python3 send_emails_final.py --send --test-to ghaniabdul712@gmail.com --limit 1
   ```
3. **Send for real:**
   ```bash
   python3 send_emails_final.py --send
   ```
   Optional: `--limit 10` to send fewer.

When asked, paste your **Gmail app password** (16 letters, no spaces). Nothing shows while you paste — that's normal.

## Good to know

- Run once a day; each run sends up to 25 and skips anyone already emailed.
- Keep the laptop awake while it runs (25 emails ≈ 40 minutes).
- Close `Email_Sent_Tracker.xlsx` in Excel while sending so it can be updated.
- Real sends appear in the tracker as **SENT**; failures as **FAILED** with the reason. Test sends appear as **TEST** if you use the TEST version of the script.
- Fill in "Reply received?", "Reply date" and "Notes" in the tracker yourself.
- Rows with no email ("Email - NA") are skipped automatically.

## Troubleshooting

| Message | Fix |
|---|---|
| `Username and Password not accepted` | Create a new app password at myaccount.google.com/apppasswords and paste it without spaces. |
| `No such file or directory` | Run `cd` into the email-sender folder, and check `SHEET_FILE` matches the file name exactly. |
| `No module named openpyxl` | `python3 -m pip install --user openpyxl` |
| Tracker file not showing | Click refresh in the VS Code Explorer. |
