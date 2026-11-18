"""
check_replies.py - Knotic Labs
Reads your mailbox (IMAP, Python built-ins only) and updates Email_Sent_Tracker.xlsx:
  - "Reply received?" -> Yes, "Reply date" filled, Notes = subject of the reply
  - Status -> BOUNCED when a "delivery failed" message mentions that address
Run:  python3 check_replies.py
Close the tracker in Excel/Numbers before running.
"""
import datetime as dt, email, getpass, imaplib, os, re, sys
from email.header import decode_header, make_header
from email.utils import parsedate_to_datetime
from openpyxl import load_workbook

SENDER_EMAIL = os.environ.get("SENDER_EMAIL", "")         # same address you send from
IMAP_HOST    = os.environ.get("IMAP_HOST", "imap.gmail.com")  # Outlook: outlook.office365.com, Zoho: imap.zoho.in
TRACKER_FILE = "Email_Sent_Tracker.xlsx"
# Free mail domains: for these, only an exact address match counts as a reply
FREE = {"gmail.com", "yahoo.com", "hotmail.com", "outlook.com", "icloud.com", "live.com"}

C_DATE, C_TO, C_STATUS, C_REPLY, C_RDATE, C_NOTES = 1, 5, 8, 12, 13, 14

def imap_date(d): return d.strftime("%d-%b-%Y")

def pick_folder(M):
    for f in ('"[Gmail]/All Mail"', "INBOX"):
        if M.select(f, readonly=True)[0] == "OK": return

def headers(M, num):
    _, data = M.fetch(num, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE)])")
    msg = email.message_from_bytes(data[0][1])
    subj = str(make_header(decode_header(msg.get("Subject", ""))))
    try: when = parsedate_to_datetime(msg["Date"]).date()
    except Exception: when = None
    return msg.get("From", ""), subj, when

def search(M, *criteria):
    typ, data = M.search(None, *criteria)
    return data[0].split() if typ == "OK" and data[0] else []

def main():
    if not os.path.exists(TRACKER_FILE): sys.exit(f"{TRACKER_FILE} not found - run this in the email-sender folder.")
    user = SENDER_EMAIL or input("Your email address: ").strip()
    pwd = os.environ.get("SMTP_PASSWORD") or getpass.getpass("App password (not shown): ")
    M = imaplib.IMAP4_SSL(IMAP_HOST); M.login(user, pwd); pick_folder(M)

    wb = load_workbook(TRACKER_FILE); ws = wb.active
    replies = bounces = 0
    for r in range(2, ws.max_row + 1):
        status = str(ws.cell(r, C_STATUS).value or "").upper()
        to = str(ws.cell(r, C_TO).value or "").strip().lower()
        if status != "SENT" or "@" not in to or ws.cell(r, C_REPLY).value == "Yes": continue
        try: since = dt.date.fromisoformat(str(ws.cell(r, C_DATE).value)[:10])
        except ValueError: continue
        domain = to.split("@")[1]

        # bounce?
        for n in search(M, "SINCE", imap_date(since), "FROM", "mailer-daemon", "BODY", to):
            ws.cell(r, C_STATUS).value = "BOUNCED"; ws.cell(r, C_NOTES).value = "Delivery failed - find another address"
            bounces += 1; print("BOUNCED ", to); break
        if ws.cell(r, C_STATUS).value == "BOUNCED": continue

        # reply? exact address, or anyone at the company's domain
        who = to if domain in FREE else "@" + domain
        for n in reversed(search(M, "SINCE", imap_date(since), "FROM", who)):
            frm, subj, when = headers(M, n)
            ws.cell(r, C_REPLY).value = "Yes"
            ws.cell(r, C_RDATE).value = when.isoformat() if when else ""
            ws.cell(r, C_NOTES).value = f"Reply: {subj}"[:200]
            replies += 1; print("REPLIED ", to, "-", subj); break

    M.logout()
    try: wb.save(TRACKER_FILE)
    except PermissionError: sys.exit(f"Close {TRACKER_FILE} first, then run again.")
    print(f"Done. {replies} new replies, {bounces} bounces. Tracker updated.")

if __name__ == "__main__":
    main()
