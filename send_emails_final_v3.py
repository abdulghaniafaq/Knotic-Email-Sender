"""
Knotic Labs - outreach email sender.
Reads each row of the outreach sheet and sends its email template to that row's
email address from YOUR own email account, over standard SMTP. No third-party
email service is used.

Usage:
  python send_emails.py                 # preview only (sends nothing)
  python send_emails.py --send          # really send
  python send_emails.py --send --limit 5
"""
import argparse, datetime as dt, csv, getpass, os, re, smtplib, ssl, sys, time
from datetime import datetime
from email.message import EmailMessage
from email.utils import make_msgid, formatdate
from openpyxl import load_workbook

# ---------- SETTINGS (edit these) ----------
SHEET_FILES  = [                       # every prospect sheet to read from
    "UAE_Email_Outreach_new_only.xlsx",
    "v2_UAE_Outreach(1).xlsx","UAE_Email_Outreach_final.xlsx"
]
SENDER_NAME  = "Abdul Ghani"
SENDER_EMAIL = os.environ.get("SENDER_EMAIL", "abdulghani@knotic.co")      # e.g. abdul@knotic.co
SMTP_HOST    = os.environ.get("SMTP_HOST", "smtp.gmail.com")  # Outlook: smtp.office365.com, Zoho: smtp.zoho.in
SMTP_PORT    = int(os.environ.get("SMTP_PORT", "587"))
DELAY_SECONDS = 90          # wait between emails so it looks human and avoids spam flags
DAILY_LIMIT   = 25          # max emails per run
SKIP_UNCONFIRMED = True     # skip rows whose email type says "unconfirmed"
CC_EMAILS = "abdullah.shahid@knotic.co, kafeel001khan@knotic.co"              # put your 3 CC emails here, e.g. "a@x.com, b@y.com, c@z.com"
LOG_FILE = "sent_log.csv"   # remembers who was emailed so nobody gets it twice
TRACKER_FILE = "Email_Sent_Tracker.xlsx"   # Excel log of every email sent / failed
FOLLOWUP_1_DAYS, FOLLOWUP_2_DAYS = 4, 10
# -------------------------------------------

EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")

def load_rows(path):
    ws = load_workbook(path).active
    head = [str(c.value or "").strip().lower() for c in ws[1]]
    col = lambda name: next(i for i, h in enumerate(head) if h.startswith(name))
    ci, ce, ct, ctype = col("company"), col("email id"), col("email template"), col("email type")
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[ci]: continue
        yield {"company": row[ci], "sector": row[col("sector")] or "", "emails": EMAIL_RE.findall(str(row[ce] or "")),
               "template": str(row[ct] or ""), "type": str(row[ctype] or "")}

def split_template(t):
    lines = t.strip().splitlines()
    subject = "Quick dea from Knotic Labs"
    if lines and lines[0].lower().startswith("subject:"):
        subject = lines[0].split(":", 1)[1].strip()
        lines = lines[1:]
    return subject, "\n".join(lines).strip()

def already_sent():
    if not os.path.exists(LOG_FILE): return sent_in_tracker()
    with open(LOG_FILE, newline="", encoding="utf-8") as f:
        done = {r["email"].strip().lower() for r in csv.DictReader(f) if r["status"] == "sent"}
    return done | sent_in_tracker()

def sent_in_tracker():
    # Also skip anyone already marked "sent" in the Excel sent tracker.
    if not os.path.exists(TRACKER_FILE): return set()
    ws = load_workbook(TRACKER_FILE, read_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    if not rows: return set()
    h = [str(x or "").strip().lower() for x in rows[0]]
    if "sent to" not in h or "status" not in h: return set()
    ti, si = h.index("sent to"), h.index("status")
    return {str(r[ti]).strip().lower() for r in rows[1:] if r[ti] and str(r[si] or "").lower() == "sent"}

TRACKER_HEAD = ["Date", "Time", "Company", "Sector", "Sent to", "CC", "Subject", "Status",
                "Error", "Follow-up 1 due", "Follow-up 2 due", "Reply received?", "Reply date", "Notes"]

def _open_tracker():
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    if os.path.exists(TRACKER_FILE):
        return load_workbook(TRACKER_FILE)
    wb = Workbook(); ws = wb.active; ws.title = "Sent log"
    ws.append(TRACKER_HEAD)
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF"); c.fill = PatternFill("solid", start_color="1F4E78")
    for col, w in zip("ABCDEFGHIJKLMN", [12, 10, 36, 20, 34, 30, 50, 10, 30, 15, 15, 15, 12, 30]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    # bring in anything sent before the tracker existed
    if os.path.exists(LOG_FILE):
        with open(LOG_FILE, newline="", encoding="utf-8") as f:
            for r in csv.DictReader(f):
                t = datetime.fromisoformat(r["time"])
                ws.append(_row(t, r["company"], "", r["email"], "", "", r["status"], r.get("note", "")))
    return wb

def _row(t, company, sector, email, cc, subject, status, note):
    ok = status == "sent"
    f1 = (t + dt.timedelta(days=FOLLOWUP_1_DAYS)).date().isoformat() if ok else ""
    f2 = (t + dt.timedelta(days=FOLLOWUP_2_DAYS)).date().isoformat() if ok else ""
    return [t.date().isoformat(), t.strftime("%H:%M"), company, sector, email, cc, subject,
            status.upper(), note, f1, f2, "No" if ok else "", "", ""]

def log(company, email, status, note="", sector="", subject="", cc=""):
    t = datetime.now()
    new = not os.path.exists(LOG_FILE)
    wb = _open_tracker()          # create (and import old CSV) before adding the new line
    with open(LOG_FILE, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new: w.writerow(["time", "company", "email", "status", "note"])
        w.writerow([t.isoformat(timespec="seconds"), company, email, status, note])
    wb.active.append(_row(t, company, sector, email, cc, subject, status, note))
    try:
        wb.save(TRACKER_FILE)
    except PermissionError:
        print(f"  (close {TRACKER_FILE} in Excel so it can be updated - the CSV log was still saved)")

def build(to, subject, body):
    m = EmailMessage()
    m["From"] = f"{SENDER_NAME} <{SENDER_EMAIL}>"
    m["To"], m["Subject"] = to, subject
    if CC_EMAILS: m["Cc"] = CC_EMAILS
    m["Date"], m["Message-ID"] = formatdate(localtime=True), make_msgid(domain=SENDER_EMAIL.split("@")[-1])
    m.set_content(body)
    html = body.replace("&", "&amp;").replace("<", "&lt;").replace("\n", "<br>")
    html = html.replace("https://labs.knotic.co/", '<a href="https://labs.knotic.co/">labs.knotic.co</a>')
    m.add_alternative(f'<div style="font-family:Arial,sans-serif;font-size:14px">{html}</div>', subtype="html")
    return m

# ---------- SCHEDULE (India time) ----------
# Which sectors go out on which day, and the time to start sending (IST).
# A sector matches if any keyword appears in the row's Sector column.
SCHEDULE = {
    "Tuesday":   {"start": "10:15", "sectors": ["bank", "exchange", "freight", "logistic", "3pl", "fulfil"]},
    "Wednesday": {"start": "10:30", "sectors": ["health", "dental", "lab", "insurance", "account", "clinic", "physio"]},
    "Thursday":  {"start": "10:15", "sectors": ["educat", "school", "nurser", "training", "tuition", "real estate", "property", "trading", "fintech",
                                                "food", "produce", "seafood", "fit-out", "facility", "fm services"]},
}
LATEST_START = "13:00"   # don't start a run after this time (IST)
IST = dt.timezone(dt.timedelta(hours=5, minutes=30))

def wait_for_slot(day_name):
    plan = SCHEDULE.get(day_name)
    if not plan:
        print(f"No sending planned for {day_name}. Planned days: {', '.join(SCHEDULE)}."); return None
    now = dt.datetime.now(IST)
    sh, sm = map(int, plan["start"].split(":")); lh, lm = map(int, LATEST_START.split(":"))
    start = now.replace(hour=sh, minute=sm, second=0, microsecond=0)
    if now > now.replace(hour=lh, minute=lm, second=0, microsecond=0):
        print(f"It's past {LATEST_START} IST - too late for today. Run again on the next planned day."); return None
    if now < start:
        print(f"Waiting until {plan['start']} IST to start sending... (keep the laptop awake)")
        time.sleep((start - now).total_seconds())
    return plan["sectors"]

def main():
    p = argparse.ArgumentParser()
    p.add_argument("--send", action="store_true", help="actually send (default is preview)")
    p.add_argument("--limit", type=int, default=DAILY_LIMIT)
    p.add_argument("--test-to", help="send every email to this address instead (for testing)")
    p.add_argument("--schedule", action="store_true", help="only send today's sectors, starting at the planned IST time")
    a = p.parse_args()

    sectors = None
    if a.schedule:
        today = dt.datetime.now(IST).strftime("%A")
        sectors = wait_for_slot(today) if a.send else SCHEDULE.get(today, {}).get("sectors")
        if not sectors: return
        print(f"{today}: sending only sectors matching {sectors}")

    done, queue = already_sent(), []
    rows = []
    for f in SHEET_FILES:
        if os.path.exists(f): rows += load_rows(f)
        else: print(f"WARNING: {f} not found - skipped")
    for r in rows:
        if sectors and not any(k in str(r["sector"]).lower() for k in sectors):
            continue
        if not r["emails"]:
            print(f"SKIP  {r['company']}: no email"); continue
        if SKIP_UNCONFIRMED and "unconfirmed" in r["type"].lower():
            print(f"SKIP  {r['company']}: email unconfirmed"); continue
        to = r["emails"][0]
        if to.lower() in done:
            print(f"SKIP  {r['company']}: already sent to {to}"); continue
        done.add(to.lower())   # stops the same email being queued twice from two sheets
        if "[" in r["template"]:
            print(f"SKIP  {r['company']}: template still has a [placeholder]"); continue
        queue.append((r["company"], to, *split_template(r["template"]), r["sector"]))
    queue = queue[: a.limit]

    print(f"\n{len(queue)} email(s) ready.\n")
    if not a.send:
        for c, to, s, b, _ in queue:
            print("=" * 70, f"\nTO: {to}  ({c})\nSUBJECT: {s}\n\n{b}\n")
        print("Preview only. Run again with --send to send.")
        return

    global SENDER_EMAIL
    SENDER_EMAIL = SENDER_EMAIL or input("Your email address: ").strip()
    password = os.environ.get("SMTP_PASSWORD") or getpass.getpass("App password (not shown): ")
    ctx = ssl.create_default_context()
    def connect():
        srv = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30)
        srv.starttls(context=ctx); srv.login(SENDER_EMAIL, password)
        return srv

    s = connect()
    for i, (c, to, subj, body, sec) in enumerate(queue, 1):
        target = a.test_to or to
        try:
            msg = build(target, subj, body)
            try:
                s.send_message(msg)
            except (smtplib.SMTPServerDisconnected, smtplib.SMTPSenderRefused, OSError):
                print("  connection lost - reconnecting...")
                try: s.close()
                except Exception: pass
                time.sleep(10)
                s = connect()
                s.send_message(msg)
            print(f"SENT  {i}/{len(queue)}  {c} -> {target}")
            log(c, target, "test" if a.test_to else "sent", "", sec, subj, CC_EMAILS)
        except Exception as e:
            print(f"FAIL  {c} -> {target}: {e}"); log(c, target, "failed", str(e), sec, subj, CC_EMAILS)
        if i < len(queue): time.sleep(DELAY_SECONDS)
    try: s.quit()
    except Exception: pass
    print("Done. Tracker updated:", TRACKER_FILE)

if __name__ == "__main__":
    sys.exit(main())