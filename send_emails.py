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
import argparse, csv, getpass, os, re, smtplib, ssl, sys, time
from datetime import datetime
from email.message import EmailMessage
from email.utils import make_msgid, formatdate
from openpyxl import load_workbook

# ---------- SETTINGS (edit these) ----------
SHEET_FILE   = "UAE_Email_Outreach_v3.xlsx"
SENDER_NAME  = "Abdul Ghani | Knotic Labs"
SENDER_EMAIL = os.environ.get("SENDER_EMAIL", "")      # e.g. abdul@knotic.co
SMTP_HOST    = os.environ.get("SMTP_HOST", "smtp.gmail.com")  # Outlook: smtp.office365.com, Zoho: smtp.zoho.in
SMTP_PORT    = int(os.environ.get("SMTP_PORT", "587"))
DELAY_SECONDS = 90          # wait between emails so it looks human and avoids spam flags
DAILY_LIMIT   = 25          # max emails per run
SKIP_UNCONFIRMED = True     # skip rows whose email type says "unconfirmed"
LOG_FILE = "sent_log.csv"   # remembers who was emailed so nobody gets it twice
# -------------------------------------------

EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")

def load_rows(path):
    ws = load_workbook(path).active
    head = [str(c.value or "").strip().lower() for c in ws[1]]
    col = lambda name: next(i for i, h in enumerate(head) if h.startswith(name))
    ci, ce, ct, ctype = col("company"), col("email id"), col("email template"), col("email type")
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[ci]: continue
        yield {"company": row[ci], "emails": EMAIL_RE.findall(str(row[ce] or "")),
               "template": str(row[ct] or ""), "type": str(row[ctype] or "")}

def split_template(t):
    lines = t.strip().splitlines()
    subject = "Quick idea from Knotic Labs"
    if lines and lines[0].lower().startswith("subject:"):
        subject = lines[0].split(":", 1)[1].strip()
        lines = lines[1:]
    return subject, "\n".join(lines).strip()

def already_sent():
    if not os.path.exists(LOG_FILE): return set()
    with open(LOG_FILE, newline="", encoding="utf-8") as f:
        return {r["email"].lower() for r in csv.DictReader(f) if r["status"] == "sent"}

def log(company, email, status, note=""):
    new = not os.path.exists(LOG_FILE)
    with open(LOG_FILE, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new: w.writerow(["time", "company", "email", "status", "note"])
        w.writerow([datetime.now().isoformat(timespec="seconds"), company, email, status, note])

def build(to, subject, body):
    m = EmailMessage()
    m["From"] = f"{SENDER_NAME} <{SENDER_EMAIL}>"
    m["To"], m["Subject"] = to, subject
    m["Date"], m["Message-ID"] = formatdate(localtime=True), make_msgid(domain=SENDER_EMAIL.split("@")[-1])
    m.set_content(body)
    html = body.replace("&", "&amp;").replace("<", "&lt;").replace("\n", "<br>")
    html = html.replace("https://labs.knotic.co/", '<a href="https://labs.knotic.co/">labs.knotic.co</a>')
    m.add_alternative(f'<div style="font-family:Arial,sans-serif;font-size:14px">{html}</div>', subtype="html")
    return m

def main():
    p = argparse.ArgumentParser()
    p.add_argument("--send", action="store_true", help="actually send (default is preview)")
    p.add_argument("--limit", type=int, default=DAILY_LIMIT)
    p.add_argument("--test-to", help="send every email to this address instead (for testing)")
    a = p.parse_args()

    done, queue = already_sent(), []
    for r in load_rows(SHEET_FILE):
        if not r["emails"]:
            print(f"SKIP  {r['company']}: no email"); continue
        if SKIP_UNCONFIRMED and "unconfirmed" in r["type"].lower():
            print(f"SKIP  {r['company']}: email unconfirmed"); continue
        to = r["emails"][0]
        if to.lower() in done:
            print(f"SKIP  {r['company']}: already sent to {to}"); continue
        if "[" in r["template"]:
            print(f"SKIP  {r['company']}: template still has a [placeholder]"); continue
        queue.append((r["company"], to, *split_template(r["template"])))
    queue = queue[: a.limit]

    print(f"\n{len(queue)} email(s) ready.\n")
    if not a.send:
        for c, to, s, b in queue:
            print("=" * 70, f"\nTO: {to}  ({c})\nSUBJECT: {s}\n\n{b}\n")
        print("Preview only. Run again with --send to send.")
        return

    global SENDER_EMAIL
    SENDER_EMAIL = SENDER_EMAIL or input("Your email address: ").strip()
    password = os.environ.get("SMTP_PASSWORD") or getpass.getpass("App password (not shown): ")
    ctx = ssl.create_default_context()
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as s:
        s.starttls(context=ctx); s.login(SENDER_EMAIL, password)
        for i, (c, to, subj, body) in enumerate(queue, 1):
            target = a.test_to or to
            try:
                s.send_message(build(target, subj, body))
                print(f"SENT  {i}/{len(queue)}  {c} -> {target}")
                if not a.test_to: log(c, to, "sent")
            except Exception as e:
                print(f"FAIL  {c} -> {target}: {e}"); log(c, to, "failed", str(e))
            if i < len(queue): time.sleep(DELAY_SECONDS)
    print("Done. See", LOG_FILE)

if __name__ == "__main__":
    sys.exit(main())
