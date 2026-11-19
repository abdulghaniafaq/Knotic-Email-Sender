"""
check_replies.py - Knotic Labs
Reads your mailbox (IMAP, Python built-ins only) and updates Email_Sent_Tracker.xlsx:
  - "Reply received?" -> Yes, "Reply date" filled
  - Notes = detected intent + short summary of the reply, e.g.
      "INTERESTED - Wants a call next week. | 'Happy to talk, send a slot Tuesday.'"
    Intent is detected on this computer with keyword rules (no AI/3rd-party service).
Run:  python3 check_replies.py            (new replies only)
      python3 check_replies.py --redo     (also re-read replies already marked Yes)
  - Status -> BOUNCED when a "delivery failed" message mentions that address
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

# ---------- intent detection (edit keywords freely) ----------
INTENTS = [  # checked top to bottom, first match wins
    ("UNSUBSCRIBE",   "Asked not to be contacted - stop all follow-ups",
     ["unsubscribe", "remove me", "do not contact", "don't contact", "stop emailing", "take me off", "not to contact"]),
    ("OUT OF OFFICE", "Auto-reply / out of office - follow up after they return",
     ["out of office", "out of the office", "on leave", "on vacation", "annual leave", "automatic reply", "auto-reply", "autoreply", "away from the office", "limited access to email"]),
    ("NOT INTERESTED","Not interested - close politely",
     ["not interested", "no thank", "no, thank", "not looking", "no requirement", "no need", "not required", "already have", "in-house", "we are covered", "decline", "not a fit"]),
    ("MEETING",       "Wants a call/meeting - book it",
     ["schedule a call", "set up a call", "book a call", "a quick call", "let's talk", "lets talk", "meeting", "available on", "free on", "calendly", "teams link", "zoom", "google meet", "call me", "give me a call"]),
    ("INTERESTED",    "Interested - reply soon",
     ["interested", "sounds good", "tell me more", "more details", "more information", "send details", "share details", "proposal", "pricing", "quote", "cost", "case study", "portfolio", "brochure", "happy to", "would like to"]),
    ("REFERRAL",      "Forwarded/referred to someone else - contact that person",
     ["forwarded", "forwarding", "right person", "relevant team", "concerned team", "concerned department", "in copy", "cc'd", "copied", "please contact", "reach out to", "get in touch with"]),
    ("NOT NOW",       "Not now - follow up in a few months",
     ["not at the moment", "not right now", "at this time", "future", "next quarter", "next year", "later", "keep in touch", "revisit", "budget"]),
    ("QUESTION",      "Asked a question - answer it", ["?"]),
]

def body_text(M, num):
    _, data = M.fetch(num, "(BODY.PEEK[])")
    msg = email.message_from_bytes(data[0][1]); text = ""
    parts = msg.walk() if msg.is_multipart() else [msg]
    for p in parts:
        if p.get_content_type() in ("text/plain", "text/html") and "attachment" not in str(p.get("Content-Disposition", "")):
            raw = p.get_payload(decode=True) or b""
            t = raw.decode(p.get_content_charset() or "utf-8", "ignore")
            if p.get_content_type() == "text/html": t = re.sub(r"<[^>]+>", " ", re.sub(r"(?is)<(style|script).*?</\1>", " ", t))
            text = t
            if p.get_content_type() == "text/plain": break
    return text

def clean_reply(text):
    """Keep only what they wrote: drop quoted earlier email and signature."""
    out = []
    for line in text.splitlines():
        s = line.strip()
        if s.startswith(">"): continue
        if re.match(r"^(On .+wrote:|From:|-----Original Message|Sent from my|Get Outlook|________________)", s): break
        if re.match(r"^(regards|best regards|kind regards|thanks and regards|warm regards|best|thanks|thank you|cheers|sincerely)[,!.]?$", s, re.I) and out: break
        out.append(s)
    return re.sub(r"\s+", " ", " ".join(out)).strip()

def detect_intent(subj, text):
    low = (subj + " " + text).lower()
    for name, action, words in INTENTS:
        if any(w in low for w in words):
            if name == "QUESTION" and "?" not in text: continue
            return name, action
    return "REPLIED", "Read and respond"

def summarise(text, n=2, limit=220):
    sents = [s.strip() for s in re.split(r"(?<=[.!?])\s+", text) if len(s.strip()) > 3]
    sents = [s for s in sents if not re.match(r"^(hi|hello|dear|hey|good (morning|afternoon|evening))\b", s, re.I) or len(s) > 40]
    summ = " ".join(sents[:n]) or text
    return summ[:limit].rstrip() + ("..." if len(summ) > limit else "")

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
        if status not in ("SENT",) or "@" not in to or (ws.cell(r, C_REPLY).value == "Yes" and "--redo" not in sys.argv): continue
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
            text = clean_reply(body_text(M, n))
            intent, action = detect_intent(subj, text)
            ws.cell(r, C_NOTES).value = f"{intent} - {action}. | \"{summarise(text)}\""[:400]
            replies += 1; print(f"{intent:<14}", to, "-", summarise(text, 1, 90)); break

    M.logout()
    try: wb.save(TRACKER_FILE)
    except PermissionError: sys.exit(f"Close {TRACKER_FILE} first, then run again.")
    print(f"Done. {replies} new replies, {bounces} bounces. Tracker updated.")

if __name__ == "__main__":
    main()