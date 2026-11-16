"""Alerting tests against a mocked Twilio endpoint. Run: python tests/test_alerts.py"""
import datetime as dt, http.server, json, os, shutil, subprocess, sys, tempfile, threading, urllib.parse

HERE = os.path.dirname(os.path.abspath(__file__))
PKG = os.path.dirname(HERE)
sys.path.insert(0, PKG)
import dayreport as d

POSTS, STATUS = [], [201]


class Mock(http.server.BaseHTTPRequestHandler):
    def do_POST(self):
        body = urllib.parse.parse_qs(self.rfile.read(int(self.headers["Content-Length"])).decode())
        POSTS.append({"path": self.path, "To": body["To"][0], "Body": body["Body"][0], "auth": self.headers.get("Authorization")})
        self.send_response(STATUS[0]); self.end_headers(); self.wfile.write(b"{}")

    def log_message(self, *a):
        pass


srv = http.server.HTTPServer(("127.0.0.1", 0), Mock)
threading.Thread(target=srv.serve_forever, daemon=True).start()
ENV = dict(os.environ, DAYREPORT_TWILIO_API="http://127.0.0.1:%d" % srv.server_port,
           TWILIO_ACCOUNT_SID="ACtest", TWILIO_AUTH_TOKEN="tok", TWILIO_FROM="+15550000000",
           PYTHON_KEYRING_BACKEND="keyring.backends.null.Keyring")
os.environ.update({k: ENV[k] for k in ("DAYREPORT_TWILIO_API", "TWILIO_ACCOUNT_SID", "TWILIO_AUTH_TOKEN",
                                       "TWILIO_FROM", "PYTHON_KEYRING_BACKEND")})
XML = sys.argv[1] if len(sys.argv) > 1 else "/workspace/dayreport/xml/merged"
TEMPLATE = sys.argv[2] if len(sys.argv) > 2 else "/workspace/dayreport/test_blank_template.xlsm"
RECIPS = """
[recipient.a]
name = A
number = +15551110001
alerts = error, pos_unreachable
[recipient.b]
name = B
number = +15551110002
alerts = success, missing_day
[recipient.c]
name = C
number = +15551110003
alerts = collision, already_running, missed_run
"""


def env_dir(recipients=True, quiet="", src_files=True):
    t = tempfile.mkdtemp(prefix="drtest")
    for x in ("src", "dr"):
        os.makedirs(os.path.join(t, x))
    if src_files:
        for f in os.listdir(XML):
            shutil.copy(os.path.join(XML, f), os.path.join(t, "src"))
    shutil.copy(TEMPLATE, os.path.join(t, "dr", "day report template.xlsm"))
    ini = """[dayreport]
source_outbox = {t}/src
local_inbox = {t}/inbox
archive_dir = {t}/archive
day_report_dir = {t}/dr
template = {t}/dr/day report template.xlsm
dailysales_dbf =
log_file = {t}/logs/dayreport.log
remove_from_source = yes
[alerts]
enabled = yes
dedup_hours = 12
quiet_hours = {q}
quiet_hours_apply_to = success, missing_day
""".format(t=t, q=quiet) + (RECIPS if recipients else "")
    with open(os.path.join(t, "t.ini"), "w") as f:
        f.write(ini)
    return t


def run(t, *extra):
    del POSTS[:]
    p = subprocess.run([sys.executable, os.path.join(PKG, "dayreport.py"), "--config", os.path.join(t, "t.ini")] + list(extra),
                       env=ENV, capture_output=True, text=True)
    return p.returncode, list(POSTS), p.stdout


def who(posts):
    return sorted((p["To"][-1], p["Body"].split(":")[0].split(" ", 2)[2].rsplit(" ", 1)[0]) for p in posts)


ok = True


def check(name, cond, info=""):
    global ok
    ok &= bool(cond)
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else "  -> %s" % info))


# 1 empty recipient list -> nothing sent, even with pos down
t = env_dir(recipients=False)
rc, posts, _ = run(t)
check("empty recipients: full run sends nothing", rc == 0 and posts == [], (rc, posts))
shutil.rmtree(os.path.join(t, "src"))
rc, posts, _ = run(t)
check("empty recipients: POS down sends nothing, exit 3", rc == 3 and posts == [], (rc, posts))

# 2 routing: success + missing_day -> B only
t = env_dir()
rc, posts, out = run(t)
w = who(posts)
check("success+missing_day go only to B", rc == 0 and {x[0] for x in w} == {"2"} and ("2", "SUCCESS") in w, w)
check("all bodies <=160 chars and start 'Day Report'", all(len(p["Body"]) <= 160 and p["Body"].startswith("Day Report") for p in posts))
check("Twilio API path + basic auth", posts and posts[0]["path"] == "/2010-04-01/Accounts/ACtest/Messages.json" and posts[0]["auth"].startswith("Basic "))
print("   sample:", posts[0]["Body"] if posts else None)

# 3 POS unreachable -> A; repeat -> dedup
shutil.rmtree(os.path.join(t, "src"))
rc, posts, _ = run(t)
check("pos_unreachable -> A, exit 3", rc == 3 and ("1", "POS UNREACHABLE") in who(posts), who(posts))
print("   sample:", [p["Body"] for p in posts if p["To"].endswith("1")])
rc, posts, _ = run(t)
check("pos_unreachable de-dup on 2nd run", rc == 3 and not any(p["To"].endswith("1") for p in posts), who(posts))

# 4 collision -> C (and error not sent for pure collision)
t2 = env_dir(src_files=False)
f = sorted(x for x in os.listdir(XML) if x.startswith("MSM34026100"))[2]
m = sorted(x for x in os.listdir(XML) if x.startswith("MCM34026100"))[2]
shutil.copy(os.path.join(XML, f), os.path.join(t2, "src")); shutil.copy(os.path.join(XML, m), os.path.join(t2, "src"))
run(t2)
import re
x = re.sub(r"<MiscellaneousSummaryAmount>([0-9.]+)<", lambda mm: "<MiscellaneousSummaryAmount>%s<" % (float(mm.group(1)) + 1),
           open(os.path.join(XML, f), encoding="iso-8859-1").read(), count=3)
open(os.path.join(t2, "src", f), "w", encoding="iso-8859-1").write(x); shutil.copy(os.path.join(XML, m), os.path.join(t2, "src"))
rc, posts, _ = run(t2)
check("collision -> C only, exit 1", rc == 1 and who(posts) == [("3", "COLLISION")], who(posts))

# 5 already running -> C
os.makedirs(os.path.join(t2, "inbox"), exist_ok=True)
open(os.path.join(t2, "inbox", ".dayreport.lock"), "w").write("1")
rc, posts, _ = run(t2, "--no-fetch")
check("already_running -> C, exit 4", rc == 4 and who(posts) == [("3", "ALREADY RUNNING")], who(posts))
os.remove(os.path.join(t2, "inbox", ".dayreport.lock"))

# 6 error (unreadable file) -> A
open(os.path.join(t2, "inbox", "MSM_broken.xml"), "w").write("<x/>")
rc, posts, _ = run(t2, "--no-fetch")
check("error -> A, exit 1", rc == 1 and ("1", "ERROR") in who(posts) and all(p["To"].endswith("1") or "MISSING" in p["Body"] for p in posts), who(posts))
os.remove(os.path.join(t2, "inbox", "MSM_broken.xml"))

# 7 SMS failure never changes the exit code
STATUS[0] = 500
t3 = env_dir()
rc, posts, out = run(t3)
check("Twilio HTTP 500: run still exit 0", rc == 0 and posts, rc)
STATUS[0] = 201
ENV_BAD = dict(ENV, DAYREPORT_TWILIO_API="http://127.0.0.1:1")
p = subprocess.run([sys.executable, os.path.join(PKG, "dayreport.py"), "--config", os.path.join(t3, "t.ini"), "--no-fetch"], env=ENV_BAD, capture_output=True, text=True)
check("Twilio unreachable: run unaffected (exit 0)", p.returncode == 0, p.returncode)

# 8 quiet hours: held then flushed (in-process with fixed clocks)
t4 = env_dir(quiet="22:00-07:00")
import configparser
cp = configparser.ConfigParser(interpolation=None); cp.read(os.path.join(t4, "t.ini"))
sp = os.path.join(t4, "state.json")
del POSTS[:]
a = d.Alerter(cp, sp, dt.datetime(2026, 10, 14, 23, 0))
a.alert("success", "OK test", key="q1")
a.alert("error", "boom", key="q2")
check("quiet hours: success held, error still sent at 23:00", who(POSTS) == [("1", "ERROR")], who(POSTS))
del POSTS[:]
d.Alerter(cp, sp, dt.datetime(2026, 10, 15, 6, 30)).flush_pending()
check("quiet hours: still held at 06:30", POSTS == [])
d.Alerter(cp, sp, dt.datetime(2026, 10, 15, 7, 30)).flush_pending()
check("quiet hours: released at 07:30 to B", who(POSTS) == [("2", "SUCCESS")], who(POSTS))
del POSTS[:]
d.Alerter(cp, sp, dt.datetime(2026, 10, 15, 8, 0)).alert("success", "OK test", key="q1")
check("dedup: same alert within 12h not resent", POSTS == [])
d.Alerter(cp, sp, dt.datetime(2026, 10, 15, 20, 0)).alert("success", "OK test", key="q1")
check("dedup: resent after 12h", who(POSTS) == [("2", "SUCCESS")], who(POSTS))

# 9 watchdog
t5 = env_dir()
rc, posts, _ = run(t5, "--watchdog")
check("watchdog: no run today -> missed_run to C, exit 1", rc == 1 and who(posts) == [("3", "MISSED RUN")], who(posts))
print("   sample:", posts[0]["Body"] if posts else None)
run(t5)
rc, posts, _ = run(t5, "--watchdog")
check("watchdog: after good run -> silent, exit 0", rc == 0 and posts == [], (rc, who(posts)))
st = d.load_state(os.path.join(t5, "logs", "dayreport_state.json"))
st["last_end"] = ""; st["sms_sent"] = {}; d.save_state(os.path.join(t5, "logs", "dayreport_state.json"), st)
rc, posts, _ = run(t5, "--watchdog")
check("watchdog: started but not finished -> alert", rc == 1 and "never finished" in (posts[0]["Body"] if posts else ""), who(posts))

# 10 --test-sms
rc, posts, _ = run(t5, "--test-sms", "+15559990000")
check("--test-sms sends one text", rc == 0 and len(posts) == 1 and posts[0]["To"] == "+15559990000", posts)

print("ALL PASS" if ok else "SOME TESTS FAILED")
sys.exit(0 if ok else 1)
