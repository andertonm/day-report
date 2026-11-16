"""gastanks.watch page -> tank-snapshot CSV. Run: python tests/test_grab_tanks.py"""
import configparser, datetime as dt, logging, os, shutil, sys, tempfile, time
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import dayreport as d

logging.basicConfig(level=logging.INFO)
HERE = os.path.dirname(os.path.abspath(__file__))
FIXTURE = os.path.join(HERE, "fixtures", "gastanks_home_2026-11-11.html")
# Fixture timestamps are UTC; make local time = the store's (Eastern) for stable asserts.
os.environ["TZ"] = "America/New_York"
if hasattr(time, "tzset"):
    time.tzset()


def html():
    return open(FIXTURE, encoding="utf-8").read()


def cfg(**kw):
    cp = configparser.ConfigParser(interpolation=None)
    base = {"tank_fill": "yes", "tank_grab_retry_seconds": "0"}
    base.update(kw)
    cp["dayreport"] = base
    return cp["dayreport"]


def test_parse_fixture():
    rows = d.parse_tank_page(html())
    by = {r["tank_number"]: r for r in rows}
    assert sorted(by) == ["01", "02", "03", "04"]
    assert by["01"]["product"] == "Regular" and by["01"]["height"] == 40
    assert by["01"]["volume"] == 3900 and by["01"]["ullage"] == 5503
    assert by["02"]["product"] == "Diesel" and by["02"]["height"] == 24
    assert by["03"]["product"] == "Premium" and by["03"]["height"] == 54
    assert by["04"]["height"] == 14 and round(by["04"]["water"], 2) == 0.65
    assert round(by["04"]["water_volume"], 2) == 2.31
    # 2026-11-11T22:19:07Z = 5:19:07 PM Eastern
    assert by["01"]["reading_at"] == dt.datetime(2026, 11, 11, 17, 19, 7), by["01"]["reading_at"]
    print("parse_fixture OK")


def test_parse_table_fallback():
    page = """<table><tr><th>Tank #</th><th>Product</th><th>F1</th><th>F2</th><th>F3</th>
    <th>Volume</th><th>TC Volume</th><th>Ullage</th><th>Height</th><th>Water</th><th>Temp</th>
    <th>Water Volume</th></tr>
    <tr><td>01</td><td>Regular</td><td></td><td></td><td></td><td>4,018</td><td>3983</td><td>5385</td><td>41</td><td>0.00</td><td>72.36</td><td>0.00</td></tr>
    <tr><td>02</td><td>Diesel</td><td></td><td></td><td></td><td>1796</td><td>1785</td><td>7607</td><td>24</td><td>0.00</td><td>72.81</td><td>0.00</td></tr>
    <tr><td>03</td><td>Premium</td><td></td><td></td><td></td><td>3562</td><td>3531</td><td>2189</td><td>55</td><td>0.00</td><td>72.51</td><td>0.00</td></tr>
    <tr><td>04</td><td>Premium</td><td></td><td></td><td></td><td>509</td><td>505</td><td>5242</td><td>14</td><td>0.65</td><td>70.28</td><td>2.31</td></tr>
    </table>"""
    rows = d.parse_tank_page(page)
    by = {r["tank_number"]: r for r in rows}
    assert by["01"]["height"] == 41 and by["01"]["volume"] == 4018
    assert by["04"]["water"] == 0.65 and by["04"]["reading_at"] is None
    print("table_fallback OK")


def test_parse_errors():
    for bad in ("", "<html><body>Maintenance</body></html>",
                html().replace("&quot;tank_number&quot;:&quot;04&quot;", "&quot;tank_number&quot;:&quot;09&quot;")):
        try:
            d.parse_tank_page(bad)
        except d.TankPageError:
            continue
        raise AssertionError("expected TankPageError")
    print("parse_errors OK")


def test_grab_writes_csv_reader_accepts():
    t = tempfile.mkdtemp()
    c = cfg(tank_csv_dir=t)
    now = lambda: dt.datetime(2026, 11, 11, 17, 20, 0)
    rc = d.grab_tanks(c, fetcher=lambda url: html(), sleep=lambda s: None, now=now)
    assert rc == 0
    p = os.path.join(t, "tank-snapshot-2026-11-11.csv")
    text = open(p, encoding="utf-8").read().splitlines()
    assert text[0] == ",".join(d.TANK_CSV_COLUMNS), text[0]
    assert text[1] == "01,Regular,3900,3866,5503,40,0,72.38,0,2026-11-11 17:19:07", text[1]
    assert text[4].startswith("04,Premium,509,505,5242,14,0.65,70.24,2.31,"), text[4]
    bdate, when, heights = d.parse_tank_csv(p)
    assert bdate == dt.date(2026, 11, 11)
    assert heights == {"IR_Regular": 40, "Diesel": 24, "IR_Supreme": 54, "IR_Supreme2": 14}, heights
    assert d.load_tank_csv_index(t)[dt.date(2026, 11, 11)][0] == p
    print("grab_writes_csv OK")


def test_retry_then_success_and_total_failure():
    t = tempfile.mkdtemp()
    calls, sleeps = [], []
    def flaky(url):
        calls.append(url)
        if len(calls) < 3:
            raise OSError("network down")
        return html()
    now = lambda: dt.datetime(2026, 11, 11, 17, 20, 0)
    c = cfg(tank_csv_dir=t, tank_grab_attempts="5", tank_grab_retry_seconds="60")
    assert d.grab_tanks(c, fetcher=flaky, sleep=sleeps.append, now=now) == 0
    assert len(calls) == 3 and sleeps == [60, 60]

    t2 = tempfile.mkdtemp()
    sleeps = []
    def down(url):
        raise OSError("no route to host")
    c = cfg(tank_csv_dir=t2, tank_grab_attempts="5", tank_grab_retry_seconds="60")
    assert d.grab_tanks(c, fetcher=down, sleep=sleeps.append, now=now) == d.EXIT_TANK_GRAB == 6
    assert len(sleeps) == 4 and os.listdir(t2) == []
    print("retry OK")


def test_stale_readings_fail():
    t = tempfile.mkdtemp()
    now = lambda: dt.datetime(2026, 11, 12, 1, 0, 0)  # page still says 5:19 PM yesterday
    rc = d.grab_tanks(cfg(tank_csv_dir=t, tank_grab_attempts="2"), fetcher=lambda u: html(),
                      sleep=lambda s: None, now=now)
    assert rc == 6 and os.listdir(t) == []
    print("stale OK")


def test_no_timestamp_uses_fetch_time_and_later_kept():
    t = tempfile.mkdtemp()
    page = html().replace("&quot;created_at&quot;:&quot;2026-11-11T22:19:07.000000Z&quot;,", "")
    rows = d.parse_tank_page(page)
    assert all(r["reading_at"] is None for r in rows)
    late = lambda: dt.datetime(2026, 11, 11, 23, 55, 0)
    assert d.grab_tanks(cfg(tank_csv_dir=t), fetcher=lambda u: page, sleep=lambda s: None, now=late) == 0
    p = os.path.join(t, "tank-snapshot-2026-11-11.csv")
    assert d.parse_tank_csv(p)[1] == dt.datetime(2026, 11, 11, 23, 55, 0)
    # an earlier reading for the same business day must not replace the 23:55 one
    early = lambda: dt.datetime(2026, 11, 11, 17, 20, 0)
    assert d.grab_tanks(cfg(tank_csv_dir=t), fetcher=lambda u: html(), sleep=lambda s: None, now=early) == 0
    assert d.parse_tank_csv(p)[1] == dt.datetime(2026, 11, 11, 23, 55, 0)
    # 00:30 reading counts as the previous business day and replaces it (latest wins)
    page2 = page
    after_midnight = lambda: dt.datetime(2026, 11, 12, 0, 30, 0)
    assert d.grab_tanks(cfg(tank_csv_dir=t), fetcher=lambda u: page2, sleep=lambda s: None, now=after_midnight) == 0
    assert d.parse_tank_csv(p)[1] == dt.datetime(2026, 11, 12, 0, 30, 0)
    assert os.listdir(t) == ["tank-snapshot-2026-11-11.csv"]
    print("fetch_time + keep_later OK")


def test_default_dir_and_missing_warning():
    t = tempfile.mkdtemp()
    c = cfg(day_report_dir=os.path.join(t, "day_reports"))
    assert d.tank_csv_dir(c) == os.path.join(t, "day_reports", "tank_snapshots")
    c2 = cfg(day_report_dir="x", tank_csv_dir=os.path.join(t, "elsewhere"))
    assert d.tank_csv_dir(c2) == os.path.join(t, "elsewhere")
    assert d.warn_missing_tank_snapshot(c, today=dt.date(2026, 11, 12)) is False
    now = lambda: dt.datetime(2026, 11, 11, 17, 20, 0)
    assert d.grab_tanks(c, fetcher=lambda u: html(), sleep=lambda s: None, now=now) == 0
    assert os.path.exists(os.path.join(t, "day_reports", "tank_snapshots", "tank-snapshot-2026-11-11.csv"))
    assert d.warn_missing_tank_snapshot(c, today=dt.date(2026, 11, 12)) is True
    print("default_dir + warning OK")


def test_main_grab_mode_touches_nothing_else():
    t = tempfile.mkdtemp()
    dr = os.path.join(t, "day_reports")
    ini = os.path.join(t, "dayreport.ini")
    open(ini, "w").write("[dayreport]\nday_report_dir = %s\nlog_file = %s\ntank_fill = yes\n"
                         "local_inbox = %s\nsource_outbox = %s\ntank_grab_attempts = 1\n"
                         % (dr, os.path.join(dr, "logs", "dayreport.log"),
                            os.path.join(t, "inbox"), os.path.join(t, "no_such_pos")))
    orig = d.fetch_tank_page
    d.fetch_tank_page = lambda url=d.TANK_URL, timeout=30: html()
    try:
        # live clock: make the fixture fresh enough by disabling the age check
        open(ini, "a").write("tank_max_age_minutes = 0\n")
        rc = d.main(["--config", ini, "--grab-tanks"])
    finally:
        d.fetch_tank_page = orig
        for h in list(d.log.handlers):
            d.log.removeHandler(h); h.close()
    assert rc == 0, rc
    assert sorted(os.listdir(dr)) == ["logs", "tank_snapshots"], os.listdir(dr)
    assert not os.path.exists(os.path.join(t, "inbox"))     # no lock / staging touched
    log = open(os.path.join(dr, "logs", "dayreport.log")).read()
    assert "tank snapshot written" in log and "tank grab end, exit code 0" in log
    print("main --grab-tanks OK")


if __name__ == "__main__":
    test_parse_fixture()
    test_parse_table_fallback()
    test_parse_errors()
    test_grab_writes_csv_reader_accepts()
    test_retry_then_success_and_total_failure()
    test_stale_readings_fail()
    test_no_timestamp_uses_fetch_time_and_later_kept()
    test_default_dir_and_missing_warning()
    test_main_grab_mode_touches_nothing_else()
    print("ALL GRAB TANKS TESTS PASSED")
