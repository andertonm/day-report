"""Tank Watcher CSV -> stick inches. Run: python tests/test_tank_csv.py"""
import calendar, datetime as dt, logging, os, shutil, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import dayreport as d

logging.basicConfig(level=logging.INFO)
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SAMPLES = os.path.join(os.path.dirname(ROOT), "tankwatcher")
# samples live next to rebuild/ in the workspace; also accept a local copy
if not os.path.isdir(SAMPLES):
    SAMPLES = os.path.join(ROOT, "tankwatcher_samples")
TEMPLATE = os.path.join(os.path.dirname(ROOT), "mark_template_2026-10-26.xlsm")
if not os.path.exists(TEMPLATE):
    TEMPLATE = os.path.join(ROOT, "mark_template_2026-10-26.xlsm")


def test_business_date():
    assert d.tank_business_date(dt.datetime(2026, 11, 9, 12, 34, 7)) == dt.date(2026, 11, 9)
    assert d.tank_business_date(dt.datetime(2026, 11, 9, 6, 0, 0)) == dt.date(2026, 11, 9)
    assert d.tank_business_date(dt.datetime(2026, 11, 10, 0, 15, 0)) == dt.date(2026, 11, 9)
    assert d.tank_business_date(dt.datetime(2026, 11, 10, 5, 59, 59)) == dt.date(2026, 11, 9)
    assert d.tank_business_date(dt.datetime(2026, 11, 1, 1, 0, 0)) == dt.date(2026, 10, 31)
    print("business_date OK")


def test_parse_and_index():
    p = os.path.join(SAMPLES, "tank-snapshot-2026-11-09-1234.csv")
    assert os.path.exists(p), p
    bdate, when, heights = d.parse_tank_csv(p)
    assert bdate == dt.date(2026, 11, 9)
    assert heights["IR_Regular"] == 50
    assert heights["Diesel"] == 24
    assert heights["IR_Supreme"] == 57
    assert heights["IR_Supreme2"] == 14
    assert "Tank 3 and 4" not in heights

    # early-morning reading -> previous day
    early = os.path.join(tempfile.mkdtemp(), "early.csv")
    open(early, "w").write(
        'tank_number,product,volume,tc_volume,ullage,height,water,temperature,water_volume,reading_at\n'
        '01,Regular,1,1,1,33.456,0,70,0,"2026-11-10 01:05:00"\n'
        '02,Diesel,1,1,1,22.1,0,70,0,"2026-11-10 01:05:00"\n'
        '03,Premium,1,1,1,40,0,70,0,"2026-11-10 01:05:00"\n'
        '04,Premium,1,1,1,15.999,0,70,0,"2026-11-10 01:05:00"\n'
    )
    bdate, when, heights = d.parse_tank_csv(early)
    assert bdate == dt.date(2026, 11, 9)
    assert heights["IR_Regular"] == 33.46   # rounded to 2 decimals
    assert heights["IR_Supreme2"] == 16.0 or heights["IR_Supreme2"] == 16

    idx = d.load_tank_csv_index(SAMPLES)
    # later 13:23 snapshot should win over 12:34 for the same day
    assert dt.date(2026, 11, 9) in idx
    path, when, heights = idx[dt.date(2026, 11, 9)]
    assert when.hour == 13 and when.minute == 23
    assert heights["IR_Regular"] == 50  # same heights in both samples
    print("parse_and_index OK")


def _make_month_wb(dirpath, year, month, blank_days=None, prefill=None):
    """Copy template to a month workbook. prefill: {(day, sheet): inches}."""
    name = "day report %s %d.xlsm" % (d.MONTHS[month - 1], year)
    path = os.path.join(dirpath, name)
    shutil.copy2(TEMPLATE, path)
    if prefill:
        upd = {}
        for (day, sheet), inches in prefill.items():
            upd.setdefault(sheet, {})["D%d" % (7 + day)] = inches
        tmp = d.write_workbook(path, None, cell_updates=upd)
        os.replace(tmp, path)
    return path


def test_fill_blank_only():
    assert os.path.exists(TEMPLATE), TEMPLATE
    t = tempfile.mkdtemp()
    csv_dir = os.path.join(t, "csv")
    os.makedirs(csv_dir)
    # one CSV for Nov 9
    shutil.copy2(os.path.join(SAMPLES, "tank-snapshot-2026-11-09-1324.csv"),
                 os.path.join(csv_dir, "tank-snapshot-2026-11-09.csv"))
    dr = os.path.join(t, "day_reports")
    os.makedirs(dr)
    wb = _make_month_wb(dr, 2026, 11, prefill={(9, "Diesel"): 99.9})  # Diesel already typed

    cfg = {
        "day_report_dir": dr,
        "tank_fill": "yes",
        "tank_csv_dir": csv_dir,
        "template": TEMPLATE,
    }
    # ConfigParser-like getboolean
    class C(dict):
        def getboolean(self, k, default=False):
            v = self.get(k, default)
            if isinstance(v, bool):
                return v
            return str(v).lower() in ("yes", "true", "1", "on")
        def get(self, k, default=""):
            return dict.get(self, k, default)
    cfg = C(cfg)

    n = d.fill_tank_sticks(cfg, dates=[dt.date(2026, 11, 9)], dry_run=False, today=dt.date(2026, 11, 10))
    assert n == 3, n  # Regular, Supreme, Supreme2 - not Diesel
    sheets = d._read_sheets(wb, d.STICK_SHEETS)
    assert sheets["IR_Regular"][(7 + 9, 4)][1] == 50
    assert sheets["Diesel"][(7 + 9, 4)][1] == 99.9   # untouched
    assert sheets["IR_Supreme"][(7 + 9, 4)][1] == 57
    assert sheets["IR_Supreme2"][(7 + 9, 4)][1] == 14
    # backup made
    assert any(x.startswith("day report November 2026.xlsm.") for x in os.listdir(os.path.join(dr, "backups")))

    # second run: all filled -> writes 0, no overwrite
    n2 = d.fill_tank_sticks(cfg, dates=[dt.date(2026, 11, 9)], dry_run=False, today=dt.date(2026, 11, 10))
    assert n2 == 0
    sheets = d._read_sheets(wb, d.STICK_SHEETS)
    assert sheets["Diesel"][(7 + 9, 4)][1] == 99.9
    print("fill_blank_only OK")


def test_missing_csv_is_info_not_error():
    t = tempfile.mkdtemp()
    dr = os.path.join(t, "day_reports")
    os.makedirs(dr)
    _make_month_wb(dr, 2026, 11)
    class C(dict):
        def getboolean(self, k, default=False):
            v = self.get(k, default)
            return str(v).lower() in ("yes", "true", "1", "on") if not isinstance(v, bool) else v
        def get(self, k, default=""):
            return dict.get(self, k, default)
    cfg = C({"day_report_dir": dr, "tank_fill": "yes",
             "tank_csv_dir": os.path.join(t, "empty_csv"), "template": TEMPLATE})
    os.makedirs(cfg["tank_csv_dir"])
    n = d.fill_tank_sticks(cfg, dates=[dt.date(2026, 11, 5)], dry_run=False)
    assert n == 0
    print("missing_csv OK")


def test_disabled():
    class C(dict):
        def getboolean(self, k, default=False):
            return False
        def get(self, k, default=""):
            return ""
    assert d.fill_tank_sticks(C(), dates=[dt.date(2026, 11, 9)]) == 0
    print("disabled OK")


if __name__ == "__main__":
    test_business_date()
    test_parse_and_index()
    test_fill_blank_only()
    test_missing_csv_is_info_not_error()
    test_disabled()
    print("ALL TANK CSV TESTS PASSED")
