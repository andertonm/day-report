"""PDF printing tests with a mocked SumatraPDF. Run: python tests/test_print_pdfs.py"""
import configparser, logging, os, sys, tempfile, types
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import dayreport as d
logging.basicConfig(level=logging.INFO)


def setup():
    t = tempfile.mkdtemp()
    src, arch = os.path.join(t, "out"), os.path.join(t, "arch")
    os.makedirs(src)
    for n in ("a.pdf", "B.PDF", "msm110126.xml", "x.XLS"):
        open(os.path.join(src, n), "w").write(n)
    exe = os.path.join(t, "SumatraPDF.exe"); open(exe, "w").close()
    cp = configparser.ConfigParser(interpolation=None)
    cp["dayreport"] = {"source_outbox": src, "archive_dir": arch, "print_pdfs": "yes",
                       "pdf_printer": "Lexmark in store", "sumatra_exe": exe}
    return cp["dayreport"], src, arch


def run(rcs, dry=False):
    cfg, src, arch = setup()
    calls = []
    def runner(cmd, timeout, capture_output):
        calls.append(cmd); return types.SimpleNamespace(returncode=rcs[len(calls) - 1])
    import datetime as dt
    res = d.print_pdfs(cfg, types.SimpleNamespace(dry_run=dry), runner, today=dt.date(2026, 11, 2))
    return res, calls, sorted(os.listdir(src)), os.path.join(arch, "printed", "2026-11-02")


(ok, fail), calls, left, a = run([0, 0])
assert ok == 2 and fail == [] and left == ["msm110126.xml", "x.XLS"], left
assert sorted(os.listdir(a)) == ["B.PDF", "a.pdf"]
assert calls[0][1:4] == ["-print-to", "Lexmark in store", "-silent"]

(ok, fail), calls, left, a = run([1, 0])          # first (B.PDF) fails -> kept
assert ok == 1 and fail == ["B.PDF"] and "B.PDF" in left and "a.pdf" not in left

(ok, fail), calls, left, a = run([], dry=True)    # dry run: nothing printed or deleted
assert calls == [] and "a.pdf" in left and "B.PDF" in left

def boom(*a, **k): raise d.__dict__.get("subprocess", __import__("subprocess")).TimeoutExpired("x", 1)
cfg, src, arch = setup()
ok, fail = d.print_pdfs(cfg, types.SimpleNamespace(dry_run=False), boom)
assert ok == 0 and len(fail) == 2 and len([n for n in os.listdir(src) if n.lower().endswith(".pdf")]) == 2

cfg, src, arch = setup(); cfg["sumatra_exe"] = "/nope"
d.find_sumatra = lambda c: None
ok, fail = d.print_pdfs(cfg, types.SimpleNamespace(dry_run=False))
assert ok == 0 and len(fail) == 2
print("ALL PDF PRINT TESTS PASSED")
