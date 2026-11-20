#!/usr/bin/env python3
"""
dayreport.py - Python replacement for dayreport.exe (VFP Day Report v03) + its batch file.

What it does (same output as main.prg v03, without Excel):
  1. Pulls Passport MSM/MCM/ISM XML files from the POS outbox (UNC path), archives the
     originals, and stages them in a local inbox.
  2. For every business date found (oldest first):
       - creates "day report <Month> <YYYY>.xlsm" from the template if it does not exist yet
         (never overwrites an existing month workbook),
       - resets sheet "Day N" from Template!A1:T150, writes the MSM block at L1 and the
         MCM block at T1 exactly as v03 did,
       - appends the ISM item lines to dailysales.dbf (optional).
  3. Logs everything; exits non-zero on any error. Never shows a dialog.

Only the "Day N" sheet XML is rewritten. Every other part of the .xlsm (all other sheets
including the DEQ tank reconciliation sheets, VBA project, ActiveX button, comments,
charts, styles, shared strings) is copied through byte-for-byte. workbook.xml gets
fullCalcOnLoad="1" and calcChain.xml is dropped so Excel recalculates everything on open.

Exit codes: 0 ok / nothing to do, 1 error while processing, 2 config/usage error,
            3 source outbox unreachable (local files still processed), 4 another run active,
            5 PDF print failed, 6 --grab-tanks could not save tank readings.
"""
import argparse, configparser, copy, datetime as dt, hashlib, io, logging, os, re, shutil
import sys, tempfile, zipfile
from lxml import etree

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PNS = "http://schemas.openxmlformats.org/package/2006/relationships"
N = "{%s}" % NS
log = logging.getLogger("dayreport")

MSM_HEADER = ["mscode", "tender", "tsubcode", "amount", "count", "date"]
MCM_HEADER = ["merchcode", "descript", "discamount", "disccount", "promamount", "promcount",
              "refamount", "refcount", "salesqty", "salesamt", "transcount", "openamount",
              "opencount"]
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August",
          "September", "October", "November", "December"]


class DayReportError(Exception):
    pass


class CollisionError(DayReportError):
    """Day sheet already holds different data for the same date."""


# --------------------------------------------------------------------------- VFP emulation
def strextract(s, start, end, occ=1, nocase=False):
    """VFP STREXTRACT(s, start, end, occ[, 1]): text between the occ-th start and next end.
    Returns '' if either delimiter is missing (VFP default behaviour)."""
    hay = s.lower() if nocase else s
    a, b = (start.lower(), end.lower()) if nocase else (start, end)
    pos = -1
    for _ in range(occ):
        pos = hay.find(a, pos + 1)
        if pos < 0:
            return ""
    i = pos + len(a)
    j = hay.find(b, i)
    return "" if j < 0 else s[i:j]


_num = re.compile(r"^\s*[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?")


def val(s):
    """VFP VAL(): leading numeric part, 0 if none."""
    m = _num.match(s or "")
    return float(m.group(0)) if m else 0.0


def num(x):
    """Write integers as ints (Excel shows 2, not 2.0)."""
    return int(x) if float(x).is_integer() else x


def business_date(xml):
    """Date the report's data applies to: <EndDate> in the MovementHeader (same as v03).
    Passport closes the day around 21:00-22:00; Mark confirmed EndDate is the data date.
    Falls back to <BusinessDate> with a warning if EndDate is missing."""
    d = strextract(xml, "<EndDate>", "</EndDate>").strip()
    if not d:
        d = strextract(xml, "<BusinessDate>", "</BusinessDate>").strip()
        log.warning("no <EndDate> in file - using <BusinessDate> %s", d)
    try:
        return dt.date.fromisoformat(d)
    except ValueError:
        raise DayReportError("no valid <EndDate>/<BusinessDate> (got %r)" % d)


def parse_msm(xml):
    """Rows exactly as GET_MSM builds msm_data.dbf (incl. the trailing blank row)."""
    d = business_date(xml)
    fdate = d.strftime("%m%d%y")
    rows, occ = [], 1
    while True:
        det = strextract(xml, "<MSMDETAIL>", "</MSMDETAIL>", occ, nocase=True)
        code = (strextract(det, "<MiscellaneousSummaryCode>", "</MiscellaneousSummaryCode>").strip()
                + strextract(det, "<MiscellaneousSummarySubCode>", "</MiscellaneousSummarySubCode>").strip()
                + strextract(det, "<MiscellaneousSummarySubCodeModifier>", "</MiscellaneousSummarySubCodeModifier>").strip())
        rows.append([code,
                     strextract(det, "<TenderCode>", "</TenderCode>").strip(),
                     strextract(det, "<TenderSubCode>", "</TenderSubCode>").strip(),
                     num(val(strextract(det, "<MiscellaneousSummaryAmount>", "</MiscellaneousSummaryAmount>"))),
                     num(val(strextract(det, "<MiscellaneousSummaryCount>", "</MiscellaneousSummaryCount>"))),
                     fdate])
        occ += 1
        if det == "":          # VFP appends the blank record before testing -> keep it
            break
    return d, rows


def parse_mcm(xml):
    """Rows exactly as GET_MCM builds mcm_data.dbf: departments 1..32, gaps = NOT USED."""
    d = business_date(xml)
    fdate = d.strftime("%m%d%y")
    rows, last, occ = [], 0, 1
    g = lambda det, t, sp="": strextract(det, "<%s>" % t, "</%s>%s" % (t, sp))
    while last <= 32:
        det = strextract(xml, "<MCMDETAIL>", "</MCMDETAIL>", occ, nocase=True)
        cur = int(val(g(det, "MerchandiseCode")))
        if det == "" and last < 32:
            # v03 would loop forever here; we pad to 32 instead and log it.
            log.warning("MCM file ends before department 32 (last=%d); padding with NOT USED", last)
            cur = 33
        while cur - last > 1:
            last += 1
            rows.append([str(last), "NOT USED"] + [0] * 11 + [fdate])
            if last == 32:
                break
        if last == 32:
            break
        rows.append([g(det, "MerchandiseCode").strip(), g(det, "MerchandiseCodeDescription").strip()]
                    + [num(val(g(det, t))) for t in ("DiscountAmount", "DiscountCount", "PromotionAmount",
                                                     "PromotionCount", "RefundAmount", "RefundCount",
                                                     "SalesQuantity", "SalesAmount", "TransactionCount")]
                    # v03 searched for "</OpenDepartmentSalesAmount> " (trailing space), so these
                    # were always 0 unless the XML had that exact space. Reproduced on purpose.
                    + [num(val(g(det, "OpenDepartmentSalesAmount", " "))),
                       num(val(g(det, "OpenDepartmentSalesCount", " "))), fdate])
        last = cur
        if last == 32:
            break
        occ += 1
    # VFP column order: merchcode, descript, discamount, disccount, promamount?...
    return d, rows


def reorder_mcm(r):
    # parse order above: code, desc, discamt, disccnt, promamt, promcnt, refamt, refcnt, qty,
    # amt, trans, openamt, opencnt, date. VFP struct (from Sept workbook header):
    # merchcode descript discamount disccount promamount promcount refamount refcount salesqty
    # salesamt transcount openamount opencount  -> same order; date is not pasted (A1:M34).
    return r[:13]


def parse_ism(xml):
    d = business_date(xml)
    out, occ = [], 1
    while True:
        det = strextract(xml, "<ISMDETAIL>", "</ISMDETAIL>", occ, nocase=True)
        if det == "":
            break
        g = lambda t: strextract(det, "<%s>" % t, "</%s>" % t).strip()
        out.append(dict(SELLDATE=d.strftime("%Y%m%d"), UPC=g("POSCode"), UPCMOD=g("POSCodeModifier"),
                        ITEMID=g("ItemID"), ITEMDESC=g("Description"), MERCHCODE=g("MerchandiseCode"),
                        SELLUNIT=int(val(g("SellingUnits"))), SELLPRICE=val(g("ActualSalesPrice")),
                        QTYSOLD=int(val(g("SalesQuantity"))), SALESAMT=val(g("SalesAmount")),
                        DISCAMT=val(g("DiscountAmount")), DISCCOUNT=int(val(g("DiscountCount"))),
                        PROMOAMT=val(g("PromotionAmount")), PROMOCOUNT=int(val(g("PromotionCount"))),
                        REFUNDAMT=val(g("RefundAmount")), REFCOUNT=int(val(g("RefundCount"))),
                        TRANSCOUNT=int(val(g("TransactionCount")))))
        occ += 1
    return d, out


# --------------------------------------------------------------------------- xlsx helpers
def col2n(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


def n2col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


_ref = re.compile(r"^([A-Z]+)(\d+)$")


def split_ref(r):
    m = _ref.match(r)
    return col2n(m.group(1)), int(m.group(2))


def sheet_paths(z):
    wb = etree.fromstring(z.read("xl/workbook.xml"))
    rels = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    tgt = {r.get("Id"): r.get("Target") for r in rels}
    out = {}
    for s in wb.iter(N + "sheet"):
        t = tgt[s.get("{%s}id" % RNS)]
        out[s.get("name")] = "xl/" + t.lstrip("/").replace("xl/", "", 1) if not t.startswith("/xl/") else t[1:]
    return out


def shared_strings(z):
    try:
        root = etree.fromstring(z.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return ["".join(si.itertext()) for si in root.iter(N + "si")]


def cell_text(c, sst):
    t = c.get("t")
    if t == "inlineStr":
        return "".join(c.find(N + "is").itertext())
    v = c.find(N + "v")
    if v is None:
        return None
    return sst[int(v.text)] if t == "s" else v.text


def read_cell(z, sheet_xml, ref, sst):
    root = etree.fromstring(z.read(sheet_xml))
    for c in root.iter(N + "c"):
        if c.get("r") == ref:
            return cell_text(c, sst)
    return None


def make_cell(ref, value, style=None):
    c = etree.Element(N + "c", r=ref)
    if style:
        c.set("s", style)
    if isinstance(value, str):
        c.set("t", "inlineStr")
        is_ = etree.SubElement(c, N + "is")
        t = etree.SubElement(is_, N + "t")
        t.text = value
        if value != value.strip():
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    else:
        etree.SubElement(c, N + "v").text = repr(value) if isinstance(value, float) else str(value)
    return c


def build_day_sheet(day_xml, tmpl_xml, msm_rows, mcm_rows, extra=None):
    """Return new XML bytes for the Day sheet: Template A1:T150 pasted, then MSM at L1, MCM at T1."""
    day = etree.fromstring(day_xml)
    tmpl = etree.fromstring(tmpl_xml)
    sd = day.find(N + "sheetData")
    cells = {}            # (row, col) -> element
    rowattrs = {}
    for row in sd.findall(N + "row"):
        r = int(row.get("r"))
        rowattrs[r] = {k: v for k, v in row.attrib.items() if k not in ("r", "spans")}
        for c in row.findall(N + "c"):
            cells[split_ref(c.get("r"))[::-1]] = c

    # 1. Template!A1:T150 paste (values, formulas, styles); clears what was there
    for k in [k for k in cells if k[0] <= 150 and k[1] <= 20]:
        del cells[k]
    used_si = [int(f.get("si")) for c in cells.values() for f in c.findall(N + "f") if f.get("si") is not None]
    si_off = (max(used_si) + 1) if used_si else 0
    tcells = {}
    for c in tmpl.find(N + "sheetData").iter(N + "c"):
        col, r = split_ref(c.get("r"))
        tcells[(r, col)] = c
    masters = {}
    for (r, col), c in tcells.items():
        f = c.find(N + "f")
        if f is not None and f.get("t") == "shared" and f.get("ref"):
            masters[f.get("si")] = (r, col)
    for (r, col), c in tcells.items():
        if r > 150 or col > 20:
            continue
        nc = copy.deepcopy(c)
        f = nc.find(N + "f")
        if f is not None:
            if f.get("t") == "shared":
                m = masters.get(f.get("si"))
                if m is None or m[0] > 150 or m[1] > 20:
                    raise DayReportError("Template shared formula %s has master outside A1:T150" % c.get("r"))
                f.set("si", str(int(f.get("si")) + si_off))
            v = nc.find(N + "v")          # drop stale cached value; Excel recalcs on load
            if v is not None:
                nc.remove(v)
            if nc.get("t") in ("str", "e", "b", "n"):
                del nc.attrib["t"]
        cells[(r, col)] = nc

    if msm_rows is None:                   # reset-only: clear program blocks entirely
        for k in [k for k in cells if (12 <= k[1] <= 18) or (20 <= k[1] <= 32 and k[0] <= 34)]:
            if not (k[0] <= 150 and k[1] <= 20):
                del cells[k]
        msm_rows, mcm_rows, blank = [], [], True
    else:
        blank = False
    # 2. MSM block at L1 (cols L..R, VFP pasted A1:G100 -> also blanks rows up to 100)
    block = [] if blank else [MSM_HEADER] + msm_rows
    last = max(100, len(block))
    for k in [k for k in cells if 12 <= k[1] <= 18 and k[0] > 150]:
        del cells[k]                       # leftovers from a longer earlier MSM
    for i in range(last):
        r = i + 1
        for j in range(7):
            col = 12 + j
            old = cells.pop((r, col), None)
            style = old.get("s") if old is not None else None
            if i < len(block) and j < len(block[i]):
                cells[(r, col)] = make_cell(n2col(col) + str(r), block[i][j], style)
            elif style:
                cells[(r, col)] = etree.Element(N + "c", r=n2col(col) + str(r), s=style)

    # 3. MCM block at T1:AF34
    block = [] if blank else [MCM_HEADER] + [reorder_mcm(x) for x in mcm_rows]
    for i in range(34):
        for j in range(13):
            r, col = i + 1, 20 + j
            old = cells.pop((r, col), None)
            style = old.get("s") if old is not None else None
            if i < len(block):
                cells[(r, col)] = make_cell(n2col(col) + str(r), block[i][j], style)
            elif style:
                cells[(r, col)] = etree.Element(N + "c", r=n2col(col) + str(r), s=style)

    # 3b. optional extra cell values (not used by the nightly run)
    for ref, value in (extra or {}).items():
        col, r = split_ref(ref)
        old = cells.get((r, col))
        cells[(r, col)] = make_cell(ref, value, old.get("s") if old is not None else None)

    # rebuild sheetData
    for row in sd.findall(N + "row"):
        sd.remove(row)
    byrow = {}
    for (r, col), c in cells.items():
        byrow.setdefault(r, []).append((col, c))
    maxc = 1
    for r in sorted(set(byrow) | set(rowattrs)):
        row = etree.SubElement(sd, N + "row", r=str(r))
        for k, v in rowattrs.get(r, {}).items():
            row.set(k, v)
        for col, c in sorted(byrow.get(r, []), key=lambda x: x[0]):
            row.append(c)
            maxc = max(maxc, col)
    maxr = max(byrow) if byrow else 1
    dim = day.find(N + "dimension")
    if dim is not None:
        dim.set("ref", "A1:%s%d" % (n2col(maxc), maxr))

    # 4. merges: Excel paste replaces merges inside A1:T150 with Template's
    def inside(ref):
        a, b = (ref.split(":") + [ref])[:2]
        (c1, r1), (c2, r2) = split_ref(a), split_ref(b)
        return r2 <= 150 and c2 <= 20, (r1 <= 150 and c1 <= 20)
    mc = day.find(N + "mergeCells")
    keep = [m.get("ref") for m in (mc if mc is not None else []) if not inside(m.get("ref"))[1]]
    tm = tmpl.find(N + "mergeCells")
    add = [m.get("ref") for m in (tm if tm is not None else []) if inside(m.get("ref"))[0]]
    refs = keep + add
    if mc is not None:
        day.remove(mc)
    if refs:
        mc = etree.Element(N + "mergeCells", count=str(len(refs)))
        for ref in refs:
            etree.SubElement(mc, N + "mergeCell", ref=ref)
        # mergeCells goes right after sheetData / sheetCalcPr / sheetProtection / protectedRanges / scenarios / autoFilter / sortState / dataConsolidate / customSheetViews
        anchor = sd
        for tag in ("sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter",
                    "sortState", "dataConsolidate", "customSheetViews"):
            e = day.find(N + tag)
            if e is not None:
                anchor = e
        anchor.addnext(mc)

    # 5. column widths: VFP set A:O = 10, A = 36 (Excel stores 10.7109375 / 36.7109375)
    cols = day.find(N + "cols")
    if cols is None:
        cols = etree.Element(N + "cols")
        day.find(N + "sheetFormatPr").addnext(cols) if day.find(N + "sheetFormatPr") is not None else day.find(N + "sheetData").addprevious(cols)
    others = []
    for c in cols.findall(N + "col"):
        a, b = int(c.get("min")), int(c.get("max"))
        cols.remove(c)
        if b > 15:
            c.set("min", str(max(a, 16)))
            others.append(c)
    new = [etree.Element(N + "col", min="1", max="1", width="36.7109375", customWidth="1"),
           etree.Element(N + "col", min="2", max="15", width="10.7109375", customWidth="1")]
    for c in new + sorted(others, key=lambda e: int(e.get("min"))):
        cols.append(c)
    return etree.tostring(day, xml_declaration=True, encoding="UTF-8", standalone=True)


def write_workbook(path, day_name, msm_rows=None, mcm_rows=None, days=None, cell_updates=None):
    """days: {day_name: (msm_rows|None, mcm_rows|None, extra|None)}; msm None = reset to Template.
    cell_updates: {sheet_name: {ref: value}} plain values written into other sheets (opening inventory)."""
    if days is None and day_name:
        days = {day_name: (msm_rows, mcm_rows, None)}
    days = days or {}
    with zipfile.ZipFile(path) as z:
        sp = sheet_paths(z)
        replace = {}
        for dn, (m1, m2, ex) in days.items():
            if dn not in sp or "Template" not in sp:
                raise DayReportError("%s: sheet %r or 'Template' missing" % (path, dn))
            replace[sp[dn]] = build_day_sheet(z.read(sp[dn]), z.read(sp["Template"]), m1, m2, ex)
        for sn, cells in (cell_updates or {}).items():
            part = sp[sn]
            replace[part] = set_cells_xml(replace.get(part) or z.read(part), cells)
        wb = etree.fromstring(z.read("xl/workbook.xml"))
        cp = wb.find(N + "calcPr")
        if cp is None:
            cp = etree.SubElement(wb, N + "calcPr")
            # calcPr must follow definedNames; good enough to move it after definedNames if present
            dn = wb.find(N + "definedNames")
            if dn is not None:
                dn.addnext(cp)
        cp.set("fullCalcOnLoad", "1")
        new_wb = etree.tostring(wb, xml_declaration=True, encoding="UTF-8", standalone=True)
        rels = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        for r in list(rels):
            if r.get("Type", "").endswith("/calcChain"):
                rels.remove(r)
        new_rels = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
        ct = etree.fromstring(z.read("[Content_Types].xml"))
        for o in list(ct):
            if o.get("PartName") == "/xl/calcChain.xml":
                ct.remove(o)
        new_ct = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)
        replace.update({"xl/workbook.xml": new_wb,
                        "xl/_rels/workbook.xml.rels": new_rels, "[Content_Types].xml": new_ct})
        fd, tmp = tempfile.mkstemp(suffix=".xlsm", dir=os.path.dirname(os.path.abspath(path)))
        os.close(fd)
        try:
            with zipfile.ZipFile(tmp, "w") as out:
                for info in z.infolist():
                    if info.filename == "xl/calcChain.xml":
                        continue
                    data = replace.get(info.filename, None)
                    if data is None:
                        data = z.read(info.filename)        # byte-for-byte passthrough
                    out.writestr(info, data, compress_type=info.compress_type)
            with zipfile.ZipFile(tmp) as chk:               # sanity check
                if chk.testzip() is not None:
                    raise DayReportError("written workbook failed zip test")
                for part in replace:
                    etree.fromstring(chk.read(part))
        except Exception:
            os.unlink(tmp)
            raise
    return tmp


def day_msm_pairs(path, day_name):
    """[(mscode, amount)] currently in Day N's MSM block (see sheet_msm)."""
    return [(k, a) for k, a, c in sheet_msm(path, day_name)[0]]


def day_state(path, day_name):
    """(L1, Q2) of the Day sheet: tells whether the day was already filled and for which date."""
    with zipfile.ZipFile(path) as z:
        sp = sheet_paths(z)
        sst = shared_strings(z)
        return read_cell(z, sp[day_name], "L1", sst), read_cell(z, sp[day_name], "Q2", sst)


# --------------------------------------------------------------------------- dailysales.dbf
def append_dailysales(dbf_path, rows, force):
    import dbf  # pip install dbf
    t = dbf.Table(dbf_path)
    with t.open(mode=dbf.READ_WRITE):
        dates = {r["SELLDATE"] for r in rows}
        existing = [rec for rec in t if rec.selldate.strip() in dates and not dbf.is_deleted(rec)]
        if existing and not force:
            log.info("dailysales.dbf already has %d lines for %s - skipped", len(existing), ",".join(sorted(dates)))
            return 0
        for rec in existing:
            dbf.delete(rec)
        for r in rows:
            t.append({k.lower(): v for k, v in r.items()})
        if existing:
            t.pack()
    return len(rows)


# --------------------------------------------------------------------------- file handling
KINDS = ("MSM", "MCM", "ISM")
FUEL_KEYS = ("fuelSalesByGradefuel001", "fuelSalesByGradefuel002", "fuelSalesByGradefuel004",
             "fuelSalesByGradefuel019")


def fetch(cfg, args):
    """Mirror of the old batch file: copy MSM/MCM/ISM to archive + local inbox, move other XML
    to archive, remove from the POS outbox. Returns False if the source was unreachable."""
    src = cfg["source_outbox"]
    inbox, arch = cfg["local_inbox"], os.path.join(cfg["archive_dir"], "received")
    os.makedirs(inbox, exist_ok=True)
    os.makedirs(arch, exist_ok=True)
    try:
        names = sorted(os.listdir(src))
    except OSError as e:
        log.error("cannot read POS outbox %s: %s", src, e)
        return False
    n = 0
    for name in names:
        p = os.path.join(src, name)
        if not os.path.isfile(p) or not name.lower().endswith(".xml"):
            continue
        a = os.path.join(arch, name)
        if not os.path.exists(a):
            shutil.copy2(p, a)
        if name[:3].upper() in KINDS:
            shutil.copy2(p, os.path.join(inbox, name))
            n += 1
        if cfg.getboolean("remove_from_source", True) and not args.dry_run:
            os.remove(p)
    log.info("fetched %d MSM/MCM/ISM files from %s", n, src)
    return True


def collect(inbox):
    """{date: {'MSM':[path..], 'MCM':[..], 'ISM':[..]}} plus list of bad files."""
    by_date, bad = {}, []
    for name in sorted(os.listdir(inbox)):
        p = os.path.join(inbox, name)
        kind = name[:3].upper()
        if not os.path.isfile(p) or name.startswith("."):
            continue
        if kind not in KINDS or not name.lower().endswith(".xml"):
            log.warning("ignoring unexpected file in inbox: %s", name)   # v03 aborted here
            continue
        try:
            with open(p, encoding="iso-8859-1") as f:
                d = business_date(f.read())
        except Exception as e:
            log.error("%s: %s", name, e)
            bad.append(p)
            continue
        by_date.setdefault(d, {k: [] for k in KINDS})[kind].append(p)
    return by_date, bad


def read(p):
    with open(p, encoding="iso-8859-1") as f:
        return f.read()


# Tank sheets that pull each day's gallons from the Day sheets by formula (column B, row 7+day).
TANK_SHEETS = ("IR_Regular", "Diesel", "IR_Supreme", "Tank 3 and 4")


def check_tank_links(path, day):
    """Read-only check: on each tank sheet, Sales cell B(7+day) should be a formula that
    points at 'Day <day>' only. Returns a list of problem strings (logged as warnings)."""
    probs = []
    ref = "B%d" % (7 + day)
    with zipfile.ZipFile(path) as z:
        paths = sheet_paths(z)
        for name in TANK_SHEETS:
            if name not in paths:
                continue
            root = etree.fromstring(z.read(paths[name]))
            cell = next((c for c in root.iter(N + "c") if c.get("r") == ref), None)
            f = cell.find(N + "f") if cell is not None else None
            if f is None:
                probs.append("%s!%s has no formula (typed value) - Day %d gallons will not show there"
                             % (name, ref, day))
                continue
            if not f.text:
                continue  # shared-formula child; cannot check text
            days = set(re.findall(r"'?Day (\d+)'?!", f.text))
            if days and days != {str(day)}:
                probs.append("%s!%s = %s refers to Day %s, expected Day %d only"
                             % (name, ref, f.text, ",".join(sorted(days, key=int)), day))
    return probs


# --------------------------------------------------------------------------- opening inventory carry-forward
# Tank sheets: row 7 col E = opening gallons (typed). Rows 8.. = days; col D = stick inches (typed),
# col E = gallons from the IR_Parameters strapping chart. "Tank 3 and 4" has no sticks: E = IR_Supreme + IR_Supreme2.
STICK_SHEETS = ("IR_Regular", "Diesel", "IR_Supreme", "IR_Supreme2")
COMBINED_SHEET = ("Tank 3 and 4", ("IR_Supreme", "IR_Supreme2"))


def set_cells_xml(sheet_xml, cells):
    """Write plain values into existing/new cells, keeping each cell's style and row order."""
    root = etree.fromstring(sheet_xml)
    sd = root.find(N + "sheetData")
    for ref, value in cells.items():
        col, r = split_ref(ref)
        row = next((x for x in sd.findall(N + "row") if int(x.get("r")) == r), None)
        if row is None:
            row = etree.Element(N + "row", r=str(r))
            after = [x for x in sd.findall(N + "row") if int(x.get("r")) < r]
            (after[-1].addnext(row) if after else sd.insert(0, row))
        old = next((c for c in row.findall(N + "c") if c.get("r") == ref), None)
        new = make_cell(ref, value, old.get("s") if old is not None else None)
        if old is not None:
            old.addprevious(new); row.remove(old)
        else:
            later = [c for c in row.findall(N + "c") if split_ref(c.get("r"))[0] > col]
            (later[0].addprevious(new) if later else row.append(new))
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _read_sheets(path, names):
    with zipfile.ZipFile(path) as z:
        sp = shared_strings(z), sheet_paths(z)
        return {n: _sheet_cells(z, sp[1][n], sp[0]) for n in names if n in sp[1]}


def _strap_col(cells):
    f = cells.get((8, 5), ("", ""))[1] or ""
    m = re.search(r"IR_Parameters!\$?([A-Z]+)\$?13:\$?[A-Z]+\$?747\s*\)?\s*$", f)
    return m.group(1) if m else None


def _strap_gallons(params, col, inches):
    """Excel LOOKUP(inches, A13:A747, col13:col747); A13 = 0.125 and each row adds 0.125."""
    if inches is None or inches < 0.125:
        return None
    r = min(747, 12 + int(round(inches / 0.125, 6) // 1))
    v = params.get((r, col2n(col)), ("V", None))[1]
    return v if isinstance(v, (int, float)) else None


def month_closing(prev_path, ndays):
    """Closing gallons per tank sheet from the prior month's last calendar day row (7+ndays).
    Returns ({sheet: gallons}, [sheets missing a reading])."""
    names = STICK_SHEETS + (COMBINED_SHEET[0], "IR_Parameters")
    sh = _read_sheets(prev_path, names)
    params = sh.get("IR_Parameters", {})
    row = 7 + ndays
    out, missing = {}, []
    for n in STICK_SHEETS:
        if n not in sh:
            continue
        kind, inches = sh[n].get((row, 4), ("V", None))
        col = _strap_col(sh[n])
        if kind != "V" or not isinstance(inches, (int, float)) or col is None:
            missing.append(n)
            continue
        g = _strap_gallons(params, col, inches)
        if g is None:
            missing.append(n)
            continue
        out[n] = num(float(g))
    parts = COMBINED_SHEET[1]
    if COMBINED_SHEET[0] in sh and all(p in out for p in parts):
        out[COMBINED_SHEET[0]] = num(float(sum(out[p] for p in parts)))
    return out, missing


def carry_forward_opening(cfg, wbpath, d, dry_run=False):
    """Copy last month's closing gallons into this month's opening row (E7) on each tank sheet.
    Writes a cell only if it is blank or still holds the template's value; never touches typed entries."""
    first = d.replace(day=1)
    prev = first - dt.timedelta(days=1)
    pname = "day report %s %d.xlsm" % (MONTHS[prev.month - 1], prev.year)
    ppath = os.path.join(cfg["day_report_dir"], pname)
    wbname = os.path.basename(wbpath)
    if not os.path.exists(ppath):
        return
    names = STICK_SHEETS + (COMBINED_SHEET[0],)
    cur = _read_sheets(wbpath, names)
    tmpl = _read_sheets(cfg["template"], names) if os.path.exists(cfg.get("template", "")) else {}
    todo = []
    for n in names:
        if n not in cur:
            continue
        v = cur[n].get((7, 5), ("V", None))
        t = tmpl.get(n, {}).get((7, 5), ("V", None))
        if v[1] is None or (isinstance(v[1], str) and not v[1].strip()) or v == t:
            todo.append(n)
    if not todo:
        return
    closing, missing = month_closing(ppath, prev.day)
    upd = {n: {"E7": closing[n]} for n in todo if n in closing}
    waiting = [n for n in todo if n not in closing]
    if waiting:
        log.info("%s: opening inventory for %s waits for %s day %d stick reading in %s - will retry next run",
                 wbname, ", ".join(waiting), pname, prev.day, ", ".join(missing) or "(sheets)")
    if not upd:
        return
    desc = ", ".join("%s!E7=%s" % (n, upd[n]["E7"]) for n in upd)
    if dry_run:
        log.info("[dry-run] would carry opening inventory from %s day %d: %s", pname, prev.day, desc)
        return
    bdir = os.path.join(cfg["day_report_dir"], "backups")
    os.makedirs(bdir, exist_ok=True)
    shutil.copy2(wbpath, os.path.join(bdir, "%s.%s.bak" % (wbname, dt.datetime.now().strftime("%Y%m%d-%H%M%S"))))
    tmp = write_workbook(wbpath, None, cell_updates=upd)
    try:
        os.replace(tmp, wbpath)
    except PermissionError:
        os.unlink(tmp)
        log.warning("%s is open/locked - opening inventory not carried yet, will retry", wbname)
        return
    log.info("%s: opening inventory carried from %s day %d closing: %s", wbname, pname, prev.day, desc)


# --------------------------------------------------------------------------- Tank Watcher stick inches
# Nightly CSV from gastanks.watch (saved by --grab-tanks). Column D on each stick
# sheet is typed inches; Excel's IR_Parameters chart turns them into gallons in E.
# "Tank 3 and 4" has no sticks - it sums IR_Supreme + IR_Supreme2 by formula/carry.
TANK_CSV_MAP = {
    "01": "IR_Regular",
    "02": "Diesel",
    "03": "IR_Supreme",
    "04": "IR_Supreme2",
}
# reading_at before 06:00 counts as the previous calendar day's closing inventory
# (a late-night/early-morning snapshot still belongs to the day that just ended).
TANK_DAY_CUTOFF_HOUR = 6


def tank_business_date(reading_at):
    """Map a TLS snapshot timestamp to the day-report business date.

    A reading at or after 06:00 keeps its calendar date. A reading between
    midnight and 05:59 is treated as the previous day's closing inventory.
    """
    if reading_at.hour < TANK_DAY_CUTOFF_HOUR:
        return (reading_at - dt.timedelta(days=1)).date()
    return reading_at.date()


def parse_tank_csv(path):
    """Return (business_date, reading_at, {sheet: height_inches}) or None if unusable."""
    import csv
    try:
        with open(path, newline="", encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
    except OSError as e:
        log.warning("tank csv unreadable %s: %s", path, e)
        return None
    if not rows:
        return None
    heights, reading_at = {}, None
    for row in rows:
        try:
            ts = (row.get("reading_at") or "").strip().strip('"')
            when = dt.datetime.strptime(ts, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            log.warning("tank csv %s: bad reading_at %r", path, row.get("reading_at"))
            return None
        if reading_at is None:
            reading_at = when
        elif when != reading_at:
            # keep the latest timestamp in the file as the snapshot time
            if when > reading_at:
                reading_at = when
        tnum = (row.get("tank_number") or "").strip().zfill(2)
        sheet = TANK_CSV_MAP.get(tnum)
        if not sheet:
            continue
        try:
            inches = float(row.get("height") or "")
        except ValueError:
            log.warning("tank csv %s: bad height for tank %s", path, tnum)
            continue
        # Sheet values are hand-typed to hundredths (e.g. 26.49, 14.43).
        heights[sheet] = num(round(inches, 2))
    if not heights or reading_at is None:
        return None
    return tank_business_date(reading_at), reading_at, heights


def load_tank_csv_index(csv_dir):
    """Scan csv_dir for *.csv; return {business_date: (path, reading_at, heights)}.

    When more than one file maps to the same business date, the one with the
    latest reading_at wins (nightly closing beats earlier test snapshots).
    """
    out = {}
    if not csv_dir or not os.path.isdir(csv_dir):
        return out
    try:
        names = sorted(n for n in os.listdir(csv_dir) if n.lower().endswith(".csv"))
    except OSError as e:
        log.warning("tank csv dir unreadable %s: %s", csv_dir, e)
        return out
    for n in names:
        parsed = parse_tank_csv(os.path.join(csv_dir, n))
        if not parsed:
            continue
        bdate, when, heights = parsed
        prev = out.get(bdate)
        if prev is None or when >= prev[1]:
            out[bdate] = (os.path.join(csv_dir, n), when, heights)
    return out


def _blank_stick(cells, row):
    """True if column D at row is empty (safe to fill). Never overwrite a typed value."""
    kind, val = cells.get((row, 4), ("V", None))
    if kind == "F":
        return False
    if val is None:
        return True
    if isinstance(val, str) and not val.strip():
        return True
    return False


def fill_tank_sticks(cfg, dates=None, dry_run=False, today=None):
    """Write Tank Watcher heights (inches) into blank column-D cells on stick sheets.

    For each business date in `dates` (or every day of the current and previous
    month workbooks when dates is None), look up that day's CSV and fill
    IR_Regular / Diesel / IR_Supreme / IR_Supreme2 at D(7+day) only when blank.
    Missing CSV = info log, not an error. Returns number of cells written.
    """
    if not cfg.getboolean("tank_fill", False):
        return 0
    csv_dir = tank_csv_dir(cfg)
    if not csv_dir:
        log.info("tank_fill is on but tank_csv_dir is empty - skipping stick fill")
        return 0
    index = load_tank_csv_index(csv_dir)
    today = today or dt.date.today()
    if dates is None:
        # Current + previous calendar months cover carry-forward of last month's
        # closing sticks and any blank days staff have not typed yet.
        first = today.replace(day=1)
        prev_last = first - dt.timedelta(days=1)
        months = {(today.year, today.month), (prev_last.year, prev_last.month)}
        dates = []
        for y, m in sorted(months):
            import calendar
            for day in range(1, calendar.monthrange(y, m)[1] + 1):
                dates.append(dt.date(y, m, day))
    else:
        dates = list(dates)
    # Group dates by workbook
    by_wb = {}
    for d in dates:
        wbname = "day report %s %d.xlsm" % (MONTHS[d.month - 1], d.year)
        by_wb.setdefault(wbname, []).append(d)

    written = 0
    for wbname, days in by_wb.items():
        wbpath = os.path.join(cfg["day_report_dir"], wbname)
        if not os.path.exists(wbpath):
            continue
        try:
            sheets = _read_sheets(wbpath, STICK_SHEETS)
        except Exception as e:
            log.warning("%s: cannot read stick sheets (%s)", wbname, e)
            continue
        upd = {}  # sheet -> {ref: value}
        desc = []
        for d in days:
            row = 7 + d.day
            need = [n for n in STICK_SHEETS
                    if n in sheets and _blank_stick(sheets[n], row)]
            if not need:
                continue
            hit = index.get(d)
            if not hit:
                log.info("%s: no Tank Watcher CSV for %s (blank stick cells left for staff)",
                         wbname, d.isoformat())
                continue
            path, when, heights = hit
            for n in need:
                if n not in heights:
                    continue
                ref = "D%d" % row
                upd.setdefault(n, {})[ref] = heights[n]
                desc.append("%s!%s=%s" % (n, ref, heights[n]))
        if not upd:
            continue
        if dry_run:
            log.info("[dry-run] would fill stick inches in %s from Tank Watcher: %s",
                     wbname, ", ".join(desc))
            written += sum(len(v) for v in upd.values())
            continue
        bdir = os.path.join(cfg["day_report_dir"], "backups")
        os.makedirs(bdir, exist_ok=True)
        stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
        shutil.copy2(wbpath, os.path.join(bdir, "%s.%s.bak" % (wbname, stamp)))
        tmp = write_workbook(wbpath, None, cell_updates=upd)
        try:
            os.replace(tmp, wbpath)
        except PermissionError:
            os.unlink(tmp)
            log.warning("%s is open/locked - stick inches not filled yet, will retry", wbname)
            continue
        written += sum(len(v) for v in upd.values())
        log.info("%s: filled stick inches from Tank Watcher: %s", wbname, ", ".join(desc))
    return written


# --------------------------------------------------------------------------- Tank Watcher web grab
# Replaces the Gmail -> Apps Script -> Drive route. A separate nightly task
# (DayReport_tanks_task.xml, 11:55 PM) runs  run_dayreport.bat --grab-tanks,
# which first reads the store Pi's local two-decimal CSV
# (http://192.168.0.28:8088/tanks-YYYY-MM-DD.csv) and falls back to the public
# page https://gastanks.watch/home if the Pi is unreachable, incomplete, or
# stale. It writes one tank-snapshot-YYYY-MM-DD.csv into tank_csv_dir. The
# 5:30 AM run then fills column D from it with fill_tank_sticks() above
# (heights rounded to 2 decimals). The grab never touches a workbook or the
# POS outbox.
TANK_URL = "https://gastanks.watch/home"
TANK_PI_URL = "http://192.168.0.28:8088"
TANK_PI_TIMEOUT = 10
TANK_PI_MAX_AGE_MINUTES = 60
TANK_CSV_COLUMNS = ["tank_number", "product", "volume", "tc_volume", "ullage", "height",
                    "water", "temperature", "water_volume", "reading_at"]
# Product names are not in the Pi CSV; these match gastanks.watch / the TLS-350.
TANK_PRODUCTS = {"01": "Regular", "02": "Diesel", "03": "Premium", "04": "Premium"}
EXIT_TANK_GRAB = 6


class TankPageError(Exception):
    pass


def tank_csv_dir(cfg):
    """tank_csv_dir from the ini, or <day_report_dir>\\tank_snapshots when unset."""
    d = (cfg.get("tank_csv_dir", "") or "").strip()
    if d:
        return d
    base = (cfg.get("day_report_dir", "") or "").strip()
    return os.path.join(base, "tank_snapshots") if base else ""


def _utc_to_local(ts):
    """'2026-11-11T22:19:07.000000Z' -> naive local datetime (PC's time zone)."""
    ts = (ts or "").strip()
    if not ts:
        return None
    if ts.endswith("Z"):
        ts = ts[:-1] + "+00:00"
    m = re.match(r"^(.*T\d\d:\d\d:\d\d)(\.\d+)?(.*)$", ts)
    if m:                                   # drop fractions (older Pythons choke on 6+ digits)
        ts = m.group(1) + m.group(3)
    try:
        when = dt.datetime.fromisoformat(ts)
    except ValueError:
        return None
    if when.tzinfo is not None:
        when = when.astimezone().replace(tzinfo=None)
    return when


def _tank_float(x):
    if x is None:
        return None
    if isinstance(x, (int, float)):
        return float(x)
    x = str(x).replace(",", "").strip()
    try:
        return float(x)
    except ValueError:
        return None


def _parse_tank_json(page):
    rows = []
    for item in (page.get("props") or {}).get("tanks") or []:
        tank = item.get("tank") or {}
        data = item.get("data") or {}
        rows.append({
            "tank_number": str(tank.get("tank_number") or "").strip().zfill(2),
            "product": str(tank.get("tank_type") or "").strip(),
            "volume": _tank_float(data.get("volume")),
            "tc_volume": _tank_float(data.get("tc_volume")),
            "ullage": _tank_float(data.get("ullage")),
            "height": _tank_float(data.get("height")),
            "water": _tank_float(data.get("water")),
            "temperature": _tank_float(data.get("temperature")),
            "water_volume": _tank_float(data.get("water_volume")),
            "reading_at": _utc_to_local(data.get("created_at")),
        })
    return rows


def _parse_tank_table(tables):
    """Fallback for a server-rendered <table>: header row needs 'Tank' and 'Height'."""
    keys = {"tank": "tank_number", "tank #": "tank_number", "tank number": "tank_number",
            "product": "product", "volume": "volume", "tc volume": "tc_volume",
            "ullage": "ullage", "height": "height", "water": "water", "temp": "temperature",
            "temperature": "temperature", "water volume": "water_volume"}
    for t in tables:
        if not t:
            continue
        head = [c.strip().lower() for c in t[0]]
        if not any(h.startswith("tank") for h in head) or "height" not in head:
            continue
        idx = {}
        for i, h in enumerate(head):
            k = keys.get(h)
            if k and k not in idx:
                idx[k] = i
        rows = []
        for r in t[1:]:
            if len(r) < len(head):
                continue
            row = {k: (r[i].strip() if k in ("tank_number", "product") else _tank_float(r[i]))
                   for k, i in idx.items()}
            num_m = re.match(r"\s*(\d+)\s*(.*)", row.get("tank_number") or "")
            if not num_m:
                continue
            row["tank_number"] = num_m.group(1).zfill(2)
            if not row.get("product") and num_m.group(2):
                row["product"] = num_m.group(2).strip(" -")
            row["reading_at"] = None
            rows.append(row)
        if rows:
            return rows
    return []


def parse_tank_page(html_text):
    """Return a list of tank dicts (TANK_CSV_COLUMNS keys) from the gastanks.watch home page.

    The page is an Inertia.js app: the readings ride in <div id="app" data-page="{json}">,
    including each reading's own UTC timestamp (data.created_at). reading_at is that time
    in local time, or None if the page did not give one. Raises TankPageError.
    """
    import html.parser, json

    class P(html.parser.HTMLParser):
        def __init__(self):
            super().__init__(convert_charrefs=True)
            self.data_page, self.tables, self._cell = None, [], None

        def handle_starttag(self, tag, attrs):
            a = dict(attrs)
            if self.data_page is None and "data-page" in a:
                self.data_page = a["data-page"]
            if tag == "table":
                self.tables.append([])
            elif tag == "tr" and self.tables:
                self.tables[-1].append([])
            elif tag in ("td", "th") and self.tables and self.tables[-1]:
                self._cell = []

        def handle_endtag(self, tag):
            if tag in ("td", "th") and self._cell is not None:
                self.tables[-1][-1].append(" ".join("".join(self._cell).split()))
                self._cell = None

        def handle_data(self, data):
            if self._cell is not None:
                self._cell.append(data)

    p = P()
    p.feed(html_text or "")
    rows = []
    if p.data_page:
        try:
            rows = _parse_tank_json(json.loads(p.data_page))
        except (ValueError, AttributeError, TypeError) as e:
            raise TankPageError("tank page data unreadable: %s" % e)
    if not rows:
        rows = _parse_tank_table(p.tables)
    rows = [r for r in rows if r.get("tank_number") and r.get("height") is not None]
    if not rows:
        raise TankPageError("no tank readings found on the page (site layout changed?)")
    missing = sorted(set(TANK_CSV_MAP) - {r["tank_number"] for r in rows})
    if missing:
        raise TankPageError("tank(s) %s missing from the page" % ", ".join(missing))
    return rows


def fetch_tank_page(url=TANK_URL, timeout=30):
    import urllib.request
    req = urllib.request.Request(url, headers={
        "User-Agent": "DayReport/1.0 (SDS Office PC tank snapshot)",
        "Accept": "text/html"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        if getattr(r, "status", 200) != 200:
            raise TankPageError("HTTP %s from %s" % (r.status, url))
        raw = r.read()
        cs = r.headers.get_content_charset() or "utf-8"
    return raw.decode(cs, errors="replace")


def _fmt_csv_num(x):
    if x is None:
        return ""
    x = round(float(x), 2)
    return str(int(x)) if x == int(x) else ("%.2f" % x).rstrip("0").rstrip(".")


def write_tank_snapshot(csv_dir, rows, reading_at):
    """Write tank-snapshot-<business date>.csv atomically. Keeps an existing file for that
    business date if it holds a LATER reading. Returns (path, written: bool)."""
    import csv
    bdate = tank_business_date(reading_at)
    os.makedirs(csv_dir, exist_ok=True)
    path = os.path.join(csv_dir, "tank-snapshot-%s.csv" % bdate.isoformat())
    if os.path.exists(path):
        old = parse_tank_csv(path)
        if old and old[1] > reading_at:
            log.info("tank snapshot %s already holds a later reading (%s) - kept", path, old[1])
            return path, False
    stamp = reading_at.strftime("%Y-%m-%d %H:%M:%S")
    tmp = path + ".tmp"
    with open(tmp, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(TANK_CSV_COLUMNS)
        for r in sorted(rows, key=lambda r: r["tank_number"]):
            w.writerow([r["tank_number"], r.get("product") or ""] +
                       [_fmt_csv_num(r.get(k)) for k in
                        ("volume", "tc_volume", "ullage", "height", "water", "temperature",
                         "water_volume")] + [stamp])
    os.replace(tmp, path)
    return path, True


def tank_pi_base_url(cfg):
    """Pi HTTP base, no trailing slash. Empty tank_pi_url in the ini disables the Pi."""
    if cfg is not None and "tank_pi_url" in cfg:
        return (cfg.get("tank_pi_url") or "").strip().rstrip("/")
    return TANK_PI_URL


def fetch_text(url, timeout=TANK_PI_TIMEOUT):
    """GET url and return decoded text. Used for the Pi CSV (short timeout)."""
    import urllib.request
    req = urllib.request.Request(url, headers={
        "User-Agent": "DayReport/1.0 (SDS Office PC tank snapshot)",
        "Accept": "text/csv,text/plain,*/*"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        if getattr(r, "status", 200) != 200:
            raise TankPageError("HTTP %s from %s" % (r.status, url))
        raw = r.read()
        cs = r.headers.get_content_charset() or "utf-8"
    return raw.decode(cs, errors="replace")


def parse_tank_pi_csv(text):
    """Parse a Pi tanks-YYYY-MM-DD.csv into row dicts (same keys as parse_tank_page).

    Every data row is returned; call latest_rows_per_tank() to keep the newest
    reading of each tank. Product comes from the file if present, else TANK_PRODUCTS.
    """
    import csv, io
    rows = []
    reader = csv.DictReader(io.StringIO(text or ""))
    if not reader.fieldnames:
        return rows
    for raw in reader:
        tank = (raw.get("tank") or raw.get("tank_number") or "").strip()
        if not tank:
            continue
        tnum = tank.zfill(2)
        ts = (raw.get("timestamp") or raw.get("reading_at") or "").strip().strip('"')
        try:
            when = dt.datetime.strptime(ts, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            continue
        height = _tank_float(raw.get("height"))
        if height is None:
            continue
        rows.append({
            "tank_number": tnum,
            "product": (raw.get("product") or "").strip() or TANK_PRODUCTS.get(tnum, ""),
            "volume": _tank_float(raw.get("volume")),
            "tc_volume": _tank_float(raw.get("tc_volume")),
            "ullage": _tank_float(raw.get("ullage")),
            "height": height,
            "water": _tank_float(raw.get("water")),
            "temperature": _tank_float(raw.get("temp") or raw.get("temperature")),
            "water_volume": _tank_float(raw.get("water_volume")),
            "reading_at": when,
        })
    return rows


def latest_rows_per_tank(rows):
    """Keep the latest reading_at row for each tank_number, sorted by tank."""
    best = {}
    for r in rows:
        t = r["tank_number"]
        prev = best.get(t)
        if prev is None or (r.get("reading_at") or dt.datetime.min) >= (prev.get("reading_at") or dt.datetime.min):
            best[t] = r
    return [best[k] for k in sorted(best)]


def _fetch_one_pi_csv(get, url, errors):
    try:
        text = get(url)
    except Exception as e:
        errors.append("%s (%s: %s)" % (url, type(e).__name__, e))
        return []
    rows = parse_tank_pi_csv(text)
    if not rows:
        errors.append("%s (empty or unreadable)" % url)
    return rows


def fetch_pi_tank_rows(cfg, fetched, fetcher=None):
    """Latest-per-tank rows from the Pi, or None if they should not be used.

    Today's tanks-YYYY-MM-DD.csv is fetched first. Yesterday's file is also
    considered when the run is before 06:00 or today's file is missing/empty.
    Rows are used only when every tank in TANK_CSV_MAP is present and the
    newest reading is within tank_pi_max_age_minutes (default 60).
    """
    base = tank_pi_base_url(cfg)
    if not base:
        log.info("tank grab: Pi URL is empty - skipping Pi, using gastanks.watch")
        return None
    try:
        max_age = cfg.getint("tank_pi_max_age_minutes", TANK_PI_MAX_AGE_MINUTES)
    except (TypeError, ValueError):
        max_age = TANK_PI_MAX_AGE_MINUTES
    get = fetcher or (lambda url: fetch_text(url, timeout=TANK_PI_TIMEOUT))
    today = fetched.date()
    yesterday = today - dt.timedelta(days=1)
    errors = []
    today_url = "%s/tanks-%s.csv" % (base, today.isoformat())
    collected = _fetch_one_pi_csv(get, today_url, errors)
    if fetched.hour < TANK_DAY_CUTOFF_HOUR or not collected:
        y_url = "%s/tanks-%s.csv" % (base, yesterday.isoformat())
        collected.extend(_fetch_one_pi_csv(get, y_url, errors))
    if not collected:
        log.info("tank grab: Pi unreachable or no data (%s) - falling back to gastanks.watch",
                 "; ".join(errors) or "no files")
        return None
    rows = latest_rows_per_tank(collected)
    missing = sorted(set(TANK_CSV_MAP) - {r["tank_number"] for r in rows})
    if missing:
        log.info("tank grab: Pi data missing tank(s) %s - falling back to gastanks.watch",
                 ", ".join(missing))
        return None
    stamps = [r["reading_at"] for r in rows if r.get("reading_at")]
    if not stamps:
        log.info("tank grab: Pi data has no timestamps - falling back to gastanks.watch")
        return None
    reading_at = max(stamps)
    age = (fetched - reading_at).total_seconds() / 60
    if max_age > 0 and age > max_age:
        log.info("tank grab: Pi readings are stale (last update %s, %.0f min old) - "
                 "falling back to gastanks.watch", reading_at, age)
        return None
    log.info("tank grab: using Raspberry Pi %s (reading %s, %.0f min old)",
             base, reading_at, max(age, 0))
    return rows


def _tank_snapshot_summary(rows):
    return ", ".join("%s %s %sin" % (r["tank_number"], r.get("product") or "",
                                     _fmt_csv_num(r["height"])) for r in rows)


def grab_tanks(cfg, fetcher=None, sleep=None, now=None, pi_fetcher=None):
    """--grab-tanks: read the Pi (or gastanks.watch), write the day's snapshot CSV.

    Tries the store Pi first (short timeout). If that data is missing, incomplete,
    or older than tank_pi_max_age_minutes, falls back to gastanks.watch with the
    existing retry loop. Returns an exit code.
    """
    import time
    fetcher = fetcher or fetch_tank_page
    sleep = time.sleep if sleep is None else sleep
    url = (cfg.get("tank_url", "") or "").strip() or TANK_URL
    csv_dir = tank_csv_dir(cfg)
    if not csv_dir:
        log.error("tank grab: no tank_csv_dir (and no day_report_dir) in dayreport.ini")
        return 2
    attempts = max(1, cfg.getint("tank_grab_attempts", 5))
    wait = max(0, cfg.getint("tank_grab_retry_seconds", 60))
    max_age = cfg.getint("tank_max_age_minutes", 180)
    pi_base = tank_pi_base_url(cfg)
    log.info("==== tank grab start (Pi %s, fallback %s -> %s)",
             pi_base or "off", url, csv_dir)
    fetched = now() if now else dt.datetime.now().replace(microsecond=0)
    pi_rows = fetch_pi_tank_rows(cfg, fetched, fetcher=pi_fetcher)
    if pi_rows:
        reading_at = max(r["reading_at"] for r in pi_rows if r.get("reading_at"))
        path, wrote = write_tank_snapshot(csv_dir, pi_rows, reading_at)
        if wrote:
            log.info("tank snapshot written: %s (reading %s, business date %s, source Pi): %s",
                     path, reading_at, tank_business_date(reading_at),
                     _tank_snapshot_summary(pi_rows))
        log.info("==== tank grab end, exit code 0")
        return 0
    last_err = None
    for n in range(1, attempts + 1):
        try:
            fetched = now() if now else dt.datetime.now().replace(microsecond=0)
            rows = parse_tank_page(fetcher(url))
            stamps = [r["reading_at"] for r in rows if r.get("reading_at")]
            reading_at = max(stamps) if stamps else fetched
            if not stamps:
                log.info("tank page has no reading time - using fetch time %s", fetched)
            age = (fetched - reading_at).total_seconds() / 60
            if max_age > 0 and age > max_age:
                raise TankPageError("readings are stale: last update %s (%.0f min old)"
                                    % (reading_at, age))
            path, wrote = write_tank_snapshot(csv_dir, rows, reading_at)
            if wrote:
                log.info("tank snapshot written: %s (reading %s, business date %s, "
                         "source gastanks.watch): %s",
                         path, reading_at, tank_business_date(reading_at),
                         _tank_snapshot_summary(rows))
            log.info("==== tank grab end, exit code 0")
            return 0
        except Exception as e:      # network, HTTP, parse, disk - all retried
            last_err = e
            log.warning("tank grab attempt %d/%d failed: %s: %s", n, attempts,
                        type(e).__name__, e)
            if n < attempts:
                sleep(wait)
    if "stale" in str(last_err):
        hint = ("The site is up but its numbers stopped updating - check the Raspberry Pi "
                "(192.168.0.28) and the TLS-350 at the store.")
    else:
        hint = ("Check that the office PC can reach the store Pi at %s and that "
                "%s opens in a browser." % (pi_base or TANK_PI_URL, url))
    log.error("tank grab FAILED after %d attempts - no snapshot written. Last error: %s. %s",
              attempts, last_err, hint)
    log.info("==== tank grab end, exit code %d", EXIT_TANK_GRAB)
    return EXIT_TANK_GRAB


def warn_missing_tank_snapshot(cfg, today=None):
    """5:30 AM run: warn (only) if last night's closing snapshot is not there."""
    today = today or dt.date.today()
    want = today - dt.timedelta(days=1)
    idx = load_tank_csv_index(tank_csv_dir(cfg))
    if want not in idx:
        log.warning("no Tank Watcher snapshot for %s in %s - did the 11:55 PM "
                    "\"Day Report Tanks\" task run? Stick inches for that day stay blank.",
                    want.isoformat(), tank_csv_dir(cfg))
        return False
    return True


def process_date(cfg, d, files, args):
    day_name = "Day %d" % d.day
    for k in ("MSM", "MCM"):
        if not files[k]:
            raise DayReportError("%s: no %s file - day not processed" % (d, k))
        if len(files[k]) > 1:
            log.warning("%s: %d %s files, using the newest name %s", d, len(files[k]), k,
                        os.path.basename(sorted(files[k])[-1]))
    _, msm = parse_msm(read(sorted(files["MSM"])[-1]))
    _, mcm = parse_mcm(read(sorted(files["MCM"])[-1]))
    info = {"date": d, "status": "skipped",
            "gallons": round(sum(float(r[4]) for r in msm if r[0] in FUEL_KEYS), 1),
            "sales": round(sum(float(r[3]) for r in msm if r[0] == "salestotal"), 2)}
    if len(msm) > 119:
        log.warning("%s: %d MSM lines; Day-sheet formulas only look at L2:P120 - check the sheet",
                    d, len(msm))
    wbname = "day report %s %d.xlsm" % (MONTHS[d.month - 1], d.year)
    wbpath = os.path.join(cfg["day_report_dir"], wbname)
    if not os.path.exists(wbpath):
        if args.dry_run:
            log.info("[dry-run] would create %s from template", wbpath)
            return info
        shutil.copy2(cfg["template"], wbpath)
        log.info("created new month workbook %s from %s", wbpath, cfg["template"])
    try:
        carry_forward_opening(cfg, wbpath, d, args.dry_run)
    except Exception as e:  # never fail the day over this
        log.warning("%s: opening inventory carry-forward skipped (%s)", wbname, e)
    l1, q2 = day_state(wbpath, day_name)
    mmddyy = d.strftime("%m%d%y")
    same = None
    if l1 == "mscode" and q2 == mmddyy:
        have = day_msm_pairs(wbpath, day_name)[:len(msm)]
        # compare rounded: Excel/VFP stored values like 4039.0400000000004
        same = [(k, round(a, 4)) for k, a in have] == [(r[0], round(float(r[3]), 4)) for r in msm]
    if same and not args.force:
        log.info("%s: %s already filled for %s - skipped (use --force to redo)", wbname, day_name, mmddyy)
    elif same is False and not args.force:
        raise CollisionError("%s: %s says %s but holds different numbers (filled from a different "
                             "Passport file?). Not overwritten; check it and "
                             "rerun with --force if this file is right." % (wbname, day_name, mmddyy))
    else:
        if l1 == "mscode" and q2 != mmddyy:
            log.warning("%s: %s held data for %s; overwriting with %s", wbname, day_name, q2, mmddyy)
        if args.dry_run:
            log.info("[dry-run] would fill %s!%s (%d MSM rows, %d MCM rows)", wbname, day_name, len(msm), len(mcm))
        else:
            bdir = os.path.join(cfg["day_report_dir"], "backups")
            os.makedirs(bdir, exist_ok=True)
            stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
            shutil.copy2(wbpath, os.path.join(bdir, "%s.%s.bak" % (wbname, stamp)))
            tmp = write_workbook(wbpath, day_name, msm, mcm)
            try:
                os.replace(tmp, wbpath)
            except PermissionError:
                os.unlink(tmp)
                raise DayReportError("%s is locked (open in Excel on some PC?) - close it and rerun" % wbpath)
            log.info("%s: filled %s for %s (%d MSM rows, %d departments)", wbname, day_name, mmddyy,
                     len(msm), len(mcm))
            info["status"] = "filled"
    if not args.dry_run and os.path.exists(wbpath):
        try:
            for pr in check_tank_links(wbpath, d.day):
                log.warning("%s: tank sheet formula check: %s", wbname, pr)
        except Exception as e:  # never fail the run over the check
            log.warning("%s: tank sheet formula check skipped (%s)", wbname, e)
    dbfp = cfg.get("dailysales_dbf", "").strip()
    if dbfp and files["ISM"] and not args.dry_run:
        if os.path.exists(dbfp):
            _, items = parse_ism(read(sorted(files["ISM"])[-1]))
            n = append_dailysales(dbfp, items, args.force)
            log.info("dailysales.dbf: added %d item lines for %s", n, d)
        else:
            log.warning("dailysales.dbf not found at %s - ISM not stored", dbfp)
    if not args.dry_run:
        done = os.path.join(cfg["archive_dir"], "processed", d.isoformat())
        os.makedirs(done, exist_ok=True)
        for k in KINDS:
            for p in files[k]:
                shutil.move(p, os.path.join(done, os.path.basename(p)))
    return info


# --------------------------------------------------------------------------- sheet readers
def _typed(c, sst):
    t = c.get("t")
    v = cell_text(c, sst)
    if v is None:
        return None
    if t in ("s", "inlineStr", "str"):
        return v
    try:
        return num(float(v))
    except ValueError:
        return v


def _sheet_cells(z, part, sst):
    out = {}
    for c in etree.fromstring(z.read(part)).iter(N + "c"):
        col, r = split_ref(c.get("r"))
        f = c.find(N + "f")
        out[(r, col)] = ("F", f.text) if f is not None else ("V", _typed(c, sst))
    return out


def _program_cell(r, col):
    return (12 <= col <= 18 and r <= 100) or (20 <= col <= 32 and r <= 34)


def sheet_msm(path, day_name):
    with zipfile.ZipFile(path) as z:
        sp, sst = sheet_paths(z), shared_strings(z)
        cells = _sheet_cells(z, sp[day_name], sst)
    rows, r = [], 2
    while cells.get((r, 17), ("V", None))[1] not in (None, ""):     # every MSM row carries a date in Q
        g = lambda col: cells.get((r, col), ("V", None))[1]
        rows.append((g(12) or "", float(g(15) or 0), float(g(16) or 0)))
        r += 1
    return rows, cells.get((2, 17), ("V", None))[1]


# --------------------------------------------------------------------------- SMS alerts (Twilio)
ALERT_TYPES = ("error", "pos_unreachable", "collision", "missing_day", "already_running",
               "success", "missed_run")
KEYRING_SERVICE = "dayreport-twilio"
CRED_KEYS = (("account_sid", "TWILIO_ACCOUNT_SID"), ("auth_token", "TWILIO_AUTH_TOKEN"),
             ("from_number", "TWILIO_FROM"))
E164 = re.compile(r"^\+[1-9]\d{6,14}$")


def twilio_creds():
    """(sid, token, from) from Windows Credential Manager (keyring), else environment."""
    vals = {}
    try:
        import keyring
        for k, _ in CRED_KEYS:
            v = keyring.get_password(KEYRING_SERVICE, k)
            if v:
                vals[k] = v
    except Exception as e:                      # keyring missing / no backend
        log.debug("keyring unavailable: %s", e)
    for k, env in CRED_KEYS:
        vals.setdefault(k, os.environ.get(env, ""))
    if not all(vals.values()):
        return None
    return vals["account_sid"], vals["auth_token"], vals["from_number"]


def send_sms(to, body):
    """POST to Twilio Messages API. Returns True/False; never raises."""
    try:
        import requests
        creds = twilio_creds()
        if not creds:
            log.warning("SMS not sent to %s: Twilio credentials not set (run --setup-twilio)", to)
            return False
        sid, token, frm = creds
        base = os.environ.get("DAYREPORT_TWILIO_API", "https://api.twilio.com").rstrip("/")
        r = requests.post("%s/2010-04-01/Accounts/%s/Messages.json" % (base, sid),
                          data={"To": to, "From": frm, "Body": body[:160]},
                          auth=(sid, token), timeout=20)
        if r.status_code >= 300:
            log.warning("SMS to %s failed: HTTP %s %s", to, r.status_code, r.text[:200])
            return False
        log.info("SMS sent to %s: %s", to, body)
        return True
    except Exception as e:
        log.warning("SMS to %s failed: %s", to, e)
        return False


def load_recipients(cp):
    """[recipient.NAME] sections: number = +1..., alerts = error, collision, ..."""
    out = []
    for sec in cp.sections():
        if not sec.lower().startswith("recipient."):
            continue
        c = cp[sec]
        num_ = c.get("number", "").strip()
        types = {t.strip() for t in c.get("alerts", "").split(",") if t.strip()}
        if not E164.match(num_):
            log.warning("[%s] number %r is not E.164 (+15551234567) - skipped", sec, num_)
            continue
        bad = types - set(ALERT_TYPES)
        if bad:
            log.warning("[%s] unknown alert types %s ignored", sec, ",".join(sorted(bad)))
        if c.getboolean("enabled", True):
            out.append({"name": c.get("name", sec[10:]), "number": num_, "types": types & set(ALERT_TYPES)})
    return out


def in_quiet_hours(spec, now):
    """spec 'HH:MM-HH:MM' (may wrap midnight); empty = never quiet."""
    if not spec or not spec.strip():
        return False
    a, b = [dt.datetime.strptime(x.strip(), "%H:%M").time() for x in spec.split("-")]
    t = now.time()
    return (a <= t < b) if a < b else (t >= a or t < b)


def load_state(path):
    import json
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def save_state(path, st):
    import json
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(st, f, indent=1, default=str)
    os.replace(tmp, path)


class Alerter:
    """Routes alerts to recipients by type, with quiet hours (held, sent later) and de-dup."""

    def __init__(self, cp, state_path, now=None):
        a = cp["alerts"] if cp.has_section("alerts") else {}
        self.enabled = str(a.get("enabled", "yes")).lower() in ("yes", "true", "1", "on")
        self.quiet = a.get("quiet_hours", "")
        self.quiet_types = {t.strip() for t in a.get("quiet_hours_apply_to", "success,missing_day").split(",") if t.strip()}
        self.dedup_h = float(a.get("dedup_hours", "12") or 0)
        self.site = a.get("site_name", "SDS Office PC")
        self.recipients = load_recipients(cp) if self.enabled else []
        self.state_path = state_path
        self.now = now or dt.datetime.now()

    def alert(self, kind, text, key=None):
        """kind in ALERT_TYPES; text after 'Day Report KIND mm/dd: '. Never raises."""
        try:
            body = ("Day Report %s %s: %s" % (kind.upper().replace("_", " "), self.now.strftime("%m/%d"), text))[:160]
            targets = [r for r in self.recipients if kind in r["types"]]
            if not targets:
                log.info("alert %s (no recipients): %s", kind, body)
                return 0
            st = load_state(self.state_path)
            sent = st.setdefault("sms_sent", {})
            pend = st.setdefault("sms_pending", [])
            n = 0
            for r in targets:
                dk = "%s|%s|%s" % (kind, key or text, r["number"])
                last = sent.get(dk)
                if last and self.dedup_h and (self.now - dt.datetime.fromisoformat(last)).total_seconds() < self.dedup_h * 3600:
                    log.info("alert %s to %s suppressed (sent %s, dedup %sh)", kind, r["name"], last, self.dedup_h)
                    continue
                if kind in self.quiet_types and in_quiet_hours(self.quiet, self.now):
                    if not any(p["dk"] == dk for p in pend):
                        pend.append({"dk": dk, "to": r["number"], "body": body})
                    log.info("alert %s to %s held for quiet hours (%s)", kind, r["name"], self.quiet)
                    continue
                if send_sms(r["number"], body):
                    sent[dk] = self.now.isoformat()
                    n += 1
            save_state(self.state_path, st)
            return n
        except Exception as e:
            log.warning("alerting failed (%s): %s", kind, e)
            return 0

    def flush_pending(self):
        try:
            if in_quiet_hours(self.quiet, self.now):
                return 0
            st = load_state(self.state_path)
            pend = st.get("sms_pending", [])
            keep, n = [], 0
            for p in pend:
                if send_sms(p["to"], p["body"]):
                    st.setdefault("sms_sent", {})[p["dk"]] = self.now.isoformat()
                    n += 1
                else:
                    keep.append(p)
            st["sms_pending"] = keep
            save_state(self.state_path, st)
            return n
        except Exception as e:
            log.warning("flushing held alerts failed: %s", e)
            return 0


def state_path_for(cfg):
    return cfg.get("state_file") or os.path.join(os.path.dirname(cfg.get("log_file")), "dayreport_state.json")


def update_state(path, **kw):
    try:
        st = load_state(path)
        st.update(kw)
        save_state(path, st)
    except Exception as e:
        log.warning("could not write state file %s: %s", path, e)


def run_watchdog(cp, cfg, now=None):
    """Alert (missed_run) if today's main run did not finish successfully. Returns exit code."""
    now = now or dt.datetime.now()
    sp = state_path_for(cfg)
    st = load_state(sp)
    al = Alerter(cp, sp, now)
    al.flush_pending()
    today = now.date().isoformat()
    start, end, rc = st.get("last_start", ""), st.get("last_end", ""), st.get("last_exit")
    if not start.startswith(today):
        msg = "no run today (last run %s). Check Task Scheduler on %s." % (start[:16] or "never", al.site)
    elif not end.startswith(today) or end < start:
        msg = "today's run started %s but never finished. See log on %s." % (start[11:16], al.site)
    elif rc not in (0, None):
        msg = "today's run ended with exit code %s. See log on %s." % (rc, al.site)
    else:
        log.info("watchdog: today's run finished OK at %s", end[11:16])
        return 0
    log.error("watchdog: %s", msg)
    al.alert("missed_run", msg, key="missed_run|" + today)
    return 1


def setup_twilio():
    import getpass
    import keyring
    print("Stores Twilio credentials in Windows Credential Manager (service '%s')." % KEYRING_SERVICE)
    print("Run this logged on as the SAME Windows user the scheduled task runs as.")
    sid = input("Account SID (AC...): ").strip()
    tok = getpass.getpass("Auth Token (hidden): ").strip()
    frm = input("From number, E.164 (+1...): ").strip()
    if not (sid.startswith("AC") and tok and E164.match(frm)):
        print("Not saved: check the values.")
        return 2
    for (k, _), v in zip(CRED_KEYS, (sid, tok, frm)):
        keyring.set_password(KEYRING_SERVICE, k, v)
    print("Saved. Test with:  python dayreport.py --test-sms +1XXXXXXXXXX")
    return 0


def find_sumatra(cfg):
    """SumatraPDF portable: config 'sumatra_exe', else SumatraPDF.exe next to dayreport.py."""
    here = os.path.dirname(os.path.abspath(__file__))
    for c in (cfg.get("sumatra_exe", "").strip(), os.path.join(here, "SumatraPDF.exe"),
              os.path.join(here, "SumatraPDF-3.5.2-64.exe")):
        if c and os.path.isfile(c):
            return c
    return None


def print_pdfs(cfg, args, runner=None, today=None):
    """Print every *.pdf in the POS outbox to pdf_printer, archive a copy first, delete from
    the outbox only after the printer accepted it. Returns (printed, failed_names)."""
    import subprocess
    runner = runner or subprocess.run
    src = cfg["source_outbox"]
    printer = cfg.get("pdf_printer", "Lexmark in store").strip()
    timeout = cfg.getint("pdf_print_timeout", 120)
    try:
        names = sorted(n for n in os.listdir(src)
                       if n.lower().endswith(".pdf") and os.path.isfile(os.path.join(src, n)))
    except OSError as e:
        log.error("PDF print: cannot read POS outbox %s: %s", src, e)
        return 0, ["outbox unreadable"]
    if not names:
        log.info("PDF print: no PDFs in %s", src)
        return 0, []
    if args.dry_run:
        for n in names:
            log.info("PDF print (dry run): would print %s to '%s' then delete it", n, printer)
        return 0, []
    exe = find_sumatra(cfg)
    if not exe:
        log.error("PDF print: SumatraPDF.exe not found next to dayreport.py - %d PDF(s) left in outbox", len(names))
        return 0, list(names)
    arch = os.path.join(cfg["archive_dir"], "printed", (today or dt.date.today()).isoformat())
    os.makedirs(arch, exist_ok=True)
    printed, failed = 0, []
    for n in names:
        p = os.path.join(src, n)
        try:
            shutil.copy2(p, os.path.join(arch, n))
            r = runner([exe, "-print-to", printer, "-silent", "-exit-when-done", p],
                       timeout=timeout, capture_output=True)
            if r.returncode != 0:
                raise RuntimeError("SumatraPDF return code %s" % r.returncode)
        except Exception as e:
            log.error("PDF print FAILED for %s (left in outbox): %s", n, e)
            failed.append(n)
            continue
        try:
            os.remove(p)
        except OSError as e:
            log.warning("PDF %s printed but could not be deleted from outbox: %s", n, e)
        log.info("PDF printed to '%s' and removed from outbox: %s", printer, n)
        printed += 1
    log.info("PDF print: %d printed, %d failed", printed, len(failed))
    return printed, failed


def main(argv=None):
    ap = argparse.ArgumentParser(description="Fill the monthly Day Report workbook from Passport XML.")
    ap.add_argument("--config", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "dayreport.ini"))
    ap.add_argument("--force", action="store_true", help="refill days that are already filled")
    ap.add_argument("--no-fetch", action="store_true", help="skip copying from the POS; use local inbox only")
    ap.add_argument("--dry-run", action="store_true", help="log what would happen, change nothing")
    ap.add_argument("--watchdog", action="store_true", help="only check that today's run finished OK; text if not")
    ap.add_argument("--setup-twilio", action="store_true", help="store Twilio credentials in Windows Credential Manager")
    ap.add_argument("--grab-tanks", action="store_true",
                    help="only save tonight's tank readings from the store Pi (or gastanks.watch) to tank_csv_dir")
    ap.add_argument("--test-sms", metavar="NUMBER", help="send one test text to NUMBER (E.164, e.g. +17575551234)")
    args = ap.parse_args(argv)
    if args.setup_twilio:
        return setup_twilio()

    cp = configparser.ConfigParser(interpolation=None)
    if not cp.read(args.config, encoding="utf-8"):
        print("config file not found: %s" % args.config, file=sys.stderr)
        return 2
    cfg = cp["dayreport"]
    logf = cfg.get("log_file")
    os.makedirs(os.path.dirname(logf), exist_ok=True)
    fmt = logging.Formatter("%(asctime)s %(levelname)-7s %(message)s", "%Y-%m-%d %H:%M:%S")
    for h in (logging.FileHandler(logf, encoding="utf-8"), logging.StreamHandler(sys.stdout)):
        h.setFormatter(fmt)
        log.addHandler(h)
    log.setLevel(logging.INFO)
    if args.test_sms:
        ok = send_sms(args.test_sms, "Day Report test message from %s. Alerts are working." %
                      (cp["alerts"].get("site_name", "SDS Office PC") if cp.has_section("alerts") else "SDS Office PC"))
        return 0 if ok else 1
    if args.watchdog:
        return run_watchdog(cp, cfg)
    if args.grab_tanks:
        return grab_tanks(cfg)
    statep = state_path_for(cfg)
    alerter = Alerter(cp, statep)
    alerter.flush_pending()
    log.info("==== dayreport start (force=%s, no_fetch=%s, dry_run=%s)", args.force, args.no_fetch, args.dry_run)

    lock = os.path.join(cfg["local_inbox"], ".dayreport.lock")
    os.makedirs(cfg["local_inbox"], exist_ok=True)
    try:
        fd = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        os.write(fd, str(os.getpid()).encode())
        os.close(fd)
    except FileExistsError:
        age = (dt.datetime.now().timestamp() - os.path.getmtime(lock)) / 3600
        if age < 2:
            log.error("another run seems active (%s, %.1f h old) - exiting", lock, age)
            alerter.alert("already_running", "another run was still active (%.1fh). Check %s." % (age, alerter.site),
                          key="already_running")
            return 4
        log.warning("removing stale lock %s (%.1f h old)", lock, age)
        os.remove(lock)
        return main(argv)
    rc = 0
    results, collisions, failures, pos_down = [], [], [], False
    if not args.dry_run:
        update_state(statep, last_start=dt.datetime.now().isoformat(timespec="seconds"))
    try:
        for k in ("day_report_dir", "template", "local_inbox", "archive_dir"):
            if not cfg.get(k):
                raise DayReportError("config: '%s' missing" % k)
        if not args.no_fetch and not fetch(cfg, args):
            rc = 3
            pos_down = True
        by_date, bad = collect(cfg["local_inbox"])
        if bad:
            rc = rc or 1
            failures.append("%d unreadable file(s)" % len(bad))
        if not by_date:
            log.info("no Passport files to process")
        for d in sorted(by_date):
            try:
                results.append(process_date(cfg, d, by_date[d], args))
            except CollisionError as e:
                rc = 1
                collisions.append(d)
                log.error("%s FAILED: %s", d, e)
            except DayReportError as e:
                rc = 1
                failures.append("%s %s" % (d.strftime("%m/%d"), str(e)[:60]))
                log.error("%s FAILED: %s", d, e)
            except Exception as e:
                rc = 1
                failures.append("%s %s" % (d.strftime("%m/%d"), type(e).__name__))
                log.exception("%s FAILED: %s", d, e)
    except Exception as e:
        rc = 1
        failures.append(str(e)[:80])
        log.exception("FAILED: %s", e)
    finally:
        try:
            os.remove(lock)
        except OSError:
            pass
    # Tank Watcher stick inches (CSV saved nightly by --grab-tanks). Runs even when there were no
    # Passport files, so a late CSV still fills blank D cells and unblocks E7 carry.
    if cfg.getboolean("tank_fill", False):
        try:
            warn_missing_tank_snapshot(cfg)
            fill_tank_sticks(cfg, dates=None, dry_run=args.dry_run)
            # Re-try opening inventory now that last month's closing sticks may be in.
            today = dt.date.today()
            for y, m in ((today.year, today.month),
                         ((today.replace(day=1) - dt.timedelta(days=1)).year,
                          (today.replace(day=1) - dt.timedelta(days=1)).month)):
                wbname = "day report %s %d.xlsm" % (MONTHS[m - 1], y)
                wbpath = os.path.join(cfg["day_report_dir"], wbname)
                if os.path.exists(wbpath):
                    try:
                        carry_forward_opening(cfg, wbpath, dt.date(y, m, 1), args.dry_run)
                    except Exception as e:
                        log.warning("%s: opening inventory carry-forward skipped (%s)", wbname, e)
        except Exception as e:
            log.warning("tank stick fill skipped (%s)", e)

    # Separate step: PDF reports from staff -> store printer. Independent of the XML work.
    if not args.no_fetch and cfg.getboolean("print_pdfs", False):
        try:
            _, pfail = print_pdfs(cfg, args)
        except Exception as e:
            log.exception("PDF print step crashed: %s", e)
            pfail = ["crash"]
        if pfail:
            rc = rc or 5
            failures.append("PDF print failed: %d file(s)" % len(pfail))
    if not args.dry_run:
        try:
            send_run_alerts(alerter, cfg, statep, rc, results, collisions, failures, pos_down, by_date_ok=True)
        except Exception as e:                      # alerting must never change the run result
            log.warning("alerting failed: %s", e)
        update_state(statep, last_end=dt.datetime.now().isoformat(timespec="seconds"), last_exit=rc)
    log.info("==== dayreport end, exit code %d", rc)
    return rc


def send_run_alerts(al, cfg, statep, rc, results, collisions, failures, pos_down, by_date_ok=True):
    st = load_state(statep)
    today = al.now.date()
    if pos_down:
        host = re.match(r"^\\\\([^\\]+)", cfg.get("source_outbox", "")) or re.match(r"^//([^/]+)", cfg.get("source_outbox", ""))
        al.alert("pos_unreachable", "POS unreachable (%s). See log on %s." % (host.group(1) if host else "POS", al.site),
                 key="pos_unreachable|" + today.isoformat())
    if collisions:
        al.alert("collision", "Day sheet already holds different data for %s. Not overwritten. See log on %s." %
                 (",".join(d.strftime("%m/%d") for d in collisions), al.site),
                 key="collision|" + ",".join(map(str, collisions)))
    if failures:
        al.alert("error", "run failed: %s. See log on %s." % ("; ".join(failures)[:70], al.site),
                 key="error|" + today.isoformat())
    # missing days: gap since last known date, or nothing new by yesterday
    done = sorted({r["date"] for r in results} | set(collisions))
    last_known = st.get("last_data_date")
    known = [dt.date.fromisoformat(last_known)] if last_known else []
    seq = sorted(set(known + done))
    gaps = []
    for a, b in zip(seq, seq[1:]):
        x = a + dt.timedelta(days=1)
        while x < b:
            gaps.append(x)
            x += dt.timedelta(days=1)
    newest = seq[-1] if seq else None
    expected = today - dt.timedelta(days=1)     # 05:30 run: last night's report has EndDate = yesterday
    if newest and newest < expected:
        x = newest + dt.timedelta(days=1)
        while x <= expected:
            gaps.append(x)
            x += dt.timedelta(days=1)
    if not results and not collisions and not pos_down:
        al.alert("missing_day", "no Passport files found for %s. See log on %s." % (expected.strftime("%m/%d"), al.site),
                 key="missing_day|none|" + today.isoformat())
    elif gaps:
        al.alert("missing_day", "no data for %s. See log on %s." % (",".join(g.strftime("%m/%d") for g in gaps[:8]), al.site),
                 key="missing_day|" + ",".join(map(str, gaps)))
    if newest and (not last_known or newest.isoformat() > last_known):
        update_state(statep, last_data_date=newest.isoformat())
    if rc == 0 and results:
        filled = [r for r in results if r["status"] == "filled"]
        if filled:
            txt = "OK. " + " ".join("%s %.0fgal $%.0f" % (r["date"].strftime("%m/%d"), r["gallons"], r["sales"])
                                    for r in filled[-3:])
            al.alert("success", txt, key="success|" + ",".join(str(r["date"]) for r in filled))


if __name__ == "__main__":
    sys.exit(main())
