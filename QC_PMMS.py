# -*- coding: utf-8 -*-
"""QC_PMMS - Namaa GIS QC System (standalone, no ArcGIS) | Read-only | all rules merged
Created by Ahmed Yasser
Deps: pyogrio, shapely>=2, pyproj, numpy, openpyxl
"""
import re, sys, json, math, logging, threading, queue
import uuid  # noqa: F401  (pyogrio's compiled code needs it; keeps PyInstaller from dropping it)
from pathlib import Path
from datetime import datetime, date
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import numpy as np
import shapely
from shapely import STRtree, Point
from pyproj import CRS
from pyogrio import list_layers
from pyogrio.raw import read as ogr_read
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

__author__ = "Ahmed Yasser"
CREDIT = "Created by Ahmed Yasser"
EXPIRY_DATE = date(2027, 2, 1)   # the tool will not open on or after this date (1 Feb 2027)

# ------------------------------------------------------------------ config
HERE = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
CFG_FILE = HERE / "config.json"
logging.basicConfig(filename=str(HERE / "qc.log"), level=logging.INFO,
                    format="%(asctime)s %(levelname)s %(message)s", encoding="utf-8")
logging.info("QC_PMMS started - %s", CREDIT)

DEFAULT = {
    "options": {"section_no_includes_serial": True},
    "layer_prefixes": ["NAMAA_GIS_SDE_"],
    "layer_map": {},   # e.g. {"LANE": "Namaa_GIS_SDE_LANE_N", "SECTION": "Section_N"}
    "field_aliases": {
        "REGION": ["REGION_NO"], "MAIN_STREETS": ["MAIN_NO"], "SECONDARY": ["SECONDARY_NO"],
        "INTERSECTION": ["INTERSECTION_NO"], "INTERSEC_SAMPLE": ["INTSAMPLE_NO", "SAMPLE_NO"],
        "LANE": ["LANE_NO"], "SECTION": ["SECTION_NO"]},
    "fields": {
        "section_id": "SECTION_NO", "region_id": "REGION_NO", "main_id": "MAIN_NO",
        "sec_type": "SEC_TYPE", "serial": "SERIAL_NO", "munic_code": "MUNIC_CODE",
        "from_street": "FROM_STREET", "street_names": ["FROM_STREET_NAME", "TO_STREET_NAME"],
        "lane_no": "LANE_NO", "lane_type": "LANE_TYPE", "lane_section_id": "SECTION_NO",
        "intersection_no": "INTERSECTION_NO", "main_no1": "MAIN_NO1", "main_no2": "MAIN_NO2",
        "district_no": "DISTRICT_NO", "intsample_no": "INTSAMPLE_NO", "intsample_code": "INTSAMPLE_CODE",
        "branch_code": "BRANCH_CODE", "dist_code": "DIST_CODE", "region_code": "REGION_CODE"},
    "thresholds": {"section_max_km": 1.3, "lane_ratio_min": 0.70, "lane_ratio_max": 1.30,
                   "lane_longer_m": 30.0, "lane_order_diff_m": 25.0, "intersection_search_m": 300.0,
                   "from_street_tol_m": 10.0, "region_side_offset_m": 5.0, "sliver_area_m2": 1.0},
    "arabic": {"alef": True, "yaa": True, "taa_marbuta_to_haa": False},
    "sec_types": ["B", "T", "N", "X", "R", "S"],
    "lane_types_service": ["S1", "S2", "S3", "S4", "S5", "PS", "PSL"],
    "lane_types_other": ["L", "R", "M1", "M2", "M3", "M4", "M5", "P1", "P2", "LP"],
    "lane_order_rules": [["L", "R", "M1", "M2", "M3", "M4", "M5"], ["P1", "P2"], ["L", "LP"],
                         ["S1", "S2", "S3", "S4", "S5"], ["S1", "PS"], ["S1", "PSL"]],
    "street_name_map": {
        "بداية الكوبرى": "بداية كوبرى", "بداية كوبري": "بداية كوبرى",
        "منتصف الكوبرى": "منتصف كوبرى", "منتصف كوبري": "منتصف كوبرى",
        "نهاية الكوبرى": "نهاية كوبرى", "نهاية كوبري": "نهاية كوبرى"},
    "rules": {
        "whitespace": True, "arabic_review": True, "zm": True, "duplicates": True,
        "section_id": True, "section_dup": True, "section_base_serial": True, "section_serial": True,
        "section_region": True, "section_munic": True, "section_type": True, "section_length": True,
        "section_intersection": True, "section_street_names": True, "section_direction": True,
        "lane_id": True, "lane_type": True, "lane_munic": True, "lane_section_ref": True,
        "lane_longer_than_section": True, "lane_length_ratio": False, "lane_type_vs_sec_type": True,
        "lane_length_order": True, "lane_direction": True,
        "intersection_id": True, "intersection_munic": True, "intersection_mains": True,
        "sample_id": True, "region_id": True, "region_district": True, "region_admin": True,
        "region_sliver": True},
}


def load_config():
    if not CFG_FILE.exists():
        CFG_FILE.write_text(json.dumps(DEFAULT, ensure_ascii=False, indent=2), encoding="utf-8")
        return json.loads(json.dumps(DEFAULT))
    d = json.loads(CFG_FILE.read_text(encoding="utf-8"))
    for k, v in DEFAULT.items():
        if k not in d:
            d[k] = v
        elif isinstance(v, dict):
            for a, b in v.items():
                d[k].setdefault(a, b)
    return d


CFG = load_config()


def enabled(k):
    return bool(CFG["rules"].get(k, False))


def nk(s):
    return re.sub(r"[^A-Z0-9]+", "", str(s).upper())


def sv(v):
    return "" if v is None else str(v).strip()


def serial2(v):
    s = sv(v)
    return s.zfill(2) if s.isdigit() else s


def last3(v):
    return sv(v)[-3:].zfill(3)


# ------------------------------------------------------------------ data access
def gdb_list(src):
    p = Path(src)
    if p.is_dir() and p.suffix.lower() == ".gdb":
        return [str(p)]
    if p.is_dir():
        return sorted(str(x) for x in p.rglob("*.gdb") if x.is_dir())
    return []


def feature_classes(gdb):
    return [str(n) for n, g in list_layers(gdb) if g is not None]


def canon(name):
    u = name.upper()
    for p in CFG["layer_prefixes"]:
        if u.startswith(p.upper()):
            return u[len(p):]
    return u


def helper_map(names):
    m = {}
    for n in sorted(names, key=lambda x: (canon(x) != x.upper(), x)):
        m.setdefault(canon(n), n)
    for k, v in CFG["layer_map"].items():
        hit = next((n for n in names if n.upper() == str(v).upper()), None)
        if hit:
            m[k.upper()] = hit
    return m


def base_of(name, hmap):
    for k, v in CFG["layer_map"].items():
        if str(v).upper() == name.upper():
            return k.upper()
    return canon(name)


def val(v):
    if v is None:
        return None
    if isinstance(v, np.generic):
        v = v.item()
    if isinstance(v, float):
        if v != v:
            return None
        if v.is_integer():
            return int(v)
    return v


class Layer:
    def __init__(self, gdb, name):
        self.gdb, self.name = gdb, name
        meta, fids, geom, data = ogr_read(gdb, layer=name, return_fids=True)
        n = len(geom) if geom is not None else (len(data[0]) if data else 0)
        self.fids = fids if fids is not None else np.arange(1, n + 1)
        self.cols = dict(zip(meta["fields"], data))
        self.fmap = {nk(f): f for f in meta["fields"]}
        self.string_fields = [f for f, d in zip(meta["fields"], meta["dtypes"]) if str(d) == "object"]
        self.geoms = shapely.from_wkb(geom) if geom is not None else None
        self.gtype = str(meta.get("geometry_type") or "")
        try:
            self.crs = CRS.from_user_input(meta["crs"]) if meta.get("crs") else None
        except Exception:
            self.crs = None

    def field(self, name):
        for n in (name if isinstance(name, (list, tuple)) else [name]):
            f = self.fmap.get(nk(n))
            if f:
                return f
        return None

    def _arrs(self, fields):
        return [self.cols[f] if f else None for f in fields]

    def rows(self, fields):
        arrs = self._arrs(fields)
        for i, fid in enumerate(self.fids):
            yield (int(fid), *[val(a[i]) if a is not None else None for a in arrs])

    def geom_rows(self, fields=()):
        arrs = self._arrs(fields)
        for i, fid in enumerate(self.fids):
            g = self.geoms[i] if self.geoms is not None else None
            yield (int(fid), g, *[val(a[i]) if a is not None else None for a in arrs])


def unit_factor(L):
    if L.crs is not None and L.crs.is_projected:
        return L.crs.axis_info[0].unit_conversion_factor or 1.0
    return None


class Ctx:
    def __init__(self, gdb, layer, base, helpers, rows):
        self.gdb, self.L, self.base, self._h, self.R = gdb, layer, base, helpers, rows
        self._cache, self.flips = {}, {}

    @property
    def name(self):
        return self.L.name

    def helper(self, key):
        if key not in self._h:
            return None
        if key not in self._cache:
            self._cache[key] = self.L if self._h[key] == self.L.name else Layer(self.gdb, self._h[key])
        return self._cache[key]

    def add(self, rid, rname, status, severity="INFO", oid="", field="", value="", msg=""):
        self.R.append({
            "RunTime": datetime.now().isoformat(timespec="seconds"), "SourceGDB": self.gdb,
            "Layer": self.name, "RuleID": rid, "RuleName": rname, "Status": status,
            "Severity": severity, "FeatureID": str(oid) if oid != "" else "", "Field": field,
            "Value": "" if value is None else str(value), "Message": msg})

    def missing(self, rid, rname, msg):
        self.add(rid, rname, "DEPENDENCY_MISSING", "WARNING", msg=msg)

    def need_helper(self, key, rid, rname):
        h = self.helper(key)
        if h is None:
            self.missing(rid, rname, f"{key} helper layer missing")
        return h

    def compat(self, other, rid, rname):
        a, b = self.L.crs, other.crs
        if a is not None and b is not None and not a.equals(b):
            self.add(rid, rname, "NOT_RUN", "WARNING", msg=f"CRS differs between {self.name} and {other.name}")
            return False
        return True


def fld(c, L, keys):
    """Resolve config field keys in layer L. Returns (fields or None, missing names)."""
    fs = [L.field(CFG["fields"][k]) for k in keys]
    miss = [str(CFG["fields"][k]) for k, f in zip(keys, fs) if not f]
    return (None if miss else fs), miss


def norm_ar(v):
    s = str(v).strip()
    a = CFG["arabic"]
    if a["alef"]:
        s = s.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    if a["yaa"]:
        s = s.replace("ى", "ي")
    if a["taa_marbuta_to_haa"]:
        s = s.replace("ة", "ه")
    return s


# ------------------------------------------------------------------ geometry helpers
def as_line(g):
    if g is None or g.is_empty:
        return None
    if g.geom_type == "LineString":
        return g
    if g.geom_type == "MultiLineString":
        m = shapely.line_merge(g)
        return m if m.geom_type == "LineString" else max(m.geoms, key=lambda x: x.length)
    return None


def local_dir(g):
    """(ax, ay, bx, by): same idea as positionAlongLine(0.49)->(0.5)."""
    ln = as_line(g)
    if ln is None or ln.length == 0:
        return None
    a, b = ln.interpolate(0.49, normalized=True), ln.interpolate(0.5, normalized=True)
    if (a.x, a.y) == (b.x, b.y):
        a, b = Point(ln.coords[0][:2]), Point(ln.coords[-1][:2])
    return a.x, a.y, b.x, b.y


def side_of(d, p, tol=0.001):
    ax, ay, bx, by = d
    cross = (bx - ax) * (p.y - ay) - (by - ay) * (p.x - ax)
    return "ON_LINE" if abs(cross) < tol else ("LEFT" if cross > 0 else "RIGHT")


def mid_point(g):
    ln = as_line(g)
    if ln is not None:
        return ln.interpolate(0.5, normalized=True)
    return g.centroid if g is not None and not g.is_empty else None


# ------------------------------------------------------------------ generic rules
def run_common(c):
    L = c.L
    sf = L.string_fields
    if enabled("whitespace"):
        found = []
        for fid, *vals in L.rows(sf):
            for f, v in zip(sf, vals):
                if isinstance(v, str) and (v != v.strip() or "  " in v or "\t" in v):
                    found.append((fid, f, v))
        for o, f, v in found:
            c.add("ALL_WS_01", "Unwanted whitespace", "FAIL", "WARNING", o, f, v, "Trim or repeated spaces")
        if not found:
            c.add("ALL_WS_01", "Unwanted whitespace", "PASS")
    if enabled("arabic_review"):
        found = []
        for fid, *vals in L.rows(sf):
            for f, v in zip(sf, vals):
                if isinstance(v, str) and v and norm_ar(v) != v.strip():
                    found.append((fid, f, v))
        for o, f, v in found:
            c.add("ALL_AR_01", "Arabic normalization", "REVIEW", "INFO", o, f, v,
                  "Suggested normalization only; no edit applied")
        if not found:
            c.add("ALL_AR_01", "Arabic normalization", "PASS")
    if enabled("zm"):
        suffix = L.gtype.split()[-1] if " " in L.gtype else ""
        z, m = "Z" in suffix, "M" in suffix
        c.add("ALL_GEO_01", "Z/M geometry", "REVIEW" if z or m else "PASS", "WARNING" if z or m else "INFO",
              field="Geometry", value=f"hasZ={z}; hasM={m}", msg="Confirm whether Z/M is intended" if z or m else "")


def duplicates(c):
    L = c.L
    opts = CFG["field_aliases"].get(c.base, [])
    f = L.field(opts)
    if not f:
        c.missing("DUP_01", "Duplicate identifier", f"Identifier field not found. Config candidates: {opts}")
        return
    seen = {}
    for o, v in L.rows([f]):
        if sv(v):
            seen.setdefault(sv(v), []).append(o)
    dup = {k: v for k, v in seen.items() if len(v) > 1}
    for v, ids in dup.items():
        for o in ids:
            c.add("DUP_01", "Duplicate identifier", "FAIL", "ERROR", o, f, v, f"Occurs {len(ids)} times")
    if not dup:
        c.add("DUP_01", "Duplicate identifier", "PASS")


def concat_rule(c, rid, rname, target, parts, pads=()):
    """target == concat(parts); pads = config keys padded to 2 digits."""
    keys = [target] + parts
    fs, miss = fld(c, c.L, keys)
    if fs is None:
        c.missing(rid, rname, f"Missing fields: {miss}")
        return
    bad = 0
    for o, tgt, *pv in c.L.rows(fs):
        exp = "".join(serial2(v) if k in pads else sv(v) for k, v in zip(parts, pv))
        if sv(tgt) != exp:
            bad += 1
            c.add(rid, rname, "FAIL", "ERROR", o, fs[0], tgt, f"Expected {exp}")
    if not bad:
        c.add(rid, rname, "PASS")


def left_rule(c, rid, rname, a_key, b_key, n):
    fs, miss = fld(c, c.L, [a_key, b_key])
    if fs is None:
        c.missing(rid, rname, f"Missing fields: {miss}")
        return
    bad = 0
    for o, a, b in c.L.rows(fs):
        if sv(a) != sv(b)[:n]:
            bad += 1
            c.add(rid, rname, "FAIL", "ERROR", o, fs[0], a, f"{fs[0]} must equal first {n} chars of {fs[1]} ({sv(b)[:n]})")
    if not bad:
        c.add(rid, rname, "PASS")


# ------------------------------------------------------------------ SECTION rules
def section_id(c):
    inc = CFG["options"]["section_no_includes_serial"]
    parts = ["region_id", "main_id", "sec_type"] + (["serial"] if inc else [])
    concat_rule(c, "SEC_ID_01", "SECTION_NO = REGION+MAIN+TYPE" + ("+SERIAL" if inc else ""),
                "section_id", parts, pads=("serial",))


def full_sections(c):
    inc = CFG["options"]["section_no_includes_serial"]
    fs, miss = fld(c, c.L, ["section_id"] + ([] if inc else ["serial"]))
    if fs is None:
        return None, miss
    out = []
    for o, sid, *sr in c.L.rows(fs):
        if sv(sid):
            out.append((o, sv(sid) + ("" if inc else serial2(sr[0]))))
    return out, []


def section_dup(c):
    full, miss = full_sections(c)
    if full is None:
        c.missing("SEC_DUP_01", "Duplicate full SECTION", f"Missing fields: {miss}")
        return
    seen = {}
    for o, f in full:
        seen.setdefault(f, []).append(o)
    bad = 0
    for f, ids in seen.items():
        if len(ids) > 1:
            for o in ids:
                bad += 1
                c.add("SEC_DUP_01", "Duplicate full SECTION", "FAIL", "ERROR", o, "SECTION_NO", f, f"Occurs {len(ids)} times")
    if not bad:
        c.add("SEC_DUP_01", "Duplicate full SECTION", "PASS")


def section_base_serial(c):
    f = c.L.field(CFG["fields"]["section_id"])
    if not f:
        c.missing("SEC_BASE_01", "Base ID and Base+Serial coexist", "SECTION_NO field missing")
        return
    vals = [(o, sv(v)) for o, v in c.L.rows([f]) if sv(v)]
    no_serial = {v for _, v in vals if not re.search(r"\d{2}$", v)}
    with_serial = {v[:-2] for _, v in vals if re.search(r"\d{2}$", v)}
    err = no_serial & with_serial
    for o, v in vals:
        if v in err:
            c.add("SEC_BASE_01", "Base ID and Base+Serial coexist", "FAIL", "ERROR", o, f, v,
                  "Base exists with serial sections")
    if not err:
        c.add("SEC_BASE_01", "Base ID and Base+Serial coexist", "PASS")


def section_serial(c):
    fs, miss = fld(c, c.L, ["region_id", "main_id", "sec_type", "serial"])
    if fs is None:
        c.missing("SEC_SEQ_01", "Serial sequence 01..n", f"Missing fields: {miss}")
        return
    groups = {}
    for o, reg, main, typ, s in c.L.rows(fs):
        groups.setdefault((sv(reg), sv(main), sv(typ)), []).append((o, serial2(s)))
    bad = 0
    for k, members in groups.items():
        n = len(members)
        expected = {f"{i:02d}" for i in range(1, n + 1)}
        found = {s for _, s in members if s}
        if found != expected or len(found) != n:
            bad += 1
            for o, s in members:
                c.add("SEC_SEQ_01", "Serial sequence 01..n", "FAIL", "ERROR", o, fs[3], s,
                      f"Group {'|'.join(k)}: expected {sorted(expected)}, found {sorted(found)}")
    if not bad:
        c.add("SEC_SEQ_01", "Serial sequence 01..n", "PASS")


def section_region(c):
    R_ = c.need_helper("REGION", "SEC_REF_01", "SECTION to REGION")
    if R_ is None:
        return
    k = CFG["fields"]["region_id"]
    a, b = c.L.field(k), R_.field(k)
    if not a or not b:
        c.missing("SEC_REF_01", "SECTION to REGION", f"Field {k} missing in SECTION or REGION")
        return
    valid = {sv(x) for _, x in R_.rows([b]) if sv(x)}
    bad = 0
    for o, v in c.L.rows([a]):
        if sv(v) and sv(v) not in valid:
            bad += 1
            c.add("SEC_REF_01", "SECTION to REGION", "FAIL", "ERROR", o, a, v, "REGION_NO not found in REGION layer")
    if not bad:
        c.add("SEC_REF_01", "SECTION to REGION", "PASS")


def section_munic(c):
    left_rule(c, "SEC_MUN_01", "MUNIC_CODE = left(SECTION_NO,2)", "munic_code", "section_id", 2)


def section_type(c):
    f = c.L.field(CFG["fields"]["sec_type"])
    if not f:
        c.missing("SEC_TYPE_01", "SEC_TYPE domain", "SEC_TYPE field missing")
        return
    allowed, bad = set(CFG["sec_types"]), 0
    for o, v in c.L.rows([f]):
        if sv(v) and sv(v).upper() not in allowed:
            bad += 1
            c.add("SEC_TYPE_01", "SEC_TYPE domain", "FAIL", "ERROR", o, f, v, f"Allowed: {sorted(allowed)}")
    if not bad:
        c.add("SEC_TYPE_01", "SEC_TYPE domain", "PASS")


def section_length(c):
    fac = unit_factor(c.L)
    if fac is None:
        c.add("SEC_LEN_01", "Maximum SECTION length", "NOT_RUN", "WARNING", msg="Projected CRS required; threshold is metres")
        return
    limit, bad = CFG["thresholds"]["section_max_km"] * 1000, 0
    for o, g in c.L.geom_rows():
        if g is not None and g.length * fac > limit:
            bad += 1
            c.add("SEC_LEN_01", "Maximum SECTION length", "FAIL", "WARNING", o, "Shape_Length",
                  round(g.length * fac, 2), f"Exceeds {limit} metres")
    if not bad:
        c.add("SEC_LEN_01", "Maximum SECTION length", "PASS")


def section_intersection(c):
    I = c.need_helper("INTERSECTION", "SEC_GEO_01", "SECTION crosses INTERSECTION")
    if I is None or not c.compat(I, "SEC_GEO_01", "SECTION crosses INTERSECTION"):
        return
    polys = [g for g in I.geoms if g is not None]
    idx = [i for i, g in enumerate(c.L.geoms) if g is not None]
    bad = 0
    if polys and idx:
        hit = STRtree(polys).query(np.array([c.L.geoms[i] for i in idx], dtype=object), predicate="crosses")
        for li in sorted(set(hit[0].tolist())):
            bad += 1
            c.add("SEC_GEO_01", "SECTION crosses INTERSECTION", "FAIL", "ERROR", int(c.L.fids[idx[li]]), "Shape", "",
                  "Crosses intersection polygon; check exceptions")
    if not bad:
        c.add("SEC_GEO_01", "SECTION crosses INTERSECTION", "PASS")


def section_street_names(c):
    names = CFG["fields"]["street_names"]
    fs = [f for f in (c.L.field(n) for n in names) if f]
    if not fs:
        c.missing("SEC_NAME_01", "Street name normalization", f"Fields not found: {names}")
        return
    mp, bad = {sv(k): v for k, v in CFG["street_name_map"].items()}, 0
    for o, *vals in c.L.rows(fs):
        for f, v in zip(fs, vals):
            if isinstance(v, str) and sv(v) in mp and mp[sv(v)] != v:
                bad += 1
                c.add("SEC_NAME_01", "Street name normalization", "REVIEW", "INFO", o, f, v, f"Suggested: {mp[sv(v)]}")
    if not bad:
        c.add("SEC_NAME_01", "Street name normalization", "PASS")


def section_direction(c):
    L, fm, th = c.L, CFG["fields"], CFG["thresholds"]
    fac = unit_factor(L)
    if fac is None:
        c.add("SEC_DIR_01", "SECTION direction", "NOT_RUN", "WARNING", msg="Projected CRS required (metres)")
        return
    off, tol = th["region_side_offset_m"] / fac, th["from_street_tol_m"] / fac
    reg_f, fs_f = L.field(fm["region_id"]), L.field(fm["from_street"])
    sec_f, typ_f = L.field(fm["section_id"]), L.field(fm["sec_type"])

    # --- region side index
    R_ = c.helper("REGION")
    rtree = None
    if R_ is None or not reg_f or not R_.field(fm["region_id"]):
        c.missing("SEC_DIR_01", "Region on right side", "REGION layer or REGION_NO field missing")
    elif c.compat(R_, "SEC_DIR_01", "Region on right side"):
        rg = [(g, sv(n)) for _, g, n in R_.geom_rows([R_.field(fm["region_id"])]) if g is not None]
        if rg:
            rtree, rnos = STRtree([g for g, _ in rg]), [n for _, n in rg]
    # --- main street index
    M = c.helper("MAIN_STREETS")
    mtree = None
    if M is None or not fs_f or not M.field(fm["main_id"]):
        c.missing("SEC_DIR_02", "FROM_STREET at start point", "MAIN_STREETS layer or FROM_STREET/MAIN_NO field missing")
    elif c.compat(M, "SEC_DIR_02", "FROM_STREET at start point"):
        mg = [(g, last3(n)) for _, g, n in M.geom_rows([M.field(fm["main_id"])]) if g is not None]
        if mg:
            mtree, mcodes = STRtree([g for g, _ in mg]), [n for _, n in mg]
    # --- lane index
    LN = c.helper("LANE")
    lanes = {}
    if LN is None or not (sec_f and typ_f) or not LN.field(fm["lane_section_id"]) or not LN.field(fm["lane_type"]):
        c.missing("SEC_DIR_03", "Lane on left side of SECTION", "LANE layer or fields missing")
    elif c.compat(LN, "SEC_DIR_03", "Lane on left side of SECTION"):
        for _, g, s, t in LN.geom_rows([LN.field(fm["lane_section_id"]), LN.field(fm["lane_type"])]):
            if g is not None:
                lanes.setdefault(sv(s), []).append((sv(t).upper(), g))

    bad = {"1": 0, "2": 0, "3": 0}
    for o, g, reg, frm, sid, typ in L.geom_rows([reg_f, fs_f, sec_f, typ_f]):
        ln = as_line(g)
        if ln is None or ln.length == 0:
            continue
        reasons, needs_flip = [], False
        if rtree is not None:
            m, n = ln.interpolate(0.5, normalized=True), ln.interpolate(0.51, normalized=True)
            dx, dy = n.x - m.x, n.y - m.y
            d = math.hypot(dx, dy) or 1.0
            tp = Point(m.x + dy / d * off, m.y - dx / d * off)
            hit = rtree.query(tp, predicate="within")
            actual = rnos[int(hit[0])] if len(hit) else None
            if actual != sv(reg):
                needs_flip = True
                bad["1"] += 1
                reasons.append("region not on right side")
                c.add("SEC_DIR_01", "Region on right side", "FAIL", "WARNING", o, reg_f, reg,
                      f"REGION at right side = {actual}; flip suggested")
        if mtree is not None:
            code = sv(frm).zfill(3)
            at_start = any(mcodes[i] == code for i in mtree.query(Point(ln.coords[0][:2]).buffer(tol), predicate="intersects"))
            at_end = any(mcodes[i] == code for i in mtree.query(Point(ln.coords[-1][:2]).buffer(tol), predicate="intersects"))
            if not at_start and not at_end:
                bad["2"] += 1
                c.add("SEC_DIR_02", "FROM_STREET at start point", "FAIL", "ERROR", o, fs_f, frm,
                      "CRITICAL: FROM_STREET not found at either end")
            elif at_end and not at_start:
                needs_flip = True
                bad["2"] += 1
                reasons.append("FROM_STREET at end")
                c.add("SEC_DIR_02", "FROM_STREET at start point", "FAIL", "WARNING", o, fs_f, frm,
                      "FROM_STREET found at end point; flip suggested")
        if lanes and sv(sid) in lanes:
            d = local_dir(g)
            valid = ["S1"] if sv(typ).upper() == "S" else ["L", "LP"]
            for lt, lg in lanes[sv(sid)]:
                if lt in valid and d is not None:
                    p = mid_point(lg)
                    if p is not None and side_of(d, p) == "RIGHT":
                        needs_flip = True
                        bad["3"] += 1
                        reasons.append(f"lane {lt} on right")
                        c.add("SEC_DIR_03", "Lane on left side of SECTION", "FAIL", "WARNING", o, "Shape", lt,
                              "Lane lies on the RIGHT of the section; flip suggested")
                        break
        if needs_flip:
            c.flips[o] = "; ".join(reasons)
    for k, rid, rname, ran in [("1", "SEC_DIR_01", "Region on right side", rtree is not None),
                               ("2", "SEC_DIR_02", "FROM_STREET at start point", mtree is not None),
                               ("3", "SEC_DIR_03", "Lane on left side of SECTION", bool(lanes))]:
        if ran and not bad[k]:
            c.add(rid, rname, "PASS")


# ------------------------------------------------------------------ LANE rules
def lane_id(c):
    concat_rule(c, "LANE_ID_01", "LANE_NO = SECTION_NO+LANE_TYPE", "lane_no", ["lane_section_id", "lane_type"])


def lane_type(c):
    f = c.L.field(CFG["fields"]["lane_type"])
    if not f:
        c.missing("LANE_TYPE_01", "LANE_TYPE domain", "LANE_TYPE field missing")
        return
    allowed = set(CFG["lane_types_service"]) | set(CFG["lane_types_other"])
    bad = 0
    for o, v in c.L.rows([f]):
        if sv(v) and sv(v).upper() not in allowed:
            bad += 1
            c.add("LANE_TYPE_01", "LANE_TYPE domain", "FAIL", "ERROR", o, f, v, f"Allowed: {sorted(allowed)}")
    if not bad:
        c.add("LANE_TYPE_01", "LANE_TYPE domain", "PASS")


def lane_munic(c):
    left_rule(c, "LANE_MUN_01", "MUNIC_CODE = left(SECTION_NO,2)", "munic_code", "lane_section_id", 2)


def section_lookup(c, rid, rname, value_field_key=None, geom=False):
    S = c.need_helper("SECTION", rid, rname)
    if S is None:
        return None, None
    sid = S.field(CFG["fields"]["section_id"])
    extra = S.field(CFG["fields"][value_field_key]) if value_field_key else None
    if not sid or (value_field_key and not extra):
        c.missing(rid, rname, "Reference fields missing in SECTION")
        return None, None
    if geom and not c.compat(S, rid, rname):
        return None, None
    d = {}
    for o, g, s, e in S.geom_rows([sid, extra]):
        if sv(s):
            d[sv(s)] = (g, e)
    return S, d


def lane_ref(c):
    a = c.L.field(CFG["fields"]["lane_section_id"])
    S, d = section_lookup(c, "LANE_REF_01", "LANE to SECTION reference")
    if S is None:
        return
    if not a:
        c.missing("LANE_REF_01", "LANE to SECTION reference", "Lane SECTION_NO field missing")
        return
    bad = 0
    for o, v in c.L.rows([a]):
        if sv(v) and sv(v) not in d:
            bad += 1
            c.add("LANE_REF_01", "LANE to SECTION reference", "FAIL", "ERROR", o, a, v, "SECTION_NO not found in SECTION layer")
    if not bad:
        c.add("LANE_REF_01", "LANE to SECTION reference", "PASS")


def lane_longer(c):
    a = c.L.field(CFG["fields"]["lane_section_id"])
    S, d = section_lookup(c, "LANE_LEN_02", "Lane longer than section", geom=True)
    if S is None:
        return
    if not a:
        c.missing("LANE_LEN_02", "Lane longer than section", "Lane SECTION_NO field missing")
        return
    fac = unit_factor(c.L)
    if fac is None:
        c.add("LANE_LEN_02", "Lane longer than section", "NOT_RUN", "WARNING", msg="Projected CRS required (metres)")
        return
    thr, bad = CFG["thresholds"]["lane_longer_m"], 0
    for o, g, s in c.L.geom_rows([a]):
        sg = d.get(sv(s), (None,))[0]
        if g is not None and sg is not None:
            diff = (g.length - sg.length) * fac
            if diff >= thr:
                bad += 1
                c.add("LANE_LEN_02", "Lane longer than section", "FAIL", "WARNING", o, a, s,
                      f"Lane is {diff:.2f} m longer than its section (limit {thr} m)")
    if not bad:
        c.add("LANE_LEN_02", "Lane longer than section", "PASS")


def lane_ratio(c):
    a = c.L.field(CFG["fields"]["lane_section_id"])
    S, d = section_lookup(c, "LANE_LEN_01", "LANE/SECTION length ratio", geom=True)
    if S is None or not a:
        return
    lo, hi, bad = CFG["thresholds"]["lane_ratio_min"], CFG["thresholds"]["lane_ratio_max"], 0
    for o, g, s in c.L.geom_rows([a]):
        sg = d.get(sv(s), (None,))[0]
        if g is not None and sg is not None and sg.length:
            r = g.length / sg.length
            if r < lo or r > hi:
                bad += 1
                c.add("LANE_LEN_01", "LANE/SECTION length ratio", "REVIEW", "WARNING", o, a, s, f"Ratio={r:.3f}; accepted {lo}-{hi}")
    if not bad:
        c.add("LANE_LEN_01", "LANE/SECTION length ratio", "PASS")


def lane_vs_sec_type(c):
    fm = CFG["fields"]
    a, t = c.L.field(fm["lane_section_id"]), c.L.field(fm["lane_type"])
    S, d = section_lookup(c, "LANE_TYP_02", "Lane type vs section type", "sec_type")
    if S is None:
        return
    if not (a and t):
        c.missing("LANE_TYP_02", "Lane type vs section type", "Lane fields missing")
        return
    svc, oth, bad = set(CFG["lane_types_service"]), set(CFG["lane_types_other"]), 0
    for o, s, lt in c.L.rows([a, t]):
        st = sv(d.get(sv(s), (None, ""))[1]).upper()
        lt_ = sv(lt).upper()
        wrong = (st == "S" and lt_ not in svc) or (st in {"B", "T", "N", "X", "R"} and lt_ not in oth)
        if wrong:
            bad += 1
            c.add("LANE_TYP_02", "Lane type vs section type", "FAIL", "ERROR", o, t, lt, f"SEC_TYPE={st} does not allow lane type {lt_}")
    if not bad:
        c.add("LANE_TYP_02", "Lane type vs section type", "PASS")


def lane_order(c):
    a, t = c.L.field(CFG["fields"]["lane_section_id"]), c.L.field(CFG["fields"]["lane_type"])
    fac = unit_factor(c.L)
    if not (a and t):
        c.missing("LANE_ORD_01", "Lane length order", "Lane fields missing")
        return
    if fac is None:
        c.add("LANE_ORD_01", "Lane length order", "NOT_RUN", "WARNING", msg="Projected CRS required (metres)")
        return
    thr = CFG["thresholds"]["lane_order_diff_m"]
    members, tl = {}, {}
    for o, g, s, lt in c.L.geom_rows([a, t]):
        if not sv(s):
            continue
        members.setdefault(sv(s), []).append((o, lt))
        if g is not None:
            k = sv(lt).upper()
            tl.setdefault(sv(s), {})[k] = max(tl.get(sv(s), {}).get(k, 0), g.length * fac)
    bad = 0
    for s, tmap in tl.items():
        issues = []
        for seq in CFG["lane_order_rules"]:
            for x, y in zip(seq, seq[1:]):
                if x in tmap and y in tmap and tmap[y] - tmap[x] >= thr:
                    issues.append(f"{x}<{y} by {tmap[y] - tmap[x]:.2f}m")
        if issues:
            bad += 1
            msg = "; ".join(issues[:5]) + (f" ...(+{len(issues) - 5})" if len(issues) > 5 else "")
            for o, lt in members[s]:
                c.add("LANE_ORD_01", "Lane length order", "REVIEW", "WARNING", o, t, lt, msg)
    if not bad:
        c.add("LANE_ORD_01", "Lane length order", "PASS")


def lane_direction(c):
    a = c.L.field(CFG["fields"]["lane_section_id"])
    S, d = section_lookup(c, "LANE_DIR_01", "Lane opposite to section", geom=True)
    if S is None or not a:
        return
    bad = 0
    for o, g, s in c.L.geom_rows([a]):
        sg = d.get(sv(s), (None,))[0]
        if g is None or sg is None:
            continue
        v1, v2 = local_dir(sg), local_dir(g)
        if v1 and v2:
            dp = (v1[2] - v1[0]) * (v2[2] - v2[0]) + (v1[3] - v1[1]) * (v2[3] - v2[1])
            if dp < 0:
                bad += 1
                c.add("LANE_DIR_01", "Lane opposite to section", "FAIL", "ERROR", o, a, s, "Lane direction is opposite to its section")
    if not bad:
        c.add("LANE_DIR_01", "Lane opposite to section", "PASS")


# ------------------------------------------------------------------ INTERSECTION / SAMPLE / REGION rules
def intersection_id(c):
    concat_rule(c, "INT_ID_01", "INTERSECTION_NO = MUNIC+MAIN1+MAIN2", "intersection_no", ["munic_code", "main_no1", "main_no2"])


def intersection_munic(c):
    left_rule(c, "INT_MUN_01", "MUNIC_CODE = left(DISTRICT_NO,2)", "munic_code", "district_no", 2)


def intersection_mains(c):
    M = c.need_helper("MAIN_STREETS", "INT_MAIN_01", "Intersection mains exist nearby")
    if M is None or not c.compat(M, "INT_MAIN_01", "Intersection mains exist nearby"):
        return
    f, mf = c.L.field(CFG["fields"]["intersection_no"]), M.field(CFG["fields"]["main_id"])
    if not f or not mf:
        c.missing("INT_MAIN_01", "Intersection mains exist nearby", "INTERSECTION_NO or MAIN_NO field missing")
        return
    fac = unit_factor(c.L)
    if fac is None:
        c.add("INT_MAIN_01", "Intersection mains exist nearby", "NOT_RUN", "WARNING", msg="Projected CRS required (metres)")
        return
    dist = CFG["thresholds"]["intersection_search_m"] / fac
    mg = [(g, last3(n)) for _, g, n in M.geom_rows([mf]) if g is not None]
    tree, codes = (STRtree([g for g, _ in mg]), [n for _, n in mg]) if mg else (None, [])
    bad = 0
    for o, g, no in c.L.geom_rows([f]):
        s = sv(no)
        if len(s) < 8:
            bad += 1
            c.add("INT_MAIN_01", "Intersection mains exist nearby", "FAIL", "ERROR", o, f, no, "Invalid code length (<8)")
            continue
        if g is None or tree is None:
            continue
        near = {codes[i] for i in tree.query(g.buffer(dist), predicate="intersects")}
        miss = [f"M{i}({m})" for i, m in ((1, s[2:5]), (2, s[5:8])) if m not in near]
        if miss:
            bad += 1
            c.add("INT_MAIN_01", "Intersection mains exist nearby", "FAIL", "ERROR", o, f, no,
                  f"Missing {', '.join(miss)} | Found nearby: {', '.join(sorted(near)) or 'None'}")
    if not bad:
        c.add("INT_MAIN_01", "Intersection mains exist nearby", "PASS")


def sample_id(c):
    concat_rule(c, "SMP_ID_01", "INTSAMPLE_NO = INTERSECTION_NO+CODE", "intsample_no", ["intersection_no", "intsample_code"])


def region_id(c):
    concat_rule(c, "REG_ID_01", "REGION_NO = MUNIC+BRANCH+DIST+REGION_CODE", "region_id",
                ["munic_code", "branch_code", "dist_code", "region_code"])


def region_district(c):
    left_rule(c, "REG_DIS_01", "DISTRICT_NO = left(REGION_NO,6)", "district_no", "region_id", 6)


def region_sliver(c):
    fac = unit_factor(c.L)
    if fac is None:
        c.add("REG_SLV_01", "Sliver polygon", "NOT_RUN", "WARNING", msg="Projected CRS required (m2)")
        return
    lim, bad = CFG["thresholds"]["sliver_area_m2"], 0
    for o, g in c.L.geom_rows():
        if g is not None and g.area * fac * fac < lim:
            bad += 1
            c.add("REG_SLV_01", "Sliver polygon", "FAIL", "WARNING", o, "Shape", round(g.area * fac * fac, 4), f"Area < {lim} m2")
    if not bad:
        c.add("REG_SLV_01", "Sliver polygon", "PASS")


def region_admin(c):
    L = c.L
    cents = [None if g is None else g.centroid for g in (L.geoms if L.geoms is not None else [])]
    for key, ck, nm in [("MUNICIPALITY", "munic_code", "Municipality"), ("DISTRICT", "dist_code", "District"), ("BRANCH", "branch_code", "Branch")]:
        rid = f"REG_ADM_{nm[:3].upper()}"
        rname = f"REGION {nm} code matches {nm} polygon"
        H = c.helper(key)
        if H is None:
            c.missing(rid, rname, f"{key} layer missing")
            continue
        if not c.compat(H, rid, rname):
            continue
        rf, hf = L.field(CFG["fields"][ck]), H.field(CFG["fields"][ck])
        if not rf or not hf:
            c.missing(rid, rname, f"Field {CFG['fields'][ck]} missing in REGION or {key}")
            continue
        hg = [(g, sv(v)) for _, g, v in H.geom_rows([hf]) if g is not None]
        if not hg:
            c.missing(rid, rname, f"{key} has no geometry")
            continue
        tree, hv = STRtree([g for g, _ in hg]), [v for _, v in hg]
        bad = 0
        for (o, g, v), cp in zip(L.geom_rows([rf]), cents):
            if cp is None:
                continue
            hit = tree.query(cp, predicate="within")
            if not len(hit):
                bad += 1
                c.add(rid, rname, "FAIL", "ERROR", o, rf, v, f"{nm} not found spatially (outside)")
            elif hv[int(hit[0])] != sv(v):
                bad += 1
                c.add(rid, rname, "FAIL", "ERROR", o, rf, v, f"Mismatch ({sv(v)} != {hv[int(hit[0])]})")
        if not bad:
            c.add(rid, rname, "PASS")


# ------------------------------------------------------------------ dispatch
RULES = {
    "SECTION": [("section_id", section_id), ("section_dup", section_dup), ("section_base_serial", section_base_serial),
                ("section_serial", section_serial), ("section_region", section_region), ("section_munic", section_munic),
                ("section_type", section_type), ("section_length", section_length),
                ("section_intersection", section_intersection), ("section_street_names", section_street_names),
                ("section_direction", section_direction)],
    "LANE": [("lane_id", lane_id), ("lane_type", lane_type), ("lane_munic", lane_munic), ("lane_section_ref", lane_ref),
             ("lane_longer_than_section", lane_longer), ("lane_length_ratio", lane_ratio),
             ("lane_type_vs_sec_type", lane_vs_sec_type), ("lane_length_order", lane_order),
             ("lane_direction", lane_direction)],
    "INTERSECTION": [("intersection_id", intersection_id), ("intersection_munic", intersection_munic),
                     ("intersection_mains", intersection_mains)],
    "INTERSEC_SAMPLE": [("sample_id", sample_id)],
    "REGION": [("region_id", region_id), ("region_district", region_district), ("region_sliver", region_sliver),
               ("region_admin", region_admin)],
}

PENDING = {
    "REGION": ["N/S/E/W boundary street adjacency", "Street direction vs region centre"],
    "MAIN_STREETS": ["Every main-road crossing inside INTERSECTION polygon"],
    "SECONDARY": ["Secondary road must not cross MAIN_STREETS; touch allowed"],
    "INTERSECTION": ["MAIN_NO1/MAIN_NO2 orientation", "At least one sample per intersection"],
    "INTERSEC_SAMPLE": ["Polygon coverage", "South counter-clockwise numbering", "Arm-pair completeness"],
    "SECTION": ["Urban endpoints", "Bridge/tunnel splitting", "Full SECTION topology edge analysis"],
    "LANE": ["Lane corridor containment", "Lane termination at intersection boundary"],
}


def execute(c):
    run_common(c)
    if c.base != "SECTION" and enabled("duplicates"):
        duplicates(c)
    for key, fn in RULES.get(c.base, []):
        if enabled(key):
            try:
                fn(c)
            except Exception as e:
                logging.exception("Rule %s failed on %s", key, c.name)
                c.add("ENGINE_ERROR", f"Rule {key}", "ERROR", "ERROR", msg=f"{type(e).__name__}: {e}")
    for n in PENDING.get(c.base, []):
        c.add("PENDING_CONFIG", n, "NOT_CONFIGURED", "INFO",
              msg="Requires confirmed field mappings, direction convention, CRS/tolerance or project exceptions; not guessed.")


# ------------------------------------------------------------------ outputs
HEADERS = ["RunTime", "SourceGDB", "Layer", "RuleID", "RuleName", "Status", "Severity", "FeatureID", "Field", "Value", "Message"]
BAD = ("FAIL", "REVIEW", "ERROR")


def export_excel(rows, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "QC Findings"
    ws.append(HEADERS)
    for r in rows:
        ws.append([r.get(h, "") for h in HEADERS])
    fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font, cell.fill = Font(bold=True, color="FFFFFF"), fill
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col in ws.columns:
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(55, max(12, max(len(str(x.value or "")) for x in col) + 2))
    sm = wb.create_sheet("Summary")
    sm.append(["Layer", "RuleID", "RuleName", "Status", "Count"])
    counts = {}
    for r in rows:
        k = (r["Layer"], r["RuleID"], r["RuleName"], r["Status"])
        counts[k] = counts.get(k, 0) + 1
    for k, n in sorted(counts.items()):
        sm.append([*k, n])
    for cell in sm[1]:
        cell.font, cell.fill = Font(bold=True, color="FFFFFF"), fill
    sm.freeze_panes = "A2"
    sm.auto_filter.ref = sm.dimensions
    for col in sm.columns:
        sm.column_dimensions[get_column_letter(col[0].column)].width = min(55, max(10, max(len(str(x.value or "")) for x in col) + 2))
    wb.save(path)


def collect_shp(layer, rows, start, flips):
    pos = {int(f): i for i, f in enumerate(layer.fids)}
    recs, fl = [], []
    if layer.geoms is not None:
        for r in rows[start:]:
            if r["Status"] in BAD and r["FeatureID"] != "":
                i = pos.get(int(r["FeatureID"]))
                if i is not None and layer.geoms[i] is not None:
                    recs.append((layer.geoms[i], r))
        for fid, why in flips.items():
            i = pos.get(fid)
            if i is not None and layer.geoms[i] is not None:
                fl.append((shapely.reverse(layer.geoms[i]), fid, why))
    if not recs and not fl:
        return None
    return {"gdb": layer.gdb, "name": layer.name, "gtype": layer.gtype, "crs": layer.crs, "recs": recs, "flips": fl}


def export_shapefiles(items, folder):
    from pyogrio.raw import write
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    fields = ["SRC_FID", "RULE_ID", "STATUS", "SEVERITY", "FIELD", "VALUE", "MESSAGE"]
    keys = ["FeatureID", "RuleID", "Status", "Severity", "Field", "Value", "Message"]
    n = 0
    for it in items:
        stem = f"{Path(it['gdb']).stem}_{it['name']}"
        crs = it["crs"].to_wkt() if it["crs"] else None
        gt = it["gtype"] or "Unknown"
        if it["recs"]:
            geoms = np.array([g for g, _ in it["recs"]], dtype=object)
            data = [np.array([int(r["FeatureID"]) for _, r in it["recs"]], dtype="int64")]
            data += [np.array([str(r[k])[:254] for _, r in it["recs"]], dtype=object) for k in keys[1:]]
            write(str(folder / f"{stem}_findings.shp"), shapely.to_wkb(geoms), data, fields,
                  driver="ESRI Shapefile", geometry_type=gt, crs=crs, encoding="UTF-8")
            n += 1
        if it["flips"]:
            geoms = np.array([g for g, _, _ in it["flips"]], dtype=object)
            data = [np.array([f for _, f, _ in it["flips"]], dtype="int64"),
                    np.array([w[:254] for _, _, w in it["flips"]], dtype=object)]
            write(str(folder / f"{stem}_suggested_flip.shp"), shapely.to_wkb(geoms), data, ["SRC_FID", "REASON"],
                  driver="ESRI Shapefile", geometry_type=gt, crs=crs, encoding="UTF-8")
            n += 1
    return n


def run(source, selected, notify):
    """selected: set of (gdb_path, layer_name). Returns (rows, shapefile_items)."""
    gdbs = gdb_list(source)
    if not gdbs:
        raise RuntimeError("No .gdb found.")
    names = {g: feature_classes(g) for g in gdbs}
    hmaps = {g: helper_map(ns) for g, ns in names.items()}
    freq = {}
    for g, ns in names.items():
        for n in ns:
            k = base_of(n, hmaps[g])
            freq[k] = freq.get(k, 0) + 1
    chosen = [(g, n) for g, ns in names.items() for n in ns if (g, n) in selected]
    rows, shp = [], []
    stub = lambda n: type("X", (), {"name": n})()
    for i, (gdb, name) in enumerate(chosen, 1):
        notify(i, len(chosen), f"QC: {name}")
        base = base_of(name, hmaps[gdb])
        if freq[base] > 1:
            Ctx(gdb, stub(name), base, {}, rows).add("SOURCE_AMBIGUOUS", "Duplicate layer name across GDBs", "NOT_RUN", "WARNING",
                                                     msg="Resolve source GDB; system did not choose one automatically.")
            continue
        try:
            start = len(rows)
            lyr = Layer(gdb, name)
            ctx = Ctx(gdb, lyr, base, hmaps[gdb], rows)
            execute(ctx)
            item = collect_shp(lyr, rows, start, ctx.flips)
            if item:
                shp.append(item)
        except Exception as e:
            logging.exception("Layer failed: %s", name)
            Ctx(gdb, stub(name), base, {}, rows).add("ENGINE_ERROR", "Layer execution", "ERROR", "ERROR", msg=f"{type(e).__name__}: {e}")
    return rows, shp


# ------------------------------------------------------------------ GUI
class App:
    def __init__(self, root):
        self.root = root
        root.title("QC_PMMS")
        root.geometry("1050x720")
        try:
            base = Path(getattr(sys, "_MEIPASS", HERE))
            root.iconbitmap(str(base / "NAMAA_PNG_main_3.ico"))
        except Exception:
            pass
        self.src = tk.StringVar()
        self.q = queue.Queue()
        self._ui()
        root.after(120, self._poll)

    def _ui(self):
        top = ttk.Frame(self.root, padding=10)
        top.pack(fill="x")
        ttk.Label(top, text="GDB / Folder").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.src, width=100).grid(row=1, column=0, sticky="ew", padx=5)
        ttk.Button(top, text="Browse", command=self.browse).grid(row=1, column=1)
        ttk.Button(top, text="Scan Layers", command=self.scan).grid(row=1, column=2, padx=5)
        top.columnconfigure(0, weight=1)
        box = ttk.LabelFrame(self.root, text="Target layers — select layers only", padding=8)
        box.pack(fill="both", expand=True, padx=10, pady=8)
        self.tree = ttk.Treeview(box, columns=("gdb", "path"), show="tree headings", selectmode="extended")
        for col, txt, w in [("#0", "Layer", 200), ("gdb", "GDB", 220), ("path", "GDB path", 520)]:
            self.tree.heading(col, text=txt)
            self.tree.column(col, width=w)
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        bar = ttk.Frame(self.root, padding=10)
        bar.pack(fill="x")
        ttk.Button(bar, text="Select All", command=lambda: self.tree.selection_set(self.tree.get_children())).pack(side="left")
        ttk.Button(bar, text="Clear", command=lambda: self.tree.selection_remove(self.tree.selection())).pack(side="left", padx=5)
        self.runbtn = ttk.Button(bar, text="Run QC", command=self.start)
        self.runbtn.pack(side="right")
        self.pb = ttk.Progressbar(self.root, mode="determinate")
        self.pb.pack(fill="x", padx=10)
        self.status = ttk.Label(self.root, text="Read-only mode. Select target layers, then Run QC.")
        self.status.pack(fill="x", padx=10, pady=8)

    def browse(self):
        p = filedialog.askdirectory()
        if p:
            self.src.set(p)

    def scan(self):
        self.tree.delete(*self.tree.get_children())
        total = 0
        try:
            for g in gdb_list(self.src.get()):
                for n in feature_classes(g):
                    self.tree.insert("", "end", text=n, values=(Path(g).name, g))
                    total += 1
        except Exception as e:
            logging.exception("Scan failed")
            messagebox.showerror("Scan error", str(e))
            return
        self.status.config(text=f"Found {total} feature classes.")

    def start(self):
        sel = {(str(self.tree.item(i, "values")[1]), str(self.tree.item(i, "text"))) for i in self.tree.selection()}
        if not sel:
            messagebox.showwarning("Select layers", "Select at least one target layer.")
            return
        out = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
                                           initialfile=f"Namaa_QC_{datetime.now():%Y%m%d_%H%M}.xlsx")
        if not out:
            return
        self.runbtn.config(state="disabled")
        src = self.src.get()

        def worker():
            try:
                rows, shp = run(src, sel, lambda i, n, s: self.q.put(("progress", i, n, s)))
                export_excel(rows, out)
                folder = Path(out).with_suffix("").as_posix() + "_shp"
                n = export_shapefiles(shp, folder) if shp else 0
                self.q.put(("done", out, len(rows), folder, n))
            except Exception as e:
                logging.exception("QC failed")
                self.q.put(("error", str(e)))
        threading.Thread(target=worker, daemon=True).start()

    def _poll(self):
        try:
            while True:
                m = self.q.get_nowait()
                if m[0] == "progress":
                    _, i, n, s = m
                    self.pb["maximum"], self.pb["value"] = max(1, n), i
                    self.status.config(text=s)
                elif m[0] == "done":
                    self.runbtn.config(state="normal")
                    self.status.config(text="Completed")
                    messagebox.showinfo("QC completed", f"Report saved:\n{m[1]}\nRows: {m[2]}\n\nShapefiles ({m[4]}):\n"
                                        f"{m[3] if m[4] else 'none - no findings with geometry'}")
                elif m[0] == "error":
                    self.runbtn.config(state="normal")
                    self.status.config(text="Failed — see qc.log")
                    messagebox.showerror("QC error", m[1])
        except queue.Empty:
            pass
        self.root.after(120, self._poll)


def check_expiry():
    if date.today() >= EXPIRY_DATE:
        logging.info("Expired - refused to start (%s)", date.today())
        r = tk.Tk()
        r.withdraw()
        messagebox.showerror("QC_PMMS", "انتهت صلاحية هذه النسخة من الأداة.\nاتواصل مع أحمد ياسر.\n\n"
                             "This tool has expired.\nPlease contact Ahmed Yasser.")
        r.destroy()
        return False
    return True


def main():
    if not check_expiry():
        sys.exit(0)
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
