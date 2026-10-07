"""Füllt die Monatsblätter einer Arbeitszeiterfassung direkt im XLSX-XML.

Es werden ausschließlich die Zellen D, E, J, K (und P für Bemerkungen) der Tageszeilen 4-34
angefasst. Formeln, bedingte Formatierung, Dropdowns und Namen bleiben unverändert.
"""
import os
import re
import zipfile
from calendar import monthrange
from html import escape, unescape

MONTH_SHEETS = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli",
                "August", "September", "Oktober", "November", "Dezember"]
FIRST_ROW = 4  # Tag 1 steht in Zeile 4
K_FALLBACK = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "k4_formula.txt"),
                  encoding="utf8").read()

HALF = "0.3333333333333333"  # 8:00 h als Tagesbruchteil


def t(h, m):
    return repr(h / 24 + m / 1440)


def _hhmm(v):
    h, m = map(int, v.split(":"))
    return t(h, m)


def build_mapping(shifts, daily):
    """shifts: {"F": {"start": "06:00", "end": "14:30"}, ...}; daily: Tagesbruchteil für BS/U/K.
    code -> (Kommt, Geht, J-Code, K-Wert, Bemerkung)"""
    m = {c: (_hhmm(v["start"]), _hhmm(v["end"]), None, None, None) for c, v in shifts.items()}
    if "S" in shifts:
        m["S/L"] = (m["S"][0], m["S"][1], None, None, "Training")
    m["BS"] = (None, None, "B", HALF, None)  # Berufsschule: bei allen Firmen fix 8:00
    m["DGR"] = (None, None, None, None, "DGR")  # nur Bemerkung, keine Zeiten
    m["U"] = (None, None, "U", daily, None)
    m["UU"] = (None, None, "U", daily, None)  # wie normaler Urlaub behandelt
    m["K"] = (None, None, "K", daily, None)
    return m


OUR_REMARKS = {"Training", "DGR"}
PROTECTED_J = {"K", "KR", "KU", "KA", "G", "UH"}  # manuell eingetragene Fehlzeiten bleiben erhalten
FILE_RE = re.compile(r"Arbeitszeiterfassung[_\- ]+(.+?)\.xlsx$", re.I)


def col_idx(col):
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


class Book:
    def __init__(self, path):
        zin = zipfile.ZipFile(path)
        self.infos = {i.filename: i for i in zin.infolist()}
        self.parts = {i.filename: zin.read(i.filename) for i in zin.infolist()}
        zin.close()
        wb = self.parts["xl/workbook.xml"].decode("utf8")
        rels = self.parts["xl/_rels/workbook.xml.rels"].decode("utf8")
        rel_map = {}
        for m in re.finditer(r"<Relationship\b[^>]*>", rels):
            tag = m.group(0)
            rid = re.search(r'\bId="([^"]+)"', tag).group(1)
            tgt = re.search(r'\bTarget="([^"]+)"', tag).group(1)
            rel_map[rid] = tgt
        self.sheet_path = {}
        for m in re.finditer(r"<sheet\b[^>]*>", wb):
            tag = m.group(0)
            name = unescape(re.search(r'\bname="([^"]+)"', tag).group(1))
            rid = re.search(r'\br:id="([^"]+)"', tag).group(1)
            tgt = rel_map[rid].lstrip("/")
            self.sheet_path[name] = tgt if tgt.startswith("xl/") else "xl/" + tgt
        self.shared = None

    def sheet(self, name):
        p = self.sheet_path.get(name)
        return p, (self.parts[p].decode("utf8") if p else None)

    def shared_strings(self):
        if self.shared is None:
            self.shared = []
            if "xl/sharedStrings.xml" in self.parts:
                x = self.parts["xl/sharedStrings.xml"].decode("utf8")
                for si in re.findall(r"<si>(.*?)</si>", x, re.S):
                    self.shared.append(unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S))))
        return self.shared

    def save(self, out_path):
        wb = self.parts["xl/workbook.xml"].decode("utf8")
        if "fullCalcOnLoad" not in wb:
            wb = re.sub(r"<calcPr\b", '<calcPr fullCalcOnLoad="1"', wb, count=1)
        self.parts["xl/workbook.xml"] = wb.encode("utf8")
        # calcChain entfernen, damit Excel keine Reparatur meldet (wird beim Öffnen neu aufgebaut)
        if "xl/calcChain.xml" in self.parts:
            del self.parts["xl/calcChain.xml"]
            ct = self.parts["[Content_Types].xml"].decode("utf8")
            ct = re.sub(r'<Override[^>]*calcChain[^>]*/>', "", ct)
            self.parts["[Content_Types].xml"] = ct.encode("utf8")
            rels = self.parts["xl/_rels/workbook.xml.rels"].decode("utf8")
            rels = re.sub(r'<Relationship\b[^>]*calcChain[^>]*/>', "", rels)
            self.parts["xl/_rels/workbook.xml.rels"] = rels.encode("utf8")
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in self.infos:
                if name in self.parts:
                    zout.writestr(name, self.parts[name])


CELL_RE = r'<c r="{ref}"((?:\s[^>]*?)?)(?:/>|>(.*?)</c>)'


def get_cell(xml, ref):
    return re.search(CELL_RE.format(ref=ref), xml, re.S)


def cell_value(book, m):
    """(wert, hat_formel) einer Zelle."""
    if not m:
        return None, False
    attrs, inner = m.group(1), m.group(2) or ""
    has_f = "<f" in inner
    v = re.search(r"<v>(.*?)</v>", inner, re.S)
    ts = re.search(r'\bt="(\w+)"', attrs)
    typ = ts.group(1) if ts else "n"
    if typ == "inlineStr":
        return unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", inner, re.S))), has_f
    if v is None:
        return None, has_f
    if typ == "s":
        return book.shared_strings()[int(v.group(1))], has_f
    if typ in ("str", "e", "b"):
        return unescape(v.group(1)), has_f
    return float(v.group(1)), has_f


def _style(attrs):
    s = re.search(r'\bs="(\d+)"', attrs)
    return f' s="{s.group(1)}"' if s else ""


def build_cell(ref, attrs, kind, value=None):
    s = _style(attrs)
    if kind == "empty":
        return f'<c r="{ref}"{s}/>'
    if kind == "num":
        return f'<c r="{ref}"{s}><v>{value}</v></c>'
    if kind == "str":
        return f'<c r="{ref}"{s} t="inlineStr"><is><t>{escape(value)}</t></is></c>'
    if kind == "formula":
        return f'<c r="{ref}"{s}><f>{value}</f></c>'
    raise ValueError(kind)


def put_cell(xml, ref, kind, value=None):
    m = get_cell(xml, ref)
    if m:
        new = build_cell(ref, m.group(1), kind, value)
        return xml[:m.start()] + new + xml[m.end():]
    col, row = re.match(r"([A-Z]+)(\d+)", ref).groups()
    rm = re.search(r'(<row r="%s"[^>]*?)(/>|>(.*?)</row>)' % row, xml, re.S)
    if not rm:
        raise ValueError(f"Zeile {row} fehlt in der Vorlage")
    new = build_cell(ref, "", kind, value)
    if rm.group(2) == "/>":
        repl = rm.group(1) + ">" + new + "</row>"
        return xml[:rm.start()] + repl + xml[rm.end():]
    body = rm.group(3)
    pos = len(body)
    for cm in re.finditer(r'<c r="([A-Z]+)\d+"', body):
        if col_idx(cm.group(1)) > col_idx(col):
            pos = cm.start()
            break
    body = body[:pos] + new + body[pos:]
    repl = rm.group(1) + ">" + body + "</row>"
    return xml[:rm.start()] + repl + xml[rm.end():]


def translate_formula(text, src_row, dst_row):
    return re.sub(r"(?<![A-Za-z$])([A-Z]{1,2})(\d+)\b",
                  lambda m: m.group(1) + str(dst_row) if int(m.group(2)) == src_row else m.group(0), text)


def expand_k_formulas(xml):
    """Ersetzt gemeinsame Formeln der K-Spalte (Zeilen 4-34) durch einzelne Formeln."""
    masters = {}
    for r in range(FIRST_ROW, FIRST_ROW + 31):
        m = get_cell(xml, f"K{r}")
        if m and m.group(2):
            fm = re.search(r'<f t="shared"[^>]*\bsi="(\d+)"[^>]*>(.+?)</f>', m.group(2), re.S)
            if fm:
                masters[fm.group(1)] = (fm.group(2), r)
    fallback = (K_FALLBACK, FIRST_ROW)
    for r in range(FIRST_ROW, FIRST_ROW + 31):
        m = get_cell(xml, f"K{r}")
        if not m or not m.group(2) or "<f" not in m.group(2):
            continue
        sm = re.search(r'<f t="shared"[^>]*\bsi="(\d+)"', m.group(2))
        if not sm:
            continue
        text, src = masters.get(sm.group(1), fallback)
        xml = put_cell(xml, f"K{r}", "formula", translate_formula(text, src, r))
    return xml


def parse_name(filename):
    m = FILE_RE.search(os.path.basename(filename))
    if not m:
        return "", ""
    parts = [p for p in re.split(r"_+", m.group(1)) if p]
    if len(parts) == 1:
        parts = parts[0].split()
    return (parts[0], " ".join(parts[1:])) if parts else ("", "")


def daily_hours(xml):
    """Tägliche SOLL-Zeit aus der Vorlage (Spalte O, z.B. IF(D4="",0,0.333...)) als Tagesbruchteil."""
    m = re.search(r'<c r="O\d+"[^>]*><f>IF\(D\d+=""[^,]*,0,([0-9.]+)\)</f>', xml)
    return m.group(1) if m else HALF


def inspect(path, month):
    """Basisdaten für die Vorschau: Jahr, Name, vorhandene Einträge im Monat."""
    b = Book(path)
    info = {"sheet": MONTH_SHEETS[month - 1], "year": None, "name": "", "hasData": False,
            "filledDays": 0, "sheetFound": MONTH_SHEETS[month - 1] in b.sheet_path, "dailyHours": 8.0}
    _, vx = b.sheet("Voreinstellungen")
    if vx:
        v, _ = cell_value(b, get_cell(vx, "C2"))
        info["year"] = int(v) if isinstance(v, float) else None
        n, _ = cell_value(b, get_cell(vx, "C3"))
        info["name"] = n if isinstance(n, str) else ""
    _, x = b.sheet(info["sheet"])
    if x:
        info["dailyHours"] = round(float(daily_hours(x)) * 24, 2)
        for r in range(FIRST_ROW, FIRST_ROW + 31):
            filled = False
            for col in "DEJ":
                v, f = cell_value(b, get_cell(x, f"{col}{r}"))
                if v not in (None, "", 0.0) and not f:
                    filled = True
            v, f = cell_value(b, get_cell(x, f"K{r}"))
            if not f and v not in (None, "", 0.0):
                filled = True
            info["filledDays"] += filled
        info["hasData"] = info["filledDays"] > 0
    return info


def fill(path_in, path_out, month, year, days, vorname_nachname="", overwrite=True, shifts=None):
    """days: {"1": "F", ...}. Gibt einen Bericht zurück."""
    b = Book(path_in)
    sheet = MONTH_SHEETS[month - 1]
    report = {"counts": {}, "hours": 0.0, "warnings": [],
              "written": 0, "sheet": sheet, "skipped": False}
    path, x = b.sheet(sheet)
    if not x:
        raise ValueError(f"Blatt '{sheet}' nicht in der Datei gefunden")
    vpath, vx = b.sheet("Voreinstellungen")
    if vx:
        yv, _ = cell_value(b, get_cell(vx, "C2"))
        if isinstance(yv, float) and int(yv) != year:
            report["warnings"].append(f"Jahr in der Datei ({int(yv)}) weicht vom Dienstplan ({year}) ab.")
        nv, _ = cell_value(b, get_cell(vx, "C3"))
        if not nv and vorname_nachname:
            b.parts[vpath] = put_cell(vx, "C3", "str", vorname_nachname).encode("utf8")

    existing = inspect(path_in, month)
    if existing["hasData"] and not overwrite:
        report["skipped"] = True
        report["warnings"].append("Monat enthält bereits Einträge - übersprungen.")
        b.save(path_out)
        return report

    x = expand_k_formulas(x)
    daily = daily_hours(x)
    daily_h = round(float(daily) * 24, 2)
    shifts = shifts or {}
    mapping = build_mapping(shifts, daily)
    n_days = monthrange(year, month)[1]
    for d in range(1, n_days + 1):
        r = FIRST_ROW + d - 1
        code = (days.get(str(d)) or "").strip().upper()
        if code and code not in mapping:
            report["warnings"].append(f"Tag {d}: Code '{code}' unbekannt - nichts geschrieben.")
            continue
        j_old, _ = cell_value(b, get_cell(x, f"J{r}"))
        if isinstance(j_old, str) and j_old.upper() in PROTECTED_J and code != j_old.upper():
            report["warnings"].append(f"Tag {d}: vorhandener Eintrag '{j_old}' bleibt erhalten (Plan: {code or '-'}).")
            continue
        for col in "DEJ":
            x = put_cell(x, f"{col}{r}", "empty")
        k_formula = translate_formula(K_FALLBACK, FIRST_ROW, r)
        p_old, _ = cell_value(b, get_cell(x, f"P{r}"))
        if isinstance(p_old, str) and p_old in OUR_REMARKS:
            x = put_cell(x, f"P{r}", "empty")
        if not code:
            x = put_cell(x, f"K{r}", "formula", k_formula)
            continue
        kommt, geht, jcode, kval, remark = mapping[code]
        if kommt:
            x = put_cell(x, f"D{r}", "num", kommt)
            x = put_cell(x, f"E{r}", "num", geht)
            x = put_cell(x, f"K{r}", "formula", k_formula)
        elif jcode is None:
            x = put_cell(x, f"K{r}", "formula", k_formula)  # nur Bemerkung
        else:
            x = put_cell(x, f"J{r}", "str", jcode)
            x = put_cell(x, f"K{r}", "num", kval)
        if remark:
            x = put_cell(x, f"P{r}", "str", remark)
        key = "U" if code == "UU" else code
        report["counts"][key] = report["counts"].get(key, 0) + 1
        report["hours"] += (shifts[code]["net"] if code in shifts else (shifts["S"]["net"] if code == "S/L" and "S" in shifts
                            else (0 if code == "DGR" else (8 if code == "BS" else daily_h))))
        report["written"] += 1
    b.parts[path] = x.encode("utf8")
    b.save(path_out)
    return report
