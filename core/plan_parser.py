"""Liest einen Dienstplan (Ausbildungsplan) als PDF. Unterstützt HC (Handling Counts) und FACL (Fiege).

Liefert pro Azubi die Codes je Tag und die Schichtzeiten. Die Zeiten werden aus der Legende des PDFs
gelesen; ist eine Legendenzeile nicht lesbar (z.B. überlappender Text), gelten die Standardzeiten der Firma.
"""
import re
import pdfplumber

MONTHS = {
    "jan": 1, "feb": 2, "mär": 3, "mar": 3, "maer": 3, "apr": 4, "mai": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "okt": 10, "nov": 11, "dez": 12,
}
STATIC_CODES = {"BS", "U", "UU", "K", "S/L", "GW", "DGR"}  # keine Schichten mit Uhrzeiten

COMPANIES = {
    "HC": {"label": "HC (Handling Counts)", "shifts": {
        "F": ("06:00", "14:30"), "S": ("12:30", "21:00"), "N": ("21:45", "06:15")}},
    "LUG": {"label": "LUG FRA", "shifts": {
        "1": ("06:55", "14:59"), "2": ("14:55", "22:59"), "3": ("22:55", "06:59"), "4": ("03:55", "11:59"),
        "5": ("08:55", "16:59"), "6": ("10:55", "18:59"), "7": ("00:55", "08:59"), "WD": ("09:00", "15:00")}},
    "FACL": {"label": "FACL (Fiege Air Cargo Logistic)", "shifts": {
        "F": ("06:00", "14:00"), "M": ("09:00", "17:00"), "S": ("14:00", "22:00"),
        "S3": ("11:30", "20:00"), "S4": ("12:00", "20:00"), "N": ("22:00", "06:00")}},
}


def _group(items, key_top, tol=3.0):
    rows = []
    for it in sorted(items, key=lambda w: (key_top(w), w["x0"])):
        if rows and abs(rows[-1]["top"] - key_top(it)) <= tol:
            rows[-1]["items"].append(it)
        else:
            rows.append({"top": key_top(it), "items": [it]})
    for r in rows:
        r["items"].sort(key=lambda w: w["x0"])
    return rows


def _chars_text(chars):
    out, prev = [], None
    for c in sorted(chars, key=lambda c: c["x0"]):
        if c["text"].isspace():
            continue
        if prev is not None and c["x0"] - prev["x1"] > 1.2:
            out.append(" ")
        out.append(c["text"])
        prev = c
    return "".join(out).strip()


def net_hours(start, end):
    """Netto-Arbeitszeit einer Schicht (Pause 30 min ab 6 h, 45 min ab 9 h)."""
    sh, sm = map(int, start.split(":"))
    eh, em = map(int, end.split(":"))
    gross = ((eh * 60 + em) - (sh * 60 + sm)) % (24 * 60) / 60
    pause = 0.75 if gross > 9 else (0.5 if gross > 6 else 0)
    return round(gross - pause, 2)


def _detect_company(text):
    t = text.lower()
    if "fiege" in t or "air cargo" in t or "facl" in t:
        return "FACL"
    if re.search(r"(?<![a-z])lug(?![a-z])", t):
        return "LUG"
    if "handling" in t or "counts" in t:
        return "HC"
    return None


def _legend_shifts(rows_chars):
    """Liest Legendenzeilen wie 'S3  11:30 Uhr - 20:00 Uhr' oder '1  06:55 Uhr  14:59 Uhr'.
    Das Kürzel steht direkt links der Startzeit; Kleinbuchstaben (überlappender Text wie 'Woche') werden ignoriert.
    Nicht lesbare Zeilen werden übersprungen."""
    found = {}
    for chars in rows_chars:
        cs = [c for c in sorted(chars, key=lambda c: c["x0"]) if not c["text"].isspace()]
        flat = "".join(c["text"] for c in cs)
        m = re.search(r"(\d{2}:\d{2})Uhr-?(\d{2}:\d{2})Uhr", flat)
        if not m:
            continue
        x_start = cs[m.start()]["x0"]
        code = "".join(c["text"] for c in cs[:m.start()] if x_start - 14 <= c["x0"] < x_start
                       and (c["text"].isdigit() or c["text"].isupper()))
        if re.fullmatch(r"[A-Z]{1,3}\d?|\d", code):
            found[code] = (m.group(1), m.group(2))
    return found


def parse_plan(path):
    with pdfplumber.open(path) as pdf:
        pages = []
        for page in pdf.pages:
            pages.append({
                "words": page.extract_words(keep_blank_chars=False, use_text_flow=False, x_tolerance=1.5, y_tolerance=2),
                "chars": page.chars, "text": page.extract_text() or ""})

    first_text = pages[0]["text"]
    m = re.search(r"Monat\s+([A-Za-zÄÖÜäöüß]+)\.?\s*'?(\d{2,4})", first_text)
    if not m:
        raise ValueError("Monat im Dienstplan nicht gefunden (Zeile 'Monat Okt'26').")
    month = MONTHS.get(m.group(1).lower()[:3])
    if not month:
        raise ValueError(f"Unbekannter Monat: {m.group(1)}")
    year = int(m.group(2))
    year += 2000 if year < 100 else 0
    company = _detect_company(first_text[:200])
    title = first_text.splitlines()[0].strip() if first_text else ""

    people, legend_rows, days_in_plan = [], [], 0
    for pg in pages:
        words, chars = pg["words"], pg["chars"]
        rows = _group(words, lambda w: w["top"])
        best = None
        for r in rows:
            nums = [w for w in r["items"] if w["text"].isdigit() and 1 <= int(w["text"]) <= 31]
            if len(nums) >= 28 and (best is None or len(nums) > len(best[1])):
                best = (r, nums)
        if not best:
            continue
        hdr_row, day_words = best
        days = {int(w["text"]): (w["x0"] + w["x1"]) / 2 for w in day_words}
        nums_sorted = sorted(days)
        days_in_plan = max(days_in_plan, nums_sorted[-1])
        centers = [days[d] for d in nums_sorted]
        pitch = (centers[-1] - centers[0]) / (len(centers) - 1)
        first_x = centers[0] - pitch / 2
        schule_w = next((w for w in words if w["text"] == "Schule"), None)
        name_limit = (schule_w["x0"] - 1) if schule_w else first_x - 120
        school_limit = first_x - 1

        def day_for(x):
            if x < first_x:
                return None
            d = int(round((x - centers[0]) / pitch)) + nums_sorted[0]
            return d if d in days else None

        gewerk, last_person_top = "", None
        for r in rows:
            if r["top"] <= hdr_row["top"] + 14:
                continue
            ws = r["items"]
            row_chars = [c for c in chars if abs(c["top"] - r["top"]) <= 3.0]
            row_text = " ".join(w["text"] for w in ws)
            name = _chars_text([c for c in row_chars if c["x0"] < name_limit])
            school = _chars_text([c for c in row_chars if name_limit <= c["x0"] < school_limit])
            if "+" not in school:
                if last_person_top is None and ws[0]["x0"] < name_limit and name:
                    gewerk = row_text
                elif last_person_top is not None:
                    legend_rows.append(row_chars)
                continue
            if not name:
                continue
            cells = {}
            for w in ws:
                if w["x0"] >= school_limit - 2:
                    d = day_for((w["x0"] + w["x1"]) / 2)
                    if d is not None:
                        cells[d] = (cells.get(d, "") + w["text"]).upper()
            last_person_top = r["top"]
            people.append({"name": name, "school": school, "gewerk": re.sub(r"\s*\(.*$", "", gewerk),
                           "days": {str(d): c for d, c in sorted(cells.items())}})
    if not people:
        raise ValueError("Keine Azubi-Zeilen im Dienstplan gefunden.")

    shifts = dict(COMPANIES[company]["shifts"]) if company else {}
    legend = _legend_shifts(legend_rows)
    shifts.update(legend)
    present = {c for p in people for c in p["days"].values()}
    if "S/L" in present and "S" not in shifts:
        shifts["S"] = ("12:30", "21:00")
    shifts_out = {c: {"start": s, "end": e, "net": net_hours(s, e)} for c, (s, e) in shifts.items()}
    known = set(shifts_out) | STATIC_CODES
    unknown = sorted(c for c in present if c not in known)
    return {"month": month, "year": year, "daysInPlan": days_in_plan, "people": people,
            "unknownCodes": unknown, "company": company, "companyLabel": COMPANIES[company]["label"] if company else title,
            "title": title, "shifts": shifts_out, "legendRead": sorted(legend)}


if __name__ == "__main__":
    import json, sys
    res = parse_plan(sys.argv[1])
    print(res["companyLabel"], res["month"], res["year"], res["daysInPlan"], len(res["people"]))
    print("shifts", {k: (v["start"], v["end"], v["net"]) for k, v in res["shifts"].items()}, "legend gelesen:", res["legendRead"])
    print("unbekannt:", res["unknownCodes"])
    for p in res["people"][:int(sys.argv[2]) if len(sys.argv) > 2 else 3]:
        print(p["gewerk"], "|", p["name"], "|", p["school"], json.dumps(p["days"]))
