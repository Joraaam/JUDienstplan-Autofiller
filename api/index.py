"""Dienstplan Autofill - zustandslose API (Vercel Serverless / lokal).

Jede Anfrage ist eigenstaendig, es wird nichts auf dem Server gespeichert:
  POST /api/plan     PDF            -> gelesener Dienstplan
  POST /api/inspect  Excel + Plan   -> Zuordnung und Zustand der Datei
  POST /api/fill     Excel + Params -> ausgefuellte Excel (base64) + Bericht
"""
import base64
import io
import json
import os
import socket
import sys
import threading
import webbrowser
from calendar import monthrange

ROOT = getattr(sys, "_MEIPASS", os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from flask import Flask, jsonify, request, send_from_directory  # noqa: E402

from core import excel_fill as ef  # noqa: E402
from core import matching, plan_parser  # noqa: E402

PUBLIC = os.path.join(ROOT, "public")
MONTH_NAMES = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August",
               "September", "Oktober", "November", "Dezember"]

app = Flask(__name__, static_folder=None)
app.config["MAX_CONTENT_LENGTH"] = 4 * 1024 * 1024  # Vercel erlaubt max. 4,5 MB pro Anfrage


def err(message, status=400):
    return jsonify(error=message), status


@app.errorhandler(413)
def too_big(_):
    return err("Datei zu groß (max. 4 MB pro Anfrage).", 413)


@app.get("/")
def index():  # auf Vercel liefert das CDN public/index.html, lokal übernimmt das hier
    return send_from_directory(PUBLIC, "index.html")


@app.get("/api/health")
def health():
    return jsonify(ok=True)


@app.post("/api/plan")
def plan():
    f = request.files.get("plan")
    if not f or not f.filename.lower().endswith(".pdf"):
        return err("Bitte den Dienstplan als PDF hochladen.")
    try:
        p = plan_parser.parse_plan(io.BytesIO(f.read()))
    except Exception as e:  # noqa: BLE001
        return err(f"Dienstplan konnte nicht gelesen werden: {e}", 422)
    return jsonify(month=p["month"], year=p["year"], monthName=MONTH_NAMES[p["month"] - 1],
                   daysInMonth=monthrange(p["year"], p["month"])[1], people=p["people"],
                   unknownCodes=p["unknownCodes"], shifts=p["shifts"], company=p["company"],
                   companyLabel=p["companyLabel"])


@app.post("/api/inspect")
def inspect_file():
    f = request.files.get("file")
    if not f:
        return err("Keine Datei.")
    try:
        month, year = int(request.form["month"]), int(request.form["year"])
        names = json.loads(request.form.get("people", "[]"))
    except (KeyError, ValueError):
        return err("Ungültige Parameter.")
    fname = os.path.basename(f.filename.replace("\\", "/"))
    entry = {"filename": fname}
    if not fname.lower().endswith(".xlsx"):
        entry.update(status="error", message="Nur .xlsx-Dateien werden unterstützt.")
        return jsonify(entry)
    try:
        nach, vor = ef.parse_name(fname)
        info = ef.inspect(io.BytesIO(f.read()), month)
        if not nach and info["name"]:
            parts = info["name"].split()
            vor, nach = " ".join(parts[:-1]), parts[-1]
        idx, cands = matching.match_person(nach, vor, [{"name": n} for n in names])
        entry.update(nachname=nach, vorname=vor, match=idx, candidates=cands, info=info, status="ok")
        if not info["sheetFound"]:
            entry.update(status="error", message=f"Blatt '{info['sheet']}' fehlt in der Datei.")
        elif info["year"] and info["year"] != year:
            entry["yearWarning"] = f"Datei ist für {info['year']}, Dienstplan für {year}."
    except Exception as e:  # noqa: BLE001
        entry.update(status="error", message=f"Datei nicht lesbar: {e}")
    return jsonify(entry)


@app.post("/api/fill")
def fill():
    f = request.files.get("file")
    if not f:
        return err("Keine Datei.")
    try:
        params = json.loads(request.form["params"])
        month, year = int(params["month"]), int(params["year"])
    except (KeyError, ValueError):
        return err("Ungültige Parameter.")
    fname = os.path.basename((params.get("filename") or f.filename).replace("\\", "/"))
    try:
        nach, vor = ef.parse_name(fname)
        out = io.BytesIO()
        report = ef.fill(io.BytesIO(f.read()), out, month, year, params.get("days", {}),
                         (vor + " " + nach).strip(), bool(params.get("overwrite", True)), params.get("shifts") or {})
        return jsonify(filename=fname, report=report, data=base64.b64encode(out.getvalue()).decode("ascii"))
    except Exception as e:  # noqa: BLE001
        return err(str(e), 422)


def free_port(start):
    for port in range(start, start + 20):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            if sock.connect_ex(("127.0.0.1", port)) != 0:
                return port
    return start


if __name__ == "__main__":  # lokal starten: python api/index.py
    port = free_port(int(os.environ.get("PORT", "5173")))
    url = f"http://127.0.0.1:{port}"
    print("=" * 60)
    print(" Dienstplan Autofill laeuft:", url)
    print(" Der Browser oeffnet sich automatisch.")
    print(" Zum Beenden dieses Fenster schliessen.")
    print("=" * 60)
    app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024
    threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=port, debug=False)
