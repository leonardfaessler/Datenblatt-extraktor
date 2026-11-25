"""
Datenblatt-Extraktor
Liest PDF-Datenblätter von Elektromotoren, lässt Claude die Kennwerte
herausziehen und speichert alles in einer Excel-Tabelle.
"""

import base64
import json
import math
from pathlib import Path

import anthropic
from openpyxl import Workbook

# ---------------------------------------------------------------
# 1) Einstellungen
# ---------------------------------------------------------------
ORDNER = Path("datenblaetter")          # hier liegen die PDFs
AUSGABE = Path("motordaten.xlsx")       # hierhin kommt die Tabelle
MODELL = "claude-sonnet-5-5"            # genauer als Haiku, immer noch günstig

# ---------------------------------------------------------------
# 2) Das Schema: welche Felder wir pro Motorvariante wollen
#    (Einheiten sind fest vorgegeben -> Claude muss umrechnen)
# ---------------------------------------------------------------
FELDER = {
    "hersteller":            ("string", "z.B. maxon, Faulhaber"),
    "baureihe":              ("string", "Motorfamilie/Typ, z.B. 'EC-i 40'"),
    "artikelnummer":         ("string", "Bestellnummer oder Typbezeichnung dieser Variante, z.B. '449469', '2232S024BX4SC', 'DB42M03'. KEINE Zeichnungs-, Dokument- oder Revisionsnummern aus dem Schriftfeld."),
    "nennspannung_V":        ("number", "Nennspannung in Volt"),
    "leistung_hersteller_W": ("number", "Leistungsangabe des Herstellers in Watt (z.B. '70 watt' im Titel oder 'Assigned power rating'), so wie sie im Datenblatt steht"),
    "leerlaufdrehzahl_rpm":  ("number", "Leerlaufdrehzahl in U/min"),
    "nenndrehzahl_rpm":      ("number", "Nenndrehzahl in U/min"),
    "nennmoment_mNm":        ("number", "Nenn-/Dauerdrehmoment in mNm"),
    "nennstrom_A":           ("number", "Nenn-/Dauerstrom in Ampere"),
    "max_wirkungsgrad_pct":  ("number", "Maximaler Wirkungsgrad in Prozent"),
    "gewicht_g":             ("number", "Motorgewicht in Gramm"),
    "bemerkung":             ("string", "Unsicherheiten, Umrechnungen, Besonderheiten"),
}

# Daraus bauen wir ein "Werkzeug", das Claude ausfüllen muss.
# So bekommen wir garantiert sauberes JSON statt freiem Text.
WERKZEUG = {
    "name": "motordaten_speichern",
    "description": "Speichert die Kennwerte aller Motorvarianten aus einem Datenblatt.",
    "input_schema": {
        "type": "object",
        "properties": {
            "varianten": {
                "type": "array",
                "description": "Eine Zeile pro Motorvariante (z.B. pro Spannung/Artikelnummer).",
                "items": {
                    "type": "object",
                    "properties": {
                        name: {"type": [typ, "null"], "description": beschr}
                        for name, (typ, beschr) in FELDER.items()
                    },
                    "required": list(FELDER.keys()),
                },
            }
        },
        "required": ["varianten"],
    },
}

ANWEISUNG = """Du bist Ingenieur und erfasst Motordaten für eine Angebotsdatenbank.
Lies das angehängte Datenblatt und gib ALLE Motorvarianten darin zurück,
jeweils eine Zeile pro Variante (z.B. pro Nennspannung oder Artikelnummer).

Regeln:
- Rechne Werte in die vorgegebenen Einheiten um (z.B. Nm -> mNm, Ncm -> mNm, kg -> g, min-1 = rpm).
- Gleiche Grössen heissen je nach Hersteller anders (z.B. 'Rated torque' = 'Nominal torque').
- Wenn ein Wert nicht im Datenblatt steht: null. Niemals raten oder erfinden.
- Notiere Umrechnungen und Unsicherheiten im Feld 'bemerkung'.
- Artikelnummer: die Nummer, mit der man den Motor bestellt. Gibt es keine Bestellnummer, nimm die Typbezeichnung. Nummern aus dem Zeichnungs-Schriftfeld sind keine Artikelnummern.
- Benutze das Werkzeug motordaten_speichern für die Antwort."""


# ---------------------------------------------------------------
# 3) Ein einzelnes PDF an Claude schicken
# ---------------------------------------------------------------
def datenblatt_auslesen(client, pdf_pfad):
    pdf_base64 = base64.standard_b64encode(pdf_pfad.read_bytes()).decode("utf-8")

    antwort = client.messages.create(
        model=MODELL,
        max_tokens=4000,
        tools=[WERKZEUG],
        tool_choice={"type": "auto"},   # Claude entscheidet selbst; der Prompt verlangt das Werkzeug
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": pdf_base64,
                    },
                },
                {"type": "text", "text": ANWEISUNG},
            ],
        }],
    )

    # Claude antwortet mit einem "tool_use"-Block, darin stecken unsere Daten
    for block in antwort.content:
        if block.type == "tool_use":
            varianten = block.input.get("varianten", [])      # Liste holen, sonst leere Liste
            if isinstance(varianten, str):                    # kam die Liste als Text?
                varianten = json.loads(varianten)             #   -> Text in echte Python-Liste umwandeln
            varianten = [v for v in varianten if isinstance(v, dict)]   # nur echte Datensätze behalten
            return varianten, antwort.usage
    return [], antwort.usage


# ---------------------------------------------------------------
# 4) Hauptprogramm: alle PDFs durchgehen, Excel schreiben
# ---------------------------------------------------------------
def main():
    client = anthropic.Anthropic()   # Key kommt aus ANTHROPIC_API_KEY

    pdfs = sorted(ORDNER.glob("*.pdf"))
    print(f"{len(pdfs)} Datenblätter gefunden.\n")

    alle_zeilen = []
    tokens_in = tokens_out = 0

    for pdf in pdfs:
        print(f"-> {pdf.name} ...", end=" ", flush=True)
        try:
            varianten, usage = datenblatt_auslesen(client, pdf)
        except Exception as fehler:
            print(f"FEHLER: {fehler}")
            continue

        tokens_in += usage.input_tokens
        tokens_out += usage.output_tokens
        for v in varianten:
            v["datei"] = pdf.name
            # Mechanische Nennleistung selbst berechnen: P = M * omega
            m = v.get("nennmoment_mNm")             # Moment in mNm
            n = v.get("nenndrehzahl_rpm")           # Drehzahl in U/min
            if m is not None and n is not None:     # nur rechnen, wenn beide Werte da sind
                omega = n * 2 * math.pi / 60        # U/min -> rad/s
                v["leistung_berechnet_W"] = round(m / 1000 * omega, 1)   # mNm -> Nm, Ergebnis in W
            else:
                v["leistung_berechnet_W"] = None    # fehlt ein Wert -> leer lassen
            alle_zeilen.append(v)
        print(f"{len(varianten)} Varianten")

    # Rohdaten zusätzlich als JSON sichern (praktisch zum Nachschauen)
    Path("motordaten.json").write_text(
        json.dumps(alle_zeilen, indent=2, ensure_ascii=False), encoding="utf-8"
    )

    # Excel schreiben
    wb = Workbook()
    ws = wb.active
    ws.title = "Motoren"
    spalten = ["datei"] + list(FELDER.keys()) + ["leistung_berechnet_W"]
    ws.append(spalten)
    for zeile in alle_zeilen:
        ws.append([zeile.get(s) for s in spalten])
    ws.freeze_panes = "A2"
    wb.save(AUSGABE)

    print(f"\nFertig: {len(alle_zeilen)} Motorvarianten -> {AUSGABE}")
    print(f"Verbrauch: {tokens_in} Input-Tokens, {tokens_out} Output-Tokens")


if __name__ == "__main__":
    main()
