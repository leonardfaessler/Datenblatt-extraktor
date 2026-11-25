# Datenblatt-Extraktor für Elektromotoren

Ein Python-Tool, das technische Kennwerte aus PDF-Datenblättern von Elektromotoren mit einem Large Language Model (Claude) ausliest, in einheitliche Einheiten umrechnet und als strukturierte Tabelle (Excel + JSON) speichert. Es ist als Baustein für eine Angebotsdatenbank für Antriebssysteme gedacht.

## Problem

Motordatenblätter sind je nach Hersteller völlig unterschiedlich aufgebaut:

- verschiedene Layouts (Tabellen mit mehreren Varianten pro Seite, Einzelblätter, Flyer)
- verschiedene Begriffe für dieselbe Grösse (*Nominal torque* vs. *Rated torque*)
- verschiedene Einheiten (mNm, Ncm, Nm; rpm vs. min⁻¹; g vs. kg)

Klassisches Auslesen per Text-Extraktion scheitert daran: Der Rohtext eines Datenblatts ist meist ein Durcheinander aus Zahlen ohne erkennbare Zuordnung zu ihren Beschriftungen.

## Lösung

```
PDF-Ordner ──► Claude (liest PDF inkl. Tabellenlayout)
                 │  füllt ein festes Schema aus (Tool Use / JSON)
                 ▼
           Validierung in Python ──► Berechnungen in Python ──► Excel + JSON
```

1. **PDF direkt an das Modell:** Das PDF wird als Dokument übergeben, sodass das Modell die Tabellenstruktur sieht und nicht nur den Rohtext.
2. **Festes Schema über Tool Use:** Das Modell füllt ein vordefiniertes Formular mit 12 Feldern in festen Einheiten aus, mit einer Zeile pro Motorvariante.
3. **Klare Regeln im Prompt:** Einheiten umrechnen, synonyme Begriffe erkennen, fehlende Werte als `null` lassen (**nie raten**), Unsicherheiten im Feld `bemerkung` dokumentieren.
4. **Validierung im Code:** Die KI-Antwort wird auf das erwartete Format geprüft und bei Bedarf korrigiert. Fehler bei einzelnen PDFs brechen den Lauf nicht ab.
5. **Berechnungen deterministisch in Python:** Die mechanische Nennleistung *P = M · 2πn/60* wird im Code berechnet, nicht von der KI.
6. **Nachvollziehbarkeit:** Jede Zeile enthält die Quelldatei, damit jeder Wert gegen das Original geprüft werden kann.

## Ergebnis

Getestet mit 6 Datenblättern von 4 Herstellern (maxon, Faulhaber, Nanotec, Dunkermotoren) → **14 Motorvarianten**. Alle stichprobenweise gegen die Original-PDFs geprüften Werte stimmen, inklusive Einheitenumrechnungen (z.B. 0,125 Nm → 125 mNm).

| Hersteller | Typ | Spannung | Nennmoment | Leistung lt. Hersteller | Leistung berechnet |
|---|---|---|---|---|---|
| Nanotec | DB42M03 | 24 V | 125 mNm | 52 W | 52,4 W |
| maxon | EC-i 40 (449470) | 36 V | 84,8 mNm | 70 W | 77,4 W |
| Faulhaber | 2232S012BX4SC | 12 V | 17 mNm | 10,5 W | 7,3 W |
| maxon | EC-max 16 (283828) | 12 V | 3,23 mNm | 5 W | 2,0 W |

Die vollständigen Resultate stehen in [`motordaten.xlsx`](motordaten.xlsx) und [`motordaten.json`](motordaten.json).

## Erkenntnisse

- **«Nennleistung» ist nicht gleich «Nennleistung».** Bei Nanotec und Dunkermotoren entspricht die Herstellerangabe der mechanischen Leistung im Nennpunkt. Bei maxon ist sie eine *Leistungsklasse* der Baureihe («Assigned power rating»): Beim EC-max 16 liegt die tatsächliche Nennleistung bei nur ~40 % der angegebenen 5 W. Deshalb speichert das Tool beide Werte getrennt. Ein Filter nur nach Herstellerangabe würde sonst ungeeignete Motoren vorschlagen.
- **KI-Ausgaben müssen validiert werden.** In einem Durchlauf kam die Variantenliste einmal als Text statt als Liste zurück, was ohne Prüfung zum Absturz führte. Seither wird das Format im Code geprüft.
- **Allgemeine Regeln statt Einzelfall-Korrekturen.** Eine Zeichnungsnummer aus dem Schriftfeld wurde anfangs als Artikelnummer erkannt. Gelöst wurde das mit einer allgemeinen Regel und Beispielen im Schema (Few-Shot), nicht mit einer Sonderregel für einen Hersteller.

## Grenzen

- Getestet mit einer kleinen Stichprobe; grosse Kataloge mit Dutzenden Motoren pro PDF wurden noch nicht systematisch geprüft.
- Die Extraktion ist nicht 100 % deterministisch, kritische Werte sollten stichprobenartig geprüft werden.
- Flyer ohne Bestellnummer liefern nur die Typbezeichnung als Artikelnummer.

## Mögliche Erweiterungen

- Abfrage-Agent: Fragen in natürlicher Sprache («Welcher Motor schafft 80 W bei 24 V?») über die erzeugte Tabelle
- Automatische Plausibilitätsprüfung (z.B. Nennmoment < Anhaltemoment, berechnete vs. angegebene Leistung)
- Speicherung in einer echten Datenbank statt Excel

## Nutzung

```bash
pip install -r requirements.txt
export ANTHROPIC_API_KEY="..."      # eigenen API-Key setzen
# PDF-Datenblätter in den Ordner datenblaetter/ legen
python3 extraktor.py
```

Die Datenblätter selbst sind aus urheberrechtlichen Gründen nicht im Repository. Verwendet wurden öffentlich verfügbare Datenblätter von maxon (EC-i 40 70 W, EC-i 40 100 W High Torque, EC-max 16), Faulhaber (2232 BX4 SC), Nanotec (DB42M03) und Dunkermotoren (BG 31).

## Technik

Python · Anthropic Claude API (PDF-Input, Tool Use) · openpyxl

---

*Entwickelt von Leonard Fässler (Maschinenbau, ETH Zürich) mit KI-Unterstützung als Lernprojekt zur Anwendung von LLMs im Engineering.*
