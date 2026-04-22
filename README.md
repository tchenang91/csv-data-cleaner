# CSV Data Cleaner → Excel (`Erzeugte Datei.xlsx`)

Dieses Projekt enthält ein Python-Skript, das CSV-Dateien bereinigt und als formatierte Excel-Datei exportiert.

## Datei

- `clean_csv_to_xlsx.py`

## Was das Skript macht

- Liest **eine oder mehrere CSV-Dateien** ein.
- Trennt Zeilen korrekt in Spalten, auch wenn eine komplette Zeile fälschlich in einer Zelle steht (komma-separiert).
- Entfernt Datensätze mit leerer `user_id`.
- Sortiert alle Datensätze alphabetisch nach `user_id`.
- Formatiert Inhalte:
  - `user_id` wird in **UPPERCASE** geschrieben.
  - Andere Textwerte: nur erster Buchstabe groß.
- Erstellt eine Excel-Datei mit:
  - Dateiname: **`Erzeugte Datei.xlsx`**
  - Tabellenblatt: **`Erzeugte Datei`**
- Excel-Formatierung:
  - Alle Zellen zentriert
  - Alle Zellen mit Rahmenlinien
  - Erste Zeile grau hinterlegt + fett
  - Spalten `price` und `final_price` als Buchhaltungsformat

## Voraussetzungen

- Python 3.10+
- Paket `openpyxl`

Installation:

```bash
pip install openpyxl
```

## Verwendung

### 1) Eine CSV-Datei

```bash
python clean_csv_to_xlsx.py input.csv
```

### 2) Mehrere CSV-Dateien zusammenführen

```bash
python clean_csv_to_xlsx.py file1.csv file2.csv file3.csv
```

### 3) Eigener Ausgabename

```bash
python clean_csv_to_xlsx.py input.csv -o "Erzeugte Datei.xlsx"
```

## Beispielablauf

1. CSV-Datei(en) bereitstellen.
2. Skript ausführen.
3. Ergebnis in `Erzeugte Datei.xlsx` prüfen.

## Hinweise

- Die Spalte `user_id` muss im Header vorhanden sein (z. B. `user_id`, `userid`, `user-id`).
- Falls `price`/`final_price` nicht existieren, wird deren Zahlenformatierung einfach übersprungen.
