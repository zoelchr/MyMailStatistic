# MailStatistic

Automatisierte **Outlook‑Postfachanalysen** mit PowerShell & Excel.

<div align="center">
  📬 + 🛠️ &nbsp;→&nbsp; 📊
</div>

---

## 📦 Repository‑Inhalt

| Datei                          | Zweck                                                                                                 |
| ------------------------------ | ----------------------------------------------------------------------------------------------------- |
| **MailStatistic.ps1**          | Sammelt Metadaten aus ausgewählten Outlook‑Postfächern und exportiert sie in eine Excel‑Arbeitsmappe. |
| **MailStatisticTemplate.xlsm** | Makro‑fähige Arbeitsmappe, die die Rohdaten in interaktive Dashboards verwandelt.                     |
| **mailboxes.psd1** *(optional)* | Konfigurationsdatei für benannte Postfächer (→ siehe `MailboxMapFile`).                               |

---

## ⚙️ Voraussetzungen

* Windows‑Rechner mit Microsoft Outlook (getestet mit Outlook 2019 / Microsoft 365)
* Excel 2016 oder neuer **mit aktivierten Makros**
* PowerShell 5.1 oder höher (funktioniert in Windows PowerShell und PowerShell 7)

---

## 🚀 Schnellstart

```powershell
# 1) Repository klonen oder ZIP herunterladen
git clone https://github.com/<org>/MailStatistic.git
cd MailStatistic

# 2) Analyse starten mit vorkonfigurierten Postfächern
powershell -ExecutionPolicy Bypass -File .\MailStatistic.ps1 `
  -MailboxMapFile .\mailboxes.psd1 `
  -YearsBack 1 `
  -FileLogging
```

Die Ergebnisdatei `out\MailStatistic_YYYYMMDD_HHmm.xlsx` öffnet sich automatisch (ansonsten doppelklicken).

---

## 📝 Skriptparameter

| Parameter           | Typ         | Standard                     | Beschreibung                                                                                    |
| ------------------- | ----------- | ---------------------------- | ----------------------------------------------------------------------------------------------- |
| `-ExcelTemplate`    | *string*    | `MailStatisticTemplate.xlsm` | Pfad zur Makro‑Arbeitsmappe, die die Daten aufnehmen soll.                                      |
| `-OutDir`           | *string*    | `./out`                      | Ordner für die datierte Ausgabedatei und optionale Logs.                                        |
| `-StartDate`        | *datetime*  | *(none)*                     | Analyse‐Start (überschreibt `YearsBack`/`MonthBack`).                                           |
| `-EndDate`          | *datetime*  | *now*                        | Analyse‐Ende.                                                                                   |
| `-YearsBack`        | *int*       | `0`                          | E‑Mails bis zu *n* Jahre rückwirkend berücksichtigen.                                           |
| `-MonthBack`        | *int*       | `1`                          | E‑Mails bis zu *n* Monate rückwirkend berücksichtigen (ignoriert, falls `StartDate` angegeben). |
| `-MailboxMapFile`   | *string*    | *(none)*                     | Optional: Pfad zu einer `.psd1`-Datei mit benannten Postfächern (→ siehe unten).                |
| `-NoMailboxQuery`   | *switch*    | `False`                      | Interaktive Postfachauswahl überspringen; nur `Mailboxes` verwenden.                            |
| `-NoProgress`       | *switch*    | `False`                      | Fortschrittsbalken unterdrücken.                                                                |
| `-NoConsoleLogging` | *switch*    | `True`                       | Ausführliche Konsolenausgabe stummschalten.                                                     |
| `-FileLogging`      | *switch*    | `False`                      | Debug‑Log (`log.txt`) im `OutDir` schreiben.                                                    |
| `-Testing`          | *switch*    | `False`                      | Lauf auf ≈ 40 Nachrichten begrenzen (Schnelltest).                                              |

> ℹ️ `-MailboxMapFile` und `-Mailboxes` schließen sich nicht aus. Es können auch beide genutzt werden.

---

## 🔧 Erweiterte Konfiguration: `mailboxes.psd1`

Für häufig wiederkehrende oder benannte Postfächer kann optional eine Konfigurationsdatei verwendet werden:

**Beispiel: `mailboxes.psd1`**
```powershell
@{
  'WInS-Projekt (LGL)'     = 'WInS-Projekt@lgl.bayern.de'
  'Shapth-Projekt (LGL)'   = 'Shapth-Projekt@lgl.bayern.de'
  'twfa-projekt (LGL)'     = 'twfa-projekt@lgl.bayern.de'
  'SHAPTH Tickets (LGL)'   = 'SHAPTH-Tickets@lgl.bayern.de'
  'WAFA-Tickets (LGL)'     = 'WAFA-Tickets@lgl.bayern.de'
}
```

Diese Datei kann dann über den Parameter `-MailboxMapFile` übergeben werden:

```powershell
powershell -File MailStatistic.ps1 -MailboxMapFile .\mailboxes.psd1 -YearsBack 1
```

Die Anzeige der Postfächer in der Auswertung erfolgt anhand der konfigurierten Namen.

---

## 📊 Excel‑Template‑Highlights

* **Email Statistic**‑Blatt: Rohdaten + Hyperlink‑Spalte → Original‑Mail in Outlook öffnen
* **Pivot‑Dashboards** (Shortcut **Ctrl + Shift + R**):

  * Verlauf je Absender
  * Empfang vs. Versand
  * Typische Versanduhrzeiten
* **Duplicate Filter**‑Makro: kennzeichnet Dubletten, damit Pivot‑Ergebnisse sauber bleiben
* **Open Mail**‑Makro: springt aus jeder Zeile direkt zur Nachricht in Outlook

> **Sicherheitshinweis**
> Das VBA‑Projekt ist mit dem Selbstzertifikat **MailStatistic** signiert. Signatur vor Aktivierung der Makros prüfen.

---

## 🧹 Pflegehinweise

* Ergebnisse sammeln sich im `OutDir`; alte Reports können gefahrlos gelöscht werden.
* Sehr große Postfächer? Erst mit `-Testing` einen Probelauf durchführen, um Laufzeit & Layout zu prüfen.

---

## 📜 Lizenz

[MIT](LICENSE)

## 🤝 Beiträge

Pull‑Requests und Issues sind willkommen!
Bitte verwende [Conventional Commits](https://www.conventionalcommits.org/) für Commit‑Nachrichten.

---
