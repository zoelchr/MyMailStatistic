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

# 2) Skript ausführen (Beispiel)
powershell -ExecutionPolicy Bypass -File .\MailStatistic.ps1 `
  -ExcelTemplate .\MailStatisticTemplate.xlsm `
  -OutDir .\out `
  -YearsBack 1 `
  -Mailboxes "Postfach A","shared@contoso.com"
```

Die Ergebnisdatei `out\\MailStatistic_YYYYMMDD_HHmm.xlsx` öffnet sich automatisch (ansonsten doppelklicken).

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
| `-Mailboxes`        | *string\[]* | `@('Postfach A')`            | Ein bis **vier** Postfach‑Anzeige­namen oder SMTP‑Adressen.                                     |
| `-NoMailboxQuery`   | *switch*    | `False`                      | Interaktive Postfachauswahl überspringen; nur `Mailboxes` verwenden.                            |
| `-NoProgress`       | *switch*    | `False`                      | Fortschrittsbalken unterdrücken.                                                                |
| `-NoConsoleLogging` | *switch*    | `True`                       | Ausführliche Konsolenausgabe stummschalten.                                                     |
| `-FileLogging`      | *switch*    | `False`                      | Debug‑Log (`log.txt`) im `OutDir` schreiben.                                                    |
| `-Testing`          | *switch*    | `False`                      | Lauf auf ≈ 40 Nachrichten begrenzen (Schnelltest).                                              |

> ℹ️ **Relativer versus absoluter Zeitraum**
> Wenn `StartDate`/`EndDate` gesetzt sind, werden `YearsBack` und `MonthBack` ignoriert.

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
