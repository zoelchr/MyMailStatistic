## 📊 MailStatisticTemplate.xlsm

Excel-Arbeitsmappe mit VBA-Makros, die die von `MailStatistic.ps1` exportierten Rohdaten in interaktive Auswertungen verwandelt.

### Tabellenblatt-Struktur
| Sheet | Zweck | Wichtigste Spalten |
|-------|-------|--------------------|
| **Email Statistic** | Ablage aller importierten Mails | `StoreID`, `EntryID`, `Email`, `SentOn`, `Sender`, `BehalfOf`, `Subject`, `Folder`, `Words`, `Recipients`, `comparisonKey`, `Doubletten`, `Mailbox`, `Senderemail` |

### Eingebaute Makros (Alt + F11 zum Einsehen)
- **Pivot-Generator** – erstellt bzw. aktualisiert vorbereitete Pivot-Tabellen:  
  - Zeitreihen für einzelne Absender  
  - Empfang vs. Versand  
  - Durchschnittliche Versand­uhrzeiten  
- **Open Mail** – öffnet die ursprünglich gespeicherte Nachricht in Outlook anhand der `EntryID`.  
- **Duplicate Filter** – kennzeichnet Dubletten, damit sie Statistiken nicht verfälschen.  

### Workflow
1. `MailStatistic.ps1` ausführen – das Skript fügt Datensätze in *Email Statistic* ein.  
2. Arbeitsmappe öffnen und **Makros aktivieren**.  
3. Mit **Ctrl + Shift + R** (oder über *Entwicklertools → Makros*) die Pivot-Generator-Routine starten.  
4. In den Pivot-Berichten nach Zeitraum, Postfach (bis zu 4), Absender usw. filtern.  
5. Zur Original-Mail springen, indem du eine Zeile markierst und das **Open-Mail-Makro** ausführst.  

### Systemvoraussetzungen
- Windows-PC mit installiertem Outlook (Makros nutzen das Outlook-COM-Objekt).  
- Excel 2016 oder neuer, Makroausführung aktiviert.  

> **Hinweis zur Sicherheit:** Die VBA-Projekte sind signiert (Selbstzertifikat *MailStatistic*). Prüfe die Signatur, bevor du Makros zulässt.
