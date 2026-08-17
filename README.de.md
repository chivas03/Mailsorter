# Mailsorter
Ein leichtgewichtinales, effizientes Tool zur Optimierung und Organisation von Outlook-Postfächern und Ordnerstrukturen. Nach dem Prinzip "Finden anstatt Suchen": Ordner per Live-Filter blitzschnell filtern, E-Mails mit einem Klick einsortieren & den Verlauf nutzen.


## Schluss mit dem ewigen Scrollen: Warum ich dieses Tool entwickelt habe

Da ich kein professioneller Entwickler bin, baue ich Tools in meiner Freizeit mit der Hilfe von AI, um echte Hürden im Büroalltag zu lösen. Dieses Projekt begann vor langer Zeit als einfache TreeView zum Sortieren von E-Mails. Mit der Zeit hat es sich zu einem funktionsreichen Produktivitätshelfer weiterentwickelt.

### Das Problem: Das Chaos tiefer Ordnerstrukturen
Jeder, der mit einer großen E-Mail-Flut in Microsoft Outlook arbeitet, kennt den Kampf:
* **Endlose Klick-Chaos:** Das Navigieren durch hunderte verschachtelte Unterordner, Projekte und Archive kostet zu viele Klicks und zerstört die Konzentration.
* **Träge Outlook-Suche:** Die integrierte Outlook-Suche oder das manuelle Durchstöbern des Baums ist oft langsam, umständlich und unterbricht den Arbeitsfluss.
* **Überlastung durch Shared Mailboxes:** In professionellen Umgebungen mit mehreren Konten oder freigegebenen Postfächern wird man mit irrelevanten Ordnern überflutet.

### Die Lösung: "Finden anstatt Suchen"
Nach dem Motto **"Finden anstatt Suchen"** verwandelt dieses Tool die mühsame Pflicht des E-Mail-Einsortierens in einen blitzschnellen Reflex:
* **Live-Suchfilter:** Einen Teil des Ordnernamens eintippen, Enter drücken, und das Verzeichnis filtert sich in Millisekunden.
* **Verlauf der letzten Ordner:** Speichert automatisch die letzten 20 besuchten Ordner, sodass folgende E-Mails mit nur einem Klick einsortiert werden können.
* **Root-Ordner-Filter:** Alles ausblenden, was nicht benötigt wird, um den Arbeitsbereich aufgeräumt zu halten.

---

## Auf Geschwindigkeit ausgelegt: Funktionen, die tausende Klicks sparen

* **Sofortiges Ordner-Caching:** Scannt und cacht die Verzeichnisstruktur einmal beim Start für eine rasante Navigation ohne Outlook-Verzögerungen.
* **Live-Verzeichnissuche:** Direkt tippen, um Zielordner sofort zu filtern und zu finden.
* **Verlauf (Recent Folders):** Behält die zuletzt verwendeten Ordner im Blick, um Mails mit einem Klick abzulegen.
* **Store-Verwaltung:** Steuern, welche Postfächer (Stores) aktiv sind, und benutzerdefinierte Expansionsebenen oder erlaubte Root-Ordner pro Konto definieren.
* **Sichere Fehlerbehandlung:** Robuste Verbindungstests zur reibungslosen Handhabung von Shared Mailboxes und Offline-/Archiv-Speichern (`store.IsOpen`).
* **Vielseitige Modi:** Ausgewählte E-Mails verschieben oder kopieren – oder das Tool rein als schnellen Navigations-Launcher ("Zu Ordner wechseln") nutzen.
* **Unterordner-Erstellung:** Neue Unterordner direkt aus dem Formular heraus im Handumdrehen erstellen.
* **Getestete Kompatibilität:** Vollständig kompatibel mit **Microsoft Outlook 2024 LTSC** auf Windows 11 64-Bit-Systemen (und kompatiblen Desktop-Versionen, die VBA unterstützen).

---

## In 5 Minuten startklar: Installations- & Sicherheitsleitfaden

1. **Repository klonen oder herunterladen**.
2. **Outlook-Sicherheitseinstellungen anpassen:** 
   * Gehe zu **Datei** -> **Optionen** -> **Trust Center** -> **Trust Center-Einstellungen** -> **Makroeinstellungen**.
   * Wähle **"Benachrichtigungen für alle Makros"**.
3. Microsoft Outlook öffnen und `Alt + F11` drücken, um den VBA-Editor zu öffnen.
4. **Dateien importieren:** Gehe auf `Datei` -> `Datei importieren...` und wähle die Dateien aus dem `/src`-Ordner aus (`.bas`, `.frm` und `.frx` Dateien).
5. **TreeView-Referenz aktivieren:** 
   * Im VBA-Editor auf **Extras** -> **Verweise** gehen.
   * Sicherstellen, dass **Microsoft Windows Common Controls 6.0 (SP6)** aktiviert ist.
6. **Makro-Signatur & Sicherheitshinweis:**
   * Da dieser Code von GitHub heruntergeladen wurde, zeigt Outlook beim ersten Ausführen eine Sicherheitswarnung. Bei Aufforderung einfach erlauben/aktivieren.
   * *Fortgeschritten (Selbstsignierung):* Falls es die Umgebung erfordert, kann das Windows-Tool `SELCRT.EXE` genutzt werden, um ein persönliches Zertifikat zu erstellen und das Projekt via **Extras -> Digitale Signatur** zu signieren.
7. **Makro zuweisen:** Weise das Makro `RunFolderPicker` aus dem Modul `modEmailSorter` einer Schnellstart-Symbolleiste oder einem benutzerdefinierten Shortcut zu.

---

## So meisterst du "Finden anstatt Suchen" im Alltag

1. **E-Mails auswählen:** E-Mails in Outlook auswählen und das Makro starten, um den **Verschieben/Kopieren-Modus** zu öffnen. Ohne ausgewählte E-Mails öffnet sich der **Navigations-Modus** ("Zu Ordner wechseln").
2. **Navigieren & Suchen:** Den Baum aufklappen oder in das Filterfeld klicken, um einen Teil des Ordnernamens einzutippen und sofort zu filtern.
3. **E-Mails einsortieren:** Doppelklick auf einen Ordner oder die Auktions-Buttons nutzen, um sofort zu verschieben/kopieren. Das Verlaufspanel für häufige Ziele verwenden.

---

## Feedback, Bugs & Mitwirken

Da es sich hierbei um ein persönliches Projekt handelt, sind Feedback, Bug-Meldungen und Optimierungsideen sehr willkommen! Bitte öffne ein **Issue** oder reiche einen **Pull Request** ein.

---

## Das Projekt unterstützen

Wenn dir dieses Tool den Arbeitsalltag erleichtert und Zeit spart, freue ich mich riesig! Wer die Weiterentwicklung und Pflege dieses Projekts unterstützen möchte, darf gerne einen Freund einen Kaffee spendieren.

---

## Lizenz

Dieses Projekt ist Open-Source und sowohl für **private als auch kommerzielle Zwecke** frei nutzbar, unter einer einzigen Bedingung: **Sie dürfen diese Software in keiner Form kommerzialisieren, bündeln oder die Idee bzw. den Code verkaufen.**
