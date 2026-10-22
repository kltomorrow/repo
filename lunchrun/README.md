# LunchRun – Mittagsbestellung fürs Büro

Kleine Web-App auf Basis von Google Apps Script und einem Google Sheet.
Kolleg:innen bestellen bis 10:45 Uhr, ein Läufer trägt sich morgens ein und
sieht die gesammelte Liste, die Office-Managerin pflegt das Menü.

- `Code.gs` – Server: liest und schreibt das Sheet, prüft Bestellschluss und PIN
- `index.html` – die Oberfläche fürs Handy

## Einrichten (einmalig, ca. 10 Minuten)

1. **Google Sheet anlegen:** Unter [sheets.new](https://sheets.new) eine neue Tabelle erstellen, z. B. „LunchRun“.
2. **Skript öffnen:** In der Tabelle *Erweiterungen → Apps Script*.
3. **Code einfügen:**
   - Den Inhalt von `Code.gs` in die vorhandene Datei `Code.gs` kopieren (alles Alte ersetzen).
   - Links auf **+ → HTML** klicken, die Datei `index` nennen (ohne `.html`) und den Inhalt von `index.html` einfügen.
   - Speichern.
4. **Einstellungen anpassen:** Oben in `Code.gs` im Block `CONFIG` den Namen des Ladens, die Uhrzeit und die Währung prüfen.
5. **Tabellen anlegen:** Oben in der Funktionsauswahl `setup` wählen und **Ausführen** klicken. Beim ersten Mal fragt Google nach Berechtigungen:
   *Berechtigungen prüfen → Konto wählen → „Erweitert“ → „Zu LunchRun wechseln (unsicher)“ → Zulassen.*
   Die Warnung erscheint, weil das Skript nicht von Google geprüft ist. Es ist dein eigenes.
   Danach gibt es die Blätter **Menu**, **People**, **Orders** und **Runners** mit Beispieldaten.
6. **Daten anpassen:**
   - Blatt **People**: echte Namen der Kolleg:innen eintragen (eine Zeile pro Person).
   - Blatt **Menu**: echtes Menü eintragen (geht auch später in der App).
7. **PIN ändern:** Im Skript-Editor *Projekteinstellungen (Zahnrad) → Skripteigenschaften* den Wert von `MANAGER_PIN` (Standard `1234`) ändern.
8. **Veröffentlichen:** *Bereitstellen → Neue Bereitstellung → Typ: Web-App*
   - Ausführen als: **Ich**
   - Zugriff: **Jeder**
   - Bereitstellen und die URL (endet auf `/exec`) kopieren.
9. **Teilen:** Die URL an alle schicken. Am Handy über *Teilen → Zum Home-Bildschirm* wie eine App ablegen.

**Nach Code-Änderungen:** *Bereitstellen → Bereitstellungen verwalten → Stift → Version: Neue Version → Bereitstellen.* Die URL bleibt gleich.

## So funktioniert es

- **Anmeldung:** Beim ersten Öffnen wählt jede Person ihren Namen. Das Handy merkt ihn sich („Not you?“ wechselt). Mit privaten Gmail-Konten kann das Skript nicht selbst erkennen, wer die Seite öffnet. Die Namenswahl ist deshalb nicht fälschungssicher, fürs Büro aber ausreichend.
- **Bestellen:** eine Bestellung pro Person und Tag; bis zum Bestellschluss änderbar oder stornierbar. Nach 10:45 lehnt der Server Änderungen ab, egal was das Handy anzeigt.
- **Läufer:** Unter *Today's orders* trägt sich die erste Person mit „I’ll do the run today“ ein. Mit „I can’t do it after all“ kann sie sich wieder austragen.
- **Menü bearbeiten:** Tab *Edit menu*, PIN eingeben, Namen/Preise ändern, Artikel als nicht verfügbar markieren oder neue hinzufügen, speichern. Zum Löschen eines Artikels die Zeile im Blatt **Menu** entfernen.
- Alle Bestellungen bleiben im Blatt **Orders** stehen, nach Datum. Das ist zugleich das Archiv.

## Grenzen

- Die PIN schützt nur das Menü-Bearbeiten; sie ist für ein kleines Büro gedacht, nicht für Fremde im Internet.
- Wer die URL hat, kann die App öffnen. Teile sie nur im Büro.
- Keine Bezahlfunktion: bezahlt wird beim Läufer.
- Wochenenden und Feiertage werden nicht gesondert behandelt.
