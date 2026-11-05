# Systemdiagnose

Prüft Windows tiefgehend (gebaut für PC 1, läuft auch auf anderen Windows-PCs), **erklärt jeden Befund** (Was ist das? Woran liegt es? Was ist zu tun?)
und kann ausgewählte Probleme **sicher beheben** – mit Bestätigung, Wiederherstellungspunkt und Rückgängig.
Getrennt davon räumt der Reiter **„Reinigung“** Caches und Temp-Dateien auf.

## Starten

Desktop-Verknüpfung **„Systemdiagnose“** (startet `Systemdiagnose.exe` → Programmfenster ohne Konsole). Die App hat einen
eigenen Taskleisteneintrag (lässt sich anheften), zeigt den Fortschritt einer Prüfung im Taskleistenknopf und meldet sich
mit einer Windows-Benachrichtigung, wenn eine längere Prüfung im Hintergrund fertig ist.

1. Oben den Modus wählen und **„Prüfung starten“**:

   | Modus | Rechte | Dauer | Zusätzlich |
   |---|---|---|---|
   | Schnell | keine | ~20 s | alles Grundlegende, mit Live-Fortschritt |
   | Gründlich | Admin | ~5–10 Min. | SMART, Fehlerzähler, TPM, Wiederherstellungspunkte, Aktivierung, DISM (Prüfung, Analyse, Tiefenscan `/ScanHealth`), SFC, Dateisystem (chkdsk im Lesemodus, NTFS-Ereignisse, Reparaturstatus, TRIM), Startkonfiguration, WMI-Datenbank, Signaturen aller Treiber, Dienste, Aufgaben und Autostart-Programme, Ergebnis der Windows-Speicherdiagnose, komplette Begehung jedes Laufwerks (größte Ordner) |

   Unter „Hinweise“ steht nur, was einen Blick lohnt; reine Übersicht steht unter „Alle Prüfungen“. Das Fortschrittsfenster zeigt
   oben den Fortschritt in Prozent und darunter einen dünnen Balken, der nur läuft, solange die Prüfung wirklich arbeitet.

   Umschalten mit dem Schalter **Schnell ⇄ Gründlich** unter „Prüfung starten“; ⓘ oder Rechtsklick auf den Schalter erklärt die Stufen.
   Die Reinigungskategorien werden bei jeder Prüfung mitgemessen. Die frühere Stufe „Tief“ ist in „Gründlich“ aufgegangen.

   Die App lässt sich nur einmal öffnen; ein zweiter Start holt das vorhandene Fenster nach vorn.
   Eigenes Hintergrundbild fürs Star-Wars-Design: `Design\StarWars.png` bzw. `Design\CloneWars.png` (oder `.jpg`) in diesen Ordner legen, App neu starten.

2. Drei Reiter: **Befunde · Reinigung · Verlauf**. Im Reiter Befunde oben Prüfmodus und Kacheln (filtern nach Status),
   links die Befunde (Filter: Handlungsbedarf / Mit Lösung / Alle Prüfungen, Suche; daneben leise „Archiv (n)“).
   Rechts die Karte mit **Was ist das? · Woran liegt es? · Was ist zu tun? · Empfehlung · Lösung · Details**.
3. Unter **„Lösung“** ankreuzen, was behoben werden soll (oder „Empfohlene auswählen“), dann
   **„Beheben“** → Bestätigungsdialog → Windows fragt nach Adminrechten.
   **Archivieren:** Jeder Befund mit Status Hinweis/Warnung/Kritisch hat unten in der Karte „Archivieren“. Das ist
   für Befunde gedacht, die du selbst erledigt hast, die die App nicht beheben kann (dann steht dort „In der App nicht
   behebbar“, z. B. WHEA-Hardwarefehler) oder die nicht mehr relevant sind. Archivierte Befunde zählen nicht mehr in
   den Kacheln und unter „Handlungsbedarf“. Sie stehen unter „Archiv“ neben der Suche (dort „Wiederherstellen“) und abgeblendet
   unter „Alle Prüfungen“. Sie kommen **von selbst zurück**, sobald ein neues Ereignis auftritt (z. B. ein neuer
   WHEA-Eintrag) oder der Befund schlimmer wird – mit dem Hinweis „Wieder aufgetaucht“. Ist ein Befund wieder OK,
   fällt er aus dem Archiv. Gespeichert wird in `Archiv.json`.
4. **Danach prüft die App selbst neu** (nach einer Reinigung: neu messen, ~8 s; nach Beheben oder Rückgängig: neu prüfen,
   immer im Modus Schnell ~30 s, ohne zweite Adminabfrage) und zeigt einen **Vergleich mit
   vorher**: welche Befunde behoben sind, welche offen bleiben (mit der Meldung der Maßnahme), was neu ist und je
   Reinigungskategorie „vorher → nachher“, z. B. „Chrome weiterhin 1.4 GB · läuft gerade“. Gelöschtes bleibt so nicht
   stehen. Einen eigenen Ergebnisdialog gibt es nur, wenn etwas fehlgeschlagen ist, abgelehnt wurde oder simuliert war.
5. Neu prüfen von Hand: Modus wählen und **„Prüfung starten“** (vergleicht nach einer Maßnahme ebenfalls) oder **F5**.
   Berichte älter als 60 Tage löscht die App beim Start selbst (die neueste Datei jeder Art bleibt immer).
6. Reiter **„Verlauf“** mit drei Ansichten: **Maßnahmen** (alle Behebungen und Reinigungen; umkehrbare lassen sich mit
   einem Klick zurücknehmen), **Systemänderungen** (was sich seit der vorigen Prüfung geändert hat) und **Entwicklung**
   (Zuverlässigkeitsindex und freier Speicher je Laufwerk als Kurve). **Datenquellen & Sicherheit** öffnet das
   Schild-Symbol oben rechts („Zurück“ oder Esc führt zurück).
   Neben dem Titel steht die **Gesamtnote** (z. B. „1 Warnung · geprüft vor 2 Std.“). **„Berichte“** öffnet den Ordner
   mit allen Berichten – je Prüfung eine HTML-Datei `Diagnose-<Datum>.html` zum Nachschauen.
7. **Einstellungen** (Zahnrad oben rechts): Design **Standard** (hell/dunkel wie Windows), **Immer hell**, **Immer dunkel**
   oder **Star Wars · Imperium** (Schwarz, Stahlgrau, Sturmtruppen-Weiß und Imperiumsrot; im Hintergrund ein berechnetes
   Weltraumbild mit roter Gaswolke, Sternzerstörern, Planet und einer Sith-Gestalt mit rotem Lichtschwert; Überschriften
   in Bahnschrift; die Reiter tragen rote Lichtschwert-Rahmen, hell bei Maus darauf oder geöffnet) oder
   **Star Wars · The Clone Wars** (dasselbe in Lichtschwert-Blau, Hintergrund aus `Design\CloneWars.png`). Wird sofort
   übernommen und in `Einstellungen.json` gespeichert. Ebenfalls dort: **automatisch prüfen beim Start**, wenn der letzte
   Stand älter als 3 Tage ist (Schnell, ohne Adminrechte; Standard an), und **„Dieser PC“** mit „Sollwerte neu einlernen“.

### Auf anderen PCs installieren

`Installer\Setup-erstellen.ps1` baut **`Installer\Ausgabe\Systemdiagnose-Setup-<Version>.exe`** (eine Datei, ~1,7 MB):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Installer\Setup-erstellen.ps1
```

Die EXE auf den anderen PC kopieren (USB-Stick, Cloud) und doppelklicken → **„Installieren“**. Keine Adminrechte nötig;
installiert wird nach `%LOCALAPPDATA%\Programs\Systemdiagnose`, mit Verknüpfung im Startmenü (und auf Wunsch auf dem
Desktop). Eine neuere EXE auf demselben PC **aktualisiert** nur das Programm – Berichte, Archiv und Einstellungen bleiben.
Entfernen: Einstellungen → Apps → Installierte Apps → „Systemdiagnose“ (fragt extra, ob auch die Berichte weg sollen).
Kommt die EXE aus dem Internet, warnt Windows SmartScreen (unsigniert): „Weitere Informationen“ → „Trotzdem ausführen“.

**Die App erkennt den PC selbst:** Auf PC 1 gelten die bekannten Sollwerte (Monitore, RAM, Energieplan, Dienste). Auf jedem
anderen PC startet sie neutral, lernt beim ersten Lauf dessen Monitore, RAM und Energieplan ein (`Profile\<PC>.json`) und
vergleicht danach nur noch mit eigenen Ständen – Berichte, Archiv und Vergleiche verschiedener PCs mischen sich nie.

### Reinigung (wie CCleaner)

Reiter **„Reinigung“**: 19 Kategorien in vier Gruppen, jede mit Größe, Dateizahl und Balken.

| Gruppe | Kategorien |
|---|---|
| Windows | Temp (Benutzer/Windows), Update-Downloads, Übermittlungsoptimierung, Upgrade-Protokolle, CBS/DISM-Protokolle, Miniaturansichten, Fehlerberichte, Absturzabbilder > 14 Tage, Papierkorb (auch verwaiste Reste) |
| Browser | Chrome, Edge, Firefox, Brave - **nur Cache**, nie Cookies, Passwörter, Verlauf, Tabs |
| Programme | Discord, Blitz, Steam (Browser-Cache) |
| Grafik & Treiber | DirectX-/AMD-Shader-Cache, C:\AMD (entpackte Treiber-Installer) |

- **„Neu analysieren“** misst in wenigen Sekunden ohne Adminrechte (Windows-Temp und Kernel-Abbilder nur mit Admin messbar).
  Nach einer Reinigung zeigt es dabei auch den Vergleich, was tatsächlich weniger geworden ist.
- Vorausgewählt (**„Empfohlene“**) sind nur Kategorien ohne Nachteile. Orange Hinweise vor dem Ankreuzen lesen
  (Shader-Cache = kurze Ruckler beim nächsten Spielstart, Papierkorb = endgültig, C:\AMD enthält eigene Sicherungen).
- **„Reinigen“** → Bestätigung → Bereinigungsteil mit Adminrechten. Läuft das Programm einer Kategorie (z. B. Chrome),
  wird sie übersprungen; gerade benutzte Dateien bleiben liegen. Danach zeigt der Dialog den freigegebenen Platz.
- **Programme schließen:** Läuft beim Klick auf „Reinigen“ ein Programm, dessen Daten gelöscht werden sollen
  (Chrome, Edge, Firefox, Brave, Discord, Blitz, Steam), kommt vorher der Hinweis **„Erst Programme schließen“** mit
  Anleitung je Programm (z. B. Steam über „Steam > Beenden“, weil das X es nur in den Infobereich schickt). Die App prüft
  alle 1,5 s nach; sobald alles zu ist, geht es von selbst zur Bestätigung weiter. Alternativ „Ohne diese reinigen“
  oder „Abbrechen“. Schon im Reiter steht unter „Ausgewählt“, welche Programme vorher zu schließen sind; die Hinweise
  „läuft gerade“ werden bei jeder Rückkehr in die App frisch geprüft.
- Welche Ordner zu einer Kategorie gehören, steht ausschließlich in `Reinigung.psd1` (reine Datendatei).

Konsolen-Alternative ohne Fenster: `Systemdiagnose starten.cmd` (nur Diagnose + HTML-Bericht).

## Aufbau – drei getrennte Teile

| Datei | Aufgabe | Rechte |
|---|---|---|
| `Systemdiagnose.ps1` | Diagnose-Engine: liest alles, schreibt nur Berichte | nur lesend |
| `Systemdiagnose-App.ps1` | Programmfenster: zeigt Befunde, sammelt die Auswahl | ändert nichts* |
| `Bereinigung.ps1` | führt bestätigte Maßnahmen aus festem Katalog aus | Admin, nur auf Anforderung |
| `Reinigung.psd1` | Reinigungskategorien (Ordner, Muster, Mindestalter) - nur Daten | - |

Engine-Modus `-Modus Reinigung` misst nur die Kategorien (schreibt `Berichte\Reinigung-*.json`).

\* außer harmlosen Sofort-Aktionen: Einstellungsseite öffnen, Mauszeiger-Sperre aufheben.

Ausgaben: `Berichte\` (HTML-Bericht, Befunddatei JSON, Vergleichsdatei) und `Bereinigung\` (Pläne, Ergebnisse, `Protokoll.txt`).
Alte Dateien in beiden Ordnern können jederzeit gelöscht werden.

## Sicherheit

**Diagnose (nur lesend)**
1. Selbstprüfung vor jedem Start: liest den eigenen Code und startet nicht, wenn darin ein nicht freigegebener
   Befehl, eine schreibende Methode, eine Datei-Umleitung, ein dynamischer Aufruf oder ein zerbrochener Text steht.
2. WhatIf-Netz: ändernde Windows-Befehle würden nur simulieren.
3. Registry nur über Lese-Handles; externe Programme nur aus System32 mit freigegebenen, lesenden Parametern.
4. Lösungen werden nur als Daten beschrieben, nie ausgeführt.

**Bereinigung (getrennter Teil)**
1. Nur Aktionen aus einem festen Katalog – der Plan liefert Kennung und Parameter, niemals Befehle.
2. Strenge Parameterprüfung und **erneute Prüfung am System direkt vor jeder Aktion** (ist das Problem noch da?).
3. Harte Grenzen: Windows-Dienste und Microsoft-Aufgaben werden nie angefasst; Treiberpakete nur, wenn eine aktive
   Ersatzversion gebunden ist, das Paket nicht neuer ist und ein höchstens 24 h alter Wiederherstellungspunkt existiert.
4. Umkehrbare Aktionen speichern den alten Zustand → Rückgängig.
5. Benutzer-Einstellungen nur, wenn die Bereinigung unter demselben Konto läuft wie die App.
6. Getestet mit einem absichtlich bösartigen Plan (Defender abschalten, Systemdateien löschen, eingeschleuste
   Parameter, unbekannte Aktionen): alles abgelehnt.

**Reinigung (zusätzlich)**
1. Der Plan nennt nur eine Kategorie-Kennung - Ordner löst der Bereinigungsteil selbst aus `Reinigung.psd1` auf,
   eingeschleuste Pfade im Plan werden ignoriert.
2. Harte Sperren unabhängig von der Kategorienliste: Laufwerkswurzel, Windows-Systemordner, Programmordner und
   persönliche Ordner (Dokumente, Desktop, Downloads, Bilder, OneDrive …) - weder der Ordner selbst noch darin oder darüber.
3. Löschen Datei für Datei über ein Handle, dessen **tatsächlicher Pfad** vorher geprüft wird: Ordner-Verknüpfungen
   (Junctions/Symlinks) können nichts nach außen umleiten. Nie `Remove-Item -Recurse` (folgt in PowerShell 5.1 Junctions).
4. Getestet mit 48 Fallen im Testordner (Verknüpfung nach außen, Umleitungstrick direkt am Löscher, gesperrte,
   schreibgeschützte und junge Dateien, alle Sperrordner) - alles korrekt.

Katalog: Treiberpaket entfernen · verwaisten Treiber/Dienst abschalten · Profil-Dienst starten/abschalten ·
geplante Aufgabe deaktivieren · Autostart deaktivieren · **Reinigen (Kategorie)** · Wiederherstellungspunkt anlegen ·
DISM/SFC-Reparatur · Energieplan · Schnellstart aus · GameDVR aus · Dynamic Lighting aus · Minidumps an · UAC an ·
UAC-Nachfrage an · Defender-Signaturen · Firewall an · Zeit synchronisieren · DNS-Cache leeren · doppeltes
Hintergrund-Skript beenden.

## Anpassen

`Systemdiagnose.ps1`, Abschnitt **„1. PC-PROFIL“**: Sollwerte (Monitore mit Hz/HDR, RAM, Energieplan,
Pflicht-/Verbots-Dienste, Laufwerks-Ausnahmen), bekanntes Log-Rauschen und die Erklärtexte („Wissen“).
Die Selbstprüfung kontrolliert nach jeder Änderung automatisch, ob alles weiterhin nur lesend ist.

`Reinigung.psd1`: Kategorien ergänzen oder ändern (Basis ist immer ein fester Name wie `LocalAppData`, nie ein
absoluter Pfad; `*` steht für ganze Ordnernamen, z. B. Browser-Profile). Diagnose und Bereinigung lesen dieselbe Datei.
