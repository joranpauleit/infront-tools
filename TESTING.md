# Infront Toolkit – Testplan

Manuelle Checkliste für alle Features des Infront Toolkits.
Vor jedem Test: PowerPoint neu starten, Add-in laden, neue leere Präsentation öffnen.

---

## Allgemein

- [ ] Ribbon-Tab "Infront" erscheint nach Add-in-Installation
- [ ] Alle Gruppen (Font, Text, Formatting, Format, Design, Slides, Review, Quality, Tables, Advanced …) sichtbar
- [ ] Single-Tab-View und Multi-Tab-View umschaltbar
- [ ] Kein Fehler beim Starten von PowerPoint mit geladenem Add-in

---

## Schritt 1 – Rebranding

- [ ] Tab-Label lautet "Infront" (nicht "Instrumenta")
- [ ] About-Dialog zeigt "Infront Toolkit v…"
- [ ] Settings-Dialog zeigt keine "Instrumenta"-Texte in Benutzeroberfläche
- [ ] Feature-Suche listet Einträge mit "Infront >" statt "Instrumenta >"

---

## Schritt 2 – Eckenradius in Pixeln (modCornerRadius)

### Normalfall
- [ ] Rechteck selektieren → Button "Eckenradius (px)" → InputBox erscheint
- [ ] Eingabe "8" → Eckenradius wird auf 6 pt gesetzt (8 × 0,75)
- [ ] Mehrere Shapes selektieren → alle erhalten Radius
- [ ] Gruppe selektieren → enthaltene Shapes werden rekursiv behandelt

### Kantenfälle
- [ ] Eingabe "0" → Radius = 0 (keine Rundung)
- [ ] Eingabe > halbe Seite → Wert wird auf 0,5 (Maximum) geklemmt
- [ ] Shape ohne Adjustments (z.B. Linie) → Shape wird übersprungen, kein Absturz
- [ ] Abbruch in InputBox → nichts passiert

---

## Schritt 3 – Screen Color Picker (modColorPicker)

### Windows
- [ ] Button öffnet Color-Picker-Dialog
- [ ] Farbe aufnehmen → frmColorPicker erscheint mit Vorschau
- [ ] Optionen Fill/Line/Font auswählen → Farbe wird angewendet
- [ ] Abbruch in Form → keine Änderung

### Mac
- [ ] Button öffnet macOS NSColorPanel (`MacScript("choose color")`)
- [ ] Gewählte Farbe erscheint in frmColorPicker
- [ ] Fallback InputBox bei Fehler (z.B. Skript abgebrochen) → kein Absturz

### Kantenfälle
- [ ] Kein Shape selektiert → Button deaktiviert (`getEnabled`)
- [ ] Shape ohne Fill → FillColor-Option hat keine Wirkung (kein Absturz)

---

## Schritt 4 – Brand Compliance Checker (modBrandCompliance)

### Konfiguration
- [ ] Kein `Infront_BrandConfig.ini` vorhanden → Vorlage wird erstellt, Meldung erscheint
- [ ] Nach Anpassen der INI: Button nochmals → Prüfung startet

### Prüfung
- [ ] Präsentation mit falscher Schriftart → Verstoß "Font" erscheint in Liste
- [ ] Präsentation mit nicht erlaubter Füllfarbe → Verstoß "FillColor"
- [ ] Schriftgröße < MinFontSizePt → Verstoß "FontSize"
- [ ] Gruppe mit innerem Shape → Verstoß wird korrekt erkannt (rekursiv)
- [ ] Tabelle mit nicht erlaubter Zellfarbe → Verstoß erkannt

### Form
- [ ] frmBrandCompliance öffnet sich mit Verstössen in lstViolations
- [ ] "Zur Folie" → navigiert zu korrekter Folie
- [ ] "Auswahl beheben" → Fix angewendet, Eintrag aus Liste entfernt
- [ ] "Als CSV exportieren" → CSV-Datei erstellt

### Kantenfälle
- [ ] Keine Präsentation offen → Fehlermeldung, kein Absturz
- [ ] Keine Verstösse → "Keine Verstöße gefunden"-Meldung
- [ ] Leere Präsentation (0 Folien) → kein Absturz

---

## Schritt 5 – Format Painter Plus (modFormatPainterPlus)

### Normalfall
- [ ] Genau 1 Shape selektieren → Button aktiv → Form öffnet sich
- [ ] lblSourceInfo zeigt korrekte Farben/Schrift-Werte
- [ ] Checkboxen standardmäßig aktiviert (außer Breite/Höhe)
- [ ] Ziel-Shapes selektieren → "Anwenden" → Formatierung übertragen
- [ ] "Alle" / "Keine" Buttons funktionieren

### Eigenschaften einzeln testen
- [ ] Nur FillColor → Füllfarbe übertragen, Schrift unverändert
- [ ] Nur FontName → Schriftart übertragen, Fill unverändert
- [ ] Breite/Höhe aktivieren → Größe wird übertragen

### Kantenfälle
- [ ] 0 Shapes selektiert → Button deaktiviert
- [ ] Mehrere Shapes als Quelle → Button deaktiviert
- [ ] Shape ohne TextFrame → Font-Optionen haben keine Wirkung (kein Absturz)

---

## Schritt 6 – Global Find & Replace (modFindReplace)

### Normalfall
- [ ] Form öffnet sich über Button
- [ ] Suchtext eingeben → "Vorschau" zeigt korrekte Trefferzahl
- [ ] "Alle ersetzen" → Bestätigung bei Scope=Alle → Ersetzungen durchgeführt
- [ ] lblResult zeigt korrekte Anzahl

### Optionen
- [ ] Groß-/Kleinschreibung aus → "Hallo" findet "HALLO"
- [ ] Groß-/Kleinschreibung an → "Hallo" findet nicht "HALLO"
- [ ] Nur ganze Wörter → "test" findet nicht "testen"
- [ ] Scope "Aktuelle Folie" → nur aktuelle Folie wird durchsucht
- [ ] Scope "Selektierte Folien" → nur markierte Folien
- [ ] Notizen einschließen → Text in Sprechernotizen wird ersetzt

### Formatierungserhalt
- [ ] Fett markiertes Wort ersetzen → Fettformatierung bleibt erhalten
- [ ] Wort mit Farbe ersetzen → Farbe bleibt erhalten

### Kantenfälle
- [ ] Leerer Suchtext → keine Aktion, Hinweis in lblResult
- [ ] Kein Treffer → "Keine Treffer gefunden"
- [ ] Gruppen / Tabellen → Text darin wird ebenfalls durchsucht

---

## Schritt 7 – Agenda Wizard (modAgendaWizard)

### Normalfall
- [ ] Form öffnet sich mit Standardwerten
- [ ] 3 Agendapunkte eingeben → Generieren → Übersichtsfolie eingefügt
- [ ] Modus "Übersicht + Fortschrittsfolien" → N+1 Folien eingefügt
- [ ] Farben/Schriftgrößen aus Form werden korrekt übernommen
- [ ] Fortschrittsfolien: Punkt i = aktiv (fett), <i = erledigt (grau), >i = inaktiv (hellgrau)

### Idempotenz
- [ ] Nochmals "Generieren" → alte Folien gelöscht, neue eingefügt (keine Duplikate)
- [ ] "Agenda löschen" → alle markierten Folien entfernt, Bestätigung

### Kantenfälle
- [ ] Keine Agendapunkte → Fehlermeldung
- [ ] Leere Zeilen in txtItems → werden übersprungen
- [ ] Einfügen nach Folie 0 → vor Folie 1 eingefügt

---

## Schritt 8 – Master-Importer (modMasterImport)

### Normalfall
- [ ] "Durchsuchen" → Datei-Dialog öffnet sich
- [ ] "Masters laden" → ListBox zeigt Master-Namen aus gewählter Datei
- [ ] Master auswählen → "Importieren" → Master in aktiver Präsentation vorhanden
- [ ] "Auf alle Folien anwenden" → alle Folien nutzen neuen Master

### Optionen
- [ ] "Ungenutzte Masters entfernen" → alte ungenutzte Masters gelöscht
- [ ] chkApplyAll und chkApplySelected schließen sich gegenseitig aus

### Kantenfälle
- [ ] Quelldatei hat 0 Folien → Fehlermeldung "enthält keine Folien", kein Absturz
- [ ] Passwortgeschützte Datei → Fehlermeldung, kein Absturz
- [ ] Nicht-existente Datei → Fehlermeldung

---

## Schritt 9 – User-Name-Stempel (modUserStamp)

### Normalfall
- [ ] "Stempel setzen" → YesNoCancel-Dialog erscheint
- [ ] "Ja" → Stempel auf aktuelle Folie gesetzt (unten rechts, grau)
- [ ] "Nein" → Stempel auf alle Folien gesetzt
- [ ] Stempel enthält Username + aktuelles Datum/Zeit
- [ ] "Stempel entfernen" → Bestätigung → alle Stempel gelöscht

### Idempotenz
- [ ] Zweimal "Stempel setzen" auf gleicher Folie → nur ein Stempel (alter wird ersetzt)

### Kantenfälle
- [ ] Application.UserName leer → InputBox erscheint
- [ ] Abbruch in Dialog → keine Aktion
- [ ] Keine Stempel vorhanden → "Keine User-Stempel" Meldung

---

## Schritt 10 – Smart Gap Equalizer (modGapEqualizer)

### Normalfall
- [ ] 3 Shapes selektieren → Button aktiv → Form öffnet sich
- [ ] "Aktualisieren" → lblCurrentInfo zeigt Min/Max/Avg
- [ ] Modus "Durchschnitt" → Gaps angeglichen
- [ ] Modus "Custom 10" → alle Gaps = 10 pt

### Modi
- [ ] Anchor "Erstes Shape fixiert" → erstes Shape bleibt, Rest verschiebt sich
- [ ] Anchor "Bounds" → Shapes gleichmäßig innerhalb Gesamtbounds verteilt
- [ ] txtGapPt deaktiviert wenn Modus ≠ Custom ODER Anchor = Bounds
- [ ] Vertical → Shapes nach oben/unten verschoben

### Kantenfälle
- [ ] 2 Shapes → funktioniert (1 Gap)
- [ ] 1 Shape → Button deaktiviert
- [ ] Überlappende Shapes (negativer Gap) → wird korrekt angezeigt und kann auf 0 gesetzt werden
- [ ] Kein Shape selektiert → Button deaktiviert

---

## Schritt 11 – Red Box (modRedBox)

### Normalfall
- [ ] Shapes selektieren → "Red Box" → Outline-Rahmen um Selektion (+ 6 pt Padding)
- [ ] "Red Box (Filled)" → halbtransparente rote Fläche
- [ ] Kein Shape selektiert → Box in Folienmitte
- [ ] "Red Boxes entfernen" → alle Boxes auf aktiver Folie gelöscht

### Kantenfälle
- [ ] Mehrere Red Boxes auf einer Folie → alle entfernt
- [ ] > 3 Boxes → Bestätigung vor Löschen
- [ ] Red Box ist ein normales Shape (kann nachträglich bearbeitet werden)
- [ ] Speichern und neu öffnen → Boxes bleiben mit Tag erhalten

---

## Plattformtests (Windows + Mac)

| Feature | Windows | Mac |
|---|---|---|
| Eckenradius | ☐ | ☐ |
| Color Picker | ☐ (WinAPI) | ☐ (MacScript) |
| Brand Check CSV-Export | ☐ (FileDialog) | ☐ (InputBox) |
| Find & Replace | ☐ | ☐ |
| Master-Import Dateidialog | ☐ (FileDialog) | ☐ (InputBox) |
| Format Painter Plus | ☐ | ☐ |
| Gap Equalizer | ☐ | ☐ |
| Red Box | ☐ | ☐ |

---

## Regressionstests (Instrumenta-Basisfunktionen)

Nach allen Änderungen sicherstellen dass keine Instrumenta-Grundfunktionen beschädigt wurden:

- [ ] Slide Grader funktioniert
- [ ] Script Editor öffnet sich
- [ ] Color Manager funktioniert
- [ ] Slide Library Insert funktioniert
- [ ] Mail Merge Funktionen fehlerfrei
- [ ] Settings Dialog speichert und lädt korrekt
- [ ] Feature-Suche findet alle Features
