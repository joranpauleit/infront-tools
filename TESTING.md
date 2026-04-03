# Infront Toolkit – Testdokumentation

Manuelle Testcheckliste und API-Limitdokumentation für das Infront Toolkit.
Zielplattform: **PowerPoint für Mac (Microsoft 365)**.

---

## Testumgebungen

| Parameter | Wert |
|---|---|
| Primäre Plattform | macOS (PowerPoint für Mac) |
| Office-Version | Microsoft 365, ≥ 16.70 |
| Browser-Engine (Add-in) | WebKit (WKWebView auf Mac) |
| Node.js | ≥ 18 |
| TypeScript | ≥ 5.5 |
| Office.js Requirement Set | PowerPoint 1.5+ (angestrebt) |

---

## Bekannte API-Limits und Mac-Einschränkungen

### Kategorie C – Nicht robust umsetzbar

| Feature | Einschränkung | Dokumentiert seit |
|---|---|---|
| Screen-Pixel-Farbpicker | EyeDropper API nicht in WebKit/Safari (Stand 2026) | v1.0.0 |
| Master vollständig ersetzen | `SlideMaster` API in Office.js sehr eingeschränkt, kein vollständiger Ersatz | v1.0.0 |
| Natives Undo | `Application.Undo()` nicht in Office.js verfügbar | v1.0.0 |
| Globale Shortcuts | Office Add-ins können keine systemweiten Shortcuts registrieren | v1.0.0 |
| Tabellen-Zellen-Merge/Split | Kein API in Office.js PowerPoint | v1.0.0 |
| Animationen programmatisch | `Shape.animations` API sehr eingeschränkt | v1.0.0 |

### Kategorie B – Umsetzbar mit Workaround

| Feature | Einschränkung | Workaround |
|---|---|---|
| Corner Radius präzise | `adjustments`-Array nur für roundedRectangle dokumentiert | Normierung via Shape-Breite/-Höhe, andere Typen überspringen |
| Shape Tags persistent | Requirement Set PowerPoint 1.5+; auf alten Mac-Versionen nicht verfügbar | Shape-Name-Encoding als Fallback |
| Agenda Auto-Update | Keine Event-API für Slide-Wechsel auf Mac | Manueller Update-Button |
| Find & Replace Farbe | Keine Farbsuche-API | Alle Shapes iterieren, RGB-Vergleich mit Toleranz |
| Seitenzahlen aktualisieren | Platzhalter-Zugriff eingeschränkt | Textbox nach bekanntem Namen suchen |

---

## Modul-/Service-Übersicht

| Modul | Pfad | Zweck |
|---|---|---|
| SelectionService | `src/services/powerpoint/SelectionService.ts` | Selektion lesen |
| ShapeService | `src/services/powerpoint/ShapeService.ts` | Shape-Suche, -Löschung |
| SlideService | `src/services/powerpoint/SlideService.ts` | Slide-Iteration |
| TextService | `src/services/powerpoint/TextService.ts` | Text lesen/schreiben |
| ConfigService | `src/services/config/ConfigService.ts` | Document Settings |
| BrandConfig | `src/services/config/BrandConfig.ts` | Typen + Default-Konfiguration |
| SessionState | `src/services/state/SessionState.ts` | Undo-Fallback |
| colorUtils | `src/utils/colorUtils.ts` | Farb-Konvertierung, Distanz |
| geometryUtils | `src/utils/geometryUtils.ts` | Positionen, Abstände in pt |
| logger | `src/utils/logger.ts` | Zentrales Logging |
| errorHandler | `src/utils/errorHandler.ts` | Office.js Fehler-Mapping |
| notifications | `src/utils/notifications.ts` | Benachrichtigungs-Typen |

---

## Happy-Path-Tests (Grundgerüst)

### Add-in laden

- [ ] Dev-Server läuft auf `https://localhost:3000`
- [ ] Zertifikat im Mac-Browser als vertrauenswürdig markiert
- [ ] `manifest.xml` erfolgreich in PowerPoint geladen (Sideload)
- [ ] Tab „Infront Toolkit" erscheint im Ribbon
- [ ] Alle 7 Gruppen sichtbar: Shapes, Format, Quality, Ausrichten, Struktur, Design, Review

### Ribbon-Buttons

- [ ] Eckenradius → öffnet Task Pane mit CornerRadius-Panel
- [ ] Farbwähler → öffnet Task Pane mit ColorPicker-Panel
- [ ] Brand Check → öffnet Task Pane mit BrandCheck-Panel
- [ ] Format Painter+ → öffnet Task Pane mit FormatPainter-Panel
- [ ] Suchen & Ersetzen → öffnet Task Pane mit FindReplace-Panel
- [ ] Gap... → öffnet Task Pane mit GapEqualizer-Panel
- [ ] Agenda → öffnet Task Pane mit Agenda-Panel
- [ ] Master importieren → öffnet Task Pane mit MasterImport-Panel
- [ ] Red Box Einstellungen → öffnet Task Pane mit RedBox-Panel
- [ ] Kommentar → öffnet Task Pane mit Review-Panel
- [ ] Meine Kommentare → öffnet Task Pane mit Review-Panel (my-comments)

### Direkt-Befehle (ExecuteFunction)

- [ ] Gap H: 3 selektierte Shapes → horizontale Abstände werden angeglichen
- [ ] Gap V: 3 selektierte Shapes → vertikale Abstände werden angeglichen
- [ ] Red Box: INFRONT_REDBOX erscheint auf aktiver Slide
- [ ] Red Box (nochmals): INFRONT_REDBOX wird entfernt (Toggle)
- [ ] Red Box: Alle Slides → INFRONT_REDBOX auf allen Slides
- [ ] Red Box entfernen → alle INFRONT_REDBOX-Shapes gelöscht
- [ ] Markieren → INFRONT_HIGHLIGHT_*-Shape erscheint
- [ ] Kommentare entfernen → alle INFRONT_COMMENT_*-Shapes gelöscht

### Corner Radius Panel (partiell)

- [ ] Eingabe „8" → Rounded Rectangle wird angepasst
- [ ] Eingabe „0" → kein Fehler, Radius auf 0
- [ ] Ungültige Eingabe → Fehlermeldung erscheint
- [ ] Kein Shape selektiert → Warnmeldung erscheint
- [ ] Reguläres Rechteck → wird übersprungen, Meldung zeigt „0 angepasst, 1 übersprungen"

### Color Picker Panel (partiell)

- [ ] Hex-Eingabe „#FF0000" → rote Vorschau erscheint
- [ ] Kurzform „#F00" → wird zu „#FF0000" normalisiert
- [ ] „Aus Shape" → Füllfarbe des selektierten Shapes wird übernommen
- [ ] Ungültiger Hex → Fehlermeldung
- [ ] Anwenden auf Füllung → Shape-Füllung ändert sich
- [ ] Anwenden auf Linie → Shape-Linie ändert sich
- [ ] Zuletzt-verwendet-Farben erscheinen nach erstem Anwenden

### Find & Replace (Text-Tab)

- [ ] Suchtext eingeben, Ersetzen-Text eingeben → „Alle ersetzen" läuft
- [ ] Leeres Suchfeld → „Alle ersetzen"-Button deaktiviert
- [ ] Text nicht gefunden → Meldung „0 Textvorkommen ersetzt"

### Review Panel

- [ ] Kommentar einfügen → INFRONT_COMMENT_*-Shape erscheint auf aktiver Slide
- [ ] Zeitstempel korrekt formatiert (DE: TT.MM.JJJJ HH:MM)

---

## Edge-Case-Tests

- [ ] Keine Shapes selektiert → alle relevanten Features zeigen Warnmeldung
- [ ] Leere Präsentation (0 Slides) → kein Absturz
- [ ] Shape-Gruppe selektiert → defensiv behandelt, kein Absturz
- [ ] Tabellen-Shape selektiert → Corner Radius überspringt ohne Fehler
- [ ] Sehr viele Shapes (50+) → kein Timeout

---

## Deployment-Checkliste (Produktion)

- [ ] `manifest.xml`: localhost-URLs durch Produktions-Domain ersetzen
- [ ] HTTPS-Zertifikat auf Produktions-Server gültig
- [ ] Icons in `assets/icons/` vorhanden (16/32/80 px)
- [ ] `npm run build` erfolgreich
- [ ] `dist/` auf Server deployt
