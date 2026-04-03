# Changelog – Infront Toolkit

Alle wesentlichen Änderungen an diesem Projekt werden in dieser Datei dokumentiert.

---

## [1.0.0] – 2026-04-03

### Neu
- Komplette Neuarchitektur als Office.js Add-in (React + TypeScript + Webpack)
- `manifest.xml` mit vollständiger Ribbon-Struktur (7 Gruppen, 18 Buttons)
- Ribbon-Tab „Infront Toolkit" mit deutschen Labels, Screentips und Supertips
- Task Pane mit URL-basiertem Routing (`?view=`)
- Panel-Grundgerüste für alle Features (Schritt 4–13):
  - CornerRadiusPanel (partiell implementiert)
  - ColorPickerPanel (partiell implementiert)
  - BrandCheckPanel (Stub)
  - FormatPainterPanel (Stub)
  - FindReplacePanel (Text-Tab partiell implementiert)
  - GapEqualizerPanel (Stub)
  - AgendaPanel (Stub)
  - MasterImportPanel (Stub)
  - ReviewPanel (Kommentar-Einfügung implementiert)
  - RedBoxPanel (Stub)
- ExecuteFunction-Befehle (commands.ts):
  - `gapHorizontal` – horizontale Abstände angleichen (3+ Shapes)
  - `gapVertical` – vertikale Abstände angleichen (3+ Shapes)
  - `toggleRedBox` – Red Box auf aktiver Slide ein-/ausschalten
  - `redBoxAllSlides` – Red Box auf allen Slides einfügen
  - `removeRedBoxAll` – alle INFRONT_REDBOX-Shapes entfernen
  - `removeComments` – alle INFRONT_COMMENT_*/INFRONT_HIGHLIGHT_*-Shapes entfernen
  - `addHighlight` – Highlight-Rechteck einfügen
- Service-Layer:
  - `SelectionService` – Selektion lesen
  - `ShapeService` – Shape-Operationen
  - `SlideService` – Slide-Iteration
  - `TextService` – TextFrame-Operationen
  - `ConfigService` – Document Settings / Brand Config
  - `SessionState` – Undo-Fallback (Session-Snapshots)
- Utilities:
  - `colorUtils` – Hex/RGB-Konvertierung, Farb-Distanz, Toleranz-Vergleich
  - `geometryUtils` – Positions-/Größen-Berechnungen in pt
  - `logger` – Zentrales Logging
  - `errorHandler` – Office.js-Fehler-Mapping (DE)
  - `notifications` – Benachrichtigungs-Typen und Factory-Funktionen
- `config/Infront_BrandConfig.json` – Brand-Compliance-Konfiguration (Profile: Default, Strict)
- `README.md` – Setup, Dev, Build, Deployment, Einschränkungen
- `TESTING.md` – API-Limits, Mac-Einschränkungen, Testfälle
- `SHORTCUTS.md` – Tastaturkürzel-Dokumentation

### Architektur-Entscheidungen
- Klassisches XML-Manifest (stabiler auf Mac als Unified Manifest)
- office-js vom CDN geladen (kein Bundling)
- URL-Parameter-Routing statt React Router (keine externe Abhängigkeit)
- Fluent UI v8 für konsistente Office-UI

### Bekannte Einschränkungen (Kategorie C)
- Screen-Pixel-Picker: nicht möglich auf Mac (WebKit/EyeDropper nicht verfügbar)
- Vollständiger Master-Import: Office.js API zu eingeschränkt
- Natives Undo: nicht verfügbar in Office.js
- Globale Tastaturkürzel: nicht möglich in Office Add-ins

---

## Geplante Schritte

| Schritt | Feature | Status |
|---------|---------|--------|
| 4 | Corner Radius (vollständig) | Offen |
| 5 | Color Picker (vollständig) | Offen |
| 6 | Brand Compliance Checker | Offen |
| 7 | Format Painter+ | Offen |
| 8 | Find & Replace (vollständig) | Offen |
| 9 | Agenda Wizard | Offen |
| 10 | Master importieren | Offen |
| 11 | Review / Annotationen | Offen |
| 12 | Gap Equalizer (vollständig) | Offen |
| 13 | Red Box (vollständig) | Offen |
| 14 | Stabilität & Bugfixing | Offen |
| 15 | Shortcut-Kompatibilität | Offen |
