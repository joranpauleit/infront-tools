# Infront Toolkit

Professionelles PowerPoint Add-in für Strategieberater — entwickelt für **PowerPoint auf dem Mac**.

Technologie: Office.js · React · TypeScript · Webpack

---

## Voraussetzungen

- Node.js ≥ 18
- npm ≥ 9
- PowerPoint für Mac (Microsoft 365, Version ≥ 16.70 empfohlen)
- macOS 12 oder neuer

---

## Setup

```bash
# Abhängigkeiten installieren
npm install
```

### Icons einrichten

Das Manifest referenziert Icons unter `assets/icons/`. Folgende Dateien müssen vorhanden sein:

```
assets/icons/icon-16.png   (16×16 px)
assets/icons/icon-32.png   (32×32 px)
assets/icons/icon-80.png   (80×80 px)
```

Platzhalter-Icons können aus `src/CustomUI/icons/png/` kopiert und skaliert werden.

---

## Entwicklung

### Dev-Server starten

```bash
npm run dev-server
```

Der webpack-dev-server läuft auf `https://localhost:3000`.

**HTTPS-Zertifikat vertrauen (einmalig, Mac):**

```bash
# Browser öffnen und Zertifikat akzeptieren
open "https://localhost:3000"
# Im Safari/Chrome-Dialog "Vertrauen" wählen
# Oder über: Systemeinstellungen → Datenschutz & Sicherheit → Zertifikate
```

### Add-in in PowerPoint laden (Sideload)

1. PowerPoint öffnen
2. Menü: **Einfügen → Add-ins → Meine Add-ins → Add-in hochladen**
3. `manifest.xml` aus dem Repository-Root auswählen
4. Der Tab **Infront Toolkit** erscheint im Ribbon

---

## Build

```bash
# Production Build (in /dist)
npm run build

# Development Build (mit Source Maps)
npm run build:dev
```

Output liegt in `dist/`. Für die Produktion werden `dist/` + `manifest.xml` benötigt.

---

## Typecheck & Lint

```bash
npm run typecheck   # TypeScript prüfen ohne Build
npm run lint        # ESLint ausführen
```

---

## Projektstruktur

```
infront-tools/
├── manifest.xml              # Office Add-in Manifest
├── package.json
├── tsconfig.json
├── webpack.config.js
├── assets/icons/             # Add-in Icons (PNG, mind. 16/32/80 px)
├── config/
│   └── Infront_BrandConfig.json   # Brand Compliance Konfiguration
├── src/
│   ├── commands/             # ExecuteFunction-Ribbon-Handler
│   │   ├── commands.html
│   │   └── commands.ts
│   ├── taskpane/             # React Task Pane App
│   │   ├── index.html
│   │   ├── index.tsx
│   │   ├── App.tsx           # Root + URL-Router
│   │   ├── App.css
│   │   └── components/       # Feature-Panel-Komponenten
│   │       ├── shared/       # Gemeinsame UI-Komponenten
│   │       ├── CornerRadius/
│   │       ├── ColorPicker/
│   │       ├── BrandCheck/
│   │       ├── FormatPainter/
│   │       ├── FindReplace/
│   │       ├── GapEqualizer/
│   │       ├── Agenda/
│   │       ├── MasterImport/
│   │       ├── Review/
│   │       └── RedBox/
│   ├── services/
│   │   ├── powerpoint/       # Office.js Service-Layer
│   │   ├── config/           # Konfiguration
│   │   └── state/            # Session-State / Undo-Fallback
│   └── utils/                # Farb-, Geometrie-, Logger-Utilities
├── README.md
├── CHANGELOG.md
├── TESTING.md
└── SHORTCUTS.md
```

---

## Ribbon-Struktur

| Gruppe     | Button                                              | Typ          |
|------------|-----------------------------------------------------|--------------|
| Shapes     | Eckenradius                                         | Task Pane    |
| Format     | Farbwähler, Format Painter+                         | Task Pane    |
| Quality    | Brand Check, Suchen & Ersetzen, Kommentare entfernen, Red Box entfernen | Task Pane / Direkt |
| Ausrichten | Gap H, Gap V, Gap...                                | Direkt / Task Pane |
| Struktur   | Agenda                                              | Task Pane    |
| Design     | Master importieren, Red Box, Red Box: Alle Slides, Red Box Einstellungen | Task Pane / Direkt |
| Review     | Kommentar, Markieren, Meine Kommentare             | Task Pane / Direkt |

---

## Bekannte Einschränkungen (Mac / Office.js)

Vollständige Liste: [TESTING.md](TESTING.md)

| Einschränkung | Grund | Fallback |
|---|---|---|
| Screen-Pixel-Picker | WebKit/Safari unterstützt EyeDropper API nicht | Hex-Eingabe + Shape-Farbübernahme |
| Master vollständig ersetzen | Kein SlideMaster-Replacement in Office.js | Theme-Farben/-Fonts anpassen |
| Natives Undo | Kein `Application.Undo()` in Office.js | Session-Snapshot + Revert-Button |
| Globale Shortcuts | Office Add-ins können keine globalen Shortcuts registrieren | Ribbon-Buttons als Primärpfad |

---

## Legacy-Referenz

`src/Modules/`, `src/Forms/`, `src/Classes/`, `bin/`, `v/` enthalten den ursprünglichen VBA-Code (Instrumenta). Diese Dateien sind **keine aktive Codebasis** und dienen nur als Feature-Referenz.
