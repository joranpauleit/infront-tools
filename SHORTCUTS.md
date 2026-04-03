# Infront Toolkit – Tastaturkürzel

PowerPoint VBA bietet keine `Application.OnKey`-API (nur Word/Excel).
Direkte Hotkeys wie `Ctrl+Shift+X` sind daher **nicht programmatisch zuzuweisen**.

Stattdessen stehen zwei Wege zur Verfügung:

---

## 1. Alt-Tasten-Navigation (KeyTips)

Alle Infront-eigenen Schaltflächen sind mit `keytip`-Attributen versehen.
Nach Drücken von **Alt** erscheinen die Buchstaben-Kürzel im Ribbon.

### Tastenfolge (Beispiel Single-Tab-View)

```
Alt  →  [Tab-Kürzel]  →  [Button-Kürzel]
```

Der Tab-Buchstabe hängt von der Office-Version und Installationsreihenfolge ab.
Sobald der Tab aktiv ist, gelten die folgenden Button-Kürzel:

| Kürzel | Feature                  | Gruppe     |
|--------|--------------------------|------------|
| `CR`   | Eckenradius (px)         | Format     |
| `CP`   | Color Picker             | Format     |
| `FP`   | Format Painter+          | Format     |
| `GE`   | Gap Equalizer            | Align      |
| `RB`   | Red Box (Outline)        | Design     |
| `RF`   | Red Box (Filled)         | Design     |
| `RX`   | Red Boxes entfernen      | Design     |
| `AW`   | Agenda Wizard            | Slides     |
| `MI`   | Master importieren       | Slides     |
| `US`   | Stempel setzen           | Review     |
| `UX`   | Stempel entfernen        | Review     |
| `BC`   | Brand Check              | Quality    |
| `FR`   | Find & Replace+          | Advanced   |

Dieselben Kürzel gelten im Multi-Tab-View (TabView-Buttons).

---

## 2. Quick Access Toolbar (QAT) – Strg+1 … Strg+9

Für häufig genutzte Funktionen empfiehlt sich die **Schnellzugriffsleiste**:

1. Rechtsklick auf einen Infront-Button → **"Zur Schnellzugriffsleiste hinzufügen"**
2. Die ersten 9 Einträge in der QAT sind über `Strg+1` … `Strg+9` erreichbar.

Empfohlene Belegung:

| Taste    | Feature             |
|----------|---------------------|
| `Strg+1` | Find & Replace+     |
| `Strg+2` | Format Painter+     |
| `Strg+3` | Color Picker        |
| `Strg+4` | Brand Check         |
| `Strg+5` | Gap Equalizer       |
| `Strg+6` | Red Box (Outline)   |
| `Strg+7` | Agenda Wizard       |
| `Strg+8` | Stempel setzen      |
| `Strg+9` | Master importieren  |

Die QAT-Reihenfolge lässt sich unter
**Datei → Optionen → Symbolleiste für den Schnellzugriff** anpassen.

---

## 3. Mac-Hinweis

Auf macOS zeigt PowerPoint keine KeyTips an (Alt-Taste hat andere Belegung).
Für Mac-Nutzer ist die QAT der einzige Weg zu echten Tastaturkürzeln:

- QAT-Einträge 1–9 werden mit **⌘+1** … **⌘+9** aufgerufen
- Alternative: Ribbon-Suche über **⌘+/** (Office 365 Mac ab Version 16.x)

---

## Feature-Übersicht (Schritte 1–13)

| Schritt | Feature                        | Modul / Form                         |
|---------|--------------------------------|--------------------------------------|
| 1       | Rebranding (Infront)           | CustomUI.xml, modSettings, …         |
| 2       | Eckenradius in Pixeln          | modCornerRadius                      |
| 3       | Screen Color Picker            | modColorPicker, frmColorPicker       |
| 4       | Brand Compliance Checker       | modBrandCompliance, frmBrandCompliance|
| 5       | Format Painter Plus            | modFormatPainterPlus, frmFormatPainterPlus |
| 6       | Global Find & Replace          | modFindReplace, frmFindReplace       |
| 7       | Agenda Wizard                  | modAgendaWizard, frmAgendaWizard     |
| 8       | Master-Importer                | modMasterImport, frmMasterImport     |
| 9       | User-Name-Stempel              | modUserStamp                         |
| 10      | Smart Gap Equalizer            | modGapEqualizer, frmGapEqualizer     |
| 11      | Red Box                        | modRedBox                            |
| 12      | Bug-Fixes & TESTING.md         | (diverse), TESTING.md                |
| 13      | Shortcuts & SHORTCUTS.md       | CustomUI.xml (keytips), SHORTCUTS.md |
