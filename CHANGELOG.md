# Changelog – Infront Toolkit

Alle wesentlichen Änderungen an diesem Projekt werden in dieser Datei dokumentiert.

---

## [Schritt 3] – Screen Color Picker (Windows + Mac) (2026-04-02)

### Neu

- **`src/Modules/modColorPicker.bas`**: Color Picker mit plattformspezifischer Implementierung:
  - `ShowColorPicker` (Public, Ribbon-Callback): Haupt-Entry-Point, ruft Plattform-Picker, öffnet Ergebnisform.
  - `PickScreenColorWindows` (Private, `#If Not Mac`): Zeigt Anweisungs-Dialog, liest nach OK Pixel-Farbe an Mausposition via Windows API. `ReleaseDC` in allen Code-Pfaden sichergestellt. COLORREF-Rückgabe wird direkt als VBA-RGB genutzt (identisches Byte-Layout, keine Konvertierung nötig).
  - `PickScreenColorMac` (Private, `#If Mac`): Nutzt `MacScript("choose color")` für macOS NSColorPanel, skaliert 0–65535 → 0–255, defensives Parsing. Fallback: manuelle Hex-Eingabe.
  - `ApplyColorToSelection` (Public): Wendet Farbe auf Fill / Line / Font aller selektierten Shapes an. Child-ShapeRange-Support.
  - `ApplyColorToShape` (Private): Fehlergesichertes Anwenden auf einzelne Shape.
  - `ColorToHex` / `ColorToRGBString` (Public): Hilfsfunktionen für die Ergebnisform.
  - API-Deklarationen: Identisches Muster zu `ModuleEyedropper.bas` (`Private`, `#If VBA7 And Win64`), kein Namenskonflikt.
- **`src/Forms/frmColorPicker.frm`**: Ergebnisform mit Farbvorschau (`lblPreview`), Hex/RGB-Anzeige (`lblInfo`), OptionButtons (Fill/Line/Font), Apply- und Close-Button. Controls müssen in VBA-IDE angelegt werden (kein .frx vorhanden – entspricht Projektkonvention).
- **`src/CustomUI/CustomUI.xml`**: Neue Gruppe `InfrontFormatGroup` / `TabViewInfrontFormatGroup` ("Format") in beiden Tab-Views nach der Shapes-Gruppe. Erster Button: `ColorPickerButton` → `ShowColorPicker`. Weitere Buttons folgen in Schritt 5 (Format Painter+).

### Technische Entscheidungen und Einschränkungen

| Plattform | Implementierung | Einschränkung |
|---|---|---|
| Windows | `GetCursorPos` + `GetDC(0)` + `GetPixel` nach OK-Klick | Kein Real-Time-Eyedropper; Maus muss vor OK positioniert sein |
| Mac | `MacScript("choose color")` → NSColorPanel | Kein Screen-Eyedropper; zeigt System-Farbauswahl-Dialog |
| Mac Fallback | Manuelle Hex-Eingabe | Wenn MacScript nicht verfügbar |

- `AppleScriptTask()` **bewusst nicht verwendet** (erfordert Deployment von .applescript-Datei nach `~/Library/Application Scripts/com.microsoft.Powerpoint/`)
- COLORREF ↔ VBA-RGB: identisches Byte-Layout, keine Konvertierung nötig
- Undo: automatisch durch PowerPoint, kein UndoRecord verfügbar

---

## [Schritt 2] – Corner Radius in Pixeln (2026-04-02)

### Neu

- **`src/Modules/modCornerRadius.bas`**: Neues Modul mit zwei Funktionen:
  - `SetCornerRadiusPx` (Public, Ribbon-Callback): Fragt Eckenradius in Pixel ab, konvertiert nach Punkten (Faktor 0,75 bei 96 DPI), berechnet den normierten Adjustment-Wert (`radiusPt / (Min(Width, Height) / 2)`, gedeckelt auf 0,5) und setzt ihn auf alle ausgewählten Shapes. Shapes ohne Justierungs-Support werden übersprungen. Zeigt Ergebnismeldung mit Anzahl angepasster/übersprungener Shapes.
  - `ApplyCornerRadius` (Private): Wendet den Radius auf ein einzelnes Shape an; gibt `False` zurück wenn das Shape keinen Eckenradius unterstützt (statt Fehler zu werfen).
- **`src/CustomUI/CustomUI.xml`**: Neuer Button `CornerRadiusButton` / `TabViewCornerRadiusButton` in der Shapes-Gruppe (Single-Tab-View + Multi-Tab-View), direkt vor den bestehenden Rounded-Corner-Buttons. Deutscher `screentip`/`supertip`.

### Technische Hinweise

- Formel für Adjustment-Wert identisch zu `ModuleObjectsRoundedCorners.bas` (Konsistenz)
- PowerPoint VBA kennt kein `UndoRecord` (nur Word/Excel) – Undo funktioniert automatisch pro Shape-Änderung (ggf. mehrere Ctrl+Z nötig)
- Kein Windows-API-Aufruf, kein FSO, keine plattformspezifischen Konstrukte
- Child-ShapeRange (Shapes innerhalb Gruppe selektiert) wird korrekt behandelt

---

## [Schritt 1] – Rebranding zu "Infront Toolkit" (2026-04-02)

### Geändert

- **`src/CustomUI/CustomUI.xml`**: Ribbon-Tab-Label von `"Instrumenta"` auf `"Infront"` umgestellt.
- **`src/Modules/ModuleAbout.bas`**: About-Dialog zeigt jetzt `"Infront Toolkit v1.71"` statt `"Instrumenta Powerpoint Toolbar v1.71"`.
- **`src/Modules/ModuleSettings.bas`**: Sub `DeleteAllInstrumentaSettings` umbenannt zu `DeleteAllInfrontSettings`.
- **`src/Forms/SettingsForm.frm`**: Aufruf von `DeleteAllInfrontSettings` angepasst.
- **`src/Modules/ModuleFeatureSearch.bas`**:
  - Alle Feature-Pfad-Strings (`"Instrumenta > ..."`, `"Instrumenta [Text] > ..."`) auf `"Infront > ..."` bzw. `"Infront [Text] > ..."` umgestellt.
  - Anzeige-Labels `"Instrumenta script editor"`, `"Instrumenta script"` auf `"Infront Script Editor"`, `"Infront Script"` umgestellt.
  - Anzeige-Label `"Find Instrumenta features"` auf `"Find Infront features"` umgestellt.
  - Anzeige-Label `"Instrumenta settings"` auf `"Infront settings"` umgestellt.
  - Interne Hilfsfunktion `LoadInstrumentaFeatures` zu `LoadInfrontFeatures` umbenannt.
- **`src/Modules/ModuleStyleSheets.bas`**: Alle MsgBox-Texte und Dialog-Titel, die `"Instrumenta"` als lesbaren Namen nannten, auf `"Infront"` umgestellt. Layout-Tag-Namen (`"InstrumentaStylesheet"`, `"InstrumentaStyle"`, `"InstrumentaWarning"`) bleiben unverändert (in .pptx-Dateien gespeichert, Änderung würde bestehende Präsentationen brechen).
- **`src/Forms/PyramidForm.frm`**: MsgBox-Text `"All Instrumenta Pyramid tags have been removed."` auf `"All Infront Pyramid tags have been removed."` umgestellt. Presentation-/Shape-Tag-Namen (`"InstrumentaPyramid..."`) bleiben unverändert.
- **`src/Forms/ScriptEditorForm.frm`**: Kommentar-Texte im Beispiel-Script auf `"Infront"` umgestellt.
- **`src/Forms/InsertSlideLibrarySlide.frm`**: MsgBox-Text für fehlende Slide-Library auf `"Infront Toolkit settings"` umgestellt.

### Bewusst NICHT umbenannt (mit Begründung)

| Bezeichner | Grund |
|---|---|
| `InstrumentaInitialize`, `InstrumentaGetVisible`, `InstrumentaGetVisibleOneTabView`, `InstrumentaGetVisibleMultiTabView`, `InstrumentaRefresh` | Callback-Funktionsnamen, die exakt in `customUI.xml` referenziert sind. Umbenennung würde Ribbon-Ausfall verursachen. |
| `InstrumentaRibbon`, `InstrumentaVisible`, `InstrumentaVersion` | Globale Variablennamen, in mehreren Modulen referenziert; interne Bezeichner, für Endnutzer nicht sichtbar. |
| `GetSetting("Instrumenta", ...)` / `SaveSetting("Instrumenta", ...)` | Windows-Registry-Namespace. Umbenennung würde alle bestehenden Nutzereinstellungen löschen. Eine Migration ist als optionaler Build-Schritt denkbar. |
| Presentation-/Shape-Tags (`"InstrumentaPyramidSlideIndex"`, `"InstrumentaStyle"`, etc.) | In `.pptx`-Dateien gespeichert. Umbenennung bricht alle bestehenden Nutzerpräsentationen. |
| Layout-Name `"InstrumentaStylesheet"`, Shape-Name `"InstrumentaWarning"` | Ebenfalls in `.pptx` gespeichert. |
| AppleScript-Plugin-Dateiname `InstrumentaAppleScriptPlugin.applescript` | Deployment-Abhängigkeit; muss separat im macOS-Plugin-Verzeichnis abgelegt werden. |
| GitHub-URLs in `AboutDialog.frm`, `ScriptEditorForm.frm` | Verweisen auf das Original-Repository. Sobald ein Fork-URL bekannt ist, können diese angepasst werden. |

### Hinweis: Add-in-Name in Dokumenteigenschaften der `.ppam`

Der interne Name der `.ppam`-Datei sowie die Dokumenteigenschaften (Titel, Betreff) können nicht durch Datei-Änderungen allein umgestellt werden. Dies muss manuell in PowerPoint erfolgen:

1. `.ppam` öffnen (als `.pptm` umbenennen oder direkt in VBA-Editor öffnen).
2. `Datei > Eigenschaften > Erweiterte Eigenschaften` → Titel auf `"Infront Toolkit"` setzen.
3. Datei als `.ppam` speichern.

Alternativ kann dieser Schritt in einen Build-Prozess integriert werden, der das `.ppam` nach dem Compile automatisch mit korrekten Metadaten versieht.

---

*Basis: Instrumenta PowerPoint Toolbar v1.71 (Fork von iappyx/Instrumenta, MIT License)*
