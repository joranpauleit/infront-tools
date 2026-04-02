# Changelog – Infront Toolkit

Alle wesentlichen Änderungen an diesem Projekt werden in dieser Datei dokumentiert.

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
