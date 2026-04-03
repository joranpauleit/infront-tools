# Changelog – Infront Toolkit

Alle wesentlichen Änderungen an diesem Projekt werden in dieser Datei dokumentiert.

---

## [Schritt 13] – Shortcuts + SHORTCUTS.md (2026-04-03)

### Neu

- **`src/CustomUI/CustomUI.xml`**: `keytip`-Attribut auf alle 13 Infront-eigenen Buttons gesetzt – sowohl im Single-Tab-View als auch im Multi-Tab-View (TabView):

  | Kürzel | Button-ID               | Feature              |
  |--------|-------------------------|----------------------|
  | `CR`   | CornerRadiusButton      | Eckenradius (px)     |
  | `CP`   | ColorPickerButton       | Color Picker         |
  | `FP`   | FormatPainterPlusButton | Format Painter+      |
  | `GE`   | GapEqualizerButton      | Gap Equalizer        |
  | `RB`   | InsertRedBoxButton      | Red Box (Outline)    |
  | `RF`   | InsertFilledRedBoxButton| Red Box (Filled)     |
  | `RX`   | RemoveRedBoxesButton    | Red Boxes entfernen  |
  | `AW`   | AgendaWizardButton      | Agenda Wizard        |
  | `MI`   | MasterImportButton      | Master importieren   |
  | `US`   | InsertUserStampButton   | Stempel setzen       |
  | `UX`   | RemoveUserStampsButton  | Stempel entfernen    |
  | `BC`   | BrandCheckButton        | Brand Check          |
  | `FR`   | FindReplaceButton       | Find & Replace+      |

- **`SHORTCUTS.md`**: Neue Datei mit:
  - Erklärung warum `Application.OnKey` in PPT VBA nicht verfügbar ist
  - Alt-KeyTip-Tabelle mit allen 13 Kürzeln
  - QAT-Anleitung (Strg+1…9 / ⌘+1…9 auf Mac) mit empfohlener Belegung
  - Mac-spezifischer Hinweis (keine KeyTips, QAT verwenden)
  - Feature-Übersicht aller Schritte 1–13

### Technische Entscheidungen

| Thema | Entscheidung |
|---|---|
| Keine `Application.OnKey` | In PPT VBA nicht vorhanden (nur Word/Excel) → keytips als einzige XML-seitige Lösung |
| Zwei-Buchstaben-Kürzel | Vermeiden Kollisionen mit einbuchstabigen Office-Standard-KeyTips |
| QAT als primärer Weg für Mac | macOS PowerPoint zeigt keine KeyTips an |

---

## [Schritt 12] – Bug Fixing + TESTING.md (2026-04-03)

### Bugfixes

| Datei | Bug | Fix |
|---|---|---|
| `modRedBox.bas` | Tote Konstante `REDBOX_COLOR = &H0000CC00` (nie verwendet, falsch kommentiert) | Entfernt, Klärungskommentar hinzugefügt |
| `modMasterImport.bas` | `srcPres.Slides(1).Copy` wirft Runtime Error wenn Quelldatei 0 Folien hat | Guard `If srcPres.Slides.Count = 0` → Close + Meldung vor Copy |
| `frmGapEqualizer.frm` | `txtGapPt` blieb aktiv wenn `optAnchorBounds` gewählt → Custom-Gap wird in Bounds-Modus ignoriert (irreführend) | `optAnchorBounds_Click` deaktiviert `txtGapPt`; `optAnchorFirst_Click` reaktiviert wenn `optCustom` aktiv |
| `modFormatPainterPlus.bas` | `Public Type ApplyOptions` deklariert nach erster Verwendung (Zeile 414 vs. 175) | Typ an den Anfang des Moduls verschoben (nach `FormatSnapshot`); doppelte Deklaration am Ende entfernt |

### Neu
- **`TESTING.md`**: Vollständiger manueller Testplan mit Checkboxen für alle 11 Feature-Schritte; Platform-Matrix (Windows/Mac); Regressionstest-Abschnitt für Instrumenta-Basisfunktionen

---

## [Schritt 11] – Red Box (2026-04-03)

### Neu

- **`src/Modules/modRedBox.bas`**: Red Box Modul (keine Form):
  - `InsertRedBox` (Public, Ribbon-Callback): Outline-Variante (kein Fill, roter Rahmen 2.5 pt). Shapes selektiert → Bounding-Box der Selektion + 6 pt Padding; sonst → Folienmitte (200 × 120 pt)
  - `InsertFilledRedBox` (Public, Ribbon-Callback): Filled-Variante (RGB 204,0,0, 80% Transparenz + Rahmen)
  - `RemoveRedBoxes` (Public, Ribbon-Callback): Entfernt alle Red Boxes (Tag `InfrontRedBox=1`) von der aktuellen Folie; Bestätigung nur wenn > 3
  - `CreateRedBoxShape(sld, left, top, w, h, filled)` (Public): Erstellt `msoShapeRectangle`, setzt Tag, formatiert Linie und optional Fill
  - `GetSelectionBounds(padding, left, top, w, h)` (Public): Berechnet Bounding-Box aller selektierten Shapes inkl. Padding; Fallback auf CenterBoxOnSlide
  - `CenterBoxOnSlide(...)` (Private): Berechnet zentrierte Position; Fallback-Werte 720×540 pt wenn PageSetup nicht lesbar
  - `HasShapesSelected()` (Private): Prüft `Selection.Type = ppSelectionShapes`
  - `CountRedBoxesOnSlide` / `DeleteRedBoxesFromSlide` (Private): Tag-basierte Suche und rückwärts-Löschung
- **`src/CustomUI/CustomUI.xml`**: Neue Gruppe `InfrontDesignGroup` (label="Design") **vor** `InfrontSlidesGroup` mit drei Buttons: `InsertRedBoxButton`, `InsertFilledRedBoxButton`, `RemoveRedBoxesButton`; `TabViewInfrontDesignGroup` entsprechend.

### Technische Entscheidungen

| Thema | Entscheidung |
|---|---|
| Keine Form | Zwei Ribbon-Varianten (Outline/Filled) als separate Buttons statt Auswahl-Dialog |
| Positionierung | Selektion vorhanden → Bounding-Box + Padding; sonst → Folienmitte |
| Bestätigung | Nur bei > 3 Boxes (1–3 offensichtlich gewollt) |
| Farbe | RGB(204, 0, 0) – kräftiges Rot; als VBA `RGB()`-Aufruf (kein COLORREF-Literal) |
| Tag | `InfrontRedBox=1` – persistiert in .pptx, zuverlässige Wiedererkennung |

---

## [Schritt 10] – Smart Gap Equalizer (2026-04-03)

### Neu

- **`src/Modules/modGapEqualizer.bas`**: Gap Equalizer Modul:
  - Typ `GapOptions` (Public): `Horizontal`, `Vertical`, `GapMode` (0=Custom/1=Avg/2=Min/3=Max), `CustomGapPt`, `AnchorMode` (0=Erstes fixiert/1=Bounds)
  - `ShowGapEqualizer` (Public, Ribbon-Callback): Prüft min. 2 Shapes selektiert, öffnet Form modeless
  - `EqualizeGaps(opts)` (Public): Sortiert Shapes, bestimmt Ziel-Gap, ruft `ApplyGaps`
  - `GetGapInfo(isHorizontal)` (Public): Gibt Min/Max/Avg der aktuellen Gaps als String zurück (für Form-Preview)
  - `GetShapesSorted(sr, isHorizontal, shapes())` (Public): Bubble-Sort nach Left (H) oder Top (V); kein WorksheetFunction
  - `CalculateCurrentGaps(shapes(), isHorizontal, gaps())` (Public): Gap = Abstand zwischen Kante[i] und Kante[i+1] (negative Werte = Überlappung möglich)
  - `ResolveTargetGap(shapes(), isHorizontal, opts)` (Private): Wählt Custom/Avg/Min/Max per eigener Schleife
  - `ApplyGaps(shapes(), targetGap, isHorizontal, anchorMode)` (Public): AnchorMode=0 → erstes Shape fixiert, jedes folgende relativ positioniert; AnchorMode=1 → Gesamtbounds beibehalten, gleichmäßig verteilen
- **`src/Forms/frmGapEqualizer.frm`**: Steuerform:
  - `InitForm()`: Defaults (H-Richtung, Average, Anchor=First)
  - `optCustom/Average/Minimum/Maximum_Click`: Aktiviert/deaktiviert `txtGapPt`
  - `btnRefresh_Click`: Ruft `GetGapInfo` und aktualisiert `lblCurrentInfo`
  - `btnApply_Click`: Baut `GapOptions`, ruft `EqualizeGaps`, refresh danach
  - Controls müssen in VBA-IDE angelegt werden (kein .frx)
- **`src/CustomUI/CustomUI.xml`**: `GapEqualizerButton` nach `ObjectsDecreaseSpacingVertical` in `ObjectsGroup`; `TabViewGapEqualizerButton` entsprechend.

### Technische Entscheidungen

| Thema | Entscheidung |
|---|---|
| Gap-Definition | Abstand zwischen rechter/unterer Kante Shape[i] und linker/oberer Kante Shape[i+1] |
| Negative Gaps | Korrekt berechnet (überlappende Shapes) und können auf 0 gesetzt werden |
| Sort | Bubble-Sort (kein WorksheetFunction in PPT-VBA verfügbar) |
| AnchorMode=1 | Gleichmäßige Verteilung innerhalb Gesamtbounds (wie PPT Distribute, aber mit Ziel-Gap) |
| Mindestanzahl | 2 Shapes (1 Gap), kein Minimum von 3 – auch mit 2 Shapes sinnvoll für Custom-Gap |

---

## [Schritt 9] – User-Name-Stempel (2026-04-03)

### Neu

- **`src/Modules/modUserStamp.bas`**: User-Stempel Modul (keine Form):
  - `InsertUserStamp` (Public, Ribbon-Callback): Fragt Scope per Ja/Nein/Abbruch-Dialog (Aktuelle / Alle / Abbruch), ruft `BuildStampText` und `AddStampToSlide`
  - `RemoveUserStamps` (Public, Ribbon-Callback): Zählt Stempel, bestätigt per MsgBox, ruft `DeleteAllStamps`
  - `AddStampToSlide(sld, stampText)` (Public): Entfernt vorhandenen Stempel auf der Folie, fügt neue Textbox ein (unten rechts, 7.5 pt Calibri grau, kein Rahmen/Hintergrund), setzt Tag `InfrontUserStamp=1`
  - `BuildStampText()` (Public): `Application.UserName + " | " + Format(Now, "DD.MM.YYYY  HH:MM")`; falls Username leer → InputBox
  - `RemoveStampFromSlide(sld)` (Private): Rückwärts-Löschung aller Shapes mit Tag `InfrontUserStamp=1`
  - `DeleteAllStamps()` (Private): Entfernt alle Stempel aus allen Folien, gibt Anzahl zurück
  - `CountStamps()` (Private): Zählt vorhandene Stempel
  - `AskScope(dlgTitle)` (Private): Ja=Aktuelle / Nein=Alle / Abbruch → -1
  - `GetScopeSlides(scope)` (Private): SlideRange für Scope 0/1/2
- **`src/CustomUI/CustomUI.xml`**: Neue Gruppe `InfrontReviewGroup` (label="Review") **vor** `InfrontQualityGroup` mit zwei Buttons: `InsertUserStampButton` (→ `InsertUserStamp`) und `RemoveUserStampsButton` (→ `RemoveUserStamps`); `TabViewInfrontReviewGroup` entsprechend vor `TabViewInfrontQualityGroup`.

### Technische Entscheidungen

| Thema | Entscheidung |
|---|---|
| Keine Form nötig | Scope-Abfrage per `vbYesNoCancel`-MsgBox (Ja=Aktuell / Nein=Alle / Abbruch) |
| Idempotenz | Vorhandener Stempel auf derselben Folie wird vor Neusetzen gelöscht |
| UserName | `Application.UserName`; leer → InputBox einmalig |
| Stempel-Design | 7.5 pt Calibri, RGB(150,150,150), kein Rahmen, kein Hintergrund, rechts-bündig |
| Tag-Erkennung | `Shape.Tags("InfrontUserStamp") = "1"` – persistiert in .pptx |

---

## [Schritt 8] – Infront Master-Importer (2026-04-03)

### Neu

- **`src/Modules/modMasterImport.bas`**: Master-Importer Modul:
  - Typ `ImportOptions` (Public): `ApplyToAllSlides`, `ApplyToSelectedSlides`, `RemoveUnusedAfterImport`
  - `ShowMasterImport` (Public, Ribbon-Callback): Öffnet Form modeless
  - `LoadMastersFromFile(filePath, masterNames(), masterCount)` (Public): Öffnet Quelldatei `ReadOnly + WithWindow:=False`, liest `SlideMasters.Count` und Namen aus, schließt sofort
  - `ImportMaster(srcPath, masterIndex, opts)` (Public): Öffnet Quelldatei, kopiert Master über Folienkopie-Trick (`Slides(1).Copy` + `Paste` → neuer Master), benennt Master, löscht temporäre Folie, wendet Master optional auf Folien an, entfernt optional ungenutzte Masters
  - `ApplyMasterToSlides(newMaster, opts)` (Private): Wendet `CustomLayouts(1)` des neuen Masters auf alle / selektierte Folien an
  - `CleanUpUnusedMasters()` (Public): Ermittelt verwendete Masters über `sld.CustomLayout.Parent.Index`, löscht rückwärts alle ungenutzten
  - `BrowseForFile()` (Public): Windows `msoFileDialogFilePicker` gefiltert auf PPT-Dateien; Mac InputBox; Fallback InputBox bei Fehler
- **`src/Forms/frmMasterImport.frm`**: Steuerform:
  - `InitForm()`: Setzt Defaults, deaktiviert `btnImport`
  - `btnBrowse_Click`: Ruft `BrowseForFile`, befüllt `txtSourceFile`
  - `btnLoadMasters_Click`: Ruft `LoadMastersFromFile`, befüllt `lstMasters`, aktiviert `btnImport`
  - `chkApplyAll_Click` / `chkApplySelected_Click`: Gegenseitiger Ausschluss
  - `btnImport_Click`: Baut `ImportOptions`, ruft `ImportMaster`, zeigt Ergebnis in `lblStatus`
  - Controls müssen in VBA-IDE angelegt werden (kein .frx)
- **`src/CustomUI/CustomUI.xml`**: `MasterImportButton` nach `AgendaWizardButton` in `InfrontSlidesGroup`; `TabViewMasterImportButton` entsprechend in `TabViewInfrontSlidesGroup`.

### Technische Entscheidungen

| Thema | Entscheidung |
|---|---|
| Master-Kopier-Methode | Folienkopie-Trick: `srcPres.Slides(1).Copy` + `Paste` zieht den Master mit; temporäre Folie wird gelöscht |
| ReadOnly-Öffnen | `WithWindow:=msoFalse` verhindert UI-Flash und verhindert versehentliche Bearbeitung der Quelldatei |
| Gegenseitiger Ausschluss | `chkApplyAll` und `chkApplySelected` deaktivieren sich gegenseitig im Click-Handler |
| `CleanUpUnusedMasters` | Rückwärts-Löschung verhindert Index-Verschiebung; prüft via `CustomLayout.Parent.Index` |
| Mac-Dateiauswahl | InputBox statt AppleScriptTask (kein Deployment-Aufwand) |

---

## [Schritt 7] – Agenda Wizard (2026-04-03)

### Neu

- **`src/Modules/modAgendaWizard.bas`**: Agenda Wizard Modul:
  - Typ `AgendaConfig` (Public): `Title`, `Items()`, `ItemCount`, `ActiveColor`, `InactiveColor`, `DoneColor`, `TitleFontSize`, `ItemFontSize`, `InsertionMode` (0/1), `InsertAfterSlide`
  - `ShowAgendaWizard` (Public, Ribbon-Callback): Öffnet Form modeless
  - `GenerateAgenda(cfg)` (Public): Löscht vorhandene Agenda-Folien, fügt Master-Übersicht und (bei Modus 1) Fortschrittsfolien vor jeder Sektion ein
  - `InsertAgendaSlide(insertPos, cfg, activeIdx)` (Private): Erstellt Folie ohne Layout-Abhängigkeit (alle Shapes per VBA positioniert); markiert Folie mit Tag `InfrontAgenda=1`; Farblogik: activeIdx=-1 → alle gleich aktiv, activeIdx=i → i aktiv/hervorgehoben, <i → erledigt (DoneColor), >i → inaktiv (InactiveColor)
  - `DeleteExistingAgendaSlides()` (Public): Löscht rückwärts alle Folien mit Tag `InfrontAgenda=1`; idempotent
  - `ParseItemList(raw, items(), itemCount)` (Public): Normalisiert `vbCrLf`/`vbCr`/`vbLf`, splittet, überspringt Leerzeilen
  - `CountAgendaSlides()` (Public): Zählt vorhandene Agenda-Folien (für Form-Status)
- **`src/Forms/frmAgendaWizard.frm`**: Steuerform:
  - `InitForm()`: Setzt Standardwerte, ruft `UpdateStatus`
  - `btnGenerate_Click`: Parst Items, baut AgendaConfig, bestätigt nicht (kein Massen-Risiko), ruft `GenerateAgenda`
  - `btnDelete_Click`: Bestätigung, ruft `DeleteExistingAgendaSlides`
  - `btnPickActive/Inactive/Done_Click`: Ruft `PickColorHex` (InputBox mit aktuellem Wert als Vorschlag)
  - Controls müssen in VBA-IDE angelegt werden (kein .frx)
- **`src/CustomUI/CustomUI.xml`**: Neue Gruppe `InfrontSlidesGroup` (label="Slides") vor `TableGroup` im Single-Tab-View; `TabViewInfrontSlidesGroup` am Anfang des `TabViewInstrumentaTables`-Tabs. Button: `AgendaWizardButton` → `ShowAgendaWizard`.

### Technische Entscheidungen

| Thema | Entscheidung |
|---|---|
| Idempotenz | Tag `InfrontAgenda=1` auf jeder Agenda-Folie → Neugenerierung löscht alte zuverlässig |
| Layout-Unabhängigkeit | Alle Shapes per `AddTextbox` / `AddLine` positioniert – kein ppLayoutTitle o.ä. |
| Farbangabe | Hex-InputBox (kein nativer Color-Dialog auf Mac); `btnPick*` als Wrapper |
| Fortschrittsfolien | activeIdx=i: aktueller Punkt fett+ActiveColor, abgehakte Punkte DoneColor, zukünftige InactiveColor |
| Folientag | `Slide.Tags.Add "InfrontAgenda", "1"` – persistiert in .pptx/.ppam |

---

## [Schritt 6] – Global Find & Replace (2026-04-03)

### Neu

- **`src/Modules/modFindReplace.bas`**: Suchen & Ersetzen Modul:
  - Typ `FindReplaceOptions` (Public): `FindText`, `ReplaceText`, `MatchCase`, `WholeWord`, `Scope` (0/1/2), `IncludeNotes`, `TargetShapes` (0/1/2)
  - `ShowFindReplace` (Public, Ribbon-Callback): Öffnet Form modeless
  - `ExecuteReplace(opts)` (Public): Traversiert Slides per Scope, je Shape rekursiv/zellenweise/run-weise, gibt Anzahl Ersetzungen zurück
  - `CountMatches(opts)` (Public): Identische Traversierung ohne Ersatz (Preview)
  - `ReplaceInShape` / `CountInShape` (Private): Gruppen rekursiv, Tabellen zellenweise, TextFrame an `ReplaceInTextRange` delegiert
  - `ReplaceInTextRange` (Public): Run-weiser Ersatz – Formatierung (Fett/Kursiv/Farbe) jedes Runs bleibt erhalten; Treffer über Run-Grenzen werden bewusst nicht ersetzt
  - `ReplaceString` (Private): Eigene InStr-Schleife mit MatchCase + WholeWord-Unterstützung (kein `VBA.Replace` da kein WholeWord)
  - `IsWordChar` (Private): Wortzeichen-Test für WholeWord-Logik
  - `GetScopeSlides` (Private): Gibt SlideRange für Scope 0/1/2 zurück
  - `ShapeMatchesTarget` (Private): Filter für Platzhalter vs. Textboxen vs. Alle
- **`src/Forms/frmFindReplace.frm`**: Steuerform:
  - `UserForm_Initialize`: Setzt Defaults (Alle Folien, Alle Shapes)
  - `btnPreview_Click`: Zählt Treffer, zeigt in `lblResult`
  - `btnReplaceAll_Click`: Bestätigung bei Scope=Alle, führt aus, zeigt Ergebnis
  - Enter-Key-Handling: txtFind→txtReplace→Ersetzen
  - Controls müssen in VBA-IDE angelegt werden (kein .frx)
- **`src/CustomUI/CustomUI.xml`**: `FindReplaceButton` nach `ReplaceDialog` in `AllGroup` (Single + Multi-Tab)

### Technische Entscheidungen

| Thema | Entscheidung |
|---|---|
| Formatierungserhalt | Run-weiser Ersatz: Formatierung bleibt je Run erhalten; Treffer über Run-Grenzen werden nicht ersetzt (dokumentiert) |
| WholeWord | Eigene `IsWordChar`-Prüfung statt Regex (nicht in PPT-VBA verfügbar) |
| Sprechernotizen | Optional über `chkIncludeNotes`; greift auf `sld.NotesPage.Shapes` zu |
| Bestätigung | Nur bei Scope=Alle Folien um versehentliche Massen-Ersetzungen zu verhindern |

---

## [Schritt 5] – Format Painter Plus (2026-04-03)

### Neu

- **`src/Modules/modFormatPainterPlus.bas`**: Format Painter Plus Modul:
  - Typen `FormatSnapshot` und `ApplyOptions` (Public) – von der Form lesbar/befüllbar
  - `ShowFormatPainterPlus` (Public, Ribbon-Callback): Prüft genau 1 selektiertes Source-Shape, ruft `CaptureFormat` auf, öffnet Form modeless
  - `CaptureFormat(shp)` (Public): Liest Fill (Type/Color/Transparency), Line (Visible/Color/Weight/Dash), Font (Name/Size/Bold/Italic/Underline/Color, aus erstem Run des ersten Paragraphen), TextAlign (H/V), ShapeWidth/Height in `g_Snapshot`
  - `ApplyFormatToSelection(opts)` (Public): Iteriert selektierte Shapes, ruft `ApplyToShape` auf, Zusammenfassung per MsgBox
  - `ApplyToShape(shp, opts)` (Private): Wendet jede Eigenschaft einzeln mit `On Error Resume Next` an – keine Abstürze bei nicht unterstützten Shape-Typen
- **`src/Forms/frmFormatPainterPlus.frm`**: Steuerform mit 14 Checkboxen in 5 Frames (Füllung / Linie / Schrift / Ausrichtung / Größe):
  - `InitForm()`: Befüllt `lblSourceInfo` mit Kurzübersicht der gecapturten Werte, setzt Checkboxen (alle außer Breite/Höhe standardmäßig aktiviert)
  - `btnApply_Click`: Liest Checkboxen, baut `ApplyOptions`, ruft `modFormatPainterPlus.ApplyFormatToSelection`
  - `btnSelectAll_Click` / `btnNone_Click`: Alle aktivieren / deaktivieren
  - Controls müssen in VBA-IDE angelegt werden (kein .frx – Projektkonvention)
- **`src/CustomUI/CustomUI.xml`**: `FormatPainterPlusButton` nach `ColorPickerButton` in `InfrontFormatGroup`; `TabViewFormatPainterPlusButton` in `TabViewInfrontFormatGroup`. `getEnabled="EnableWhenExactlyOneShape"` (Quell-Selektion).

### Technische Entscheidungen

| Thema | Entscheidung |
|---|---|
| Font-Capture | Nur erster Run des ersten Paragraphen – repräsentiert den dominanten Stil |
| `On Error Resume Next` pro Property | Vermeidet Abbruch bei Shapes ohne Fill/Line/TextFrame |
| Größe (Breite/Höhe) | Standardmäßig deaktiviert – ungewolltes Resize-Risiko zu hoch |
| Modeless Form | Nutzer kann Shapes selektieren ohne Form zu schließen |
| `EnableWhenExactlyOneShape` | Callback aus Instrumenta-Basis – stellt sicher, dass genau 1 Quell-Shape vorliegt |

---

## [Schritt 4] – Brand Compliance Checker (2026-04-02)

### Neu

- **`Infront_BrandConfig.ini`** (Repo-Wurzel): Kommentierte Beispiel-Konfigurationsdatei mit zwei Profilen (`Default`, `Strict`). Felder: `ActiveProfile`, `Name`, `AllowedFonts`, `AllowedColors`, `ColorTolerance`, `MinFontSizePt`.
- **`src/Modules/modBrandCompliance.bas`**: Vollständiges Modul für den Brand Compliance Check:
  - `ShowBrandCheck` (Public, Ribbon-Callback): Lädt Profil aus INI, iteriert alle Slides/Shapes, öffnet `frmBrandCompliance` im modeless-Modus.
  - `GetConfigPath` (Public): Ermittelt INI-Pfad über `ThisPresentation.Path`, Fallback: AddIns-Kollektion nach "infront" durchsuchen.
  - `LoadProfile` / `CreateDefaultConfig` (Private): INI-Profil laden bzw. Vorlage erstellen wenn keine INI vorhanden.
  - `ReadIniValue` / `WriteIniValue` (Public): Vollständiger INI-Parser/Writer mit `Open/Line Input/Print#/Close` – kein FSO.
  - `ParseColorList` / `ParseFontList` (Private): Kommagetrennte Listen aus INI in Array parsen.
  - `CheckShape` (Private, rekursiv): Traversiert Gruppen (`shp.GroupItems`), leitet Tabellen an `CheckTable` weiter, prüft FillColor/LineColor/TextFrame.
  - `CheckTable` (Private): Prüft alle Zellen einer Tabelle (Fill + Text).
  - `CheckTextFrame` (Private): Iteriert Paragraphen und Runs; prüft `Font.Name` und `Font.Size`.
  - `IsColorAllowed` / `ColorMaxChannelDiff` / `NearestAllowedColor` (Private/Public): Farb-Toleranzprüfung per maximaler Kanal-Differenz (0–30 Range).
  - `IsFontAllowed` (Private): Groß-/Kleinschreibungstoleranter Schriftart-Vergleich.
  - `AddViolation` (Private): Dynamisches Array `g_Violations` mit automatischem `ReDim Preserve`.
  - `ExportViolationsToCSV` (Public): CSV-Export mit Windows-FileDialog / Mac-InputBox-Fallback, Semikolon-Trenner, `CsvEscape`-Maskierung.
  - `FixViolation` (Public): Behebt einzelnen Verstoß automatisch (nächste erlaubte Farbe / ersten erlaubten Font / MinFontSizePt).
  - `ColorToHexStr` / `AllowedColorsAsString` / `AllowedFontsAsString` / `CsvEscape` (Public/Private): Hilfs- und Formatierungsfunktionen.
- **`src/Forms/frmBrandCompliance.frm`**: Ergebnisform mit 6-spaltiger ListBox (`lstViolations`), Zusammenfassungs-Label (`lblSummary`), Buttons: `btnGoToSlide`, `btnFixSelected`, `btnExportCSV`, `btnClose`. Controls müssen in VBA-IDE angelegt werden (kein .frx – Projektkonvention).
- **`src/CustomUI/CustomUI.xml`**: Neue Gruppe `InfrontQualityGroup` (label="Quality") vor der Advanced-Gruppe im Single-Tab-View; `TabViewInfrontQualityGroup` entsprechend im Multi-Tab-View. Erster Button: `BrandCheckButton` → `ShowBrandCheck`.

### Technische Entscheidungen

| Thema | Entscheidung |
|---|---|
| Farb-Toleranz | Maximale Kanal-Differenz (nicht Euklidisch), Bereich 0–30 |
| INI-Parser | Open/Line Input/Close – kein FSO (Plattform-Anforderung) |
| Mac CSV-Export | InputBox mit Desktop-Vorschlag (kein SaveAs-Dialog ohne AppleScriptTask) |
| Gruppen-Traversierung | Rekursiv über `GroupItems` |
| Tabellen | `shp.HasTable` → `shp.Table.Cell(r,c)` |
| UndoRecord | Nicht verfügbar in PowerPoint VBA; `FixViolation` erstellt automatisch Undo-Einträge per Shape-Änderung |
| Konfigurationspfad | `ThisPresentation.Path` + `Application.PathSeparator` (kein hardcodierter Separator) |

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
