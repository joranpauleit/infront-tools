# Infront Toolkit – Tastaturkürzel

Dokumentation der verfügbaren Tastaturpfade für das Infront Toolkit auf dem Mac.

---

## Wichtige Einschränkung: Keine globalen Shortcuts

**Office Add-ins können auf dem Mac keine globalen Tastaturkürzel für PowerPoint registrieren.**

| Technik | Verfügbar auf Mac | Grund |
|---|---|---|
| `Application.OnKey` (VBA-Äquivalent) | Nein | Kein API in Office.js |
| `KeyboardShortcuts` in VersionOverrides (Office Add-in API) | Nein\* | Nur in Outlook vollständig unterstützt |
| Task-Pane-interne Shortcuts | **Ja** | Standard-Web-Shortcuts im WKWebView |
| Ribbon KeyTip-Navigation | **Ja** | Mac-Accessibility-Modus |
| Custom Ribbon-Tastenkürzel | Nein | Nicht von Add-ins konfigurierbar |

\* *Stand April 2026. Für PowerPoint-Desktop auf Windows existiert ein experimentelles `KeyboardShortcuts`-API – auf Mac noch nicht verfügbar.*

Dokumentiert in `TESTING.md`: Kategorie A – grundsätzlich nicht umsetzbar.

---

## Primärer Zugangspfad: Ribbon

Alle Features sind über den Tab **„Infront Toolkit"** erreichbar.
Ribbon-Buttons sind die empfohlene primäre Interaktionsform.

### Gruppe: Shapes

| Button | Typ | Aktion |
|---|---|---|
| Eckenradius | ShowTaskpane | Öffnet Corner-Radius-Panel (`?view=corner-radius`) |

### Gruppe: Format

| Button | Typ | Aktion |
|---|---|---|
| Farbwähler | ShowTaskpane | Öffnet Color-Picker-Panel (`?view=color-picker`) |
| Format Painter+ | ShowTaskpane | Öffnet Format-Painter-Panel (`?view=format-painter`) |

### Gruppe: Quality

| Button | Typ | Aktion |
|---|---|---|
| Brand Check | ShowTaskpane | Öffnet Brand-Check-Panel (`?view=brand-check`) |
| Suchen & Ersetzen | ShowTaskpane | Öffnet Find-Replace-Panel (`?view=find-replace`) |
| Kommentare entfernen | ExecuteFunction | Entfernt alle INFRONT_COMMENT_* + INFRONT_HIGHLIGHT_* |
| Red Box entfernen | ExecuteFunction | Entfernt alle INFRONT_REDBOX-Shapes |

### Gruppe: Ausrichten

| Button | Typ | Aktion |
|---|---|---|
| Gap H | ExecuteFunction | Gleicht horizontale Abstände an (≥3 Shapes) |
| Gap V | ExecuteFunction | Gleicht vertikale Abstände an (≥3 Shapes) |
| Gap... | ShowTaskpane | Öffnet Gap-Equalizer-Panel (`?view=gap-equalizer`) |

### Gruppe: Struktur

| Button | Typ | Aktion |
|---|---|---|
| Agenda | ShowTaskpane | Öffnet Agenda-Wizard-Panel (`?view=agenda`) |

### Gruppe: Design

| Button | Typ | Aktion |
|---|---|---|
| Master importieren | ShowTaskpane | Öffnet Master-Import-Panel (`?view=master-import`) |
| Red Box (Toggle) | ExecuteFunction | Schaltet INFRONT_REDBOX auf aktiver Folie ein/aus |
| Red Box: Alle Folien | ExecuteFunction | Fügt INFRONT_REDBOX auf allen Folien ein |
| Red Box Einstellungen | ShowTaskpane | Öffnet Red-Box-Panel (`?view=red-box`) |

### Gruppe: Review

| Button | Typ | Aktion |
|---|---|---|
| Kommentar | ShowTaskpane | Öffnet Review-Panel (`?view=review`) |
| Markieren | ExecuteFunction | Fügt INFRONT_HIGHLIGHT_*-Rechteck ein |
| Meine Kommentare | ShowTaskpane | Öffnet Kommentar-Übersicht (`?view=my-comments`) |

---

## Task-Pane-interne Shortcuts

Wenn die Task Pane geöffnet und fokussiert ist, stehen Standard-Browser-Shortcuts zur Verfügung:

| Shortcut (Mac) | Funktion |
|---|---|
| `Tab` | Zum nächsten Bedienelement |
| `Shift+Tab` | Zum vorherigen Bedienelement |
| `Return` / `Space` | Aktuellen Button/Checkbox aktivieren |
| `Escape` | Dropdown / Dialog schließen |
| `⌘+A` | Gesamten Text im Eingabefeld auswählen |
| `⌘+C` / `⌘+V` / `⌘+X` | Kopieren / Einfügen / Ausschneiden |
| `⌘+Z` | Undo im Eingabefeld (nicht PowerPoint-Undo) |
| Pfeiltasten | Navigation in ChoiceGroup / Dropdown |

### Pivot-Tab-Navigation (Fluent UI)

In Panels mit Tabs (z.B. Find & Replace, Review, Gap Equalizer):

| Shortcut | Funktion |
|---|---|
| `←` / `→` | Zwischen Tabs wechseln (wenn Tab-Bar fokussiert) |
| `Tab` | Fokus von Tab-Bar in Tab-Inhalt verschieben |

---

## Ribbon-Navigation via Mac Accessibility

Für schnellen Zugriff ohne Maus:

### Vorbereitung

1. **Vollzugriff aktivieren**: Systemeinstellungen → Bedienungshilfen → Tastatur → „Vollzugriff aktivieren"
2. Mit `Ctrl+F2` Menüleiste fokussieren (oder `F6` für Ribbon in Office)
3. Mit Pfeiltasten und `Return` durch Ribbon-Tabs und Buttons navigieren

### Navigation in PowerPoint Mac

| Tastenkürzel | Funktion |
|---|---|
| `Ctrl+F6` | Zwischen Bereichen wechseln (Ribbon, Folie, Task Pane) |
| `F6` | Zur nächsten Gruppe im Ribbon |
| `←` / `→` | Im Ribbon zwischen Buttons navigieren |
| `Return` | Button aktivieren |
| `Escape` | Ribbon verlassen |

---

## Workflow-Empfehlungen

### Schnellster Weg zu den wichtigsten Features

| Ziel | Empfohlener Weg |
|---|---|
| Abstände schnell angleichen | Ribbon → Gap H oder Gap V (ExecuteFunction, kein Task-Pane-Öffnen) |
| Red Box ein-/ausschalten | Ribbon → Design → Red Box (Toggle) |
| Brand Check starten | Ribbon → Quality → Brand Check → „Check starten" |
| Kommentar hinterlassen | Ribbon → Review → Kommentar |
| Farbe ersetzen (deck-weit) | Ribbon → Quality → Suchen & Ersetzen → Tab „Farbe" |

### Task Pane offen halten

Die Task Pane merkt sich die zuletzt geöffnete Ansicht während der Sitzung.
Beim Schließen + Wiederöffnen wird die Standard-Ansicht (Welcome) geladen.

---

## Geplante Verbesserungen (zukünftige Versionen)

| Feature | Status | Bedingung |
|---|---|---|
| `KeyboardShortcuts` API für PowerPoint Mac | Ausstehend | Wenn Microsoft API auf Mac erweitert |
| `keytip`-Attribute im Ribbon (Alt+Buchstabe) | Technisch möglich\* | Bereits in manifest.xml vorbereitet |
| Fokus-Trap in Modals | Implementiert (Fluent UI) | — |

\* *`keytip`-Attribute sind in `manifest.xml` für alle Ribbon-Buttons definiert (z.B. `keytip="E"` für Eckenradius). Auf Mac ist die Aktivierung über `Alt` jedoch systembedingt eingeschränkt.*
