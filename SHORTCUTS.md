# Infront Toolkit – Tastaturkürzel

Dokumentation der verfügbaren Tastaturpfade für das Infront Toolkit auf dem Mac.

---

## Wichtige Einschränkung

**Office Add-ins können auf dem Mac keine globalen Tastaturkürzel für PowerPoint registrieren.**

- Es gibt keine API in Office.js für `Application.OnKey` (VBA-Äquivalent)
- Shortcuts können nur innerhalb der Task Pane (wenn aktiv und fokussiert) genutzt werden
- Ribbon-Buttons sind der primäre Zugangspfad für alle Features

Status in TESTING.md dokumentiert: Kategorie C – aktuell nicht robust umsetzbar.

---

## Primärer Zugangspfad: Ribbon

Alle Features sind über den **Infront Toolkit** Ribbon-Tab erreichbar:

| Feature | Gruppe | Schritt |
|---|---|---|
| Eckenradius | Shapes | 4 |
| Farbwähler | Format | 5 |
| Format Painter+ | Format | 7 |
| Brand Check | Quality | 6 |
| Suchen & Ersetzen | Quality | 8 |
| Kommentare entfernen | Quality | 11 |
| Red Box entfernen | Quality | 13 |
| Gap H | Ausrichten | 12 |
| Gap V | Ausrichten | 12 |
| Gap... | Ausrichten | 12 |
| Agenda | Struktur | 9 |
| Master importieren | Design | 10 |
| Red Box | Design | 13 |
| Red Box: Alle Slides | Design | 13 |
| Red Box Einstellungen | Design | 13 |
| Kommentar | Review | 11 |
| Markieren | Review | 11 |
| Meine Kommentare | Review | 11 |

---

## Task-Pane-interne Shortcuts

Wenn die Task Pane geöffnet und fokussiert ist, reagiert sie auf Standard-Browser-Shortcuts:

| Shortcut (Mac) | Funktion |
|---|---|
| `Tab` / `Shift+Tab` | Zwischen Bedienelementen wechseln |
| `Return` / `Space` | Aktuellen Button aktivieren |
| `Escape` | Dropdown/Dialog schließen |
| `⌘+A` | Text in Eingabefeld alles auswählen |
| `⌘+C` / `⌘+V` | Kopieren/Einfügen in Eingabefeldern |

---

## Keyboard Access via Mac Accessibility

Als Workaround für schnellen Ribbon-Zugriff auf dem Mac:

1. **Vollzugriff aktivieren**: Systemeinstellungen → Bedienungshilfen → Tastatur → Vollzugriff aktivieren
2. Mit `F6` / `Ctrl+F6` zwischen Ribbon, Task Pane und Presentation wechseln
3. Im Ribbon: mit Pfeiltasten und `Return` navigieren

---

## Geplant (Schritt 15)

Schritt 15 analysiert welche weiteren Tastaturpfade innerhalb der Office Add-in Architektur realistisch umsetzbar sind und dokumentiert diese vollständig.

| Implementierungsform | Realistisch auf Mac |
|---|---|
| Globale PowerPoint-Shortcuts | Nein (keine API) |
| Add-in-Keyboard-Shortcuts (VersionOverrides) | Begrenzt (nur Outlook unterstützt KeyboardShortcuts vollständig) |
| Task-Pane-interne Shortcuts | Ja (Standard-Web-Shortcuts) |
| Ribbon-KeyTip-Navigation | Ja (über Accessibility-Modus) |
