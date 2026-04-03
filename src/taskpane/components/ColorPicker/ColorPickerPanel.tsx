/**
 * ColorPickerPanel.tsx
 * Feature-Panel: Farbwähler.
 * Vollständige Implementierung: Schritt 5.
 *
 * HINWEIS: Ein systemweiter Screen-Pixel-Picker ist im Office Add-in
 * auf Mac nicht möglich (WebKit/Safari unterstützt EyeDropper API nicht).
 * Fallback: Hex/RGB-Eingabe, Farbübernahme aus Shape, Recent-Colors-Palette.
 */

import * as React from "react";
import { Stack } from "@fluentui/react/lib/Stack";
import { Text } from "@fluentui/react/lib/Text";
import { PrimaryButton, DefaultButton } from "@fluentui/react/lib/Button";
import { TextField } from "@fluentui/react/lib/TextField";
import { ChoiceGroup, IChoiceGroupOption } from "@fluentui/react/lib/ChoiceGroup";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import ColorSwatch from "../shared/ColorSwatch";
import NotificationBar, { NotificationType } from "../shared/NotificationBar";

type ApplyTarget = "fill" | "line" | "font";

interface Notification {
  message: string;
  type: NotificationType;
}

const TARGET_OPTIONS: IChoiceGroupOption[] = [
  { key: "fill", text: "Füllfarbe" },
  { key: "line", text: "Linienfarbe" },
  { key: "font", text: "Schriftfarbe" },
];

const ColorPickerPanel: React.FC = () => {
  const [hexInput, setHexInput]           = React.useState<string>("#003366");
  const [applyTarget, setApplyTarget]     = React.useState<ApplyTarget>("fill");
  const [recentColors, setRecentColors]   = React.useState<string[]>([]);
  const [notification, setNotification]   = React.useState<Notification | null>(null);
  const [isRunning, setIsRunning]         = React.useState(false);

  /** Normalisiert und validiert einen Hex-Farbwert. */
  const normalizeHex = (raw: string): string | null => {
    const cleaned = raw.trim().replace(/^#/, "");
    if (/^[0-9A-Fa-f]{6}$/.test(cleaned)) return `#${cleaned.toUpperCase()}`;
    if (/^[0-9A-Fa-f]{3}$/.test(cleaned)) {
      const [r, g, b] = cleaned.split("");
      return `#${r}${r}${g}${g}${b}${b}`.toUpperCase();
    }
    return null;
  };

  /** Liest die Farbe des ersten selektierten Shapes und übernimmt sie ins Eingabefeld. */
  const handlePickFromShape = async () => {
    setIsRunning(true);
    try {
      await PowerPoint.run(async (context) => {
        const selection = context.presentation.getSelectedShapes();
        selection.load("items");
        await context.sync();

        if (selection.items.length === 0) {
          setNotification({ message: "Kein Shape selektiert.", type: "warning" });
          return;
        }

        const shape = selection.items[0];
        shape.fill.load("foregroundColor");
        await context.sync();

        const color = shape.fill.foregroundColor;
        if (color) {
          setHexInput(color.startsWith("#") ? color : `#${color}`);
          setNotification({ message: `Farbe übernommen: ${color}`, type: "success" });
        }
      });
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setNotification({ message: `Fehler: ${message}`, type: "error" });
    } finally {
      setIsRunning(false);
    }
  };

  /** Wendet die eingegebene Farbe auf alle selektierten Shapes an. */
  const handleApply = async () => {
    const hex = normalizeHex(hexInput);
    if (!hex) {
      setNotification({ message: "Ungültiger Hex-Wert. Format: #RRGGBB", type: "error" });
      return;
    }

    setIsRunning(true);
    setNotification(null);

    try {
      await PowerPoint.run(async (context) => {
        const selection = context.presentation.getSelectedShapes();
        selection.load("items");
        await context.sync();

        if (selection.items.length === 0) {
          setNotification({ message: "Keine Shapes selektiert.", type: "warning" });
          return;
        }

        let applied = 0;
        let skipped = 0;

        for (const shape of selection.items) {
          try {
            if (applyTarget === "fill") {
              shape.fill.setSolidColor(hex);
              applied++;
            } else if (applyTarget === "line") {
              shape.lineFormat.color = hex;
              applied++;
            } else if (applyTarget === "font") {
              shape.load("textFrame");
              await context.sync();
              if (shape.textFrame) {
                shape.textFrame.textRange.font.color = hex;
                applied++;
              } else {
                skipped++;
              }
            }
          } catch {
            skipped++;
          }
        }

        await context.sync();

        // Recent Colors aktualisieren
        setRecentColors((prev) => {
          const updated = [hex, ...prev.filter((c) => c !== hex)].slice(0, 8);
          return updated;
        });

        setNotification({
          message: `${applied} Shape(s) eingefärbt, ${skipped} übersprungen.`,
          type: applied > 0 ? "success" : "warning",
        });
      });
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setNotification({ message: `Fehler: ${message}`, type: "error" });
    } finally {
      setIsRunning(false);
    }
  };

  const currentHex = normalizeHex(hexInput);
  const rgbParts = currentHex
    ? [
        parseInt(currentHex.slice(1, 3), 16),
        parseInt(currentHex.slice(3, 5), 16),
        parseInt(currentHex.slice(5, 7), 16),
      ]
    : null;

  return (
    <Stack className="panel-container" tokens={{ childrenGap: 12 }}>
      <Text className="panel-title">Farbwähler</Text>

      <MessageBar messageBarType={MessageBarType.info} isMultiline styles={{ root: { marginBottom: 4 } }}>
        Systemweiter Pixel-Picker nicht verfügbar (Mac/WebKit-Einschränkung).
        Bitte Hex-Wert eingeben oder Farbe aus Shape übernehmen.
      </MessageBar>

      <NotificationBar
        message={notification?.message ?? null}
        type={notification?.type}
        onDismiss={() => setNotification(null)}
      />

      {/* Farb-Eingabe */}
      <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
        <TextField
          label="Hex-Farbwert:"
          value={hexInput}
          onChange={(_e, val) => setHexInput(val ?? "")}
          placeholder="#003366"
          styles={{ root: { flex: 1 } }}
        />
        {currentHex && (
          <ColorSwatch color={currentHex} size={32} showLabel={false} />
        )}
      </Stack>

      {/* RGB-Anzeige */}
      {rgbParts && (
        <Text variant="small" style={{ color: "#555", fontFamily: "monospace" }}>
          RGB({rgbParts[0]}, {rgbParts[1]}, {rgbParts[2]})
        </Text>
      )}

      {/* Ziel-Auswahl */}
      <ChoiceGroup
        label="Anwenden auf:"
        options={TARGET_OPTIONS}
        selectedKey={applyTarget}
        onChange={(_e, option) => option && setApplyTarget(option.key as ApplyTarget)}
      />

      {/* Aktionen */}
      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <PrimaryButton
          text={isRunning ? "Wird angewendet…" : "Anwenden"}
          onClick={handleApply}
          disabled={isRunning || !currentHex}
          styles={{ root: { flex: 1 } }}
        />
        <DefaultButton
          text="Aus Shape"
          onClick={handlePickFromShape}
          disabled={isRunning}
          title="Farbe des ersten selektierten Shapes übernehmen"
        />
      </Stack>

      {/* Zuletzt verwendete Farben */}
      {recentColors.length > 0 && (
        <Stack tokens={{ childrenGap: 6 }}>
          <Text variant="small" style={{ color: "#666" }}>Zuletzt verwendet:</Text>
          <Stack horizontal wrap tokens={{ childrenGap: 4 }}>
            {recentColors.map((c) => (
              <ColorSwatch
                key={c}
                color={c}
                size={24}
                onClick={(color) => setHexInput(color)}
                title={c}
              />
            ))}
          </Stack>
        </Stack>
      )}
    </Stack>
  );
};

export default ColorPickerPanel;
