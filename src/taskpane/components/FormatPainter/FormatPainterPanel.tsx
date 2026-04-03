/**
 * FormatPainterPanel.tsx
 * Feature-Panel: Format Painter+.
 * Vollständige Implementierung: Schritt 7.
 */

import * as React from "react";
import { Stack } from "@fluentui/react/lib/Stack";
import { Text } from "@fluentui/react/lib/Text";
import { PrimaryButton, DefaultButton } from "@fluentui/react/lib/Button";
import { Checkbox } from "@fluentui/react/lib/Checkbox";
import NotificationBar, { NotificationType } from "../shared/NotificationBar";

interface FormatOptions {
  fill:     boolean;
  line:     boolean;
  text:     boolean;
  geometry: boolean;
  shadow:   boolean;
}

interface Notification {
  message: string;
  type: NotificationType;
}

const DEFAULT_OPTIONS: FormatOptions = {
  fill:     true,
  line:     true,
  text:     false,
  geometry: false,
  shadow:   false,
};

const FormatPainterPanel: React.FC = () => {
  const [options, setOptions]           = React.useState<FormatOptions>(DEFAULT_OPTIONS);
  const [notification, setNotification] = React.useState<Notification | null>(null);
  const [isRunning, setIsRunning]       = React.useState(false);

  const toggleOption = (key: keyof FormatOptions) => {
    setOptions((prev) => ({ ...prev, [key]: !prev[key] }));
  };

  const selectAll  = () => setOptions({ fill: true, line: true, text: true, geometry: true, shadow: true });
  const selectNone = () => setOptions({ fill: false, line: false, text: false, geometry: false, shadow: false });

  const handleApply = async () => {
    setIsRunning(true);
    setNotification(null);

    try {
      await PowerPoint.run(async (context) => {
        const selection = context.presentation.getSelectedShapes();
        selection.load("items");
        await context.sync();

        if (selection.items.length < 2) {
          setNotification({ message: "Bitte Quell-Shape und Ziel-Shape(s) selektieren.", type: "warning" });
          return;
        }

        // Vollständige Implementierung in Schritt 7
        setNotification({
          message: "Format Painter+ – vollständige Implementierung folgt in Schritt 7.",
          type: "info",
        });
      });
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setNotification({ message: `Fehler: ${message}`, type: "error" });
    } finally {
      setIsRunning(false);
    }
  };

  return (
    <Stack className="panel-container" tokens={{ childrenGap: 12 }}>
      <Text className="panel-title">Format Painter+</Text>
      <Text className="panel-description">
        Selektiere zuerst das Quell-Shape, dann (mit Shift/Cmd) die Ziel-Shapes.
      </Text>

      <NotificationBar
        message={notification?.message ?? null}
        type={notification?.type}
        onDismiss={() => setNotification(null)}
      />

      <Text variant="smallPlus" style={{ fontWeight: 600 }}>Zu übertragende Eigenschaften:</Text>
      <Stack tokens={{ childrenGap: 6 }}>
        <Checkbox label="Füllung"    checked={options.fill}     onChange={() => toggleOption("fill")} />
        <Checkbox label="Linie"      checked={options.line}     onChange={() => toggleOption("line")} />
        <Checkbox label="Text"       checked={options.text}     onChange={() => toggleOption("text")} />
        <Checkbox label="Geometrie"  checked={options.geometry} onChange={() => toggleOption("geometry")} />
        <Checkbox label="Schatten"   checked={options.shadow}   onChange={() => toggleOption("shadow")} />
      </Stack>

      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <DefaultButton text="Alle auswählen" onClick={selectAll}  styles={{ root: { flex: 1 } }} />
        <DefaultButton text="Keine"          onClick={selectNone} styles={{ root: { flex: 1 } }} />
      </Stack>

      <PrimaryButton
        text={isRunning ? "Wird angewendet…" : "Auf Selektion anwenden"}
        onClick={handleApply}
        disabled={isRunning}
        styles={{ root: { width: "100%", marginTop: 8 } }}
      />

      <Text variant="small" style={{ color: "#888" }}>
        Vollständige Implementierung: Schritt 7.
      </Text>
    </Stack>
  );
};

export default FormatPainterPanel;
