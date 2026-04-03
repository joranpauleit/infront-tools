/**
 * GapEqualizerPanel.tsx
 * Feature-Panel: Gap-Equalizer (erweiterte Optionen).
 * Vollständige Implementierung: Schritt 12.
 */

import * as React from "react";
import { Stack } from "@fluentui/react/lib/Stack";
import { Text } from "@fluentui/react/lib/Text";
import { PrimaryButton } from "@fluentui/react/lib/Button";
import { TextField } from "@fluentui/react/lib/TextField";
import { ChoiceGroup, IChoiceGroupOption } from "@fluentui/react/lib/ChoiceGroup";
import NotificationBar, { NotificationType } from "../shared/NotificationBar";

type GapDirection = "horizontal" | "vertical" | "both";

interface Notification {
  message: string;
  type: NotificationType;
}

const DIRECTION_OPTIONS: IChoiceGroupOption[] = [
  { key: "horizontal", text: "Horizontal (Gap H)" },
  { key: "vertical",   text: "Vertikal (Gap V)" },
  { key: "both",       text: "Beide" },
];

const GapEqualizerPanel: React.FC = () => {
  const [direction, setDirection]       = React.useState<GapDirection>("horizontal");
  const [fixedGap, setFixedGap]         = React.useState<string>("");
  const [notification, setNotification] = React.useState<Notification | null>(null);
  const [isRunning, setIsRunning]       = React.useState(false);

  const handleApply = async () => {
    setIsRunning(true);
    setNotification(null);

    try {
      await PowerPoint.run(async (context) => {
        const selection = context.presentation.getSelectedShapes();
        selection.load("items");
        await context.sync();

        if (selection.items.length < 3) {
          setNotification({ message: "Bitte mindestens 3 Shapes selektieren.", type: "warning" });
          return;
        }

        // Vollständige Implementierung in Schritt 12
        setNotification({
          message: `Gap-Equalizer – vollständige Implementierung folgt in Schritt 12.`,
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
      <Text className="panel-title">Gap-Equalizer</Text>
      <Text className="panel-description">
        Gleicht Kantenabstände zwischen selektierten Shapes an.
        Äußere Shapes bleiben fixiert.
      </Text>

      <NotificationBar
        message={notification?.message ?? null}
        type={notification?.type}
        onDismiss={() => setNotification(null)}
      />

      <ChoiceGroup
        label="Richtung:"
        options={DIRECTION_OPTIONS}
        selectedKey={direction}
        onChange={(_e, opt) => opt && setDirection(opt.key as GapDirection)}
      />

      <TextField
        label="Fester Abstand (optional, in pt):"
        value={fixedGap}
        onChange={(_e, val) => setFixedGap(val ?? "")}
        placeholder="leer = gleichmäßig verteilen"
        type="number"
        min={0}
        suffix="pt"
        styles={{ root: { maxWidth: 220 } }}
      />

      <PrimaryButton
        text={isRunning ? "Wird angewendet…" : "Abstände angleichen"}
        onClick={handleApply}
        disabled={isRunning}
        styles={{ root: { width: "100%", marginTop: 8 } }}
      />

      <Text variant="small" style={{ color: "#888" }}>
        Vollständige Implementierung: Schritt 12.
      </Text>
    </Stack>
  );
};

export default GapEqualizerPanel;
