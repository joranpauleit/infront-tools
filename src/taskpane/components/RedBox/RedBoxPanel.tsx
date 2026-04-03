/**
 * RedBoxPanel.tsx
 * Feature-Panel: Red Box Einstellungen.
 * Vollständige Implementierung: Schritt 13.
 */

import * as React from "react";
import { Stack } from "@fluentui/react/lib/Stack";
import { Text } from "@fluentui/react/lib/Text";
import { PrimaryButton } from "@fluentui/react/lib/Button";
import { TextField } from "@fluentui/react/lib/TextField";
import NotificationBar, { NotificationType } from "../shared/NotificationBar";
import ColorSwatch from "../shared/ColorSwatch";

interface RedBoxConfig {
  top:    number;
  right:  number;
  bottom: number;
  left:   number;
  color:  string;
  weight: number;
}

interface Notification {
  message: string;
  type: NotificationType;
}

const DEFAULT_CONFIG: RedBoxConfig = {
  top:    20,
  right:  20,
  bottom: 20,
  left:   20,
  color:  "#FF0000",
  weight: 1.5,
};

const RedBoxPanel: React.FC = () => {
  const [config, setConfig]             = React.useState<RedBoxConfig>(DEFAULT_CONFIG);
  const [notification, setNotification] = React.useState<Notification | null>(null);
  const [isRunning, setIsRunning]       = React.useState(false);

  const updateField = (field: keyof RedBoxConfig, value: string) => {
    const num = parseFloat(value);
    setConfig((prev) => ({ ...prev, [field]: isNaN(num) ? value : num }));
  };

  const handleSave = async () => {
    setIsRunning(true);
    setNotification(null);

    try {
      // Vollständige Implementierung in Schritt 13
      // Hier: Konfiguration in Document Settings speichern
      await PowerPoint.run(async (context) => {
        // Platzhalter – Schritt 13 implementiert persistente Settings
        void context;
        setNotification({
          message: "Einstellungen gespeichert – vollständige Implementierung: Schritt 13.",
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
      <Text className="panel-title">Red Box Einstellungen</Text>
      <Text className="panel-description">
        Konfiguriert die Safe-Area-Begrenzungsbox (Name: INFRONT_REDBOX).
      </Text>

      <NotificationBar
        message={notification?.message ?? null}
        type={notification?.type}
        onDismiss={() => setNotification(null)}
      />

      <Text variant="smallPlus" style={{ fontWeight: 600 }}>Abstände (in pt):</Text>
      <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
        {(["top", "right", "bottom", "left"] as const).map((side) => (
          <TextField
            key={side}
            label={side === "top" ? "Oben" : side === "right" ? "Rechts" : side === "bottom" ? "Unten" : "Links"}
            value={String(config[side])}
            onChange={(_e, val) => updateField(side, val ?? "")}
            type="number"
            min={0}
            suffix="pt"
            styles={{ root: { width: 90 } }}
          />
        ))}
      </Stack>

      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
        <TextField
          label="Farbe (Hex):"
          value={config.color}
          onChange={(_e, val) => updateField("color", val ?? "")}
          styles={{ root: { flex: 1 } }}
        />
        <ColorSwatch color={config.color} size={28} />
      </Stack>

      <TextField
        label="Linienstärke (pt):"
        value={String(config.weight)}
        onChange={(_e, val) => updateField("weight", val ?? "")}
        type="number"
        min={0.5}
        step={0.5}
        suffix="pt"
        styles={{ root: { maxWidth: 150 } }}
      />

      <PrimaryButton
        text={isRunning ? "Wird gespeichert…" : "Einstellungen speichern"}
        onClick={handleSave}
        disabled={isRunning}
        styles={{ root: { width: "100%", marginTop: 8 } }}
      />
    </Stack>
  );
};

export default RedBoxPanel;
