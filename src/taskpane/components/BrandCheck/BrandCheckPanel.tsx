/**
 * BrandCheckPanel.tsx
 * Feature-Panel: Brand Compliance Check.
 * Vollständige Implementierung: Schritt 6.
 */

import * as React from "react";
import { Stack } from "@fluentui/react/lib/Stack";
import { Text } from "@fluentui/react/lib/Text";
import { PrimaryButton, DefaultButton } from "@fluentui/react/lib/Button";
import { Spinner } from "@fluentui/react/lib/Spinner";
import NotificationBar, { NotificationType } from "../shared/NotificationBar";

interface Notification {
  message: string;
  type: NotificationType;
}

const BrandCheckPanel: React.FC = () => {
  const [isRunning, setIsRunning]       = React.useState(false);
  const [notification, setNotification] = React.useState<Notification | null>(null);
  const [resultCount, setResultCount]   = React.useState<number | null>(null);

  const handleRun = async () => {
    setIsRunning(true);
    setNotification(null);
    setResultCount(null);

    try {
      await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        // Platzhalter – vollständige Logik in Schritt 6
        let violations = 0;
        for (const slide of slides.items) {
          const shapes = slide.shapes;
          shapes.load("items/name");
          await context.sync();
          violations += shapes.items.length; // Platzhalter-Zähler
        }

        setResultCount(violations);
        setNotification({
          message: `Prüfung abgeschlossen – vollständige Implementierung folgt in Schritt 6.`,
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
      <Text className="panel-title">Brand Compliance Check</Text>
      <Text className="panel-description">
        Prüft alle Shapes auf Einhaltung der Brand-Richtlinien (Schriften, Farben, Größen).
        Konfiguration in <code>config/Infront_BrandConfig.json</code>.
      </Text>

      <NotificationBar
        message={notification?.message ?? null}
        type={notification?.type}
        onDismiss={() => setNotification(null)}
      />

      {isRunning && <Spinner label="Deck wird geprüft…" />}

      {resultCount !== null && !isRunning && (
        <Text variant="small">Shapes geprüft: {resultCount}</Text>
      )}

      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <PrimaryButton
          text={isRunning ? "Prüft…" : "Jetzt prüfen"}
          onClick={handleRun}
          disabled={isRunning}
          styles={{ root: { flex: 1 } }}
        />
        <DefaultButton
          text="Konfiguration"
          disabled
          title="Konfigurationseditor – kommt in Schritt 6"
        />
      </Stack>

      <Text variant="small" style={{ color: "#888" }}>
        Vollständige Implementierung: Schritt 6.
      </Text>
    </Stack>
  );
};

export default BrandCheckPanel;
