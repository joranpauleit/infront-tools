/**
 * FindReplacePanel.tsx
 * Feature-Panel: Suchen & Ersetzen (Text, Farbe, Font).
 * Vollständige Implementierung: Schritt 8.
 */

import * as React from "react";
import { Stack } from "@fluentui/react/lib/Stack";
import { Text } from "@fluentui/react/lib/Text";
import { PrimaryButton } from "@fluentui/react/lib/Button";
import { TextField } from "@fluentui/react/lib/TextField";
import { Pivot, PivotItem } from "@fluentui/react/lib/Pivot";
import { Checkbox } from "@fluentui/react/lib/Checkbox";
import NotificationBar, { NotificationType } from "../shared/NotificationBar";

interface Notification {
  message: string;
  type: NotificationType;
}

const FindReplacePanel: React.FC = () => {
  const [searchText, setSearchText]     = React.useState("");
  const [replaceText, setReplaceText]   = React.useState("");
  const [caseSensitive, setCaseSensitive] = React.useState(false);
  const [notification, setNotification] = React.useState<Notification | null>(null);
  const [isRunning, setIsRunning]       = React.useState(false);

  const handleReplaceAll = async () => {
    if (!searchText) {
      setNotification({ message: "Bitte Suchtext eingeben.", type: "warning" });
      return;
    }

    setIsRunning(true);
    setNotification(null);

    try {
      await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        let replaced = 0;

        for (const slide of slides.items) {
          const shapes = slide.shapes;
          shapes.load("items/textFrame");
          await context.sync();

          for (const shape of shapes.items) {
            if (!shape.textFrame) continue;
            const tf = shape.textFrame;
            tf.textRange.load("text");
            await context.sync();

            const originalText = tf.textRange.text;
            if (!originalText) continue;

            const flags = caseSensitive ? "g" : "gi";
            const escaped = searchText.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
            const regex = new RegExp(escaped, flags);

            if (regex.test(originalText)) {
              tf.textRange.text = originalText.replace(regex, replaceText);
              replaced++;
            }
          }
        }

        await context.sync();
        setNotification({
          message: `${replaced} Textvorkommen ersetzt. (Vollständige Implementierung: Schritt 8)`,
          type: replaced > 0 ? "success" : "warning",
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
      <Text className="panel-title">Suchen &amp; Ersetzen</Text>

      <NotificationBar
        message={notification?.message ?? null}
        type={notification?.type}
        onDismiss={() => setNotification(null)}
      />

      <Pivot>
        <PivotItem headerText="Text">
          <Stack tokens={{ childrenGap: 10, padding: "12px 0 0 0" }}>
            <TextField
              label="Suchen:"
              value={searchText}
              onChange={(_e, val) => setSearchText(val ?? "")}
              placeholder="Suchtext…"
            />
            <TextField
              label="Ersetzen durch:"
              value={replaceText}
              onChange={(_e, val) => setReplaceText(val ?? "")}
              placeholder="Ersatztext…"
            />
            <Checkbox
              label="Groß-/Kleinschreibung beachten"
              checked={caseSensitive}
              onChange={(_e, checked) => setCaseSensitive(!!checked)}
            />
            <PrimaryButton
              text={isRunning ? "Ersetzt…" : "Alle ersetzen"}
              onClick={handleReplaceAll}
              disabled={isRunning || !searchText}
              styles={{ root: { width: "100%", marginTop: 4 } }}
            />
          </Stack>
        </PivotItem>

        <PivotItem headerText="Farbe">
          <Stack tokens={{ padding: "12px 0 0 0" }}>
            <Text variant="small" style={{ color: "#888" }}>
              Farb-Suche & -Ersetzung – vollständige Implementierung: Schritt 8.
            </Text>
          </Stack>
        </PivotItem>

        <PivotItem headerText="Font">
          <Stack tokens={{ padding: "12px 0 0 0" }}>
            <Text variant="small" style={{ color: "#888" }}>
              Font-Ersetzung – vollständige Implementierung: Schritt 8.
            </Text>
          </Stack>
        </PivotItem>
      </Pivot>
    </Stack>
  );
};

export default FindReplacePanel;
