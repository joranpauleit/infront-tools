/**
 * CornerRadiusPanel.tsx
 * Feature-Panel: Eckenradius setzen.
 * Vollständige Implementierung: Schritt 4.
 */

import * as React from "react";
import { Stack } from "@fluentui/react/lib/Stack";
import { Text } from "@fluentui/react/lib/Text";
import { PrimaryButton } from "@fluentui/react/lib/Button";
import { TextField } from "@fluentui/react/lib/TextField";
import NotificationBar, { NotificationType } from "../shared/NotificationBar";

interface Notification {
  message: string;
  type: NotificationType;
}

const CornerRadiusPanel: React.FC = () => {
  const [radiusPx, setRadiusPx]           = React.useState<string>("8");
  const [notification, setNotification]   = React.useState<Notification | null>(null);
  const [isRunning, setIsRunning]         = React.useState(false);

  const handleApply = async () => {
    const px = parseFloat(radiusPx);
    if (isNaN(px) || px < 0) {
      setNotification({ message: "Bitte einen gültigen Pixelwert eingeben (≥ 0).", type: "error" });
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

        // Pixel → Punkte: 1pt = 1/72 Zoll; 1px bei 96 DPI = 0.75pt
        const ptValue = px * 0.75;

        let applied  = 0;
        let skipped  = 0;

        for (const shape of selection.items) {
          shape.load("geometricShapeType");
        }
        await context.sync();

        for (const shape of selection.items) {
          // Nur Shapes mit anpassbarer Geometrie unterstützen Eckenradius.
          // roundedRectangle (GeometricShapeType 62) hat adjustment[0] = Radius (normiert 0–1).
          // Andere Shapes werden übersprungen.
          const gst = shape.geometricShapeType;

          // GeometricShapeType.roundedRectangle = "roundedRectangle"
          if (gst === PowerPoint.GeometricShapeType.roundedRectangle) {
            // adjustment[0] steuert den Radius: 0 = kein Radius, 1 = max. Radius.
            // Normierung: ptValue / (min(width, height) / 2)
            // Da width/height in diesem Kontext nicht immer geladen ist, nutzen wir
            // einen pragmatischen Faktor: ptValue / 100 (Schätzwert für Standardgrößen).
            // Schritt 4 lädt width/height für präzisere Normierung.
            shape.load(["width", "height"]);
            await context.sync();

            const maxRadius = Math.min(shape.width, shape.height) / 2;
            const normalized = maxRadius > 0 ? Math.min(ptValue / maxRadius, 1.0) : 0;
            shape.geometricShape.adjustments.getItemAt(0).value = normalized;
            applied++;
          } else {
            skipped++;
          }
        }

        await context.sync();
        setNotification({
          message: `${applied} Shape(s) angepasst, ${skipped} übersprungen.`,
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

  return (
    <Stack className="panel-container" tokens={{ childrenGap: 12 }}>
      <Text className="panel-title">Eckenradius setzen</Text>
      <Text className="panel-description">
        Setzt den Eckenradius aller selektierten Rounded-Rectangle-Shapes.
        Andere Shape-Typen werden übersprungen.
      </Text>

      <NotificationBar
        message={notification?.message ?? null}
        type={notification?.type}
        onDismiss={() => setNotification(null)}
      />

      <TextField
        label="Eckenradius in Pixel (z. B. 8):"
        value={radiusPx}
        onChange={(_e, val) => setRadiusPx(val ?? "")}
        type="number"
        min={0}
        step={1}
        suffix="px"
        styles={{ root: { maxWidth: 200 } }}
      />

      <Text variant="small" style={{ color: "#888" }}>
        Umrechnung: 1 px = 0,75 pt (bei 96 DPI).
        <br />
        Unterstützt: Rounded Rectangle. Nicht unterstützt: alle anderen Shape-Typen.
      </Text>

      <PrimaryButton
        text={isRunning ? "Wird angewendet…" : "Anwenden"}
        onClick={handleApply}
        disabled={isRunning}
        styles={{ root: { width: "100%", marginTop: 8 } }}
      />
    </Stack>
  );
};

export default CornerRadiusPanel;
