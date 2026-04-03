/**
 * ReviewPanel.tsx
 * Feature-Panel: Review / Kommentare / Annotationen.
 * Vollständige Implementierung: Schritt 11.
 * Wird auch für "Meine Kommentare" (view=my-comments) verwendet.
 */

import * as React from "react";
import { Stack } from "@fluentui/react/lib/Stack";
import { Text } from "@fluentui/react/lib/Text";
import { PrimaryButton, DefaultButton } from "@fluentui/react/lib/Button";
import { TextField } from "@fluentui/react/lib/TextField";
import NotificationBar, { NotificationType } from "../shared/NotificationBar";

interface Notification {
  message: string;
  type: NotificationType;
}

/** Liefert den Nutzernamen aus Office-Kontext (Fallback: "Unbekannt"). */
function getUserName(): string {
  try {
    return Office.context.userProfile?.displayName ?? "Unbekannt";
  } catch {
    return "Unbekannt";
  }
}

/** Formatiert einen deutschen Zeitstempel. */
function getTimestamp(): string {
  return new Date().toLocaleString("de-DE", {
    day:    "2-digit",
    month:  "2-digit",
    year:   "numeric",
    hour:   "2-digit",
    minute: "2-digit",
  });
}

const ReviewPanel: React.FC = () => {
  const [commentText, setCommentText]   = React.useState<string>("");
  const [notification, setNotification] = React.useState<Notification | null>(null);
  const [isRunning, setIsRunning]       = React.useState(false);

  const view = new URLSearchParams(window.location.search).get("view");
  const isMyComments = view === "my-comments";

  const stamp = `[${getUserName()} – ${getTimestamp()}]`;

  const handleAddComment = async () => {
    setIsRunning(true);
    setNotification(null);

    try {
      await PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        const shapes = slide.shapes;

        const shapeName = `INFRONT_COMMENT_${Date.now()}`;
        const fullText  = `${stamp}\n${commentText}`.trim();

        const box = shapes.addTextBox(fullText);
        box.name = shapeName;
        box.load(["left", "top", "width", "height"]);
        await context.sync();

        // Standard-Position: oben rechts
        box.left   = 400;
        box.top    = 20;
        box.width  = 200;
        box.height = 80;
        box.fill.setSolidColor("#FFFACD");
        box.lineFormat.color = "#FFA500";
        box.lineFormat.weight = 1;

        await context.sync();
        setNotification({
          message: `Kommentar "${shapeName}" eingefügt.`,
          type: "success",
        });
        setCommentText("");
      });
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setNotification({ message: `Fehler: ${message}`, type: "error" });
    } finally {
      setIsRunning(false);
    }
  };

  if (isMyComments) {
    return (
      <Stack className="panel-container" tokens={{ childrenGap: 12 }}>
        <Text className="panel-title">Meine Kommentare</Text>
        <Text className="panel-description">
          Übersicht aller eigenen Infront-Kommentare im Deck.
        </Text>
        <Text variant="small" style={{ color: "#888" }}>
          Vollständige Implementierung: Schritt 11.
        </Text>
      </Stack>
    );
  }

  return (
    <Stack className="panel-container" tokens={{ childrenGap: 12 }}>
      <Text className="panel-title">Kommentar hinzufügen</Text>
      <Text className="panel-description">
        Fügt ein Post-It-Kommentar mit Namens-Zeitstempel ein.
      </Text>

      <NotificationBar
        message={notification?.message ?? null}
        type={notification?.type}
        onDismiss={() => setNotification(null)}
      />

      <Text variant="small" style={{ color: "#555", fontStyle: "italic" }}>{stamp}</Text>

      <TextField
        label="Kommentartext (optional):"
        value={commentText}
        onChange={(_e, val) => setCommentText(val ?? "")}
        multiline
        rows={3}
        placeholder="Kommentar eingeben…"
      />

      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <PrimaryButton
          text={isRunning ? "Wird eingefügt…" : "Kommentar einfügen"}
          onClick={handleAddComment}
          disabled={isRunning}
          styles={{ root: { flex: 1 } }}
        />
        <DefaultButton
          text="Alle entfernen"
          disabled
          title="Über den Ribbon-Button 'Kommentare entfernen'"
        />
      </Stack>
    </Stack>
  );
};

export default ReviewPanel;
