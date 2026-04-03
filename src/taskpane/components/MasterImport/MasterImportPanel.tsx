/**
 * MasterImportPanel.tsx
 * Feature-Panel: Master / Theme importieren.
 * Vollständige Implementierung: Schritt 10.
 *
 * HINWEIS: Vollständiger Master-/Theme-Import ist über Office.js
 * auf Mac stark eingeschränkt (kein SlideMaster-Replacement-API).
 * Machbarer Fallback: Theme-Farben/-Fonts manuell setzen.
 */

import * as React from "react";
import { Stack } from "@fluentui/react/lib/Stack";
import { Text } from "@fluentui/react/lib/Text";
import { PrimaryButton } from "@fluentui/react/lib/Button";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import NotificationBar, { NotificationType } from "../shared/NotificationBar";

interface Notification {
  message: string;
  type: NotificationType;
}

const MasterImportPanel: React.FC = () => {
  const [notification, setNotification] = React.useState<Notification | null>(null);

  return (
    <Stack className="panel-container" tokens={{ childrenGap: 12 }}>
      <Text className="panel-title">Master / Theme importieren</Text>

      <MessageBar messageBarType={MessageBarType.warning} isMultiline>
        Vollständiger Master-Ersatz ist über Office.js auf Mac nicht möglich
        (API-Einschränkung). Verfügbar: Theme-Farben und -Fonts anpassen.
        Dokumentiert in TESTING.md.
      </MessageBar>

      <NotificationBar
        message={notification?.message ?? null}
        type={notification?.type}
        onDismiss={() => setNotification(null)}
      />

      <Text variant="small" style={{ color: "#888" }}>
        Vollständige Implementierung: Schritt 10.
        <br />
        Features: Theme-Farben, Theme-Fonts, Layout-Zuordnung, Farbvorschau.
      </Text>

      <PrimaryButton
        text="Datei auswählen"
        disabled
        onClick={() => setNotification({ message: "Kommt in Schritt 10.", type: "info" })}
        styles={{ root: { width: "100%", marginTop: 8 } }}
      />
    </Stack>
  );
};

export default MasterImportPanel;
