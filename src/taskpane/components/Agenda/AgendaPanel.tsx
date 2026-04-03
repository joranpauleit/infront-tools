/**
 * AgendaPanel.tsx
 * Feature-Panel: Agenda-Assistent.
 * Vollständige Implementierung: Schritt 9.
 */

import * as React from "react";
import { Stack } from "@fluentui/react/lib/Stack";
import { Text } from "@fluentui/react/lib/Text";
import { PrimaryButton } from "@fluentui/react/lib/Button";
import NotificationBar, { NotificationType } from "../shared/NotificationBar";

interface Notification {
  message: string;
  type: NotificationType;
}

const AgendaPanel: React.FC = () => {
  const [notification, setNotification] = React.useState<Notification | null>(null);

  return (
    <Stack className="panel-container" tokens={{ childrenGap: 12 }}>
      <Text className="panel-title">Agenda-Assistent</Text>
      <Text className="panel-description">
        Erstellt und verwaltet eine Agenda mit bis zu 8 Sektionen.
        Unterstützt Shape-Tags für persistente Metadaten.
      </Text>

      <NotificationBar
        message={notification?.message ?? null}
        type={notification?.type}
        onDismiss={() => setNotification(null)}
      />

      <Text variant="small" style={{ color: "#888" }}>
        Vollständige Implementierung: Schritt 9.
        <br />
        Features: Setup-Modus, Sektionen konfigurieren, Aktiv/Inaktiv-Formatierung,
        manueller Update-Button, Quick-Edit-Modus.
      </Text>

      <PrimaryButton
        text="Setup starten"
        disabled
        onClick={() => setNotification({ message: "Kommt in Schritt 9.", type: "info" })}
        styles={{ root: { width: "100%", marginTop: 8 } }}
      />
    </Stack>
  );
};

export default AgendaPanel;
