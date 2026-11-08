import * as React from "react";
import { ActionButton } from "@fluentui/react/lib/Button";

export interface IVoirPlusButtonProps {
  /** URL de la page déjà créée listant tous les favoris (propriété configurable du web part). */
  pageUrl: string;
  label?: string;
  disabled?: boolean;
}

/**
 * Lien "Voir tout" affiché à droite de l'en-tête, comme sur les web parts SharePoint natifs.
 * Ne s'affiche pas si aucune URL n'a été configurée dans le volet de propriétés.
 */
export const VoirPlusButton: React.FC<IVoirPlusButtonProps> = ({ pageUrl, label, disabled }) => {
  if (!pageUrl) {
    return null;
  }

  return (
    <ActionButton
      href={pageUrl}
      disabled={disabled}
      text={label || "Voir tout"}
      menuIconProps={{ iconName: "ChevronRight" }}
      styles={{
        root: { height: 32, padding: 0, fontWeight: 600 },
        label: { margin: 0 },
        menuIcon: { fontSize: 10, marginLeft: 4 }
      }}
    />
  );
};
