import * as React from "react";
import styles from "./VoirPlusButton.module.scss";

export interface IVoirPlusButtonProps {
  /** URL de la page déjà créée listant tous les favoris (propriété configurable du web part). */
  pageUrl: string;
  label?: string;
}

/**
 * Bouton "Voir plus" : renvoie vers la page déjà existante listant l'ensemble
 * des favoris de l'utilisateur (au-delà des N derniers affichés ici).
 * Ne s'affiche pas si aucune URL n'a été configurée dans le volet de propriétés.
 */
export const VoirPlusButton: React.FC<IVoirPlusButtonProps> = ({ pageUrl, label }) => {
  if (!pageUrl) {
    return null;
  }

  return (
    <div className={styles.voirPlusWrapper}>
      <a className={styles.voirPlusButton} href={pageUrl}>
        {label || "Voir tous mes favoris"}
      </a>
    </div>
  );
};
