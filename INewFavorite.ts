/**
 * Données nécessaires pour créer un favori dans la liste SharePoint.
 * L'auteur (Créé par) et la date de création sont renseignés automatiquement par SharePoint.
 */
export interface INewFavorite {
  /** Titre du favori (colonne Title). */
  title: string;
  /** URL de la page cible (champ Hyperlien "URL" › Url). */
  url: string;
  /** Texte affiché du lien (champ Hyperlien "URL" › Description). Par défaut : le titre. */
  description?: string;
  /** Emplacement dans l'arborescence (colonne "Odysse_Localisation"), ex. "Auto Pass > Règles de gestion". */
  localisation?: string;
  /**
   * Valeur(s) de famille à enregistrer sur le favori.
   * Si non fourni : reprend automatiquement la famille de la page actuelle.
   */
  famille?: string[];
}
