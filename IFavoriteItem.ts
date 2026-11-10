/**
 * Représentation « UI » d'un favori, découplée du format brut renvoyé par SharePoint/PnPjs.
 */
export interface IFavoriteItem {
  id: number;
  title: string;
  /** URL de la page cible, extraite du champ Hyperlien "URL" de la liste. */
  pageUrl: string;
  /** Texte affiché du lien : champ "Description" du Hyperlien "URL" (repli sur l'URL si vide). */
  linkLabel: string;
  /** Date de création ISO, utile pour un tri ou un affichage futur. */
  created: string;
  /** Emplacement du favori dans l'arborescence (colonne "Odysse_Localisation"), ex. "Auto Pass > Règles de gestion". */
  localisation: string;
}
