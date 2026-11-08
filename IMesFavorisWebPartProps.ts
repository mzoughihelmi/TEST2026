export interface IMesFavorisWebPartProps {
  /** Titre affiché du web part ("Mes favoris" par défaut si vide). */
  title: string;
  /** Titre de la liste SharePoint des favoris ("MesFavoris" par défaut si vide). */
  listTitle: string;
  /** Nombre maximum de favoris affichés (par défaut 5). */
  maxItems: number;
  /** URL de la page déjà créée listant tous les favoris ("Voir plus"). */
  voirPlusUrl: string;
  voirPlusLabel: string;
}
