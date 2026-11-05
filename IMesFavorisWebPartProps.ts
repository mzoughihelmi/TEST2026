export interface IMesFavorisWebPartProps {
  title: string;
  /** Nombre maximum de favoris affichés (par défaut 5). */
  maxItems: number;
  /** URL de la page déjà créée listant tous les favoris ("Voir plus"). */
  voirPlusUrl: string;
  voirPlusLabel: string;
}
