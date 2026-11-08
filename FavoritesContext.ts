import * as React from "react";
import { IFavoriteItem } from "../models/IFavoriteItem";

export interface IFavoritesContext {
  /** Les N derniers favoris de l'utilisateur connecté (déjà filtrés par Author). */
  favorites: IFavoriteItem[];
  /** Chargement initial (ou rechargement) en cours. */
  isLoading: boolean;
  /** Erreur bloquante de chargement : la liste ne peut pas être affichée. */
  loadError: string | undefined;
  /** Erreur non bloquante d'une action (ex. suppression échouée) : la liste reste affichée. */
  actionError: string | undefined;
  clearActionError: () => void;
  /** Id du favori en cours de suppression (undefined si aucune suppression en cours). */
  removingId: number | undefined;
  /** true tant qu'une suppression est en cours : toute autre interaction est bloquée. */
  isBusy: boolean;
  /** Retire un favori (clic sur l'étoile) : appel service puis mise à jour de l'état. */
  removeFavorite: (itemId: number) => Promise<void>;
  /** Permet de recharger manuellement la liste si besoin. */
  refresh: () => Promise<void>;
}

export const FavoritesContext = React.createContext<IFavoritesContext | undefined>(undefined);

/**
 * Hook d'accès au contexte Mes Favoris.
 * Doit être utilisé par un composant descendant de <FavoritesProvider>.
 */
export const useFavorites = (): IFavoritesContext => {
  const context = React.useContext(FavoritesContext);
  if (!context) {
    throw new Error("useFavorites() doit être utilisé à l'intérieur d'un <FavoritesProvider>.");
  }
  return context;
};
