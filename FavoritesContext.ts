import * as React from "react";
import { IFavoriteItem } from "../models/IFavoriteItem";

export interface IFavoritesContext {
  /** Les N derniers favoris de l'utilisateur connecté (déjà filtrés par Author). */
  favorites: IFavoriteItem[];
  isLoading: boolean;
  error: string | undefined;
  /** Retire un favori (clic sur l'étoile) : appel service + mise à jour de l'état. */
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
