import * as React from "react";
import { FavoritesContext, IFavoritesContext } from "./FavoritesContext";
import { IFavoriteItem } from "../models/IFavoriteItem";
import { IFavoritesService } from "../services/IFavoritesService";

export interface IFavoritesProviderProps {
  /** Service injecté (PnPjs en prod, mock possible en test) pour lire/supprimer les favoris. */
  service: IFavoritesService;
  /** Nombre maximum de favoris à charger (les N derniers créés par l'utilisateur connecté). */
  maxItems: number;
  children: React.ReactNode;
}

/**
 * Provider responsable de l'état "Mes Favoris" :
 *  - chargement initial des favoris de l'utilisateur connecté,
 *  - suppression optimiste d'un favori (clic sur l'étoile, avec rollback en cas d'erreur),
 *  - exposition du contexte aux composants enfants (FavoritesList, FavoriteItem, ...).
 */
export const FavoritesProvider: React.FC<IFavoritesProviderProps> = (props) => {
  const { service, maxItems, children } = props;

  const [favorites, setFavorites] = React.useState<IFavoriteItem[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | undefined>(undefined);

  const loadFavorites = React.useCallback(async (): Promise<void> => {
    setIsLoading(true);
    setError(undefined);
    try {
      const items = await service.getMyFavorites(maxItems);
      setFavorites(items);
    } catch (err: unknown) {
      setError("Impossible de charger vos favoris. Veuillez réessayer.");
      // eslint-disable-next-line no-console
      console.error("[FavoritesProvider] loadFavorites", err);
    } finally {
      setIsLoading(false);
    }
  }, [service, maxItems]);

  React.useEffect(() => {
    loadFavorites().catch(() => undefined);
  }, [loadFavorites]);

  const removeFavorite = React.useCallback(
    async (itemId: number): Promise<void> => {
      const previousFavorites = favorites;

      // Mise à jour optimiste : l'item disparaît immédiatement de la liste.
      setFavorites((current) => current.filter((item) => item.id !== itemId));

      try {
        await service.removeFavorite(itemId);
      } catch (err: unknown) {
        // Échec côté serveur : on restaure l'état précédent et on signale l'erreur.
        setFavorites(previousFavorites);
        setError("Impossible de retirer ce favori. Veuillez réessayer.");
        // eslint-disable-next-line no-console
        console.error("[FavoritesProvider] removeFavorite", err);
      }
    },
    [favorites, service]
  );

  const contextValue: IFavoritesContext = {
    favorites,
    isLoading,
    error,
    removeFavorite,
    refresh: loadFavorites
  };

  return <FavoritesContext.Provider value={contextValue}>{children}</FavoritesContext.Provider>;
};
