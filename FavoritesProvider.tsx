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
 *  - suppression NON optimiste d'un favori : l'item reste affiché pendant l'appel serveur,
 *    l'interface est bloquée (isBusy) et l'item n'est retiré qu'une fois la suppression confirmée,
 *  - exposition du contexte aux composants enfants.
 */
export const FavoritesProvider: React.FC<IFavoritesProviderProps> = (props) => {
  const { service, maxItems, children } = props;

  const [favorites, setFavorites] = React.useState<IFavoriteItem[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [loadError, setLoadError] = React.useState<string | undefined>(undefined);
  const [actionError, setActionError] = React.useState<string | undefined>(undefined);
  const [removingId, setRemovingId] = React.useState<number | undefined>(undefined);

  // Verrou synchrone : empêche deux suppressions simultanées même avant le re-render.
  const busyRef = React.useRef<boolean>(false);
  const isMountedRef = React.useRef<boolean>(true);

  React.useEffect(() => {
    isMountedRef.current = true;
    return () => {
      isMountedRef.current = false;
    };
  }, []);

  const loadFavorites = React.useCallback(async (): Promise<void> => {
    setIsLoading(true);
    setLoadError(undefined);
    try {
      const items = await service.getMyFavorites(maxItems);
      if (isMountedRef.current) {
        setFavorites(items);
      }
    } catch (err: unknown) {
      if (isMountedRef.current) {
        setLoadError("Impossible de charger vos favoris. Veuillez réessayer.");
      }
      // eslint-disable-next-line no-console
      console.error("[FavoritesProvider] loadFavorites", err);
    } finally {
      if (isMountedRef.current) {
        setIsLoading(false);
      }
    }
  }, [service, maxItems]);

  React.useEffect(() => {
    loadFavorites().catch(() => undefined);
  }, [loadFavorites]);

  const removeFavorite = React.useCallback(
    async (itemId: number): Promise<void> => {
      if (busyRef.current) {
        return;
      }
      busyRef.current = true;
      setActionError(undefined);
      setRemovingId(itemId);

      try {
        await service.removeFavorite(itemId);
        if (isMountedRef.current) {
          setFavorites((current) => current.filter((item) => item.id !== itemId));
        }
      } catch (err: unknown) {
        if (isMountedRef.current) {
          setActionError("Impossible de retirer ce favori. Veuillez réessayer.");
        }
        // eslint-disable-next-line no-console
        console.error("[FavoritesProvider] removeFavorite", err);
      } finally {
        busyRef.current = false;
        if (isMountedRef.current) {
          setRemovingId(undefined);
        }
      }
    },
    [service]
  );

  const clearActionError = React.useCallback((): void => setActionError(undefined), []);

  const contextValue: IFavoritesContext = {
    favorites,
    isLoading,
    loadError,
    actionError,
    clearActionError,
    removingId,
    isBusy: removingId !== undefined,
    removeFavorite,
    refresh: loadFavorites
  };

  return <FavoritesContext.Provider value={contextValue}>{children}</FavoritesContext.Provider>;
};
