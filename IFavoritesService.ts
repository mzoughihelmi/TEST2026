import { IFavoriteItem } from "../models/IFavoriteItem";

/**
 * Abstraction du service de favoris : le Provider dépend de cette interface,
 * pas de l'implémentation PnPjs. Permet de mocker facilement (tests, Storybook).
 */
export interface IFavoritesService {
  /** Retourne les N derniers favoris créés par l'utilisateur connecté. */
  getMyFavorites(top: number): Promise<IFavoriteItem[]>;
  /** Supprime un favori de la liste SharePoint. */
  removeFavorite(itemId: number): Promise<void>;
}
