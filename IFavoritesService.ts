import { IFavoriteItem } from "../models/IFavoriteItem";
import { INewFavorite } from "../models/INewFavorite";

/**
 * Abstraction du service de favoris : le Provider dépend de cette interface,
 * pas de l'implémentation PnPjs. Permet de mocker facilement (tests, Storybook).
 */
export interface IFavoritesService {
  /** Retourne les N derniers favoris créés par l'utilisateur connecté. */
  getMyFavorites(top: number): Promise<IFavoriteItem[]>;
  /**
   * URL de la liste des favoris filtrée sur l'utilisateur connecté et la famille de la page
   * (utilisée par "Voir tous" quand aucune URL n'est saisie dans le volet de propriétés).
   */
  getAllFavoritesUrl(): Promise<string>;
  /**
   * Ajoute un favori dans la liste SharePoint pour l'utilisateur connecté.
   * Si l'utilisateur a déjà un favori vers la même URL, aucun doublon n'est créé :
   * le favori existant est retourné.
   */
  addFavorite(favorite: INewFavorite): Promise<IFavoriteItem>;
  /** Famille(s) de la page actuelle (vide si la page n'en a pas). */
  getCurrentPageFamily(): Promise<string[]>;
  /** Supprime un favori de la liste SharePoint. */
  removeFavorite(itemId: number): Promise<void>;
}
