import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import { IFavoriteItem } from "../models/IFavoriteItem";
import { IFavoritesService } from "./IFavoritesService";

/** Nom de la liste SharePoint source des favoris (Title + colonne Hyperlien "URL"). */
const FAVORITES_LIST_TITLE = "MesFavoris";

/**
 * Implémentation PnPjs du service Mes Favoris.
 * - Ne retourne que les favoris créés par l'utilisateur connecté (filtre OData AuthorId).
 * - Trie par date de création descendante : "les N derniers" favoris de l'utilisateur.
 */
export class FavoritesService implements IFavoritesService {
  private readonly _sp: SPFI;

  constructor(context: WebPartContext) {
    this._sp = spfi().using(SPFx(context));
  }

  public async getMyFavorites(top: number): Promise<IFavoriteItem[]> {
    const currentUser = await this._sp.web.currentUser();

    const items = await this._sp.web.lists
      .getByTitle(FAVORITES_LIST_TITLE)
      .items.select("Id", "Title", "URL", "Created")
      .filter(`AuthorId eq ${currentUser.Id}`)
      .orderBy("Created", false)
      .top(top)();

    return items.map(
      (item): IFavoriteItem => ({
        id: item.Id,
        title: item.Title,
        pageUrl: item.URL?.Url ?? "",
        created: item.Created
      })
    );
  }

  public async removeFavorite(itemId: number): Promise<void> {
    await this._sp.web.lists
      .getByTitle(FAVORITES_LIST_TITLE)
      .items.getById(itemId)
      .delete();
  }
}
