import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import { IFavoriteItem } from "../models/IFavoriteItem";
import { IFavoritesService } from "./IFavoritesService";

/** Liste source par défaut si aucune liste n'est configurée dans le volet de propriétés. */
export const DEFAULT_FAVORITES_LIST_TITLE = "MesFavoris";

/**
 * ⚠️ À ADAPTER : nom INTERNE de la colonne (Texte ou Choix) dans la bibliothèque "Pages du site"
 * dont la valeur, sur la page actuelle, sert de filtre.
 */
const PAGE_FILTER_FIELD = "Thematique";

/**
 * ⚠️ À ADAPTER : nom INTERNE de la colonne (Texte ou Choix) dans la liste "MesFavoris"
 * comparée à la valeur de la page actuelle.
 */
const FAVORITES_FILTER_FIELD = "Thematique";

/** Échappe une valeur pour un littéral OData ('l'équipe' → 'l''équipe'). */
const odataString = (value: string): string => `'${value.replace(/'/g, "''")}'`;

/**
 * Implémentation PnPjs du service Mes Favoris.
 * - Ne retourne que les favoris créés par l'utilisateur connecté (filtre OData AuthorId).
 * - Filtre en plus sur la valeur de la métadonnée PAGE_FILTER_FIELD de la page actuelle :
 *   si la page n'a pas de valeur (ou hors d'une page), AUCUN favori n'est retourné.
 * - Trie par date de création descendante : "les N derniers" favoris de l'utilisateur.
 */
export class FavoritesService implements IFavoritesService {
  private readonly _sp: SPFI;
  private readonly _context: WebPartContext;
  private readonly _listTitle: string;

  /**
   * @param listTitle Titre de la liste des favoris (volet de propriétés) ;
   *                  "MesFavoris" par défaut si vide.
   */
  constructor(context: WebPartContext, listTitle?: string) {
    this._context = context;
    this._listTitle = listTitle?.trim() || DEFAULT_FAVORITES_LIST_TITLE;
    this._sp = spfi().using(SPFx(context));
  }

  /**
   * Lit la valeur de la métadonnée de la page actuelle.
   * Retourne une liste vide si le web part n'est pas sur une page (ex. workbench)
   * ou si la colonne est vide sur cette page.
   * Une colonne Choix à sélection multiple renvoie plusieurs valeurs.
   */
  private async _getCurrentPageValues(): Promise<string[]> {
    const { list, listItem } = this._context.pageContext;
    if (!list || !listItem) {
      return [];
    }

    const page = await this._sp.web.lists
      .getById(list.id.toString())
      .items.getById(listItem.id)
      .select(PAGE_FILTER_FIELD)();

    const raw: unknown = page[PAGE_FILTER_FIELD];
    if (Array.isArray(raw)) {
      return raw.filter((v): v is string => typeof v === "string" && v !== "");
    }
    return typeof raw === "string" && raw !== "" ? [raw] : [];
  }

  public async getMyFavorites(top: number): Promise<IFavoriteItem[]> {
    const pageValues = await this._getCurrentPageValues();

    // Page sans valeur de métadonnée (ou web part hors d'une page) : rien à afficher.
    if (pageValues.length === 0) {
      return [];
    }

    const currentUser = await this._sp.web.currentUser();

    const metadataFilter = pageValues
      .map((value) => `${FAVORITES_FILTER_FIELD} eq ${odataString(value)}`)
      .join(" or ");
    const filter = `AuthorId eq ${currentUser.Id} and (${metadataFilter})`;

    const items = await this._sp.web.lists
      .getByTitle(this._listTitle)
      .items.select("Id", "Title", "URL", "Created")
      .filter(filter)
      .orderBy("Created", false)
      .top(top)();

    return items.map(
      (item): IFavoriteItem => ({
        id: item.Id,
        title: item.Title,
        pageUrl: item.URL?.Url ?? "",
        linkLabel: item.URL?.Description || item.URL?.Url || "",
        created: item.Created
      })
    );
  }

  public async removeFavorite(itemId: number): Promise<void> {
    await this._sp.web.lists
      .getByTitle(this._listTitle)
      .items.getById(itemId)
      .delete();
  }
}
