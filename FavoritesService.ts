import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/views";
import "@pnp/sp/site-users/web";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import { IFavoriteItem } from "../models/IFavoriteItem";
import { INewFavorite } from "../models/INewFavorite";
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
  /** Valeur(s) de la métadonnée de la page, lue(s) une seule fois et partagée(s) entre les appels. */
  private _pageValuesPromise: Promise<string[]> | undefined;

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
   * Résultat mis en cache (une seule lecture de la page) ; le cache est vidé en cas d'erreur.
   */
  private _getCurrentPageValues(): Promise<string[]> {
    if (!this._pageValuesPromise) {
      this._pageValuesPromise = this._loadCurrentPageValues().catch((err: unknown) => {
        this._pageValuesPromise = undefined;
        throw err;
      });
    }
    return this._pageValuesPromise;
  }

  private async _loadCurrentPageValues(): Promise<string[]> {
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

    return items.map((item) => this._toFavoriteItem(item));
  }

  /**
   * URL de la vue par défaut de la liste des favoris (ex. .../Lists/MesFavoris/AllItems.aspx),
   * filtrée comme le web part :
   *   - filtre 1 : Créé par = utilisateur connecté (nom affiché) ;
   *   - filtre 2 : colonne famille = valeur(s) de la page actuelle (si la page en a une).
   * Utilise la syntaxe de filtres d'URL des listes SharePoint (FilterField / FilterValue).
   */
  public async getAllFavoritesUrl(): Promise<string> {
    const [view, pageValues] = await Promise.all([
      this._sp.web.lists.getByTitle(this._listTitle).defaultView.select("ServerRelativeUrl")(),
      this._getCurrentPageValues()
    ]);

    const params: string[] = [
      `FilterField1=Author`,
      `FilterValue1=${encodeURIComponent(this._context.pageContext.user.displayName)}`,
      `FilterType1=User`
    ];

    if (pageValues.length === 1) {
      params.push(
        `FilterField2=${encodeURIComponent(FAVORITES_FILTER_FIELD)}`,
        `FilterValue2=${encodeURIComponent(pageValues[0])}`
      );
    } else if (pageValues.length > 1) {
      // Plusieurs valeurs (Choix multiple) : FilterValues + séparateur ";#".
      params.push(
        `FilterField2=${encodeURIComponent(FAVORITES_FILTER_FIELD)}`,
        `FilterValues2=${encodeURIComponent(pageValues.join(";#"))}`
      );
    }

    return `${view.ServerRelativeUrl}?${params.join("&")}`;
  }

  /**
   * Ajoute un favori dans la liste pour l'utilisateur connecté.
   * - Title = favorite.title ; URL = { Url: favorite.url, Description: favorite.description ?? title } ;
   * - colonne famille = favorite.famille, ou à défaut la famille de la page actuelle ;
   * - anti-doublon : si l'utilisateur a déjà un favori vers la même URL, il est retourné tel quel.
   * "Créé par" et "Créé" sont renseignés automatiquement par SharePoint.
   */
  public async addFavorite(favorite: INewFavorite): Promise<IFavoriteItem> {
    const title = favorite.title?.trim();
    const url = favorite.url?.trim();
    if (!title || !url) {
      throw new Error("addFavorite : le titre et l'URL du favori sont obligatoires.");
    }
    const description = favorite.description?.trim() || title;

    const list = this._sp.web.lists.getByTitle(this._listTitle);

    const [currentUser, familleValues] = await Promise.all([
      this._sp.web.currentUser(),
      favorite.famille ? Promise.resolve(favorite.famille) : this._getCurrentPageValues()
    ]);

    // Anti-doublon : même utilisateur + même URL de page.
    // Une colonne Lien hypertexte ne peut pas être filtrée en OData : on filtre sur l'auteur
    // côté serveur, puis on compare l'URL côté client (insensible à la casse).
    const myItems = await list.items
      .select("Id", "Title", "URL", "Created")
      .filter(`AuthorId eq ${currentUser.Id}`)
      .top(5000)();
    const normalizedUrl = url.toLowerCase();
    const existing = myItems.find(
      (item) => (item.URL?.Url ?? "").trim().toLowerCase() === normalizedUrl
    );
    if (existing) {
      return this._toFavoriteItem(existing);
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const values: Record<string, any> = {
      Title: title,
      URL: { Url: url, Description: description }
    };
    const famille = familleValues.filter((v) => v !== "");
    if (famille.length === 1) {
      values[FAVORITES_FILTER_FIELD] = famille[0];
    } else if (famille.length > 1) {
      // Colonne Choix à sélection multiple.
      values[FAVORITES_FILTER_FIELD] = famille;
    }

    // PnPjs v3 : add() renvoie { data, item }.
    const result = await list.items.add(values);

    return this._toFavoriteItem({
      Id: result.data.Id,
      Title: title,
      URL: { Url: url, Description: description },
      Created: result.data.Created
    });
  }

  /** Convertit un élément brut SharePoint en modèle UI. */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private _toFavoriteItem(item: any): IFavoriteItem {
    return {
      id: item.Id,
      title: item.Title,
      pageUrl: item.URL?.Url ?? "",
      linkLabel: item.URL?.Description || item.URL?.Url || "",
      created: item.Created
    };
  }

  public async removeFavorite(itemId: number): Promise<void> {
    await this._sp.web.lists
      .getByTitle(this._listTitle)
      .items.getById(itemId)
      .delete();
  }
}
