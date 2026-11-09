import * as React from "react";
import {
  DetailsList,
  IColumn,
  SelectionMode,
  DetailsListLayoutMode,
  CheckboxVisibility
} from "@fluentui/react/lib/DetailsList";
import { IconButton } from "@fluentui/react/lib/Button";
import { Link } from "@fluentui/react/lib/Link";
import { Text } from "@fluentui/react/lib/Text";
import { Icon } from "@fluentui/react/lib/Icon";
import { Stack } from "@fluentui/react/lib/Stack";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { DefaultButton } from "@fluentui/react/lib/Button";
import { mergeStyleSets } from "@fluentui/react/lib/Styling";
import { useTheme } from "@fluentui/react/lib/Theme";
import { useFavorites } from "../../context/FavoritesContext";
import { IFavoriteItem } from "../../models/IFavoriteItem";

export interface IFavoritesListProps {
  /** Culture de la page SharePoint (ex. "fr-FR"), utilisée pour formater les dates. */
  locale: string;
}

const formatDate = (iso: string, locale: string, withTime: boolean): string => {
  if (!iso) {
    return "";
  }
  const date = new Date(iso);
  if (isNaN(date.getTime())) {
    return "";
  }
  return date.toLocaleString(
    locale,
    withTime
      ? { day: "numeric", month: "long", year: "numeric", hour: "2-digit", minute: "2-digit" }
      : { day: "numeric", month: "short", year: "numeric" }
  );
};

/**
 * Affiche les favoris dans une vue de liste Fluent UI (DetailsList), au plus proche
 * des vues de listes SharePoint : en-têtes Titre / Date de création / Lien,
 * effet "shimmer" pendant le chargement et étoile pour retirer un favori.
 */
export const FavoritesList: React.FC<IFavoritesListProps> = ({ locale }) => {
  const { favorites, isLoading, loadError, refresh, removeFavorite, removingId, isBusy } = useFavorites();
  const theme = useTheme();

  const classes = React.useMemo(
    () =>
      mergeStyleSets({
        cellText: {
          display: "block",
          overflow: "hidden",
          textOverflow: "ellipsis",
          whiteSpace: "nowrap"
        },
        secondary: {
          color: theme.semanticColors.bodySubtext
        },
        empty: {
          padding: "24px 8px",
          color: theme.semanticColors.bodySubtext
        },
        emptyIcon: {
          fontSize: 28,
          color: theme.palette.neutralTertiary
        }
      }),
    [theme]
  );

  const columns: IColumn[] = React.useMemo(
    (): IColumn[] => [
      {
        key: "title",
        name: "Titre",
        fieldName: "title",
        minWidth: 220,
        maxWidth: 350,
        isResizable: true,
        onRender: (item: IFavoriteItem) => (
          <Link
            href={item.pageUrl}
            target="_blank"
            rel="noopener noreferrer"
            data-interception="off"
            title={item.title}
            className={classes.cellText}
            styles={{ root: { fontWeight: 600, color: theme.semanticColors.bodyText } }}
          >
            {item.title}
          </Link>
        )
      },
      {
        key: "created",
        name: "Date de création",
        fieldName: "created",
        minWidth: 120,
        maxWidth: 160,
        isResizable: true,
        isSorted: true,
        isSortedDescending: true,
        onRender: (item: IFavoriteItem) => (
          <TooltipHost content={formatDate(item.created, locale, true)}>
            <span className={`${classes.cellText} ${classes.secondary}`}>
              {formatDate(item.created, locale, false)}
            </span>
          </TooltipHost>
        )
      },
      {
        key: "link",
        name: "Lien",
        fieldName: "pageUrl",
        minWidth: 240,
        maxWidth: 400,
        isResizable: true,
        onRender: (item: IFavoriteItem) =>
          item.pageUrl ? (
            <Link
              href={item.pageUrl}
              target="_blank"
              rel="noopener noreferrer"
              data-interception="off"
              title={item.pageUrl}
              className={classes.cellText}
            >
              {item.linkLabel}
            </Link>
          ) : (
            <span className={classes.secondary}>—</span>
          )
      },
      {
        // Étoile en dernière colonne (à droite) : retire le favori.
        key: "favorite",
        name: "Favori",
        iconName: "FavoriteStarFill",
        isIconOnly: true,
        minWidth: 32,
        maxWidth: 32,
        onRender: (item: IFavoriteItem) => (
          <IconButton
            iconProps={{ iconName: "FavoriteStarFill" }}
            title="Retirer des favoris"
            ariaLabel={`Retirer « ${item.title} » des favoris`}
            disabled={isBusy}
            onClick={() => {
              removeFavorite(item.id).catch(() => undefined);
            }}
            styles={{
              root: { height: 24, width: 24, marginTop: -4 },
              icon: { color: "#ffb900", fontSize: 16 },
              iconHovered: { color: "#d29200" },
              rootDisabled: { backgroundColor: "transparent" },
              iconDisabled: { color: "#ffb900", opacity: item.id === removingId ? 1 : 0.5 }
            }}
          />
        )
      }
    ],
    [classes, theme, locale, isBusy, removingId, removeFavorite]
  );

  if (loadError) {
    return (
      <MessageBar
        messageBarType={MessageBarType.error}
        isMultiline={false}
        actions={<DefaultButton text="Réessayer" onClick={() => refresh().catch(() => undefined)} />}
      >
        {loadError}
      </MessageBar>
    );
  }

  if (isLoading) {
    return (
      <Spinner
        size={SpinnerSize.medium}
        label="Chargement de vos favoris..."
        styles={{ root: { padding: "24px 0" } }}
      />
    );
  }

  if (favorites.length === 0) {
    return (
      <Stack horizontalAlign="center" tokens={{ childrenGap: 8 }} className={classes.empty}>
        <Icon iconName="FavoriteStar" className={classes.emptyIcon} />
        <Text variant="medium">Aucun favori pour cette page.</Text>
      </Stack>
    );
  }

  return (
    <DetailsList
      items={favorites}
      columns={columns}
      getKey={(item: IFavoriteItem) => String(item.id)}
      setKey="mesFavoris"
      selectionMode={SelectionMode.none}
      checkboxVisibility={CheckboxVisibility.hidden}
      // Colonnes à largeur fixe : évite le recalcul en boucle des largeurs ("justified")
      // qui faisait bouger la liste en continu dans la page SharePoint.
      layoutMode={DetailsListLayoutMode.fixedColumns}
      // 5 à 10 lignes seulement : pas de virtualisation, donc pas de re-mesure permanente.
      onShouldVirtualize={() => false}
      compact={false}
      ariaLabelForGrid="Mes favoris"
    />
  );
};
