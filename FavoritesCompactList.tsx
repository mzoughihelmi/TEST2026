import * as React from "react";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { DefaultButton } from "@fluentui/react/lib/Button";
import { mergeStyleSets } from "@fluentui/react/lib/Styling";
import { useFavorites } from "../../context/FavoritesContext";

/** Couleurs GMF de la maquette (homepage documentation, feuille de route p. 8). */
const GMF_PETROLE = "#1f6b6f";
const GMF_BLEU = "#1f4e8c";

const classes = mergeStyleSets({
  list: {
    listStyle: "none",
    margin: 0,
    padding: 0
  },
  item: {
    borderBottom: "1px solid #edebe9"
  },
  link: {
    display: "flex",
    alignItems: "center",
    minHeight: 48,
    padding: "8px 16px 8px 20px",
    borderLeft: `3px solid ${GMF_PETROLE}`,
    color: "#323130",
    fontSize: 13,
    lineHeight: "18px",
    textDecoration: "none",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
    transition: "background-color .1s ease",
    selectors: {
      ":hover, :focus-visible": {
        backgroundColor: "#f3f5f6",
        borderLeftColor: GMF_BLEU,
        color: "#323130",
        textDecoration: "none"
      },
      ":focus-visible": {
        outline: `2px solid ${GMF_BLEU}`,
        outlineOffset: -2
      }
    }
  },
  linkText: {
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap"
  },
  empty: {
    padding: "12px 0",
    fontSize: 13,
    color: "#605e5c"
  }
});

/**
 * Affichage "page d'accueil" (homepage documentation, repère C3 — feuille de route p. 8) :
 * liste simple des derniers favoris, titre seul cliquable vers la page du favori,
 * barre de couleur à gauche, sans date, sans lien ni étoile de suppression.
 */
export const FavoritesCompactList: React.FC = () => {
  const { favorites, isLoading, loadError, refresh } = useFavorites();

  if (isLoading) {
    return (
      <Spinner
        size={SpinnerSize.medium}
        label="Chargement de vos favoris..."
        styles={{ root: { padding: "24px 0" } }}
      />
    );
  }

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

  if (favorites.length === 0) {
    return <div className={classes.empty}>Vous n&apos;avez pas encore de favoris.</div>;
  }

  return (
    <ul className={classes.list} aria-label="Mes favoris">
      {favorites.map((item) => (
        <li key={item.id} className={classes.item}>
          <a
            className={classes.link}
            href={item.pageUrl}
            title={item.title}
          >
            <span className={classes.linkText}>{item.title}</span>
          </a>
        </li>
      ))}
    </ul>
  );
};
