import * as React from "react";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { DefaultButton } from "@fluentui/react/lib/Button";
import { mergeStyleSets } from "@fluentui/react/lib/Styling";
import { useFavorites } from "../../context/FavoritesContext";

/**
 * Couleurs relevées sur la maquette (homepage documentation, feuille de route p. 8).
 *  - barre : turquoise, bleu foncé sur la ligne active (survol / focus) ;
 *  - texte : gris ardoise ;  fond de la ligne active : gris-vert très clair ;  séparateur : gris clair.
 */
const GMF_TURQUOISE = "#008697";
const GMF_BLEU = "#004696";
const GMF_TEXTE = "#2e3a48";
const GMF_FOND_ACTIF = "#f5f8f8";
const GMF_SEPARATEUR = "#f3f2f1";

const classes = mergeStyleSets({
  list: {
    listStyle: "none",
    margin: 0,
    padding: 0
  },
  item: {
    // Trait clair sous chaque ligne + petit espace avant la suivante (comme la maquette).
    borderBottom: `1px solid ${GMF_SEPARATEUR}`,
    marginBottom: 4
  },
  link: {
    position: "relative",
    display: "flex",
    alignItems: "center",
    minHeight: 52,
    padding: "8px 16px 8px 28px",
    color: GMF_TEXTE,
    fontSize: 14,
    fontWeight: 500,
    lineHeight: "20px",
    textDecoration: "none",
    transition: "background-color .1s ease",
    selectors: {
      // Barre verticale : petit trait arrondi, centré, décollé du bord gauche.
      "::before": {
        content: '""',
        position: "absolute",
        left: 7,
        top: 9,
        bottom: 9,
        width: 4,
        borderRadius: 2,
        backgroundColor: GMF_TURQUOISE,
        transition: "background-color .1s ease"
      },
      ":hover, :focus-visible": {
        backgroundColor: GMF_FOND_ACTIF,
        color: GMF_TEXTE,
        textDecoration: "none"
      },
      ":hover::before, :focus-visible::before": {
        backgroundColor: GMF_BLEU
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
