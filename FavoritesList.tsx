import * as React from "react";
import { Icon } from "@fluentui/react/lib/Icon";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { DefaultButton } from "@fluentui/react/lib/Button";
import { mergeStyleSets } from "@fluentui/react/lib/Styling";
import { useFavorites } from "../../context/FavoritesContext";

/**
 * Couleurs relevées sur la maquette de la homepage famille (feuille de route p. 9, repère ⑥).
 */
const GMF_BLEU = "#004696"; // barre verticale à gauche de chaque ligne
const GMF_PETROLE = "#2e6f76"; // étoile
const GMF_TEXTE = "#2e3a48"; // titres
const GMF_TEXTE_SECONDAIRE = "#414c58"; // localisation
const GMF_ENTETE = "#323e4b"; // en-têtes de colonnes
const GMF_FOND_LIGNE = "#f5f8f8"; // une ligne sur deux
const GMF_FOND_SURVOL = "#e8f0f1"; // ligne survolée (plus soutenu que le fond alterné)
const GMF_SEPARATEUR = "#edebe9";

const classes = mergeStyleSets({
  table: {
    width: "100%",
    borderCollapse: "separate",
    borderSpacing: 0,
    tableLayout: "fixed"
  },
  th: {
    textAlign: "left",
    fontSize: 12,
    fontWeight: 600,
    color: GMF_ENTETE,
    padding: "10px 16px",
    whiteSpace: "nowrap"
  },
  thTitle: {
    paddingLeft: 32
  },
  thStar: {
    width: 48
  },
  row: {
    selectors: {
      ":nth-child(odd) td": { backgroundColor: GMF_FOND_LIGNE },
      // Survol (ou focus clavier dans la ligne) : fond plus soutenu, déclaré après pour primer sur l'alternance.
      ":hover td, :focus-within td": { backgroundColor: GMF_FOND_SURVOL },
      ":hover a": { color: GMF_PETROLE }
    }
  },
  td: {
    fontSize: 13,
    lineHeight: "20px",
    padding: "12px 16px",
    borderBottom: `1px solid ${GMF_SEPARATEUR}`,
    verticalAlign: "middle",
    transition: "background-color .1s ease",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap"
  },
  tdTitle: {
    position: "relative",
    paddingLeft: 32,
    selectors: {
      // Barre bleue arrondie, centrée verticalement, à gauche de la ligne.
      "::before": {
        content: '""',
        position: "absolute",
        left: 8,
        top: 9,
        bottom: 9,
        width: 3,
        borderRadius: 2,
        backgroundColor: GMF_BLEU
      }
    }
  },
  tdLocalisation: {
    color: GMF_TEXTE_SECONDAIRE,
    borderLeft: `1px solid ${GMF_SEPARATEUR}`
  },
  tdStar: {
    width: 48,
    textAlign: "center",
    padding: "0 8px"
  },
  titleLink: {
    color: GMF_TEXTE,
    fontWeight: 500,
    textDecoration: "none",
    selectors: {
      ":hover, :focus-visible": { color: GMF_PETROLE, textDecoration: "underline" }
    }
  },
  starButton: {
    width: 28,
    height: 28,
    padding: 0,
    border: "none",
    borderRadius: 2,
    background: "transparent",
    color: GMF_PETROLE,
    cursor: "pointer",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    selectors: {
      ":hover:not(:disabled)": { backgroundColor: "#e6eff0" },
      ":focus-visible": { outline: `2px solid ${GMF_BLEU}`, outlineOffset: 1 },
      ":disabled": { cursor: "default", opacity: 0.5 }
    }
  },
  starRemoving: {
    selectors: { ":disabled": { opacity: 1 } }
  },
  starIcon: {
    fontSize: 14
  },
  empty: {
    padding: "12px 0",
    fontSize: 13,
    color: "#605e5c"
  }
});

/**
 * Affichage "page famille de produit" (homepage famille, repère ⑥ — feuille de route p. 9) :
 * tableau Titre / Localisation / ★. Le titre ouvre le lien du favori dans un nouvel onglet ;
 * l'étoile retire le favori (spinner bloquant géré par App pendant la suppression).
 */
export const FavoritesList: React.FC = () => {
  const { favorites, isLoading, loadError, refresh, removeFavorite, removingId, isBusy } = useFavorites();

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
    return <div className={classes.empty}>Aucun favori pour cette famille.</div>;
  }

  return (
    <table className={classes.table} aria-label="Mes favoris">
      <thead>
        <tr>
          <th scope="col" className={`${classes.th} ${classes.thTitle}`}>
            Titre
          </th>
          <th scope="col" className={classes.th}>
            Localisation
          </th>
          <th scope="col" className={`${classes.th} ${classes.thStar}`}>
            <span className="ms-hiddenVisually" style={{ position: "absolute", width: 1, height: 1, overflow: "hidden", clip: "rect(0 0 0 0)" }}>
              Retirer
            </span>
          </th>
        </tr>
      </thead>
      <tbody>
        {favorites.map((item) => (
          <tr key={item.id} className={classes.row}>
            <td className={`${classes.td} ${classes.tdTitle}`}>
              <a
                className={classes.titleLink}
                href={item.pageUrl}
                target="_blank"
                rel="noopener noreferrer"
                data-interception="off"
                title={item.title}
              >
                {item.title}
              </a>
            </td>
            <td className={`${classes.td} ${classes.tdLocalisation}`} title={item.localisation}>
              {item.localisation}
            </td>
            <td className={`${classes.td} ${classes.tdStar}`}>
              <button
                type="button"
                className={`${classes.starButton} ${item.id === removingId ? classes.starRemoving : ""}`}
                title="Retirer des favoris"
                aria-label={`Retirer « ${item.title} » des favoris`}
                disabled={isBusy}
                onClick={() => {
                  removeFavorite(item.id).catch(() => undefined);
                }}
              >
                <Icon iconName="FavoriteStarFill" className={classes.starIcon} />
              </button>
            </td>
          </tr>
        ))}
      </tbody>
    </table>
  );
};
