import * as React from "react";
import type { WebPartContext } from "@microsoft/sp-webpart-base";
import type { IReadonlyTheme } from "@microsoft/sp-component-base";
import { ThemeProvider, PartialTheme, useTheme } from "@fluentui/react/lib/Theme";
import { Stack } from "@fluentui/react/lib/Stack";
import { Text } from "@fluentui/react/lib/Text";
import { Overlay } from "@fluentui/react/lib/Overlay";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { mergeStyleSets } from "@fluentui/react/lib/Styling";
import { FavoritesProvider } from "../../context/FavoritesProvider";
import { useFavorites } from "../../context/FavoritesContext";
import { FavoritesService } from "../../services/FavoritesService";
import { FavoritesList } from "../FavoritesList/FavoritesList";
import { FavoritesCompactList } from "../FavoritesCompactList/FavoritesCompactList";
import { VoirPlusButton } from "../VoirPlusButton/VoirPlusButton";
import { FavoritesDisplayMode } from "../../models/IMesFavorisWebPartProps";

export interface IAppProps {
  context: WebPartContext;
  /** "accueil" = liste simple sans filtre famille ; "famille" = tableau filtré sur la famille de la page. */
  displayMode: FavoritesDisplayMode;
  /** Thème de la section SharePoint (variantes de section : neutre, accentué...). */
  themeVariant: IReadonlyTheme | undefined;
  title: string;
  /** Titre de la liste SharePoint des favoris (déjà résolu avec sa valeur par défaut). */
  listTitle: string;
  maxItems: number;
  voirPlusUrl: string;
  voirPlusLabel: string;
}

interface IFavoritesViewProps {
  displayMode: FavoritesDisplayMode;
  title: string;
  voirPlusUrl: string;
  voirPlusLabel: string;
  locale: string;
}

/**
 * Vue principale : en-tête façon web part SharePoint (titre + "Voir tout"),
 * liste des favoris, et calque bloquant avec spinner pendant une suppression.
 */
const FavoritesView: React.FC<IFavoritesViewProps> = (props) => {
  const { isBusy, actionError, clearActionError } = useFavorites();
  const theme = useTheme();
  const contentRef = React.useRef<HTMLDivElement>(null);

  // Pendant une suppression, le contenu devient "inert" : ni clic, ni focus clavier.
  React.useEffect(() => {
    const el = contentRef.current;
    if (!el) {
      return;
    }
    if (isBusy) {
      el.setAttribute("inert", "");
      if (el.contains(document.activeElement)) {
        (document.activeElement as HTMLElement).blur();
      }
    } else {
      el.removeAttribute("inert");
    }
  }, [isBusy]);

  const classes = mergeStyleSets({
    root: {
      position: "relative",
      color: theme.semanticColors.bodyText
    },
    header: {
      minHeight: 32,
      marginBottom: 8
    },
    title: {
      margin: 0,
      fontSize: 20,
      fontWeight: 600,
      lineHeight: "28px",
      color: theme.semanticColors.bodyText,
      overflow: "hidden",
      textOverflow: "ellipsis",
      whiteSpace: "nowrap"
    },
    // En-tête "page d'accueil" (maquette feuille de route p. 8) : titre + lien "Tous mes favoris".
    homeHeader: {
      display: "flex",
      justifyContent: "space-between",
      alignItems: "baseline",
      gap: 12,
      marginBottom: 12
    },
    homeTitle: {
      margin: 0,
      fontSize: 16,
      fontWeight: 600,
      lineHeight: "22px",
      color: "#323130",
      overflow: "hidden",
      textOverflow: "ellipsis",
      whiteSpace: "nowrap"
    },
    homeLink: {
      fontSize: 12,
      fontWeight: 600,
      color: "#1f6b6f",
      textDecoration: "none",
      whiteSpace: "nowrap",
      selectors: {
        ":hover, :focus-visible": { textDecoration: "underline", color: "#1f6b6f" }
      }
    }
  });

  // Page d'accueil : liste simple, sans étoile (donc sans suppression ni calque bloquant).
  if (props.displayMode === "accueil") {
    return (
      <div className={classes.root}>
        <div className={classes.homeHeader}>
          {props.title ? <h2 className={classes.homeTitle}>{props.title}</h2> : <span />}
          {props.voirPlusUrl && (
            <a className={classes.homeLink} href={props.voirPlusUrl}>
              {props.voirPlusLabel}
            </a>
          )}
        </div>
        <FavoritesCompactList />
      </div>
    );
  }

  return (
    <div className={classes.root} aria-busy={isBusy}>
      <div ref={contentRef}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center" className={classes.header}>
          {props.title ? (
            <Text as="h2" className={classes.title}>
              {props.title}
            </Text>
          ) : (
            <span />
          )}
          <VoirPlusButton pageUrl={props.voirPlusUrl} label={props.voirPlusLabel} disabled={isBusy} />
        </Stack>

        {actionError && (
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline={false}
            onDismiss={clearActionError}
            dismissButtonAriaLabel="Fermer"
            styles={{ root: { marginBottom: 8 } }}
          >
            {actionError}
          </MessageBar>
        )}

        <FavoritesList locale={props.locale} />
      </div>

      {isBusy && (
        <Overlay
          styles={{
            root: {
              zIndex: 10,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              cursor: "wait",
              backgroundColor: "rgba(255, 255, 255, 0.6)"
            }
          }}
        >
          <Spinner size={SpinnerSize.large} label="Suppression en cours..." ariaLive="assertive" />
        </Overlay>
      )}
    </div>
  );
};

/**
 * Composant racine : applique le thème de la section SharePoint, instancie le service
 * PnPjs et l'injecte dans le FavoritesProvider.
 */
export const App: React.FC<IAppProps> = (props) => {
  // Le service est recréé si le contexte, la liste ou l'emplacement change → rechargement automatique.
  // Accueil : pas de filtre famille ; Famille : filtre sur la famille de la page.
  const filterByFamily = props.displayMode !== "accueil";
  const service = React.useMemo(
    () => new FavoritesService(props.context, props.listTitle, filterByFamily),
    [props.context, props.listTitle, filterByFamily]
  );
  const locale = props.context.pageContext.cultureInfo.currentUICultureName || "fr-FR";

  // "Voir tous" : URL saisie dans le volet si renseignée, sinon lien automatique vers la liste
  // des favoris filtrée (utilisateur connecté + famille de la page).
  const [autoVoirPlusUrl, setAutoVoirPlusUrl] = React.useState<string>("");
  React.useEffect(() => {
    if (props.voirPlusUrl) {
      return;
    }
    let cancelled = false;
    service
      .getAllFavoritesUrl()
      .then((url) => {
        if (!cancelled) {
          setAutoVoirPlusUrl(url);
        }
      })
      .catch((err: unknown) => {
        // eslint-disable-next-line no-console
        console.error("[App] getAllFavoritesUrl", err);
      });
    return () => {
      cancelled = true;
    };
  }, [service, props.voirPlusUrl]);

  const voirPlusUrl = props.voirPlusUrl || autoVoirPlusUrl;

  return (
    <ThemeProvider
      theme={props.themeVariant as unknown as PartialTheme}
      style={{ background: "transparent" }}
    >
      <FavoritesProvider service={service} maxItems={props.maxItems}>
        <FavoritesView
          displayMode={props.displayMode}
          title={props.title}
          voirPlusUrl={voirPlusUrl}
          voirPlusLabel={props.voirPlusLabel}
          locale={locale}
        />
      </FavoritesProvider>
    </ThemeProvider>
  );
};
