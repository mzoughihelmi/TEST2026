import * as React from "react";
import type { WebPartContext } from "@microsoft/sp-webpart-base";
import styles from "./App.module.scss";
import { FavoritesProvider } from "../../context/FavoritesProvider";
import { FavoritesService } from "../../services/FavoritesService";
import { FavoritesList } from "../FavoritesList/FavoritesList";
import { VoirPlusButton } from "../VoirPlusButton/VoirPlusButton";

export interface IAppProps {
  context: WebPartContext;
  title: string;
  maxItems: number;
  voirPlusUrl: string;
  voirPlusLabel: string;
}

/**
 * Composant racine : instancie le service PnPjs, l'injecte dans le FavoritesProvider,
 * puis assemble la liste des favoris et le bouton "voir plus".
 */
export const App: React.FC<IAppProps> = (props) => {
  // Le service n'est recréé que si le contexte du web part change.
  const service = React.useMemo(() => new FavoritesService(props.context), [props.context]);

  return (
    <FavoritesProvider service={service} maxItems={props.maxItems}>
      <div className={styles.app}>
        {props.title && <h2 className={styles.title}>{props.title}</h2>}
        <FavoritesList />
        <VoirPlusButton pageUrl={props.voirPlusUrl} label={props.voirPlusLabel} />
      </div>
    </FavoritesProvider>
  );
};
