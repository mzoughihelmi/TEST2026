import * as React from "react";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import styles from "./FavoritesList.module.scss";
import { useFavorites } from "../../context/FavoritesContext";
import { FavoriteItem } from "../FavoriteItem/FavoriteItem";

/**
 * Affiche la liste des favoris issus du contexte (déjà limitée aux N derniers
 * favoris créés par l'utilisateur connecté par le Provider/Service).
 */
export const FavoritesList: React.FC = () => {
  const { favorites, isLoading, error } = useFavorites();

  if (isLoading) {
    return <Spinner size={SpinnerSize.medium} label="Chargement de vos favoris..." />;
  }

  if (error) {
    return <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>;
  }

  if (favorites.length === 0) {
    return <div className={styles.emptyState}>Vous n'avez pas encore de favoris.</div>;
  }

  return (
    <ul className={styles.favoritesList}>
      {favorites.map((item) => (
        <FavoriteItem key={item.id} item={item} />
      ))}
    </ul>
  );
};
