import * as React from "react";
import styles from "./FavoriteItem.module.scss";
import { IFavoriteItem } from "../../models/IFavoriteItem";
import { useFavorites } from "../../context/FavoritesContext";

export interface IFavoriteItemProps {
  item: IFavoriteItem;
}

/**
 * Une ligne de favori : titre cliquable (lien vers la page) + bouton étoile.
 * L'étoile est jaune par défaut ; cliquer dessus retire le favori de la liste
 * (suppression de l'item dans la liste SharePoint "MesFavoris").
 */
export const FavoriteItem: React.FC<IFavoriteItemProps> = ({ item }) => {
  const { removeFavorite } = useFavorites();
  const [isRemoving, setIsRemoving] = React.useState<boolean>(false);
  const isMountedRef = React.useRef<boolean>(true);

  React.useEffect(() => {
    return () => {
      isMountedRef.current = false;
    };
  }, []);

  const onRemoveClick = async (): Promise<void> => {
    if (isRemoving) {
      return;
    }
    setIsRemoving(true);
    await removeFavorite(item.id);
    // Si la suppression a réussi, le composant sera démonté par le re-render de la liste.
    // On ne touche à l'état que s'il est toujours monté (cas d'un échec restauré).
    if (isMountedRef.current) {
      setIsRemoving(false);
    }
  };

  return (
    <li className={styles.favoriteItem}>
      <a
        className={styles.favoriteLink}
        href={item.pageUrl}
        title={item.title}
        target="_blank"
        rel="noopener noreferrer"
      >
        {item.title}
      </a>
      <button
        type="button"
        className={styles.starButton}
        aria-label={`Retirer "${item.title}" des favoris`}
        title="Retirer des favoris"
        disabled={isRemoving}
        onClick={onRemoveClick}
      >
        <span className={`${styles.starIcon} ${isRemoving ? styles.spinning : ""}`} aria-hidden="true">
          ★
        </span>
      </button>
    </li>
  );
};
