# Dofus Item Price Runner
Application pour transformer un fichier excel contenant un ou plusieurs tableaux de prix afin de générer des graphiques à intervalles journaliers et mis en formes.

## Installation:
- Télécharger le fichier excel template contenant deux tableaux de départ (il est possible de supprimer le 2e sans problème) https://github.com/Zladyy/DofusItemPriceExcel/blob/main/DofusItemPriceExcel/Assets/Parchos.xlsx
- Télécharger le contenu du dossier Build https://github.com/Zladyy/DofusItemPriceExcel/tree/main/Build
- Exécuter le setup.exe contenu dans le dossier afin d'installer l'application sur votre machine

## Utilisation:
### Le fichier template se présente comme suit :

> [!CAUTION]
> Chaque tableau doit être espacé par une colonne


![Excel template](https://github.com/Zladyy/DofusItemPriceExcel/assets/60046967/654acfda-cf03-4000-a119-736a626a0928)

### L'application se présente comme suit :

La première exécution requiert de sélectionner un fichier (les choix faits seront sauvegardés) afin de rendre le bouton `RUN` disponible.

![Exécutable](https://github.com/Zladyy/DofusItemPriceExcel/assets/60046967/84063b09-c594-459c-bf96-53d154a67604)


![Exécutable 2](https://github.com/Zladyy/DofusItemPriceExcel/assets/60046967/eaf7b4ca-2c36-4e6d-bff6-ac864c68a5c4)
![Exécutable 3](https://github.com/Zladyy/DofusItemPriceExcel/assets/60046967/16ee4a91-9ac6-40b8-a0c9-8359421cc516)


Après avoir cliqué sur le bouton `RUN`, le fichier excel est ouvert.

Une page d'aggrégation des données est créée (cette page sera masquée à la fin de l'exécution de l'application).
Si le tableau initial contient plusieurs prix pour le même jour, une moyenne est faite.
Les colonnes "Buy" et "Sell" ne sont créées que si une valeur a été fournie pour le seuil d'achat et de vente.

![AggregatedData](https://github.com/Zladyy/DofusItemPriceExcel/assets/60046967/4e487723-c31d-4972-b701-12b5596acc1c)

Ensuite la page des graphiques est créée.
Les lignes de seuil d'achat et de vente ne sont créées que si une valeur été fournie pour le seuil d'achat et de vente.

> [!NOTE]
> L'application génère les graphiques à une taille adéquate pour un écran en 1920x1080.

![Graphique](https://github.com/Zladyy/DofusItemPriceExcel/assets/60046967/292960e7-a919-4920-8009-a78d3c55c320)

