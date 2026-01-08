---
layout: default
title: "Analyser les sorties"
parent: "Formation : je n'ai jamais lancé QGIS"
nav_order: 4
---


# Formation à VOCAL : Analyser et interpréter les résultats de VOCAL


Cette section a pour objectif d’aider à **lire, comprendre et interpréter les résultats produits par les différents programmes de VOCAL**.  
Les outils proposés ne produisent pas uniquement des cartes ou des tableaux : ils fournissent des **indicateurs d’analyse** qui doivent être manipulés avec précaution et replacés dans leur contexte.

L’objectif de cette partie est de donner des **clés de lecture** permettant :
- d’identifier des tendances,
- de repérer des situations atypiques,
- d’orienter des analyses complémentaires ou des investigations de terrain.


Les données d'entrées ne sont pas fiables a 100%, l'appariement n'est également que partiel. Il est important alors de prendre les résultats de VOCAL comme une porte d'entrée pour identifier des zones/cas necessitant une analyse plus poussée.

---

## 4.0 - Visualiser les résultats 

Une fois les programmes de VOCAL exécutés, les résultats sont disponibles sous forme de **couches QGIS** (mémoire ou temporaires).  
Avant toute interprétation approfondie, il est essentiel de savoir **où trouver les informations**, **comment les lire**, et **comment les extraire** pour un usage ultérieur par exemple.

### Visualiser les résultats via la tables attributaires
Chaque programme de VOCAL produit une ou plusieurs couches enrichies d’attributs calculés.  
Ces attributs sont accessibles via la **table attributaire** de la couche de sortie.

Pour ouvrir la table attributaire :
1. Dans le panneau des couches, faire un **clic droit** sur la couche produite par VOCAL.
2. Cliquer sur **Ouvrir la table attributaire**.

![Ouverture de la table attributaire]({{ "/assets/images/formation/screen_visualisation_table_attributaire.png" | relative_url }})

*Figure X — Ouverture de la table attributaire d’une couche résultat VOCAL*

La table attributaire permet de :
- visualiser l’ensemble des indicateurs calculés
- trier les valeurs (par exemple par pente ou par ratio) comme avec un tableur
- effectuer des sélections

---

### Visualiser les résultats via la sélection d’un élément

Il est parfois plus lisible d’examiner les résultats **ouvrage par ouvrage** ou **zone par zone**.

1. Activer l’outil **Sélectionner des entités** dans la barre d’outils QGIS.
2. Vérifier que la couche de l'éléments que l'on veut inspecter est bien selectionnée.
3. Cliquer directement sur un ouvrage ou une zone dans la carte.

![Sélection d’un élément sur la carte]({{ "/assets/images/formation/screen_visualisation_selection_carte.png" | relative_url }})

*Figure X — Sélection d’un ouvrage ou d’une zone depuis la carte*

Une fois l’entité sélectionnée :
- elle apparaît surlignée en rouge sur la carte
- les informations de l'ouvrage s'affiche sur la droite de l'écran.

---

### Exporter les résultats sous forme de tableaux

Les résultats produits par VOCAL peuvent être exportés afin d’être utilisés :
- dans un tableur (Excel, LibreOffice)
- dans un rapport
- pour des analyses complémentaires

Pour cela :
1. Faire un **clic droit** sur la couche que l'on veut exporter.
2. Cliquer sur **Exporter** → **Sauvegarder les entités sous…**.

![Menu export des entités]({{ "/assets/images/formation/screen_export_menu.png" | relative_url }})

*Figure X — Menu d’export des entités dans QGIS*

Dans la fenêtre d’export :
1. Choisir le **format**. Par exemple : *Microsoft Excel (*.xlsx)*.
2. Définir le **chemin et le nom du fichier de sortie**.
3. Cocher **Exporter uniquement les entités sélectionnées** si nécessaire. Vous pouvez en effet utiliser des sélections pour exporter uniquement un sous-ensemble des résultats.
4. Cliquer sur **OK**.


Le fichier généré contient l’ensemble des champs attributaires, y compris :
- les indicateurs calculés par VOCAL
- les identifiants d’ouvrages ou de zones
- les éventuels champs descriptifs

---

## 4.1 Programme 1 — Évolution des volumes prélevés par ouvrage

Le programme **Évolution des volumes prélevés par ouvrage** permet d’analyser, pour chaque ouvrage individuellement, la dynamique historique des volumes prélevés à partir des données issues des campagnes de redevance de l’Agence.

### Principes généraux

Le calcul repose sur l’analyse de séries temporelles annuelles de volumes prélevés annuels.  
Pour chaque ouvrage, le programme produit plusieurs indicateurs permettant de caractériser :
- le **sens de l’évolution** (hausse, baisse, stabilité),
- l’**ampleur de cette évolution**,
- la **comparabilité** entre ouvrages de tailles différentes.

### Principaux indicateurs produits

- **Pente (slope)**  
  La pente mesure l’évolution moyenne absolue du volume prélevé, exprimée en m³/an.  
  Elle renseigne directement sur l’ampleur physique du changement.

- **Slope en pourcentage de la moyenne (`slope_pct_mean`)**  
  Cet indicateur exprime la pente en pourcentage de la moyenne des volumes prélevés par l’ouvrage.  
  Il permet de comparer des dynamiques relatives entre des ouvrages de tailles très différentes.

Pour illustrer : 2 ouvrages prelevant 1000m3 de plus par ans auront la même valeur de "slope", la croissance est la même. Cependant, si un ouvrages consomme 7000m3/an en moyenne et l'autre 45000m3/an en moyenne, la valeur de slope_pct_mean sera bien différente. L'évolution de 1000m3 par an estr bien plus important pour un ouvrage de 7000m3.

- **Slope en pourcentage du niveau initial (`slope_pct_first`)**  
  La pente est normalisée par rapport au niveau initial (moyenne des trois premières années).  
  Cet indicateur est utile pour évaluer l’évolution par rapport à la situation de départ. Il a pour avantage de lisser les conditions météorologiques sur la situation initiales.

- **CAGR (taux de croissance annuel composé en français)**  
  Le CAGR (Compound Annual Growth Rate) synthétise la croissance entre le début et la fin de la période étudiée en faisant la moyenne des trois premières années vs moyenne des trois dernières années.  
  Il offre une lecture simple, pratique dans le cas d'un bilan PGRE, mais ne reflète pas les fluctuations intermédiaires.

### Points de vigilance pour l’interprétation

- Les pentes calculées sur un faible nombre d’années sont plus sensibles aux anomalies. Je conseil de mettre minimum 4 années complète de données.
- La méthode **Theil-Sen** est à privilégier lorsque des valeurs aberrantes sont présentes. Sinon, les deux proposes des résultats comparables.
- Une évolution statistiquement marquée ne traduit pas nécessairement une non-conformité réglementaire.
- L’interprétation doit toujours être croisée avec le contexte météo/hydrologique et le contexte du prélèvement : sur un champ captant AEP par exemple, il est courant d'augmenter les prélèvements sur un des ouvrages pour baisser les prélèvements sur d'autres forages du champ.

---

## 4.2 Programme 2 — Évolution des volumes prélevés par zone

Ce programme applique les mêmes principes que le programme 1, mais à une **échelle territoriale**.  
Les volumes sont agrégés par zone (communes, bassins versants, polygones personnalisés) avant calcul des indicateurs.

### Différences avec l’analyse par ouvrage

Les volumes ici sont **agrégés spatialement** avant l’analyse temporelle. La dynamique observée reflète donc une **tendance collective** et non le comportement individuel des ouvrages.
L'évolution peut donc être due à :
  - une évolution de quelques ouvrages dominants
  - une évolution diffuse sur l’ensemble du territoire
  - l'ajout de nouveaux ouvrages sur la zone pendant la période donnée (Attention à ça!)
et le plus souvent, une combinaison des trois.

### Clés de lecture

Attention tout de même, une pente positive à l’échelle d’une zone peut masquer des situations contrastées entre ouvrages et à l’inverse, une stabilité globale peut dissimuler des évolutions fortes localisées.
Ce type d’analyse est particulièrement pertinent pour :
  - faire des diagnostics territoriaux
  - un regard global pour une bilan type PGRE
  - l’identification de zones à enjeux

Je conseil de toujorus croiser ces résultats avec une analyse par ouvrage et d'être très attentif aux ajouts d'ouvrages sur la période. On peut filtrer cela en obligeant tout les ouvrages à avoir un nombre d'années de redevance égale au nombre d'année de l'étude choisi ; C'est à dire, renseigner 6 ans dans le nombre d'années minimale pour que l'ouvrage soit pris en compte si on veut mesurer l'évolution entre 2018 et 2023.

---

## 4.3 Programme 3 — Ratio Volumes prélevés / Volumes autorisés par ouvrage

Le programme **Ratio VP / VA par ouvrage** permet de comparer, pour une année donnée, les volumes effectivement prélevés aux volumes autorisés issus des arrêtés DDTM.

### Principe du ratio

Le ratio est défini comme :

> **Ratio = Volume prélevé (VP) / Volume autorisé (VA)**

Il permet d’identifier :
- des situations de dépassement potentiel,
- des ouvrages proches de leur limite réglementaire,
- des incohérences ou des données manquantes.

### Indicateurs produits

- **Ratio VP / VA**  
  Valeur brute du rapport entre volume prélevé et volume autorisé.

- **Pourcentage de dépassement (`percent_overrun`)**  
  Exprime le ratio en pourcentage, facilitant la lecture cartographique.

### Points de vigilance

- Un ratio supérieur à 1 ne signifie pas automatiquement une infraction :
  - erreurs de saisie possibles,
  - décalages temporels entre autorisations et prélèvements,
  - pluralité d’arrêtés ou de régimes juridiques.
- Les ouvrages non appariés doivent être interprétés avec prudence.
- Les résultats doivent être croisés avec :
  - l’historique administratif,
  - les échanges avec les services instructeurs.

---

## 4.4 Programme 4 — Ratio Volumes prélevés / Volumes autorisés par zone

Ce programme applique le principe du ratio VP / VA à une **échelle agrégée**.  
Il vise à identifier des **déséquilibres structurels** sur un territoire donné.

### Intérêt principal

- Mettre en évidence des territoires où les volumes prélevés sont globalement proches ou supérieurs aux volumes autorisés.
- Fournir un indicateur synthétique à l’échelle d’un bassin versant ou d’un zonage de gestion.

### Limites d’interprétation

- L’agrégation peut masquer des situations individuelles très contrastées.
- Les volumes autorisés agrégés ne reflètent pas toujours la réalité réglementaire fine.
- Ce programme doit être utilisé comme **outil de diagnostic global**, non comme outil de contrôle individuel.

---

## 4.5 Programme 5 — État de connaissance des ouvrages Agence

Le programme **État de connaissance des ouvrages Agence** ne vise pas une analyse des prélèvements, mais un **diagnostic de qualité des données**.

### Objectifs principaux

- Identifier les ouvrages pour lesquels certaines informations essentielles sont manquantes.
- Repérer les incohérences ou les modifications récentes.
- Produire des indicateurs de complétude par zone.

### Types de résultats produits

- Liste des ouvrages sans coordonnées géographiques.
- Liste des ouvrages sans propriétaire ou interlocuteur renseigné.
- Indicateurs de complétude agrégés par territoire.
- Couches mettant en évidence les entités nécessitant une mise à jour.

### Utilisation recommandée

- Priorisation des actions de mise à jour des bases de données.
- Appui aux échanges entre partenaires (Agence, DDT(M), EPTB).
- Amélioration progressive de la fiabilité des analyses issues des autres programmes.

---

## Conclusion de la phase d’analyse

Les résultats produits par VOCAL constituent des **outils d’aide à la décision**.  
Ils doivent être interprétés avec méthode, esprit critique et en lien avec la connaissance métier des territoires.

VOCAL ne remplace pas l’expertise humaine : il la **structure, l’oriente et la rend plus lisible**.
