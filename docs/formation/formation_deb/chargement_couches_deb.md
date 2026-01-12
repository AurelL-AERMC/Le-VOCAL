---
layout: default
title: "Charger des données"
parent: "Formation : je n'ai jamais lancé QGIS"
nav_order: 2
---


# 2. Charger les données dans QGIS

Une fois le plugin VOCAL installé, nous allons pouvoir commencer le travail en chargeant les données nécessaires aux analyses.

Dans QGIS, les données sont organisées sous forme de **couches**.  
Chaque couche correspond à un fichier ou à une table de données, qui peut contenir :
- des informations spatiales (points, lignes, polygones),
- des informations attributaires (tableaux de valeurs).
- les deux à la fois

Les programmes de VOCAL s’appuient sur ces couches pour effectuer leurs calculs.  
Il est donc essentiel que les données soient **correctement chargées** et **bien interprétées par QGIS**.

Pour cette formation, nous verrons deux formats de données utilisés par l’Agence de l’eau et des DDTM :
- les fichiers **CSV** (tableaux de données),
- les fichiers **GeoPackage** (données spatiales).

---

## 2.1 Pas à pas : charger des données

Lorsque QGIS est lancé pour la première fois, l’interface est vide : aucune couche n’est chargée.

![Interface QGIS vide]({{ "/assets/images/formation/screen_qgis_vide.png" | relative_url }})

*Figure X — Interface de QGIS sans données chargées*

Le chargement des données se fait via le **Gestionnaire de sources de données**, accessible :
- par l’icône dédiée dans la barre d’outils,
- ou par le raccourci clavier **CTRL + L**.

![Icône du gestionnaire de données]({{ "/assets/images/formation/screen_gestionnaire_donnees_icone.png" | relative_url }})

*Figure X — Icône du gestionnaire de sources de données*

---

## 2.2 Charger des données sous forme de CSV

### 2.2.1 Qu’est-ce qu’un fichier CSV ?

Un fichier CSV (*Comma Separated Values*) est un fichier texte contenant des données séparé par un marqueur (virgule, point virgule, slash etc.).  
Il peut être ouvert dans un tableur (Excel, LibreOffice), mais il reste **un fichier texte brut**.

Les colonnes sont séparées par un caractère appelé **séparateur**, qui peut varier selon les contextes :
- point-virgule `;` (le plus courant en France),
- virgule `,`,
- tabulation,
- autre caractère spécifique.

Ce point est **fondamental** :  si le séparateur ou l’encodage sont mal interprétés, les données seront mal lues par QGIS.

---

### 2.2.2 Ouvrir un CSV dans QGIS

1. Ouvrir le **Gestionnaire de sources de données** (CTRL + L).
2. Sélectionner l’onglet **Texte délimité**.

![Onglet Texte délimité]({{ "/assets/images/formation/screen_gestionnaire_donnees_csv.png" | relative_url }})

*Figure X — Onglet « Texte délimité » du gestionnaire de données*

3. Cliquer sur **Parcourir** et sélectionner le fichier CSV à charger.

---

### 2.2.3 Vérifier l’encodage du fichier 

Les données de l’Agence de l’eau sont généralement encodées en **ISO-8859-1**.  
Un mauvais encodage entraîne :
- des caractères accentués illisibles,
- des erreurs dans les champs texte,
- parfois des échecs de jointure.

Dans le gestionnaire :
- repérer le champ **Encodage du fichier**,
- sélectionner **ISO-8859-1** si ce n’est pas déjà le cas.

---

### 2.2.4 Vérifier le séparateur de colonnes

Dans la section **Délimiteurs** :
- Selectionner un séparateur persolannisé
- tester le point-virgule `;` en premier lieu,
- vérifier dans l’aperçu que les colonnes sont correctement séparées.

Si toutes les données apparaissent dans une seule colonne, le séparateur est incorrect.

![Vérification du séparateur CSV]({{ "/assets/images/formation/screen_gestionnaire_donnees_csv.png" | relative_url }})

*Figure X — Vérification du séparateur de colonnes*

---

### 2.2.5 Vérifier le format des identifiants Agence

Les identifiants d’ouvrages Agence doivent impérativement être :
- lus comme **texte**,
- et non comme des nombres.

Dans la section **Définition des champs** :
- vérifier que le champ identifiant est de type **Texte (string)**,
- éviter toute conversion automatique en nombre.

Ceci est essentiel pour :
- les jointures,
- l’appariement des bases,
- le bon fonctionnement des programmes VOCAL.


---

### 2.2.6 Finaliser le chargement du CSV

Une fois toutes les vérifications effectuées :
- cliquer sur **Ajouter**,
- puis sur **Fermer**.

La couche CSV apparaît alors dans le panneau des couches.

---

## 2.3 Charger des données sous forme de GeoPackage

### 2.3.1 Qu’est-ce qu’un GeoPackage ?

Le **GeoPackage** (`.gpkg`) est un format de données spatiales moderne, robuste et largement utilisé dans QGIS.

Un GeoPackage peut contenir :
- plusieurs couches spatiales (points, lignes, polygones),
- plusieurs tables attributaires,
- des métadonnées.

C’est le format privilégié pour :
- les zonages,
- les référentiels géographiques,
- les bases consolidées.

---

### 2.3.2 Ouvrir un GeoPackage dans QGIS

1. Ouvrir le **Gestionnaire de sources de données**.
2. Sélectionner l’onglet **GeoPackage**.

![Onglet GeoPackage]({{ "/assets/images/formation/screen_gestionnaire_donnees_gpkg.png" | relative_url }})

*Figure X — Onglet GeoPackage du gestionnaire de données*

3. Cliquer sur **Parcourir** et sélectionner le fichier `.gpkg`.

---

### 2.3.3 Choisir les couches à charger

Une fois le GeoPackage sélectionné :
- la liste des couches qu’il contient apparaît,
- sélectionner une ou plusieurs couches selon les besoins,
- cliquer sur **Ajouter**.


Les couches sélectionnées apparaissent alors dans le panneau des couches QGIS.

---

## 2.4 Bonnes pratiques avant d’utiliser VOCAL

Avant de lancer les programmes de VOCAL, il est recommandé de vérifier :
- que toutes les couches nécessaires sont bien chargées,
- que les identifiants sont cohérents entre les différentes bases,
- que les champs requis sont présents et correctement typés,
- que les données correspondent bien à la zone d’étude envisagée.

Cette étape de préparation des données conditionne directement la **qualité des résultats produits**.

---

Une fois toutes nos données chargées, nous pouvons désormais passer à l’étape suivant. Pour cela, vous pouvez cliquer ici : [Utiliser Vocal](utilisation_vocal_deb)
