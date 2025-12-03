# Valorisateur des Ouvrages Connus A L'Agence - VOCAL - Plugin Orchestrateur de programmes périphériques de valorisation automatique

**Valorisateur des Ouvrages Connus À l'Agence (VOCAL)**

Ce dépôt contient : le plugin QGIS *Orchestrateur* ("VOCAL") et cinq scripts Processing (algorithmes) permettant d'analyser, comparer et valoriser les bases de connaissances relatives aux ouvrages de prélèvements d'eau connus de l'Agence de l'eau Rhône Méditerannée Corse ainsi que des DDT(M)s.

Le projet vise à rapprocher et valoriser deux sources d'information majeures :
- les volumes prélevés issus des campagnes de redevances de l'Agence de l'eau (volumes prélevés annuellement)
- les volumes autorisés des ouvrages sous régime IOTA issus des bases des DDTM(M)s.

L'objectif de ce plugin est de produire des indicateurs (évolutions temporelles de consommation, ratios prélevé / autorisé) et simplifier la production de cartes et exports valorisables.

---

## Structure du README

Ce fichier contient :
- Une page principale (présentation, installation, usage global du plugin).
- Une page par programme (description, paramètres, sorties, exemples d'usage, recommandations QML et bonnes pratiques).

Les pages suivent dans ce même document.

---

#Le VOCAL — Projet & Plugin

## Présentation rapide
VOCAL est un plugin QGIS qui facilite :
- la préparation d'une zone d'étude (chargement / extraction mémoire),
- la copie automatique des scripts Processing (dans `Processing/scripts` utilisateur) pour rendre accessibles plusieurs algorithmes de valorisation personnalisés,
- le lancement gui-friendly des algorithmes de traitement (pentes, ratios, etc.).
Les outils de valorisation sont portés par les algorithmes, le plugin n'est qu'un orchestrateur de ces programmes.

Le plugin fournit une interface en 2 étapes :
1. choix du programme + choix et chargement de la zone d'étude et potentiellement des sous-zones de travails (avec option pour créer une couche mémoire restreinte)
2. copie des scripts et ouverture automatique de la boîte d'outil Processing pour l'algorithme sélectionné.

## Contenu du dépôt
- `prelev_orchestrator/` : code du plugin (dialog, actions, scripts utilitaires, icônes et QML de démo)
- `scripts/` : scripts Processing (les 5 programmes (pour le moment), nommés et documentés ci-après)
- `Couches/` (optionnel) : exemples de geopackages de référence (départements, communes, BV, nappes)
- `QML/` : qml de styles utilisés par défaut
- `README.md` (ce document)

# Installation

### 1) Installation du plugin
- Option A — depuis GitHub : télécharger le dossier (cliquer sur le bouton _\<\> _Code -> _Download _ZIP). Une fois télécharger, lancer QGIS, cliquer sur l'onglet _Extension -> *Installer depuis un fichier ZIP*. Lancer le nouveau programme de votre barre des tâches QGIS *Le VOCAL*. Lancer une initiation en selectionnant une programme et une zone d'étude de votre choix, faites *Suivant*. Sur la nouvelle fenêtre qui s'ouvre faites *Ouvrir l'outil processing*. Un message d'erreur devrait apparaitre, c'est l'initialisation. Redemarrer QGIS. Le plugin est désormais prêt.
- **!! PAS ENCORE DISPONIBLE !! ** Option B — depuis QGIS Plugin Repo : (si publié) installer depuis _Extension > Installer > Chercher VOCAL dans la barre de recherche > Installer_.

### 2) Données auxiliaires (couches de référence)
Pour l'option A : les couches lourdes (GeoPackage, bases de fonds)sont embarquées dans le ZIP du plugin.
**!! PAS ENCORE DISPONIBLE !!** Pour l'option B : Un mécanisme de téléchargement à la première ouverture va être ajoutés via script interne pour automatiser la récupération des gpkg depuis GitHub Releases. outés via script interne) pour automatiser la récupération des gpkg depuis GitHub Releases.


## Utilisation basique (workflow)
1. Ouvrir le plugin VOCAL (menu / icône) — la première page permet de choisir le programme.
2. Choisir l'échelle (ex : Département, BV, Commune) et la valeur de la zone, charger la zone puis créer la couche mémoire (option recommandée).
3. Choisir des options de QML (Styles apportés aux couches dans QGIS) si souhaité (appliquer QML aux couches chargées).
4. Ajouter ou non des sous-zonages. Particulièrement utile pour les deux programmes basés sur de l'analyse par territoire. 
5. Cliquer sur *Suivant* pour copier les scripts Processing dans le dossier utilisateur (s'il manque) et ouvrir l'outil Processing correspondant. Potentiellement redemarrer QGIS lors de la première utilisation
6. Lancer l'algorithme dans la fenêtre Processing (ou modifier les paramètres) — les scripts se basent sur la zone mémoire si elle est créée.

---

# Programme 1 — Evolution des prélèvements par ouvrage (`compute_slopes_ouvrage_only`)

## Objectifs
Calculer, pour chaque ouvrage identifié, l'évolution temporelle des volumes prélevés par année. Produit des indicateurs normalisés : pentes en % par rapport à la moyenne, CAGR (growth rate) et z-score.

## Paramètres principaux
- **Couche d'entrée (points / table)** : prélèvements Agence.
- **Champ année** (numérique ou convertible) (obligatoire)
- **Champ identifiant ouvrage** (obligatoire)
- **Champ volume (assiette)** (obligatoire)
- **Champ nom de l'ouvrage (Libellé ouvrage)** (optionnel)
- **Champ nom du contribuable (Contribuable)** (optionnel)
- **Méthode** : `OLS` ou `Theil-Sen` (Thiel-Sen plus robuste) (obligatoire)
- **Années min / plage (start/end)** (obligatoire)
- **Appliquer QML** (optionnel)

## Sortie
Couche (points ou mémoire) contenant par ouvrage : `ouvrage_id`, `slope_ouvrage`, `n_years_ouvrage`, `name_ouv`, `name_petitionaire`, `mean_vol_ouv`, `slope_pct_mean`, `slope_pct_first`, `cagr_pct`, `slope_pct_z`.

## Notes & recommandations
- Theil-Sen est recommandé s'il existe des valeurs aberrantes.
- Calculer les pentes sur des séries avec un nombre minimal d'années (paramétrable).
- Conserver la géométrie du premier point trouvé par ouvrage pour cartographie.

Note d'analyse des indicateurs : 
La _pente_ _(slope)_ mesure l’évolution moyenne absolue du volume prélevé par ouvrage en m³/an (estimée par régression _OLS_ ou _Theil-Sen_) et renseigne l’ampleur physique du changement. Le _slope_pct_mean_ exprime cette pente en pourcentage de la moyenne des volumes de l’ouvrage (100 × slope / mean), ce qui permet de comparer la dynamique relative entre ouvrages de tailles différentes. Le _slope_pct_first_ normalise la pente par rapport au niveau initial, la moyenne des 3 premières années, pour évaluer la variation par rapport au point de départ. Enfin, le _CAGR_ (_taux de croissance annuel composé_) synthétise la croissance équivalente entre une période de départ et une période finale ( moyenne 3 premières vs 3 dernières années) ; il est utile pour résumer une trajectoire début→fin mais masque les fluctuations intermédiaires.

---

# Programme 2 — Evolution des prélèvements par zonage (`compute_slopes_zones`)

## Objectifs
Calculer des évolutions temporelles de prélèvements pour des zones (communes, BV, polygones projetés) en agrégeant les volumes des ouvrages situés dans chaque zone.

## Différences avec le Programme 1
- Agrégation par zone (somme des volumes par année puis estimation de la pente) et non par ouvrage.
- Utilisation d'intersections spatiales pour assigner chaque ouvrage à un ou plusieurs zones (selon la logique choisie).

## Paramètres
- **Couche zonage** (polygone) (obligatoire)
- **Couche d'entrée (points / table)** : prélèvements Agence.
- **Champ année** (numérique ou convertible) (obligatoire)
- **Champ identifiant ouvrage** (obligatoire)
- **Champ volume (assiette)** (obligatoire)
- **Méthode** : `OLS` ou `Theil-Sen` (Thiel-Sen plus robuste) (obligatoire)
- **Années min / plage (start/end)** (obligatoire)
- **Appliquer QML** (optionnel)
- Options d'agrégation (contiguïté/centré, gestion des doublons)

## Sortie
Couche des zones enrichie avec des indicateurs d'évolution par zone. Pour l'explication des indicateurs voir _Programme 1_.

---

# Programme 3 — Ratio Volumes prélevés (VP)/Volumes autorisés (VA) par ouvrage (`compare_prelevements_autorises`)

## Objectifs
Pour une année donnée, comparer le volume prélevé (VP, "assiettes" retenues à l'Agence) à un volume autorisé (VA, arrétés de déclaration ou d'autorisation DDTM). Déterminer les dépassements et fournir des indicateurs de ce ratio.

## Paramètres
- **Couche zone d'étude** (polygones) : filtre spatial facultatif mais recommandé.
- **Couche prélèvements** : champs année, id ouvrage, assiette (volume), champ type de milieu (optionnel), champ nom ouvrage & interlocuteur (optionnels).
- **Couche volumes autorisés** : champ ID ouvrage, champ volume autorisé (VA), champ ID DDTM (optionnel).
- **Année d'étude** : mettre 0 pour utiliser la dernière année disponible.
- **Inclure non-appariés** : booléen.
- **Appliquer QML** : chemin du QML (optionnel)

## Traitement
- Filtrage spatial (zone) si la couche zone a des géométries.
- Agréger les volumes par ID ouvrage pour l'année choisie.
- Joindre avec la table autorisée : prendre `MAX(VA)` si plusieurs enregistrements, concaténer champs DDTM distincts.
- Calculer `ratio = VP / VA` (si VA non nul) et `% overrun`.

## Sortie
Couche par ouvrage pour l'année choisie : `annee`, `ouvrage_id`, `ouvrage_name`, `interlocuteur`, `assiette`, `vol_autorise`, `ddtm_id`, `ratio`, `ratio_possible`, `percent_overrun`, `note`, `type_milieu`.

## Note sur les indicateurs
- Le ratio représente réellement la division du VP/VA
- Le %overrun présente le pourcentage que représente le VP/VA. 

## Recommandations QML
- Utiliser `data-defined` properties (size, outline, fill, shape) pour afficher :
  - halo si dépassement (ex : grande taille semi-transparente),
  - couleur différente si dépassement,
  - forme étoile si `modified_recent` (ou autre indicateur binaire);
- Eviter les renderers rule-based uniquement si tu veux que plusieurs propriétés (forme / couleur / taille) dépendent de champs différents : préférer un renderer categorized sur un champ, avec des data-defined properties pour taille/couleur/outline.

---

# Programme 4 — Ratio Volumes prélevés (VP)/Volumes autorisés (VA) pour des zones définies (`zones_compare_prelev_autorise`)

## Objectif
Similaire au programme 3 mais agrégé par zone (somme des volumes prélevés par zone) pour une année donnée et comparaison avec volumes autorisés agrégés si disponible. Ce calcul est simplement la pour afficher un problème structurel sur un territoire donné.

## Paramètres & Sortie
Analogue à `compare_prelevements_autorises` mais à l'échelle du zonage.
- **Couche zonage** (polygone) (obligatoire)

## Note sur les indicateurs
- Les indicateurs sont les mêmes que pour le _Programme 3_


---

# Programme 5 — État connaissance - ouvrages Agence (`compute_connaissance_ouvrages_agence`)

## Objectifs
Fournir un diagnostic de la qualité / complétude des ouvrages connus par l'Agence :
- lister les ouvrages sans propriétaire renseigné,
- lister les ouvrages sans coordonnées,
- lister les modifications récentes sur les ouvrages
- produire des métriques de complétude (nombre d'attributs essentiels manquants) par zone.

## Paramètres
- Couche ouvrages Agence (Issues des redevances ou du QGIS mutualisé Agence de l'eau)
- (Optionnel) couche zone pour agrégation
- Seuils / règles de détection (ex : date de modification)

## Sorties
- Table résumée par ouvrage (qualité des données),
- Couche mémoire des entités problématiques (manque coord, propriétaire absent, etc.),
- Rapport textuel sommaire (optionnel).



---

# Structure du Plugin

```
VOCAL_Plugin/
├── prelev_orchestrator.py     # plugin QGIS
├── scripts/   # Scripts Processing (5 fichiers .py)
│   ├── compute_connaissance_ouvrages_agence.py
│   ├── zones_compare_prelev_autorise.py
│   ├── compare_prelevements_autorises.py       
│   ├── compute_slopes_zones.py
│   └── compute_slopes_ouvrage_only.py              
├── QML/	# Dossier contenant les QML des couches de bases et des couches de sorties des algorithmes
├── __init__.py
├── README.md           # Information concernant le Plugin (ce document)
├── Couches/   # Dossier contenant des couches de bases
├── metadata.txt     # Données d'information de lancement du plugin
├── icon.png  	# Icone du Plugin
└── LICENSE  # License de distribution

```

---

# Problèmes possibles, débuggage & FAQ

### QML : `unexpected character` lors du chargement
- Cela peut venir d'un caractère non-utf8, d'une erreur de guillemets ou d'un commentaire mal placé.
- Utilise un éditeur qui montre les caractères invisibles (VSCode, Notepad++) et vérifie l'encodage UTF-8 sans BOM.
- Préfère la structure `renderer` categorised + `dataDefinedProperties` pour définir taille/couleur/forme dynamiquement.

### `pyrcc5` non trouvé
- Soit tu installes `pyqt5`/outils Qt via ton Python local, soit tu n'utilises pas `resources.py` et charges les icônes avec `QIcon(path)`.

### Ma couche projet ne s'affiche pas correctement après `loadNamedStyle`
- Vérifie la correspondance des noms de champs utilisés dans le QML et la couche réelle ; dans tes QML utilises `@layer` ou remplace le nom du champ dynamiquement.


---

# Contacts / support
- Auteur : Aurel Lashermes.
- Structure porteuse du projet : Délégation de Montpellier de l'Agence de l'Eau RMC
- Date : _Automne 2025_
- Tracker GitHub : ouvrir des *issues* pour bugs/ameliorations.

---

