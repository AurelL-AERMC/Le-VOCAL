---
layout: default
title: Documentation
parent: Accueil
nav_order: 2
---

# Documentation technique — Plugin QGIS VOCAL

**Valorisateur des Ouvrages Connus À L'Agence — Orchestrateur de programmes périphériques de valorisation automatique et d'analyse des prélèvements d'eau en milieu naturel**

*Choix méthodologiques et architecture technique*

*Aurel Lashermes — Agence de l'eau Rhône-Méditerranée-Corse — 2026*

---

## 1. Présentation générale de VOCAL

### 1.1 Contexte et objectifs

VOCAL (Valorisateur des Ouvrages Connus À l'Agence) est un plugin QGIS développé dans le cadre des missions de l'Agence de l'eau Rhône-Méditerranée-Corse et des Directions Départementales des Territoires et de la Mer (DDTM). Son objectif est de fournir un outil opérationnel d'analyse des prélèvements en eau à partir des données de redevance de l'Agence et des arrêtés de volumes autorisés des DDTM.

Le plugin vise à répondre à plusieurs besoins identifiés lors des bilans de Plans de Gestion de la Ressource en Eau (PGRE) et lors de l'Étude Données Prélèvement portée par l'Agence : automatiser les calculs d'évolution temporelle des volumes prélevés, permettre la comparaison entre volumes prélevés et volumes autorisés, et offrir une représentation cartographique immédiate des résultats.

### 1.2 Architecture générale : un plugin orchestrateur + des scripts Processing

Le choix architectural central de VOCAL est la séparation entre une interface utilisateur légère (l'orchestrateur) et des algorithmes de traitement indépendants (les scripts Processing). Cette décision répond à plusieurs contraintes identifiées dès la conception :

- **Maintenabilité** : les algorithmes métier peuvent être modifiés, corrigés ou enrichis sans toucher à l'interface principale.
- **Interopérabilité** : les scripts Processing sont utilisables directement depuis la Toolbox QGIS, sans passer par l'interface VOCAL, ce qui facilite les tests et les usages avancés.
- **Extensibilité** : l'ajout d'un nouveau programme se résume à l'ajout d'une entrée dans le dictionnaire `ALGO_INFOS` de l'orchestrateur et à l'écriture du script correspondant.

L'orchestrateur (`prelev_orchestrator.py`) joue le rôle de point d'entrée unique : il gère le choix du programme, la définition de la zone d'étude, le chargement des sous-zonages optionnels, l'application de styles visuels (QML) et le lancement du script Processing approprié via `processing.execAlgorithmDialog()`.

Les cinq programmes disponibles sont référencés dans un dictionnaire central `ALGO_INFOS` qui associe à chaque programme son identifiant d'algorithme Processing, le nom du fichier script source, et le nom du fichier de style QML associé :

```python
ALGO_INFOS = {
    'Evolution des volumes prélevés par ouvrage': {
        'alg_id': 'script:compute_slopes_ouvrage_only',
        'script_name': 'compute_slopes_qgis_ouvrages.py',
        'qml_name': 'ouvrages_slopes_QML.qml'
    },
    ... (4 autres programmes)
}
```

---

## 2. L'orchestrateur : choix de conception de l'interface

### 2.1 Interface mono-page et mémorisation de la zone d'étude

À partir de la version 0.0.3 de VOCAL, l'interface devient mono-page (`QDialog`) regroupant tous les paramètres de configuration, à la place d'une navigation multi-étapes. Ce choix vise à réduire les frictions pour l'utilisateur : la zone d'étude, paramètre central à tous les programmes, est mémorisée entre les sessions via le système `QgsSettings` de QGIS.

La mémorisation repose sur le stockage de l'échelle (ex. : "Départements") et de la valeur sélectionnée (ex. : "GARD") dans les paramètres persistants de QGIS. Au rechargement du plugin, si une couche portant le nom attendu est déjà présente dans le projet, elle est réutilisée sans rechargement. Dans le cas contraire, l'interface signale à l'utilisateur que le chargement manuel est requis, sans tenter de recréer automatiquement une couche potentiellement coûteuse.

### 2.2 Gestion des zones d'étude : couche mémoire vs couche complète

Pour chaque zone d'étude, VOCAL propose deux modes de chargement :

- **Couche mémoire filtrée (mode recommandé)** : seules les entités intersectant la zone sélectionnée sont copiées dans une couche mémoire temporaire, nommée selon le pattern `[source]_INTER_[valeur]`. Ce mode est plus performant pour les traitements ultérieurs car il réduit le volume de données chargé en mémoire QGIS.
- **Couche complète** : toute la couche source est ajoutée au projet, avec une sélection appliquée sur les entités de la zone. Ce mode est utilisé en fallback si la création de la couche mémoire échoue.

La création de la couche mémoire est gérée par la fonction `create_memory_layer_from_features()`, qui vérifie d'abord si une couche portant le même nom existe déjà dans le projet pour éviter les doublons. Cette vérification est essentielle dans un contexte d'usage multi-sessions où l'utilisateur peut rouvrir VOCAL plusieurs fois sur le même projet QGIS.

### 2.3 Intersection spatiale et index spatial

L'intersection entre les couches de données et la zone d'étude est réalisée par la fonction `intersect_layer_with_reference()`. Pour des raisons de performance, cette fonction construit un index spatial `QgsSpatialIndex` sur la couche de référence avant de parcourir les entités de la couche source. L'algorithme utilise d'abord une recherche par boîte englobante (`intersects` par bounding box) pour identifier les candidats, puis vérifie l'intersection géométrique exacte sur ce sous-ensemble réduit.

Ce choix évite le recours à l'algorithme Processing natif "Intersection" de QGIS, qui crée systématiquement de nouvelles géométries découpées. Ici, l'objectif est uniquement de filtrer les entités qui touchent la zone, en conservant intactes leurs géométries originales, ce qui est suffisant pour les besoins de VOCAL et significativement plus rapide.

### 2.4 Injection dynamique des chemins QML dans les scripts

Un problème pratique posé par la distribution du plugin sur des postes aux configurations diverses (chemins réseau, postes locaux, différentes versions Windows) est la localisation des fichiers de style QML. Plutôt que d'inscrire un chemin absolu dans chaque script Processing, VOCAL adopte une approche d'injection dynamique.

Lors du lancement d'un programme, l'orchestrateur copie le script source vers le dossier Processing de l'utilisateur (`QgsApplication.qgisSettingsDirPath()/processing/scripts/`). Pendant cette copie, il modifie à la volée le contenu du script via des expressions régulières pour remplacer le chemin QML codé en dur par le chemin dynamique calculé à partir de `PLUGIN_DIR` :

```python
script_content = re.sub(
    r'(default_qml\s*=\s*r?["\']).*?(["\'])',
    r'\1' + qml_path_normalized + r'\2',
    script_content
)
```

Ce mécanisme garantit que les scripts déployés sur le poste de l'utilisateur pointent toujours vers le bon répertoire QML, indépendamment de l'emplacement d'installation du plugin. Une vérification de contenu (comparaison du fichier existant avec le contenu à écrire) évite les écritures superflues qui invalideraient le cache Processing de QGIS.

---

## 3. Programme 1/2 : Évolution des volumes prélevés par ouvrage

### 3.1 Structure de l'algorithme Processing

Le script `compute_slopes_qgis_ouvrages.py` implémente la classe `ComputeSlopesByOuvrage`, héritant de `QgsProcessingAlgorithm`. Ce choix d'implémentation comme algorithme Processing natif, plutôt qu'un script standalone, présente plusieurs avantages : intégration dans la Toolbox QGIS, compatibilité avec le système de feedback et de progression, et possibilité d'enchaînement avec d'autres algorithmes via PyQGIS.

Le traitement se déroule en plusieurs phases séquentielles : filtrage spatial des données en entrée selon la zone d'étude, agrégation des volumes par ouvrage et par année, calcul des indicateurs de tendance, normalisation et écriture de la couche de sortie.

### 3.2 Choix des méthodes de régression : OLS et Theil-Sen

Pour estimer la tendance temporelle des volumes prélevés, VOCAL propose deux méthodes de régression au choix de l'utilisateur :

- **OLS (Ordinary Least Squares)** : régression linéaire classique, implémentée via `numpy.polyfit()` si NumPy est disponible, ou via une implémentation manuelle de la formule des moindres carrés dans le cas contraire. Cette méthode est plus sensible aux valeurs aberrantes (une année anormalement sèche ou humide peut significativement influencer la pente estimée).
- **Theil-Sen** : estimateur de la pente médiane des droites passant par toutes les paires de points. Plus robuste aux valeurs aberrantes. Implémenté en priorité via `scipy.stats.theilslopes()` si SciPy est disponible, avec un fallback vers une implémentation manuelle O(n²) de la médiane des pentes pairwise.

Cette approche à double fallback (bibliothèque tierce → implémentation native) est un choix délibéré pour garantir le fonctionnement du plugin même sur des postes QGIS sans NumPy ou SciPy installés, situation fréquente en contexte institutionnel.

### 3.3 Indicateurs produits : définitions et formules

Le programme calcule cinq indicateurs par ouvrage :

| Indicateur | Formule | Interprétation |
|---|---|---|
| `slope_ouvrage` | Pente de la régression (m³/an) | Évolution absolue annuelle |
| `slope_pct_mean` | 100 × slope / mean(volumes) | Évolution relative, comparaison inter-ouvrages |
| `slope_pct_first` | 100 × slope / mean(3 premières années) | Évolution par rapport à la situation initiale |
| `cagr_pct` | ((moy_3dernieres / moy_3premieres)^(1/n) - 1) × 100 | Taux de croissance début→fin, lisse les fluctuations |
| `slope_pct_z` | (slope_pct_mean - µ) / σ sur l'ensemble de la zone | Identification des ouvrages atypiques dans la zone |

Le z-score (`slope_pct_z`) est calculé sur l'ensemble des ouvrages de la zone d'étude. Il permet d'identifier des situations atypiques sans qu'il soit nécessaire de connaître à l'avance un seuil absolu. Un ouvrage avec un z-score supérieur à 2 ou inférieur à -2 peut être considéré comme statistiquement remarquable dans le contexte de la zone.

### 3.4 Filtrage par nombre minimum d'années

Le paramètre `MIN_YEARS` définit le nombre minimum d'années de données nécessaires pour qu'une pente soit calculée sur un ouvrage. Ce seuil est délibérément exposé à l'utilisateur plutôt que fixé en dur, car il dépend de l'usage : un minimum de 4 ans est conseillé pour une analyse exploratoire, mais un bilan PGRE sur une période de 6 ans justifie d'exiger 6 années complètes pour éviter de comparer des ouvrages sur des fenêtres temporelles incomparables.

Les ouvrages n'atteignant pas ce seuil sont tout de même présents dans la couche de sortie, avec leurs indicateurs à `NULL`, ce qui permet à l'utilisateur de les distinguer des ouvrages sans données plutôt que de les exclure silencieusement.

---

## 4. Programme 3/4 : Ratio volumes prélevés / volumes autorisés

### 4.1 Problématique de l'appariement multi-sources

Le programme de calcul du ratio VP/VA repose sur la jointure de deux tables distinctes : les volumes prélevés issus des campagnes de redevance de l'Agence de l'eau et les volumes autorisés issus des arrêtés de la DDTM. Ces deux tables utilisent des systèmes d'identifiants qui ne sont pas toujours cohérents, ce qui rend l'appariement partiel dans de nombreux cas.

La stratégie adoptée est la suivante : si un ouvrage de la table Agence ne trouve pas de correspondance dans la table DDTM, il est marqué avec la note "unmatched" mais peut tout de même être inclus dans la sortie (via le paramètre `INCLUDE_UNMATCHED`). Cela permet à l'analyste de visualiser l'ensemble des prélèvements de la zone, y compris ceux pour lesquels aucun arrêté n'a pu être apparié.

### 4.2 Logique de calcul du ratio et cas limites

Pour chaque ouvrage apparié, le ratio est défini comme VP / VA. Plusieurs cas limites sont traités explicitement :

- **Volume autorisé nul (VA = 0)** : le ratio n'est pas calculé (`ratio_possible = 0`). Ce cas se produit lorsque l'arrêté existant ne renseigne pas de volume, ou en cas d'erreur de saisie. Les ouvrages concernés sont comptabilisés séparément dans les logs du programme.
- **Volume autorisé manquant (pas d'arrêté)** : ouvrages non appariés, traités selon le paramètre `INCLUDE_UNMATCHED`.
- **Plusieurs lignes DDTM pour le même identifiant ouvrage** : le programme prend le MAX des volumes autorisés et concatène les identifiants DDTM distincts dans le champ `ddtm_id`. Ce choix du MAX est conservateur, il minimise les faux positifs de dépassement dans les cas où la répartition des autorisations entre forages d'un même champ captant ne reflète pas les usages réels.

### 4.3 Filtrage spatial et agrégation

Le filtrage spatial sur la zone d'étude est appliqué à la couche de prélèvements via un index spatial `QgsSpatialIndex`, selon le même principe que dans l'orchestrateur. Les volumes sont ensuite agrégés par identifiant ouvrage pour l'année sélectionnée : si plusieurs lignes de redevance correspondent au même ouvrage pour la même année, leurs volumes sont sommés.

L'année de référence peut être paramétrée à `0`, auquel cas le programme détermine automatiquement la dernière année disponible dans les données filtrées. Ce comportement par défaut facilite l'usage courant où l'on souhaite travailler sur l'année de redevance la plus récente.

---

## 5. Robustesse et gestion des formats de données

### 5.1 Parsing des nombres : le problème des formats mixtes

Les données de redevance de l'Agence et les arrêtés DDTM peuvent provenir de sources variées avec des formats numériques hétérogènes. La fonction `parse_number()`, commune aux deux programmes principaux, gère les cas suivants :

- Format français avec espace comme séparateur de milliers et virgule décimale : `"12 000,56"`
- Format français sans espace : `"12000,56"`
- Format mixte avec point-virgule : `"12.000,56"` (point = milliers, virgule = décimale)
- Format anglais standard : `"12000.56"`
- Présence d'unités textuelles : `"12000 m3"` (les caractères non numériques sont supprimés)
- Valeurs vides, `NULL` ou `None` : retourne `float('nan')` pour un traitement uniforme

Le fait que cette fonction soit dupliquée dans chaque script Processing plutôt que centralisée dans un module commun est un choix pragmatique lié aux contraintes de déploiement des scripts Processing de QGIS : chaque script doit être autonome et ne peut pas importer de modules locaux personnalisés sans configuration supplémentaire.

### 5.2 Gestion des dépendances optionnelles

VOCAL adopte une stratégie de dégradation gracieuse vis-à-vis des bibliothèques scientifiques Python. Au démarrage de chaque script, NumPy, Pandas et SciPy sont importés dans des blocs `try/except` séparés, et leur disponibilité est tracée dans des variables booléennes (`use_numpy`, `use_pandas`, `use_scipy`).

Les fonctions de calcul vérifient ces booléens avant d'utiliser les bibliothèques optimisées, et tombent sur des implémentations Python pures en cas d'absence. Cette approche garantit que le plugin fonctionne sur n'importe quelle installation QGIS, au prix d'une performance légèrement réduite sur les très grandes tables de données.

### 5.3 Gestion des identifiants : forcer le type texte

Un problème récurrent lors du chargement de fichiers CSV dans QGIS est la conversion automatique des identifiants d'ouvrages en types numériques. Un identifiant comme `"0123456"` devient `123456` après conversion, ce qui rompt les jointures avec des tables où le même identifiant est stocké sous forme de chaîne.

La documentation utilisateur de VOCAL insiste sur la vérification du type du champ identifiant au moment du chargement CSV (onglet "Définition des champs" du gestionnaire de données QGIS). Côté script, le programme normalise systématiquement les identifiants en chaînes de caractères (`str(key)`) avant les comparaisons et les jointures, ce qui évite les échecs silencieux dus à des incohérences de type.

---

## 6. Application des styles QML et retour visuel

Chaque programme de VOCAL est associé à un fichier de style QML qui définit la représentation cartographique de la couche de sortie (symbologie graduée sur la pente, les ratios, etc.). L'application du QML est optionnelle et contrôlée par un paramètre booléen `APPLY_QML`.

L'application du style est réalisée en fin de traitement, après récupération de la couche de sortie via `QgsProcessingUtils.mapLayerFromString()`. Cette récupération est nécessaire car la couche de sortie d'un algorithme Processing n'est pas directement accessible comme objet Python pendant l'exécution — elle est d'abord enregistrée dans le contexte Processing, puis accessible via son identifiant `dest_id`.

La méthode `loadNamedStyle()` peut retourner soit un booléen, soit un tuple `(bool, message)` selon la version de QGIS. Le code gère les deux cas via un bloc `try/except` sur `TypeError` pour assurer la compatibilité entre les versions 3.x de QGIS.

---

## 7. Perspectives d'évolution

Plusieurs axes d'amélioration ont été identifiés lors du développement et des premiers retours utilisateurs :

- **Centralisation des fonctions utilitaires** : la duplication de `parse_number()` et d'autres helpers dans chaque script Processing est une dette technique. Une solution propre serait de distribuer un module `vocal_utils.py` dans le dossier des scripts Processing lors de l'installation.
- **Gestion des canaux** : les canaux d'irrigation font l'objet d'une comptabilisation différente entre les DDTM (débit ou volume non restitué) et l'Agence (volume total prélevé). Ce point mériterait un traitement spécifique dans le programme de ratio VP/VA.
- **Programme d'état de connaissance** : le cinquième programme (état de connaissance des ouvrages Agence) est à usage interne Agence. Son architecture suit le même pattern que les autres programmes mais ne produit pas d'indicateurs de prélèvement — il s'agit d'un diagnostic de qualité des données.
- **Export des résultats** : la documentation utilisateur décrit l'export vers Excel via le menu QGIS, mais une fonctionnalité d'export intégrée au plugin faciliterait la production de bilans standardisés.

---

## Conclusion

VOCAL est conçu comme un outil opérationnel destiné à des agents sans expertise poussée en SIG ou en programmation. Les choix architecturaux — séparation orchestrateur/scripts, versions allégées des dépendances, mémorisation de la zone d'étude, injection dynamique des chemins — visent tous à réduire les frictions à l'usage et à garantir une utilisation stable dans un contexte institutionnel où les configurations de QGIS sont très diverses.

La transparence des indicateurs produits est un principe directeur : chaque métrique est documentée dans le code source et dans la formation utilisateur, avec ses limites et ses cas d'usage recommandés. Les résultats de VOCAL sont des portes d'entrée pour l'analyse, pas des conclusions — l'interprétation reste le travail de l'analyste.
