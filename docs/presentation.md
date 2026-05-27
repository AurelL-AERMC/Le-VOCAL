---
layout: default
title: Présentation du Projet
parent: Accueil
nav_order: 1
---

# Présentation du Projet et du Plugin VOCAL

**Agence de l'eau Rhône-Méditerranée-Corse · Délégation de Montpellier · Service RAGAF · Février 2026**

> L'Agence de l'eau et les DDT(M) disposent chacune d'informations relatives aux ouvrages de prélèvement d'eau, mais des incohérences et des écarts de connaissance ont été constatés entre leurs bases de données. L'Étude Données Prélèvements vise à fiabiliser, rendre interopérables et valoriser ces bases, dans une logique de mise en commun avec les partenaires.

---

## 1. Origines et objectifs

L'Agence de l'eau perçoit des redevances au titre des prélèvements effectués dans le milieu naturel par les différents usagers (collectivités, industriels, agriculteur·rices, etc.). Elle dispose, à ce titre, d'une connaissance des volumes prélevés annuellement.

De leur côté, les DDT(M), via le guichet unique de l'eau et l'instruction des procédures Loi sur l'Eau (régime IOTA), recensent les ouvrages autorisés et disposent d'informations réglementaires et techniques associées. Elles disposent à ce titre des volumes autorisés et prélevés mensuellement.

Il a cependant été constaté des incohérences et des écarts de complétude, ainsi qu'une absence d'identifiant commun permettant les croisements d'information entre les bases. L'objectif de l'Étude Données Prélèvements est de pallier ces manquements en visant une fiabilisation ainsi qu'une interopérabilité par l'appariement des bases. Le résultat de cette harmonisation des connaissances est voué à être valorisé et partagé avec l'ensemble des partenaires.

Ce projet a été lancé en janvier 2025 par l'Agence et la DDTM 34, partenaire initiale du projet pour un travail sur le département de l'Hérault. Il a ensuite été étendu aux autres départements majeurs de la délégation de Montpellier (Aude, Gard, Pyrénées-Orientales), avec l'appui des DDTM concernées.

---

## 2. Gestion du projet

Afin de définir précisément les besoins, les contours et la pertinence du projet, une phase initiale de co-construction a été menée. Ont été consultés : des EPTB (Fleuve Hérault et Orb-Libron), la DREAL Occitanie, des agents en interne Agence ainsi que des agents des DDTM 34, 11, 66 et 30. Un bilan de ces entretiens a permis d'identifier les besoins, les ressources, les exigences réglementaires et les interactions inter-acteurs de chacune des parties prenantes.

Le projet a été conduit à l'Agence par Jules Barbazanges et Aurel Lashermes, qui en a assuré la gestion, sous la supervision de Stéphanie Weill. Le service SPARC de l'Agence et l'équipe police de l'eau de la DDTM34 ont accompagné les travaux. L'ensemble de ces acteurs constituaient le Comité Technique, réuni à trois reprises en 2025 pour assurer le suivi du projet.

Le projet a suivi une gestion en 6 axes :

1. Identification des besoins, exigences réglementaires et ressources des acteurs du projet
2. État des lieux des bases existantes
3. Conception d'un outil d'appariement
4. Fiabilisation, complétion et interopérabilité
5. Analyse des résultats et des données
6. Valorisation des analyses et de la connaissance produite

Initialement centré sur le département 34, l'état initial de la base DDTM34 ne permettait pas un appariement immédiat pertinent ; une phase préalable de fiabilisation par consultation des archives papier a donc été nécessaire, avec 384 ouvrages balayés manuellement. Par conséquent, les travaux ont rapidement été élargis aux départements 30, 11 et 66, avec une gestion parallèle par territoire.

---

## 3. Situation initiale : cadre réglementaire et état des bases

Les Codes de l'environnement, minier et de la santé publique encadrent la connaissance, par les services de l'État, des ouvrages de prélèvement d'eau dans le milieu naturel. En synthèse, hors Zone de Répartition des Eaux :

| Service | Périmètre de suivi | Pas de temps |
|---|---|---|
| DDT(M) | Ouvrages souterrains > 1 000 m³/an · Ouvrages superficiels > 2 % du débit d'étiage | Mensuel |
| Agence de l'eau | Ouvrages dont les propriétaires prélèvent > 5 000 m³/an | Annuel |
| ARS | Prélèvements destinés à la consommation humaine (AEP) | — |
| DREAL | Forages > 10 m de profondeur | — |
| UD DREAL / DDPP | Prélèvements ICPE et ICPE d'élevage | — |

### État des bases DDTM

L'état des bases est très hétérogène selon les départements. Les bases 30 et 11 sont assez complètes (hors ouvrages AEP pour le 11) et plutôt incomplètes pour les 66 et 34 (en particulier sur les ouvrages AEP pour le 34). Concernant les volumes prélevés mensuels, les données ne sont pas exhaustives sur le 11 et le 30, limitées aux périodes de sécheresse pour le 34, et inexistantes sur le 66.

Dans les Pyrénées-Orientales, seuls les ouvrages en eaux superficielles sont intégrés au projet, un travail antérieur ayant déjà été réalisé sur les nappes du Roussillon.

### État de la base Agence

La qualité de l'information est globalement bonne côté Agence. Cependant, environ 20 % des ouvrages connus (dont environ 80 % dans les PO) ne sont pas localisés à une meilleure échelle que la commune d'implantation. De plus, environ un tiers des ouvrages connus par la délégation ne sont pas ou plus liés à un interlocuteur ou propriétaire identifié.

---

## 4. Appariement : méthode et algorithme

Une méthode réplicable d'appariement des bases de données a été développée spécifiquement pour ce projet. L'objectif est d'associer chaque ouvrage DDTM à son équivalent Agence via un identifiant commun. Un seuil de 4 000 m³/an autorisés par ouvrage a été fixé pour contenir le volume de travail.

Les bases ont fait l'objet d'un pré-traitement : nettoyage, normalisation, homogénéisation des formats. L'outil attribue ensuite un score de probabilité à chaque couple potentiel d'ouvrages selon la localisation (coordonnées GPS ou commune), le nom de l'ouvrage et le nom du propriétaire.

Trois issues sont possibles à l'issue du calcul de score :

- **Appariement automatique** : un couple unique avec un score élevé se dégage — l'appariement est validé automatiquement.
- **Ouvrages distincts** : score très faible — les ouvrages sont considérés comme distincts et sans lien.
- **Vérification manuelle** : doute ou compétition entre couples — une vérification manuelle est demandée.

Les ouvrages sans « partenaire » identifié ont également fait l'objet d'un traitement manuel et d'une typologie selon l'état de connaissance du propriétaire : interlocuteur inconnu côté DDTM seulement, côté Agence seulement, ou situation divergente entre les deux. Plus de 6 700 ouvrages DDTM ont été balayés à l'aide de cette méthode.

---

## 5. Fiabilisation et complétion des bases de données

### Campagne d'interrogations

Une campagne d'interrogation a été menée pour régulariser les situations divergentes ou inconnues côté Agence identifiées lors de l'appariement. En suivant les procédures de recherche redevable, des courriers ont été envoyés aux interlocuteurs inconnus : 109 courriers sur le 34 et 152 sur le 30. Des mails ont été envoyés aux interlocuteurs dont la situation était divergente entre la base DDTM et la base Agence : 67 mails sur le 34 et 73 sur le 30.

Il en résulte un taux de réponse d'environ 70 %, plus de 40 nouveaux redevables identifiés et 248 nouveaux ouvrages créés. 71 courriers de mise en demeure ont également été envoyés aux personnes ne répondant pas, sans suite engagée à ce stade. Les campagnes restent à conduire sur les départements 11 et 66.

### Appariement ARS

De nombreux ouvrages AEP manquaient dans les bases DDTM 34 et 11. Cette situation est issue d'un rattrapage inabouti des autorisations historiquement portées par l'ARS au titre du code de la santé publique. Un travail d'appariement similaire a alors été réalisé entre la base ARS et les bases DDTM sur le 34 et le 11. Grâce à cela, 384 ouvrages ont été ajoutés à la base DDTM 11 et 480 ouvrages à la base DDTM 34 — un amendement particulièrement significatif.

### Bénéfices transversaux

L'appariement permet la confrontation des informations entre bases : ajout de noms et coordonnées de pétitionnaires dans la base DDTM66, correction de localisations d'ouvrages dans la base Agence, raccordement aux masses d'eau sur le 34, actualisation des codes BSS. Une partie du potentiel de valorisation reste encore à exploiter.

---

## 6. Valorisation

### L'extension QGIS VOCAL

Pour faciliter l'exploitation des données consolidées, une extension QGIS spécifique à l'Agence a été développée : **VOCAL** (Valorisateur des Ouvrages Connus À l'Agence). C'est un outil qui permet d'obtenir des indicateurs d'analyse et de la visualisation automatisés. Il gère la comparaison des volumes autorisés (DDTM) et déclarés (Agence), l'analyse des tendances d'évolution des prélèvements et propose une visualisation à l'échelle territoriale ou individuelle.

L'outil a fait l'objet de présentations auprès de tous les partenaires, ainsi que de sessions de formation en interne et en externe auprès des EPTB, DREAL, DDTM30, 34 et 11. L'accueil réservé à VOCAL est très positif.

Documentation complète disponible en ligne : [aurellaermc.github.io/Le-VOCAL/](https://aurellaermc.github.io/Le-VOCAL/)

### Procédure long terme

L'appariement constitue un investissement lourd mais structurant, à ne réaliser qu'une seule fois, en considérant la pérennité du lien créé. Un travail autour de la communication entre DDTM et Agence a abouti à une cohésion améliorée entre services de l'État et à une procédure écrite validée collectivement pour assurer la mise à jour annuelle des liens et des informations nouvelles.

En parallèle, une fiche accompagnatrice des données Agence a été rédigée pour faciliter leur prise en main par les partenaires. Une amélioration de la convention de transmission des données est en attente de validation par les services du Siège.

---

## 7. Conclusion et prochaines étapes

Un effort particulier a été porté sur l'accompagnement des partenaires et la qualité des relations inter-services de l'État. Le projet, initialement perçu comme difficilement réalisable, a su produire des résultats très satisfaisants. Reconnu par les acteurs, il apporte une plus-value réelle dans la connaissance et la gestion de la ressource, même si une partie du potentiel reste encore à exploiter. L'ensemble des besoins exprimés en phase initiale a été couvert, voire dépassé avec la création d'un outil autonome dont la simplicité, la qualité de la documentation et la pertinence ont toutes été saluées par les partenaires.

### Acquis du projet

- Bases DDTM 34 et 30 fiabilisées et appariées à la base Agence
- 864 ouvrages AEP ajoutés aux bases DDTM 34 et 11
- Procédure de mise à jour annuelle validée collectivement
- Extension QGIS VOCAL opérationnelle et documentée
- 40+ nouveaux redevables identifiés

### Ce qui reste à faire

- Campagnes d'interrogation sur les départements 11 et 66
- Exploitation pleine des nouvelles informations et appariements
- Valorisation des coordonnées d'ouvrages encore imprécises
- Validation de la convention de transmission par le Siège
