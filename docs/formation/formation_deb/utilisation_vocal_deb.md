---
layout: default
title: "Utiliser VOCAL"
parent: "Formation : je n'ai jamais lancé QGIS"
nav_order: 3
---

# 3. Utiliser VOCAL

Cette section décrit comment utiliser les différents outils de VOCAL.

---

## 3.1 Choisir un programme

Le plugin VOCAL propose pour le moment **5 programmes de calcul** :

- l’évolution des volumes prélevés (par ouvrage),
- l’évolution des volumes prélevés (par zone),
- le ratio entre le volume prélevé et le volume autorisé (par ouvrage),
- le ratio entre le volume prélevé et le volume autorisé (par zone),
- une caractérisation de la connaissance des ouvrages de l’Agence.

Une fois VOCAL lancé, sur le premier menu qui s’affiche, la question du **choix du programme** est la première qui nous est posée.  
Choisissez alors le programme souhaité dans la liste déroulante.

![Choix du programme dans VOCAL]({{ "/assets/images/formation/screen_utilisation_choix_programme.png" | relative_url }})

*Figure X — Choix du programme dans VOCAL*

-> Pour cette formation, il est conseillé de choisir le **premier programme proposé** :  
**Évolution des volumes prélevés par ouvrage**.

---

## 3.2 Choisir une zone d’étude

Choisir une zone d’étude est nécessaire au bon fonctionnement de VOCAL.  
On distingue cependant **deux types de zonages** :

1. **La zone d’étude** : zone totale sur laquelle porte l’analyse.
2. **Les zones de travail** : sous-zonage utilisé pour les programmes nécessitant une agrégation spatiale.

Ces deux types de zonages se chargent différemment.

---

### 3.2.1 Charger une zone d’étude

Pour charger une zone d’étude, la méthode est similaire au choix du programme.  
Deux listes déroulantes sont disponibles :

- une pour choisir **l’échelle** de la zone,
- une seconde pour sélectionner **l’élément correspondant**.

**Exemple** :  
- Échelle : *Départements*  
- Élément : *GARD*

Une fois ces choix effectués, cliquer sur **Charger la zone et zoomer**.

![Chargement de la zone d’étude dans VOCAL]({{ "/assets/images/formation/screen_utilisation_choix_zone.png" | relative_url }})

*Figure X — Chargement de la zone d’étude dans VOCAL*

---

### 3.2.2 Charger des sous-zonages

Pour charger un sous-zonage, il faut activer l’option correspondante en cochant la case :

> **« Voulez-vous charger un sous-zonage ? »**

comme indiqué sur la capture ci-dessous :

![Chargement d’un sous-zonage dans VOCAL]({{ "/assets/images/formation/screen_utilisation_choix_sous_zonage.png" | relative_url }})

*Figure X — Chargement d’un sous-zonage dans VOCAL*

Une fois la case cochée, un menu supplémentaire apparaît :

- le premier menu déroulant permet de choisir un **niveau d’échelle** (communes, bassins versants, etc.) ;
- tous les éléments de ce niveau intersectant la zone d’étude seront chargés ;
- il est également possible de sélectionner :
  - une couche polygone déjà présente dans le projet QGIS,
  - ou un fichier externe via l’explorateur de fichiers.

Une fois le sous-zonage choisi, cliquer sur **Valider** (en bas à droite du menu).

---

**Exemple** :  
Si l’on souhaite charger le **département de l’Hérault** comme zone d’étude et ses **bassins versants** comme sous-zonage, l’interface doit ressembler à ceci :

![Exemple de chargement de zone dans VOCAL]({{ "/assets/images/formation/screen_utilisation_exemple_choix_zone.png" | relative_url }})

*Figure X — Exemple de chargement de zone dans VOCAL*

Puis cliquer sur **Valider**.

---

## 3.3 Compléter les champs du programme 1 :  
### Évolution des volumes prélevés par ouvrage

Une fois les zones choisies :

1. Lancer VOCAL.
2. Sélectionner le programme **Évolution des volumes prélevés par ouvrage**.
3. Charger une zone d’étude.
4. Cliquer sur **Valider**.

Le programme se lance alors.  
Si ce n’est pas le cas, revenir à l’étape **1. Installation de VOCAL** afin de vérifier que l’installation est complète.

Un nouveau menu Processing s’ouvre :

![Lancement du premier programme VOCAL]({{ "/assets/images/formation/screen_utilisation_premier_programme.png" | relative_url }})

*Figure X — Lancement du premier programme VOCAL*

---

### Remplissage des champs

Compléter les champs étape par étape :

**Indication de la zone d’étude**
1. Sélectionner, parmi les couches du projet, la couche correspondant à la zone d’étude  
   *(exemple : `departements_INTER_HERAULT [EPSG:2154]`)*.

**Indication de la base de données et des champs**
2. Sélectionner la base de données de l’Agence.
3. Sélectionner le champ correspondant à l’**année de campagne de redevance**.
4. Sélectionner le champ correspondant à l’**identifiant de l’ouvrage**.
5. Sélectionner le champ correspondant au **nom de l’ouvrage**. (optionnel)
6. Sélectionner le champ correspondant au **nom de l’interlocuteur**. (optionnel)
7. Sélectionner le champ correspondant au **volume retenu** (l'assiette).
8. Choisir la méthode de calcul de l’évolution  
   *(Theil-Sen est généralement plus robuste)*.
9. Indiquer le **nombre minimum d’années** nécessaires au calcul de la pente  
   *(minimum conseillé : 4, voire 5 ou 6 pour une étude de type PGRE)*.
10. Indiquer les **années de début et de fin** de l’étude temporelle.

Les autres options sont plus techniques et seront abordées dans la section **3.6** de ce guide.

Le menu complété doit ressembler à ceci :

![Exemple remplissage premier programme VOCAL]({{ "/assets/images/formation/screen_utilisation_premier_programme_rempli.png" | relative_url }})

*Figure X — Exemple de remplissage du premier programme VOCAL*

Il est alors possible de cliquer sur **Exécuter**.

---

## 3.4 Exercice  
### Compléter les champs du programme 2 : Évolution des volumes prélevés agrégés par zone

Ici, vous allez être en autonomie pour essayer de charger le deuxième programme de Vocal : l'**Évolution des volumes prélevés agrégés par zone**. 
La démarche à suivre va être exactement la même que pour le programme 1, avec simplement l'ajout d'une couche de sous-zonage. 
Bonne chance !  
L'objectif est d'avoir quelque chose du type : 
![Exercice]({{ "/assets/images/formation/screen_utilisation_premier_exercice_solution_trois.png" | relative_url }})
*Figure X — Objectif de l'exercice*

<details> 
<summary>Indice</summary>
Vous devez charger un sous-zonage dans le premier menu de VOCAL ! 

</details>


<details>
<summary>Solution</summary>

<p>Voici les champs à remplir pour obtenir une carte de l'évolution des volumes prélevés agrégés par zone :</p>

<img src="{{ '/assets/images/formation/screen_utilisation_premier_exercice_solution_un.png' | relative_url }}"
     alt="Solution premier menu"
     width="700">

<p><em>Figure X — Solution premier menu</em></p>

<img src="{{ '/assets/images/formation/screen_utilisation_premier_exercice_solution_deux.png' | relative_url }}"
     alt="Solution deuxième menu"
     width="700">

<p><em>Figure X — Solution deuxième menu</em></p>

<img src="{{ '/assets/images/formation/screen_utilisation_premier_exercice_solution_trois.png' | relative_url }}"
     alt="Solution résultats"
     width="700">

<p><em>Figure X — Solution : résultats</em></p>

</details>


---

## 3.5 Compléter les champs du programme 3  
### Ratio Volumes prélevés / Volumes autorisés par ouvrage
Même démarche que pour charger le premier programme,on va cependant rajouter une deuxième base, une base DDTM pour ajouter des volumes autorisés. On suit les mêmes premières étapes après avoir choisi une zone : 

1. Lancer VOCAL.
2. Sélectionner le programme **Ratio Volumes prélevés / Volumes autorisés par ouvrage**.
3. Charger une zone d’étude.
4. Cliquer sur **Valider**.

Le programme se lance. 
Compléter les champs étape par étape :

**Indication de la zone d’étude**
1. Sélectionner, parmi les couches du projet, la couche correspondant à la zone d’étude  
   *(exemple : `departements_INTER_HERAULT [EPSG:2154]`)*.

**Indication de la base de données et des champs côté Agence**
2. Sélectionner la base de données de l’Agence.
3. Sélectionner le champ correspondant à l’**année de campagne de redevance**.
4. Sélectionner le champ correspondant à l’**identifiant de l’ouvrage**.
5. Sélectionner le champ correspondant au **volume retenu** (l'assiette).
6. Sélectionner le champ correspondant au **Type de milieu**. (optionnel)
7. Sélectionner le champ correspondant au **nom de l’ouvrage**.(optionnel)
8. Sélectionner le champ correspondant au **nom de l’interlocuteur**.(optionnel)

**Indication de la base de données et des champs côté DDTM**
9. Sélectionner la base de données de la DDTM.
10. Sélectionner le champ correspondant à l’**ID de l'ouvrage Agence**.
11. Sélectionner le champ correspondant au **volume autorisé par la DDTM**.
12. Sélectionner le champ correspondant à l'**ID de l'ouvrage de la DDTM**.

**Indication de l'année de référence**
12. Indiquer l'année sur laquelle vous voulez que les assiettes agence soient prise en compte.

Les autres options sont plus techniques et seront abordées dans la section **3.6** de ce guide.

Le menu complété doit ressembler à ceci :

![Exemple remplissage premier programme VOCAL]({{ "/assets/images/formation/screen_utilisation_troisieme_programme_rempli.png" | relative_url }})

*Figure X — Exemple de remplissage du troisième programme VOCAL*

Il est alors possible de cliquer sur **Exécuter**.

---

---

## 3.6 Fonctionnalités optionnelles
