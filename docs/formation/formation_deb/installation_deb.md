---
layout: default
title: "Installer VOCAL"
parent: "Formation : je n'ai jamais lancé QGIS"
nav_order: 1
---


# 1. Installer VOCAL

Cette section décrit comment télécharger et installer le plugin **VOCAL** sur votre ordinateur.

---

## 1.1 Télécharger VOCAL

Le plugin **VOCAL** est disponible en libre accès *open source* sur la plateforme GitHub, à l’adresse suivante : [https://github.com/AurelL-AERMC/Le-VOCAL](https://github.com/AurelL-AERMC/Le-VOCAL) et donc accéder au plugin directement depuis ce répertoire GitHub.

Cependant, vous pouvez aussi télécharger directement le plugin via le lien suivant : [Télécharger le plugin VOCAL (ZIP)](https://github.com/AurelL-AERMC/Le-VOCAL/archive/refs/heads/main.zip)


Une fois le téléchargement fini, vous devez désormais être en possession d’un fichier compressé **ZIP** contenant le plugin.  
Placez ce fichier à l’emplacement de votre choix sur votre ordinateur.

---

## 1.2 Installer VOCAL

Pour installer le plugin, commencez par lancer **QGIS**.

1. Une fois QGIS ouvert, cliquer sur le menu :  
   **Extensions → Installer/Gérer les extensions**

![Onglet Installer depuis un ZIP]({{ "/assets/images/formation/screen_installation.png" | relative_url }})
*Figure 1 — Onglet « Installer depuis un ZIP » dans QGIS*


2. Dans la fenêtre qui s’ouvre, cliquer sur l’onglet **Installer depuis un ZIP**.

3. Cliquer sur l’icône **« … »**, puis rechercher dans votre explorateur de fichiers le fichier **ZIP** précédemment téléchargé.  
   Sélectionner le fichier et valider.

4. Cliquer sur **« Accepter »** lorsqu’une fenêtre de confirmation s’affiche.

Un nouvel icône **VOCAL** a désormais été ajouté à l’interface de QGIS.

5. Cliquer sur cet icône.

![Premier lancement de VOCAL pour finaliser l'installation]({{ "/assets/images/formation/screen_installation_premier_lancement.png" | relative_url }})
*Figure 2 — Premier lancement de VOCAL pour finaliser l'installation*
   

6. Sélectionner une **zone d’étude** (peu importe laquelle pour cette étape).  
   Charger la zone d’étude puis cliquer sur **Valider**.

7. Une fenêtre s’ouvre vous demandant de **redémarrer QGIS**.  
   Redémarrer QGIS.

Le plugin **VOCAL** est désormais installé sur votre ordinateur.


Nous pouvons maintenant passer à l’étape suivante : le chargement des données.  
[La suite : Charger des données](chargement_couches_deb)

---

## Si la manipulation n'a pas fonctionné


1. Veuillez ressayer de lancer VOCAL, charger une zone, lancer un programme, redemarrer QGIS. 
2. Vérifier que le Plugin de base de QGIS nommé Processing (icone engrenage gris) soit bien activé. Extention -> Gérer les extensions -> Extensions Installées -> Activer Processing.

---

## Récapitulatif rapide

1. Télécharger le fichier ZIP
2. Ouvrir QGIS
3. Menu **Extensions → Installer/Gérer les extensions**
4. Onglet **Installer depuis un ZIP**
5. Sélectionner le fichier téléchargé
6. Cliquer sur l’icône **VOCAL** et lancer un programme
7. Redémarrer QGIS