---
title: Complément braille - Documentation utilisateur
layout: my-default
---

## Table des matières

- [Version du document](#version-du-document)
- [Présentation du complément](#présentation-du-complément)
  - [Dictionnaire de travail](#dictionnaire-de-travail)
  - [Base de données statistiques](#base-de-données-statistiques)
- [Installation](#installation)
  - [(Optionnel) Installer le certificat du PAON](#optionnel-installer-le-certificat-du-paon)
  - [Installer le complément](#installer-le-complément)
- [Traitement pour le braille](#traitement-pour-le-braille)
- [Gestion du dictionnaire de travail](#gestion-du-dictionnaire-de-travail)
- [Consultation des données statistiques](#consultation-des-données-statistiques)


# Version du document

| **Version** | **Date**   | **Changement**     | **Auteurs**   |
|:-----------:|------------|--------------------|---------------|
|     1.0     | 2024/05/03 | Rédaction initiale | Nicolas Pavie |
|     1.2     | 2024/09/30 | Mise à jour        | Nicolas Pavie |
|    1.3.5    | 2024/11/26 | Mise à jour        | Nicolas Pavie |


# Présentation du complément

Le complément Word « Braille » du Pôle d’Adaptation des Ouvrages
Numériques fournis aux transcripteurs un pré-traitement d’analyse du
document, et un traitement guidé semi-interactif pour la protection de
mots abrégeables.

Ce complément installe un nouvel onglet « Braille » dans le ruban de
Microsoft Office Word.

<img src="media/image1.png" style="width:6.61695in;height:0.94347in"
alt="Une image contenant texte, Police, ligne, nombre Description générée automatiquement" />

Cet onglet est composé de différents groupes de boutons :

- Un premier groupe « **Actions **» contient le bouton « **Lancer le
  traitement** » qui permet le lancement de l’analyse du document.

- Un deuxième groupe « **Statut** » informe sur le dernier mot
  sélectionner et en attente d’action, ainsi que sur l’avancement du
  traitement de protection du document.

  - Un bouton « **Resélectionner** » permet de remettre en avant dans
    Word le dernier mot sélectionné.

- Un troisième groupe « **Navigation** » permet de naviguer entre les
  mots précédemment identifiés dans la phase d’analyse.

- Un groupe « **Données** » permet de consulter les données du
  complément :

  - Le bouton « **Dictionnaire de travail **» permet de consulter la
    liste de tous les mots et actions choisis pour ces mots dans le
    document

  - Le bouton « **Base de protection** » permet de consulter la base de
    données des statistiques de protection utiliser pour la
    pré-protection, et de rechercher dans cette base les informations de
    statistiques de protection d’un mot.

- Un groupe « **Sélection** » permet de traiter manuelle un élément dans
  le texte :

  - Le bouton « **Protéger la sélection** » insère un code de protection
    devant le mot sélectionné.

  - Le bouton « **Abréger la sélection** » supprime (s’il est présent)
    un code de protection situé devant un mot.

## Dictionnaire de travail

Le complément créé et gère un fichier dit de « dictionnaire de travail »
enregistrer au même emplacement qu’un fichier analyser et portant le
même nom que celui-ci, mais avec une extension de fichier en
« **.bdic** ».

Ce dictionnaire de travail est un fichier texte (encodé au format
Latin-1 ou Windows-cp1252) contenant l’intégralité des décisions prises
par le programme ou par le transcripteur, et ceux pour chaque occurrence
de mots abrégeables identifiés lors de l’analyse du document. Chaque
décision est conservée sur une ligne de ce fichier, sous la forme d’un
code d’action entre 0 et 3, suivi du mot auquel appliqué l’action tel
que rencontrer dans le document.

Ce dictionnaire est utilisé par le complément pour conserver les actions
choisies lors du traitement, et est relu au lancement de l’analyse si
celui-ci est présent.

**Attention, ce dictionnaire n’est pas transférable d’un document a un
autres. Des travaux sont en cours pour permettre la réutilisation des
décisions d’un manuscrit sur un autre.**

## Base de données statistiques

Le complément va de paire avec une base de données de statistiques de
protection. Cette base de données, disponible en ligne à l’adresse  
<https://gitlab.com/studio-paon-avh/protection_db>  
regroupe les informations de tous les dictionnaires de travail produit
par le PAON pour en déduire une liste de mots associé aux informations
suivantes :

- Le nombre de document dans lesquels le mot a nécessité une action du
  transcripteur

- Le nombre de décision de protection dans tout le document pour le mot

- Le nombre de décision d’abréviation dans tout le document pour le mot

- Un indicateur de demande systématique d’une analyse par le
  transcripteur

  - Par exemple pour les mots existant dans différentes langues

- Un éventuel commentaire sur le mot

  - Par exemple pour indiquer les différentes langues dans lesquels le
    mot peut être rencontré

Cette base de données est mise à jour à l’aide des nouveaux
dictionnaires de travail produit par le complément braille. La
fonctionnalité est en cours de réalisation, mais dans une prochaine
version du complément, lorsqu’une mise à jour de la base de données sera
disponible, le complément récupèrera ces mises à jour pour les utiliser.

# Installation

## (Optionnel) Installer le certificat du PAON

Le complément braille est réaliser par l’équipe du Pole d’Adaptation des
Ouvrages Numériques. L’équipe de développement utilise un certificat
interne pour signer ces développements, mais ce certificat n’est pas
issus d’une autorité de certification officielle (ce type de certificat
coute chère). Pour éviter une alerte de sécurité lors de l’installation
du complément ou de ses mises à jours, vous pouvez installer le
certificat « certificat-devs-paons.crt » fournis avec le logiciel.

## Installer le complément

Dans le répertoire de déploiement du complément, double cliquer le
fichier « setup.exe » pour lancer l’installation du complément.
L’installateur n’étant pas signé avec un certificat de haut niveau, un
avertissement de sécurité vous sera présenté.

Accepter l’installation pour que le complément Word s’installe.

Une fois l’installation valider, vous pouvez ouvrir Word pour constater
l’apparition du nouvel onglet « Braille ».

# Traitement pour le braille

Pour commencer la préparation d’un fichier pour le braille :

1)  Ouvrez le document à préparer avec Word, puis sélectionner l’onglet
    « **Braille** » dans le ruban de Word.

> <img src="media/image2.png" style="width:1.51063in;height:0.78136in"
> alt="Une image contenant texte, Police, capture d’écran, blanc Description générée automatiquement" />

2)  Sous le groupe « **Actions** », cliquez ensuite sur le bouton « **Lancer
    le traitement** »

<img src="media/image3.png" style="width:2.66704in;height:1.75024in"
alt="Une image contenant texte, capture d’écran, Police, nombre Description générée automatiquement" />

Au lancement du traitement sur un document, le complément ouvrira un
journal de traitement et procédera à un prétraitement d’analyse du
document pour en extraire les mots nécessitant un choix de protection ou
d’abréviation. Pour information, ces mots sont

- Tous les mots abrégeables contenant au moins une majuscule

- Tous les mots issus de la base de données de protection et indiqués
  comme étant à évaluer systématiquement.

- Tous les mots étrangers détectés par Word

  - **Pour activer la détection des mots étranger par le complément, il
    est nécessaire d’installer une ou plusieurs langues de création
    supplémentaires dans les options de Langue de Word.**

A noter que l’analyse du document, et en particulier la détection des
mots étranger, peut prendre plusieurs minutes dans le cas document long
(plus de cent pages).

A la première utilisation du complément ou si une mise à jour de la base
de données de protection est disponible, le complément installera la
base de données des statistiques de protection. Le complément
récupèrera, à partir de cette base de données, les statistiques
disponibles pour les mots identifiés à l’étape précédente, et procédera
à un prémarquage des mots dans le dictionnaire de travail en suivant les
règles suivantes :

- Si le mot est indiqué comme requérant systématiquement un avis du
  transcripteur, le mot est marqué comme « Ambigu » dans le dictionnaire
  de travail

- Si le mot a été répertorié dans plus de 100 documents et que plus de
  99% de ses occurrences ont eu le même statut (c.a.d. plus de 99%
  d’abrégé ou plus de 99% de protégé), ce statut est repris dans le
  dictionnaire de travail

Une fois cette analyse terminée, le traitement ouvre la fenêtre d’action
par mot :

<img src="media/image4.png" style="width:6.3in;height:4.81042in"
alt="Une image contenant texte, capture d’écran, logiciel, nombre Description générée automatiquement" />

Cette fenêtre permet de parcourir et de choisir s pour chaque mot et
pour chaque occurrence de celui-ci.

3)  Fermez l’éditeur de dictionnaire pour que commence le traitement de
    protection interactive des différentes occurrences des mots
    identifiés au sein du document.

Pour chaque occurrence identifiée, le complément va appliquer les règles
suivantes :

- Si le mot est indiqué comme « Protéger » dans le dictionnaire, le
  traitement va ajouter un code de protection duxburry indiquant que le
  mot doit rester complet en braille abrégé ;

- Si le mot est indiqué comme « Abréger » dans le dictionnaire, le
  traitement va laisser le mot inchangé ou supprimer le code de
  protection s’il était précédemment défini ;

- Si le mot est indiqué comme « Inconnu », une boite de dialogue
  d’action pour le mot sélectionné s’ouvre pour sélectionner le statut
  du mot.

> La boite de dialogue d’action reprend les informations statistiques
> disponibles pour le mot dans la base de données statistiques, et vous
> fournis 6 boutons d’actions utilisables sur le mot sélectionné :

- Un bouton « **Protéger ici** » (activable également avec la touche
  « **i** » de votre clavier) rajoute un code de protection sur le mot.

- Un bouton « **Resélectionner** » (activable également avec la touche
  « **f** » de votre clavier) permet de resélectionner le mot dans Word

- Un bouton « **Abréger ici** » (activable également avec la touche
  « **a** » de votre clavier) enlève un code de protection si celui-ci
  est présent devant le mot, et ignore le mot si ce n’est pas le cas.

- Un bouton « **Protéger partout** » (activable également avec la touche
  « **p** » de votre clavier) va ajouter un code de protection sur le
  mot sélectionné, et va indiquer au traitement que toutes les
  occurrences suivantes peuvent être automatiquement protégé. Le mot
  sera également indiqué en statut « Protéger » dans le dictionnaire de
  travail.

- Un bouton « **Ignorer** » (activable également avec la touche
  « **\$** » de votre clavier) ignore le mot temporairement : une action
  sera redemandée pour ce mot une fois la fin du document atteinte.

- Un bouton « **Abréger partout** » (activable également avec la touche
  « **b** » de votre clavier) va ignorer le mot ou supprimer le code de
  protection sur le mot sélectionné, et va indiquer au traitement que
  toutes les occurrences suivantes peuvent également être ignorées ou
  déprotégées. Le mot sera également indiqué en statut « Abréger » dans
  le dictionnaire de travail.

En plus de ces 6 actions, vous pouvez également interrompre le
traitement en cours en fermant simplement la boite de dialogue, et
reprendre le traitement en cliquant à nouveau sur le bouton « **Lancer
le traitement** ».

# Gestion du dictionnaire de travail

A rédiger

# Consultation des données statistiques

A rédiger
