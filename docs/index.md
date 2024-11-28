---
title: Complément Word pour le braille abrégé
layout: my-default
---


Le complément Word "Braille" est un projet réalisé au sein du Pôle Adaptation des Ouvrages Numériques (PAON) de l'[Association Valentin Haüy](https://avh.asso.fr). Ce complément vise à aider la préparation de documents Word en français pour leur transcription en braille abrégé dans l'outil DBT de Duxbury.

- [Télécharger la dernière version](https://github.com/studio-PAON-AVH/complement-word-braille/releases/download/v1.3.5.3/ComplementBrailleWord-1.3.5.3.zip)
- Ou accèder à la [page de publication](https://github.com/studio-PAON-AVH/complement-word-braille/releases/latest)

# Installer le complément

- Dezipper l'archive zip du complément
- (Optionnel) installer le certificat `certificat-devs-paons.crt`
  - Ce certificat identifie les développements réalisé par le PAON.
- Lancer le fichier `setup.exe`
- Accepter l'installation du VSTO
- Fermer et rouvrer Microsoft Word, et vérifier qu'un onglet **Braille** apparait dans le ruban de l'application.

# Traiter un document pour le braille

Le complément ajoute a Microsoft Word un nouvel onglet **Braille** dans le ruban.

Lorsque l'utilisateur demande à lancer le traitement via ce nouvel onglet, le complément analyse le document pour en extraire une liste des mots abrégeables pouvant nécessiter d'être "protégés" (c'est-à-dire conservés dans leur forme intégrale) dans certains contextes. L'utilisateur a alors accès à une boîte de dialogue spécifique lui permettant de protéger, d'abréger ou d'ignorer tout ou partie des occurrences du mot identifiées dans le document. L'action de protection d'une occurrence d'un mot se traduit dans le document par l'insertion du code DBT de conservation en forme intégrale `[[*i*]]` devant cette occurrence, tandis que l'action d'abréviation supprime ce code s'il existait précédemment.

Actuellement, le complément identifie les mots à traiter selon trois critères :
- Si le mot contient au moins une majuscule
- Ou si le mot fait partie d'une liste de mots "ambigus" définie par les transcripteurs du PAON comme étant à passer en revue systématiquement,
- Ou si le mot est signalé comme étant dans une langue étrangère par Microsoft Word.

Ce complément utilise également une base de données statistiques, issue de précédents traitements de documents pour le braille faits par le PAON. Lors du traitement des mots, ces statistiques sont utilisées pour présélectionner une action recommandée sur le mot en cours de traitement.

# Version 1.3.5.3 (Novembre 2024)

Cette version est la première à être mise à disposition du grand public.
