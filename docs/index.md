# Complément Word pour le traitement du braille

Le complément Word "Braille" est un projet réalisé au sein du Pôle Adaptation des Ouvrages Numériques (PAON) de l'[Association Valentin Haüy](https://avh.asso.fr). Ce complément vise à aider la préparation de documents Word en français pour leur transcription en braille abrégé dans l'outil DBT de Duxbury.

Lorsque l'utilisateur demande à lancer le traitement via le ruban de Word, le complément analyse le document pour en extraire une liste des mots abrégeables pouvant nécessiter d'être "protégés" (c'est-à-dire conservés dans leur forme intégrale) dans certains contextes. L'utilisateur a alors accès à une boîte de dialogue spécifique lui permettant de protéger, d'abréger ou d'ignorer tout ou partie des occurrences du mot identifiées dans le document. L'action de protection d'une occurrence d'un mot se traduit dans le document par l'insertion du code DBT de conservation en forme intégrale `[[*i*]]` devant cette occurrence, tandis que l'action d'abréviation supprime ce code s'il existait précédemment.

Actuellement, le complément identifie les mots à traiter selon trois critères :
- Si le mot contient au moins une majuscule
- Ou si le mot fait partie d'une liste de mots "ambigus" définie par les transcripteurs du PAON comme étant à passer en revue systématiquement,
- Ou si le mot est signalé comme étant dans une langue étrangère par Microsoft Word.

Ce complément utilise également une base de données statistiques, issue de précédents traitements de documents pour le braille faits par le PAON. Lors du traitement des mots, ces statistiques sont utilisées pour présélectionner une action recommandée sur le mot en cours de traitement.

# Version 1.3.5.2 (Novembre 2024)

Cette version est la première à être mise à disposition du grand public.
