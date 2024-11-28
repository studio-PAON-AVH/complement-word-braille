# Journal des modifications

| **Version**  | **Date**  | **Changement**  | **Auteurs**  |
|:--:|----|----|----|
| 1.1  | 2024/19/07 | Refonte du processus de traitement des mots | ROGER Yan, PAVIE Nicolas |
| 1.2 | 2024/09/30 | Ajout de fonctionnalités | PAVIE Nicolas |
| 1.3.5 | 2024/11/26 | Correctifs | PAVIE Nicolas |

# Version 1.3.5 – 26/11/2024

Cette nouvelle version apporte les modifications suivantes au
complément :

- Les raccourcis clavier de la fenêtre de traitement ont été changés
  pour être plus évident. La nouvelle liste des raccourcis clavier est
  la suivante :

  - **p** pour **p**rotéger le mot

  - **a** pour **a**bréger le mot

  - **i** pour **i**gnorer le mot

  - **r** pour aller au mot précédent (**r**etour)

  - **s** pour aller au mot **s**uivant

- La fenêtre de traitement affiche maintenant la première règle
  d'abréviation détecté pour le mot traité

- La sélection du filtre des statuts à afficher dans la boite de
  dialogue à changer pour être remplacer par une liste de case a cocher

  - Par défaut, les occurrences de mots "ignorés" ne sont pas affichés,
    mais vous pouvez les afficher en cochant le statut dans la liste des
    filtres

- L'analyse des mots étranger ne se fait maintenant qu'au premier
  lancement du traitement sur un document, c'est a dire si aucun fichier
  .bdic n'a été détecté lors du lancement du traitement. Sinon, ces mots
  seront directement rechargés depuis le dictionnaire existant

  - Le chargement d'un fichier .bdic ou .ddic issus d'un autre
    traitement n'est toujours pas possible mais on y travaille

- Une passe de fiabilisation et réoptimisation a été faite sur le
  prétraitement pour simplifier le rechargement des décisions et la
  réanalyse en cas de modification du contenu

- J'ai corrigé un probleme de décalage de sélection et des insertions de
  codes de protection, qui pouvait se produire quand le document
  comportait un tableau.

- Des corrections sur la détection des mots abrégeables et sur la
  détection des mots étranger ont été ajoutées

  - Correction d'un bug sur le découpage des groupes de consonnes
    contenant un r ou un l

  - Modification de la détection pour ne pas découper les mots composés

  - Modification de la détection pour ne pas inclure les prépositions se
    terminant par une apostrophe (type Qu' ou C')

  - La particule "Von" a été enregistré pour contrôle systématique

# Version 1.2 – 30/09/2024

Cette nouvelle version apporte les modifications suivantes au processus
de traitement :

- Utilisation du correcteur orthographique de Word pour assister à la
  détection de mots étrangers,

- Évaluation de la liste de mots à contrôler systématiquement à partir
  de la liste fournis par Geneviève.

<!-- -->

- Réactivation du prémarquage automatique

  - Ajout d’une fenêtre d’option et d’une option pour désactiver le
    prémarquage automatique

- Un sélecteur de mot pour la navigation est ajouté à la fenêtre de
  traitement principal.

# Version 1.1 - 19/07/2024

Cette nouvelle version apporte une refonte du processus du traitement et
plus généralement les modifications détaillées ci-après.

La fenêtre d'action a été complètement remodelée pour un traitement par
mot plutôt que par occurrence. Cette fenêtre d’action affiche maintenant
pour un mot sélectionner :

- Les informations sur le mot sélectionné ont été réagencées :

  - Le mot traiter est affiché en gros et en gras.

  - Un indicateur de progression est affiché en haut à droite :

    - Cette progression est donnée par le nombre de mot traité sur le
      nombre total de mot détecter.

  - Les informations suivantes de la base de protection sur le mot
    sélectionner sont rappelées :

    - Le nombre de document où le mot a été protégé.

    - Le nombre de document où le mot a été abrégé

    - Le nombre de document où le mot a été détecté.

- Les boutons d’actions “Protéger ici” et “Abréger ici” applique
  maintenant la protection et l’abréviation d’un mot à toutes ses
  occurrences dans le document.

  - Ses boutons sont associés respectivement aux raccourcis claviers “i”
    et “a”

- Les boutons d’action pour une occurrence du mot sont remplacés par un
  tableau d’action par occurrence. Ce tableau est composé de deux
  colonnes, avec pour chaque ligne correspondant à une occurrence :

  - Une cellule contenant un sélecteur de statut pour l'occurrence

  - Une cellule affichant le texte de l’occurrence mise en gras dans son
    contexte d’utilisation dans le document. Ce contexte est un extrait
    de 50 caractères avant et 50 caractères après l’occurrence dans le
    contenu du document.

> Il est possible de modifier le statut d'une seule occurrence, en
> cliquant sur le sélecteur dans le tableau en face de l’occurrence qui
> modifiera le statut de l’occurrence dans le document.

- Le nombre d’occurrence du mot sélectionner dans le document est
  affiché en haut à gauche du tableau d’occurrence.

- Vous pouvez filtrer les occurrences du mot selon leur statut à l’aide
  d’un sélecteur de statut situé en haut à droite du tableau
  d’occurrence.

- Des boutons de navigation ont été ajoutée à la fenêtre de traitement
  pour permettre de revenir au mot précédent ou d’aller au suivant :

  - Le bouton “Précédent” permet de revenir au mot précédemment traiter

    - Ce bouton est associé au raccourci clavier “p”

  - Le bouton “Suivant” applique les décisions de protection ou
    d’abréviation indiqué dans le tableau d’action par occurrences sur
    le document avant de sélectionner le mot suivant.

    - Ce bouton est associé au raccourci clavier “s”

- Une barre de progression est ajoutée pour vous indiquer l’avancement
  de l’application de vos décisions pour chaque occurrence du mot dans
  le document lorsque vous cliquez sur le bouton “Suivant”.

- Il est également possible de cliquer sur une ligne du tableau d’action
  par occurrence pour mettre en évidence une occurrence dans le document
  : Cliquer sur une occurrence dans le tableau déplace le curseur de
  sélection de Word sur l'occurrence du mot dans le document.

- Vous pouvez arrêter le traitement à tout moment en fermant la fenêtre
  et le reprendre en relançant le traitement depuis le ruban.

Le format du dictionnaire de données a également été modifié : Les
actions choisies pour chaque occurrence de chaque mot sont maintenant
sauvegardées pour améliorer la qualité des données de protection, et
également vous permettre de reprendre plus facilement le travail fait
sur un fichier.

Concernant le pré-traitement du document, la détection des mots à
traiter a été optimiser pour ne sélectionner que les mots abrégeables.

La pré-protection automatique est désactivée, mais les informations de
la base de données restent affichées dans la fenêtre de traitement.
