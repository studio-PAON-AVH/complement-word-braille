using MSWord = Microsoft.Office.Interop.Word;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using System.IO;
using fr.avh.braille.dictionnaire;
using fr.avh.braille.dictionnaire.Entities;
using fr.avh.archivage;
using System.Windows;

namespace fr.avh.braille.addin
{
    /// <summary>
    ///
    /// </summary>
    public class ProtectionWord : IProtection
    {
        /// <summary>
        ///
        /// </summary>
        private static readonly string PROTECTION_CODE = "[[*i*]]";
        private static readonly string PROTECTION_CODE_G1 = "[[*G1*]]";
        private static readonly string PROTECTION_CODE_G2 = "[[*G2*]]";
        private static readonly string SEPARATORS = "\"\'\\t\\f\\v\\r\\n\\[\\]\\(\\);,.:  ";
        private static readonly string MIN = "a-zàáâãäåæçèéêëìíîïðñòóôõöøùúûüý";
        private static readonly string MAJ = "A-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝ";
        private static readonly string NUM = "0-9";
        private static readonly string ALPHANUM = $"{MIN}{MAJ}{NUM}_"; // == a-zàáâãäåæçèéêëìíîïðñòóôõöøùúûüýA-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝ0-9_
        private static readonly string REG_CS = "\\[\\[\\*"; // Debut de code duxburry
        private static readonly string REG_CE = "\\*\\]\\]"; // Fin de code duxburry
        private static readonly string WORD_WITH_NUMBERS_PATTERN =
            $"({REG_CS}i{REG_CE})?"
            + // Code duxburry de protection optionnel (récupérer dans le groupe 1) (0 ou 1)
            $"(?<!{REG_CS})"
            + // Negative lookbehind (ne pas match le texte commençant par un indicateur de code duxburry)
            $"("
            + $"\\b"
            + // Debut de mot (word boundary)
            $"[{NUM}{MIN}{MAJ}]*"
            + // Chiffre ou lettre (0 ou n)
            $"("
            + $"[{MIN}{MAJ}][{NUM}]"
            + // Lettre suivie d'un chiffre
            $"|"
            + $"[{NUM}][{MIN}{MAJ}]"
            + // Chiffre suivie d'une lettre
            $")+"
            + // Motif Chiffre+Lettre ou lettre+chiffre (1 ou n)
            $"[{NUM}{MIN}{MAJ}]*"
            + // suivi de chiffres ou de lettres (0 ou n)
            $"(?!{REG_CE})"
            + // Negative lookahead (ne pas match de texte suivi d'une fin de code duxburry)
            $"\\b"
            + // Fin de mot (word boundary)
            $")";


        /// <summary>
        /// Motif de recherche de mots contenant au moins 1 majuscule : <br/>
        /// = début de ligne ou caractère non alphanumérique,<br/>
        /// puis optionnellement suivi d'un code de protection,<br/>
        /// puis optionnellement suivi d'un groupe contenant <br/>
        ///    0 ou N caractère majuscule ou minuscule,<br/>
        ///    suivi d'un groupe contenant au choix<br/>
        ///    - soit une ou plusieurs (minuscules ou majuscule ou tiret) suivie d'une majuscule<br/>
        ///    - soit une majuscule suivi d'une ou plusieurs(minuscules ou majuscule ou tiret) <br/>
        ///    suivi de 0 ou N caractère minuscule, majuscule, underscore ou tiret<br/>
        ///    et suivi d'un caractère (non alphanumérique y compris sans apostrophe pour éviter les prefix en Qu' et C') ou de la fin de ligne
        /// </summary>
        private static readonly string WORD_WITH_CAPITAL_PATTERN =
            $"(?<=[^{ALPHANUM}-]|^)" // group 1 linestart or non alphanum character
            + $"({REG_CS}i{REG_CE})?" // group 2 Optional protection code
            + $"(" // group 3 : searched word
            + $"[{MIN}{MAJ}_-]*" // optionnal prefix
            + $"(" // Group 4: one or more capital letters and at least one other letter
            + $"[{MIN}{MAJ}-]+[{MAJ}]"
            + $"|[{MAJ}][{MIN}{MAJ}-]+"
            + $")" // end group 4
            + $"[{MIN}{MAJ}_-]*" // optionnal suffix
            + $")+" // end group 3
            + $"(?=[^{ALPHANUM}'’-]|$)"; // group 5 separator (non alphanum character, no apostrophes, or end of line)

        /// <summary>
        /// Motif de recherche de mot(s)
        /// </summary>
        /// <param name="pattern"></param>
        /// <param name="opts"></param>
        /// <returns>
        /// Regex with
        /// - Group 1 : Code de protection optionnel
        /// - Group 2 : Motif détecté
        /// </returns>
        public static Regex SearchWord(string pattern, RegexOptions opts)
        {
            return new Regex(
                $"(?<=[^{ALPHANUM}-]|^)(\\[\\[\\*i\\*\\]\\])?({pattern})(?=[^{ALPHANUM}-]|$)",
                RegexOptions.Compiled | opts
            );
        }

        /// <summary>
        /// Recherche des mots contenant un chiffre <br/>
        /// (lettre minuscule ou majuscul précédé ou suivi d'un chiffre)<br/>
        /// Goupes :<br/>
        /// - [1] : recup du code de protection s'il existe<br/>
        /// - [2] : le mot contenant un ou plusieurs nombre
        ///
        /// </summary>
        private static readonly Regex WORD_WITH_NUMBERS = new Regex(
            WORD_WITH_NUMBERS_PATTERN,
            RegexOptions.Compiled | RegexOptions.Singleline
        );

        /// <summary>
        /// Recherche tous les mots contenant au moins une majuscule
        /// Goupes :<br/>
        /// - [1] : recup du code de protection s'il existe<br/>
        /// - [2] : le mot contenant une ou plusieurs majuscules
        /// </summary>
        private static readonly Regex WORD_WITH_CAPITAL = new Regex(
            WORD_WITH_CAPITAL_PATTERN,
            RegexOptions.Compiled | RegexOptions.Singleline
        );

        /// <summary>
        /// Liste globale des occurences de mots a protégé obligatoirement
        /// (tel que présenté dans le document dans leur ordre d'apparition) <br/>
        /// voir la regex wordWithNumbers
        /// </summary>
        public List<string> occurencesToProtect { get; } = new List<string>();

        /// <summary>
        /// Pour chaque indice de occurencesToProtect, indique sur l'occurence est déjà protégé ou non
        /// </summary>
        public List<bool> occurencesToProtectStatus { get; } = new List<bool>();

        /// <summary>
        /// Dictionnaire de travail pour conserver les actions fait sur les mots <br/>
        /// - 0 : non-traité<br/>
        /// - 1 : abrégé globalement<br/>
        /// - 2 : protégé globalement<br/>
        /// - 3 : ambigu, traitement au cas par cas par occurence dans le titre<br/>
        /// </summary>
        public DictionnaireDeTravail WorkingDictionnary = null;

        public string WorkingDictionnaryPath = null;

        public string ResumePath = null;

        public delegate void OnSelectionChanged(int selectedOccurence);

        public event OnSelectionChanged SelectionChanged;

        private int _selectedOccurence = 0;

        /// <summary>
        /// Occurence sélectionné (indice des listes WorkingDictionnary.occurences et occurenceSelectedAction)
        /// </summary>
        public int SelectedOccurence
        {
            get => _selectedOccurence;
            private set
            {
                _selectedOccurence = value;
                SelectionChanged?.Invoke(value);
            }
        }
        public Statut SelectedOccurenceStatut
        {
            get => WorkingDictionnary.StatutsOccurences[_selectedOccurence];
            set => WorkingDictionnary.StatutsOccurences[_selectedOccurence] = value;
        }

        public bool SelectedOccurenceEstTraitee
        {
            get => WorkingDictionnary.StatutsAppliquer[_selectedOccurence];
            set => WorkingDictionnary.StatutsAppliquer[_selectedOccurence] = value;
        }

        public int SelectedWordIndex
        {
            get
            {
                string wordSelected = WorkingDictionnary.Occurences[SelectedOccurence].ToLower();
                return WorkingDictionnary.CarteMotOccurences.Keys.ToList().IndexOf(wordSelected);
            }
        }

        public string SelectedWord
        {
            get { return WorkingDictionnary.Occurences[SelectedOccurence].ToLower(); }
        }

        public int SelectedWordOccurenceIndex
        {
            get
            {
                string wordSelected = WorkingDictionnary.Occurences[SelectedOccurence].ToLower();
                return WorkingDictionnary.CarteMotOccurences[wordSelected].IndexOf(
                    SelectedOccurence
                );
            }
        }

        public int SelectedWordOccurenceCount
        {
            get
            {
                string wordSelected = WorkingDictionnary.Occurences[SelectedOccurence].ToLower();
                return WorkingDictionnary.CarteMotOccurences[wordSelected].Count;
            }
        }

        public MSWord.Document _document { get; } = null;

        public IList<Mot> alreadyInDB { get; private set; }

        public string TexteEnMemoire { get; private set; } = null;

        /// <summary>
        /// Récupère le contenu texte du document, mais retraité pour matcher la selection par range
        /// - Supression des charactere 'bell' utilisés par word pour les tableaux mais ignorer dans le range
        /// </summary>
        public string DocumentMainRange
        {
            get => _document.StoryRanges[WdStoryType.wdMainTextStory].Text.Replace("\a", "");
        }

        private Utils.OnInfoCallback info = null;
        private Utils.OnErrorCallback error = null;

        /// <summary>
        /// Crée un nouveau protecteur de document qui va analyser le document word
        /// et permettre de réaliser des actions sur les mots identifiés
        /// </summary>
        /// <param name="inputPath"></param>
        /// <param name="word"></param>
        /// <exception cref="Exception"></exception>
        public ProtectionWord(
            MSWord.Document document,
            Utils.OnInfoCallback info = null,
            Utils.OnErrorCallback error = null
        )
        {
            this.info = info;
            this.error = error;
            // Check du dépot distant pour synchronization

            string tempDirectory = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Directory.CreateDirectory(tempDirectory);
            object notReadOnly = false;
            string filename = document.FullName;
            string dicName = Path.GetFileNameWithoutExtension(filename) + ".bdic";
            WorkingDictionnaryPath = Path.Combine(Path.GetDirectoryName(filename), dicName);
            ResumePath = Path.Combine(
                Path.GetDirectoryName(filename),
                Path.GetFileNameWithoutExtension(filename) + ".resume"
            );

            if (document != null)
            {
                _document = document;
                AnalyserDocument();
                ProtegerPhrasesEtrangere();
            }
            else
            {
                throw new Exception("Impossible d'analyser le document sélectionner avec word");
            }
        }

        public void AnalyserDocument()
        {
            int baseStepNumber = 9;
            if (_document != null)
            {
                // Sauvegarde du précédent dictionnaire s'il existe
                DictionnaireDeTravail existingDictionnary = null;
                try
                {
                    if (File.Exists(WorkingDictionnaryPath))
                    {
                        info?.Invoke(
                            $"Chargement du dictionnaire existant {WorkingDictionnaryPath}",
                            new Tuple<int, int>(0, baseStepNumber)
                        );
                        System.Threading.Tasks.Task<DictionnaireDeTravail> init =
                            DictionnaireDeTravail.FromDictionnaryFileJSON(WorkingDictionnaryPath);
                        init.Wait();
                        existingDictionnary = init.Result;
                        info?.Invoke(
                            $"Dictionnaire chargé",
                            new Tuple<int, int>(1, baseStepNumber)
                        );
                    }
                }
                catch (Exception)
                {
                    //WorkingDictionnary = new DictionnaireDeTravail(dicName);
                    //info?.Invoke(
                    //    $"Nouveau fichier de dictionnaire {dicName}",
                    //    new Tuple<int, int>(0, baseStepNumber)
                    //);
                    //WorkingDictionnary.Save(new DirectoryInfo(Path.GetDirectoryName(filename)));
                }
                WorkingDictionnary = new DictionnaireDeTravail(
                    Path.GetFileNameWithoutExtension(WorkingDictionnaryPath) + ".bdic"
                );

                info?.Invoke(
                    "Préparation du document pour analyse ...",
                    new Tuple<int, int>(2, baseStepNumber)
                );

                info?.Invoke(
                    $"Copie du texte en mémoire...",
                    new Tuple<int, int>(3, baseStepNumber)
                );
                TexteEnMemoire = DocumentMainRange;

                // remise en 1 ligne du contenu pour analyse on remplace tous les retours chariot ou new line par des espaces
                info?.Invoke(
                    $"Recherche des mots contenant des chiffres...",
                    new Tuple<int, int>(4, baseStepNumber)
                );

                // List des mots à analyser :
                // Tous les mots de plus de 1 caractère et contenant au choix
                // - des chiffres et des lettres (a proteger systématiquement)
                // - au moins une majuscule
                MatchCollection result = WORD_WITH_NUMBERS.Matches(TexteEnMemoire);
                if (result.Count > 0)
                {
                    info?.Invoke(
                        $"{result.Count} mots contenant des chiffres a protéger",
                        new Tuple<int, int>(5, baseStepNumber)
                    );
                    foreach (Match item in result)
                    {
                        if (!occurencesToProtect.Contains(item.Groups[2].Value))
                        {
                            occurencesToProtect.Add(item.Groups[2].Value);
                        }
                    }

                    info?.Invoke(
                        $"Protection automatique des mots contenants des chiffres ...",
                        new Tuple<int, int>(7, baseStepNumber)
                    );

                    foreach (string item in occurencesToProtect)
                    {
                        // preprotection
                        MSWord.Range protectionRange = _document.StoryRanges[
                            WdStoryType.wdMainTextStory
                        ];
                        var finder = protectionRange.Find;
                        finder.ClearFormatting();
                        finder.Text = item;
                        finder.Forward = true;
                        finder.MatchCase = true;
                        while (finder.Execute())
                        {
                            int position = protectionRange.Start;
                            if (!EstProteger(_document, protectionRange))
                            {
                                // Protéger le texte
                                Proteger(protectionRange);
                            }
                        }
                    }
                    _document.Save();
                    // remise a 0 du texte en mémoire pour les analyses qui suivent
                    TexteEnMemoire = DocumentMainRange;
                }
                else
                {
                    info?.Invoke(
                        $"Pas de mots contenant des chiffres détecté",
                        new Tuple<int, int>(5, baseStepNumber)
                    );
                }

                info?.Invoke(
                    $"Recherche des mots contenant des majuscules ...",
                    new Tuple<int, int>(5, baseStepNumber)
                );
                result = WORD_WITH_CAPITAL.Matches(TexteEnMemoire);
                List<string> motsIdentifes = new List<string>();
                if (result.Count > 0)
                {
                    info?.Invoke(
                        $"{result.Count} occurences de mots contenant des majuscule, filtrage des mots abrégeables ..",
                        new Tuple<int, int>(6, baseStepNumber)
                    );

                    int t = 0;
                    foreach (Match item in result)
                    {
                        info?.Invoke($"", new Tuple<int, int>(t, result.Count));
                        t++;
                        bool isAlreadyProtected = item.Groups[1].Success;
                        string foundWord = item.Groups[2].Value;
                        int indexInText = item.Groups[2].Index;
                        string wordKey = foundWord.ToLower().Trim();
                        // Si le mot est abrégeable
                        if (Abreviation.EstAbregeable(wordKey) && !motsIdentifes.Contains(wordKey))
                        {
                            motsIdentifes.Add(wordKey);
                        }
                    }
                    info?.Invoke(
                        $"{motsIdentifes.Count} mots abrégeables ajoutés a la recherche",
                        new Tuple<int, int>(6, baseStepNumber)
                    );
                }
                else
                {
                    info?.Invoke(
                        $"Pas de mots contenant des majuscules détectés",
                        new Tuple<int, int>(6, baseStepNumber)
                    );
                }

                info?.Invoke(
                    $"Récupération des mots abrégeables communs au français et au moins une autre langue ...",
                    new Tuple<int, int>(6, baseStepNumber)
                );

                List<string> motsAToujourRechercher = new List<string>();
                using (var session = BaseSQlite.CreateSessionFactory(info, error).OpenSession())
                {
                    motsAToujourRechercher = session
                        .QueryOver<Mot>()
                        .Where(
                            m => m.ToujoursDemander == 1 /*&& Abbreviation.EstAbregeable(m.Texte)*/
                        )
                        .List()
                        .Distinct()
                        .Select(m => m.Texte.ToLower())
                        .ToList();
                }
                motsAToujourRechercher = motsAToujourRechercher
                    .Distinct()
                    .Where(m => Abreviation.EstAbregeable(m))
                    .ToList();
                info?.Invoke(
                    $"{motsAToujourRechercher.Count} mots abrégeables ajoutés a la recherche",
                    new Tuple<int, int>(6, baseStepNumber)
                );

                // NP 2024/11/20 : Rechargement des mots ajouter dans le dictionnaire précédent
                List<string> wordsToAdd = new List<string>();
                if (existingDictionnary != null)
                {
                    info?.Invoke(
                        $"Rechargement des mots issus de la dernière sauvegarde du dictionnaire (mots ajoutés et mots étrangers) ...",
                        new Tuple<int, int>(6, baseStepNumber)
                    );

                    wordsToAdd = existingDictionnary.CarteMotOccurences.Keys.ToList();
                }
                else
                {
                    wordsToAdd = GetMotsEtrangers();
                }

                info?.Invoke(
                    $"{wordsToAdd.Count} mots abrégeables ajoutés a la recherche",
                    new Tuple<int, int>(6, baseStepNumber)
                );

                List<string> totalMotsATraiter = new List<string>();
                totalMotsATraiter = motsIdentifes
                    .Concat(motsAToujourRechercher)
                    .Concat(wordsToAdd)
                    .Distinct()
                    .ToList();

                info?.Invoke(
                    $"Recherche de toutes les occurences de {totalMotsATraiter.Count} mots répertoriés...",
                    new Tuple<int, int>(6, baseStepNumber)
                );
                AjouterListeMotsAuTraitement(totalMotsATraiter);
                // repasser les occurences des mots minuscules identifiés en ignorer
                foreach (var word in motsIdentifes)
                {
                    string wordKey = word.ToLower().Trim();
                    if (WorkingDictionnary.CarteMotOccurences.ContainsKey(wordKey))
                    {
                        foreach (
                            var occurenceFound in WorkingDictionnary.CarteMotOccurences[wordKey]
                        )
                        {
                            if (WorkingDictionnary.Occurences[occurenceFound] == wordKey)
                            {
                                // ignorer toutes les versions en minuscules du mot
                                WorkingDictionnary.StatutsOccurences[occurenceFound] =
                                    Statut.IGNORE;
                            }
                            else
                            {
                                //
                            }
                        }
                    }
                }
                // Repasser tous les mots a toujorus rechercher en inconnu
                for (int i = 0; i < motsAToujourRechercher.Count; i++)
                {
                    // repasser les mots en question en inconnu s'ils ont été précédemment traité
                    if (
                        WorkingDictionnary.CarteMotOccurences.ContainsKey(motsAToujourRechercher[i])
                    )
                    {
                        foreach (
                            var occurence in WorkingDictionnary.CarteMotOccurences[
                                motsAToujourRechercher[i]
                            ]
                        )
                        {
                            WorkingDictionnary.StatutsOccurences[occurence] = Statut.INCONNU;
                        }
                    }
                }
                foreach (var kv in WorkingDictionnary.CarteMotOccurences)
                {
                    info?.Invoke($"- {kv.Key} : {kv.Value.Count} occurences");
                }

                info?.Invoke(
                    $"{WorkingDictionnary.CarteMotOccurences.Keys.Count} mots trouvés, ({WorkingDictionnary.Occurences.Count}) occurences à traiter"
                );

                info?.Invoke(
                    $"Récupérations des statistiques des mots détectés ...",
                    new Tuple<int, int>(6, baseStepNumber)
                );
                using (var session = BaseSQlite.CreateSessionFactory(info, error).OpenSession())
                {
                    alreadyInDB = session
                        .QueryOver<Mot>()
                        .WhereRestrictionOn(m => m.Texte)
                        .IsIn(WorkingDictionnary.CarteMotOccurences.Keys.ToList())
                        .List();

                    // Pré marquage des mots
                    if (OptionsComplement.Instance.ActiverPreProtectionAuto)
                    {
                        info?.Invoke(
                            $"Pré-marquage des mots et trouvés dans la base statistique ...",
                            new Tuple<int, int>(6, baseStepNumber)
                        );
                        // Ne pas pré marquer les mots ayant été détecté moins de 100 fois
                        // et dont la certitute de résolution est inférieur a 90%
                        int documentBarrier = 100;
                        foreach (var item in alreadyInDB)
                        {
                            if (item.ToujoursDemander == 1)
                            {
                                // Mots particuliers pour lequel le transcripteur doit vérifier son utilisation contextuellement
                                // (Par exemple, les mots existant dans plusieurs langues)
                                //WorkingDictionnary.mots[item.Texte] = Statut.AMBIGU;
                            }
                            else if (
                                item.Documents > documentBarrier
                                && Math.Max(item.Abreviations, item.Protections) > 0
                            )
                            {
                                double certainty =
                                    1.00
                                    - (
                                        (double)Math.Min(item.Abreviations, item.Protections)
                                        / (double)Math.Max(item.Abreviations, item.Protections)
                                    );
                                if (
                                    certainty > 0.99
                                    && WorkingDictionnary.StatutMot(item.Texte) == Statut.INCONNU
                                )
                                {
                                    Statut selected =
                                        item.Abreviations > item.Protections
                                            ? Statut.ABREGE
                                            : Statut.PROTEGE;
                                    // Ne mettre a jour que les mots non traité au cas ou on serait sur une reprise de traitement
                                    foreach (
                                        var index in WorkingDictionnary.CarteMotOccurences[
                                            item.Texte
                                        ]
                                    )
                                    {
                                        if (
                                            WorkingDictionnary.StatutsOccurences[index]
                                            == Statut.INCONNU
                                        )
                                            WorkingDictionnary.StatutsOccurences[index] = selected;
                                    }
                                }
                            }
                        }
                    }
                }

                if (existingDictionnary != null)
                {
                    info?.Invoke(
                        $"Rechargement des décisions enregistrés dans le précédent dictionnaire...",
                        new Tuple<int, int>(6, baseStepNumber)
                    );

                    // TODO : pour la suppression par les transcripteurs, remplacer la suppression par le statut ignorer
                    // (utilisé le statut 3 et le renommer en "IGNORER")
                    // Rechargement du dictionnaire précédent
                    // Pour chaque mot du nouveau dictionnaire
                    // si le mot i est le meme a l'emplacement j du précédent
                    //  on récupère le statut du mot j
                    //  on incrément j
                    // sinon on passe au mot suivant dans le nouveau dictionnaire dictionnaire
                    for (
                        int i = 0, j = 0;
                        i < WorkingDictionnary.Occurences.Count
                            && j < existingDictionnary.Occurences.Count;
                        i++
                    )
                    {
                        string newWord = WorkingDictionnary.Occurences[i];
                        string oldWord = existingDictionnary.Occurences[j];
                        Statut oldStatut = existingDictionnary.StatutsOccurences[j];
                        if (newWord == oldWord)
                        {
                            // trouvé
                            // on recharge le statut s'il y en avait un dans l'ancien dico
                            if (oldStatut != Statut.INCONNU)
                            {
                                WorkingDictionnary.StatutsOccurences[i] = oldStatut;
                                // On recontrole si le mot a été traité (statut appliquer à vrai si protéger a l'ouverture du fichier)
                                WorkingDictionnary.StatutsAppliquer[i] =
                                    (
                                        WorkingDictionnary.StatutsAppliquer[i]
                                        && oldStatut == Statut.PROTEGE
                                    ) // statut protégé et déjà protégé dans le document
                                    || (
                                        !WorkingDictionnary.StatutsAppliquer[i]
                                        && oldStatut == Statut.ABREGE
                                    ) // Statut abrégé et déjà abrégé dans le document
                                    || oldStatut == Statut.IGNORE; // Statut ignorer, on reprend
                            }
                            // on passe au mot suivant dans l'ancien dico
                            j++;
                        }
                        // sinon on passe au mot suivant dans le dictionnaire courant
                        // (pour le moment je ne resupprime pas les mots supprimés dans le dictionnaire précédent)
                    }
                }

                // Mise en place d'un fichier de reprise de traitement
                if (File.Exists(ResumePath))
                {
                    string[] resume = File.ReadAllLines(ResumePath);
                    string mot = resume[0];
                    bool isFinished = resume.Length > 1 && resume[1].StartsWith("--");
                    if (WorkingDictionnary.CarteMotOccurences.ContainsKey(mot))
                    {
                        info?.Invoke(
                            $"Relancement du traitement au mot {mot}",
                            new Tuple<int, int>(9, baseStepNumber)
                        );
                        List<string> keys = WorkingDictionnary.CarteMotOccurences.Keys.ToList();
                        int wordCountToReprocess = isFinished
                            ? keys.Count
                            : keys.IndexOf(mot.ToLower());
                        if (wordCountToReprocess > 0)
                        {
                            info?.Invoke(
                                $"Controle des décisions précédentes sur {wordCountToReprocess} mots",
                                new Tuple<int, int>(9, baseStepNumber)
                            );
                            for (int i = 0; i < wordCountToReprocess; i++)
                            {
                                string wordToCheck = keys[i];
                                List<int> toTreat = WorkingDictionnary.CarteMotOccurences[
                                    wordToCheck
                                ]
                                    .Where(
                                        o =>
                                            WorkingDictionnary.StatutsOccurences[o] == Statut.ABREGE
                                            || WorkingDictionnary.StatutsOccurences[o]
                                                == Statut.PROTEGE
                                    )
                                    .ToList();
                                info?.Invoke(
                                    $"- Controle des décisions précédentes sur {wordToCheck} : {WorkingDictionnary.CarteMotOccurences[wordToCheck].Count} occurences, {toTreat.Count} décisions à appliquer",
                                    new Tuple<int, int>(9, baseStepNumber)
                                );

                                // Parcours des occurences avec le parcours optimisé
                                foreach (int occurence in toTreat)
                                {
                                    try
                                    {
                                        SelectionnerOccurence(occurence);
                                        AppliquerStatutSurOccurence(
                                            SelectedOccurence,
                                            SelectedOccurenceStatut
                                        );
                                    }
                                    catch (Exception e)
                                    {
                                        MessageBox.Show(
                                            $"L'erreur suivante s'est produite lors du contrôle de {wordToCheck}\r\n"
                                                + e.Message
                                        );
                                    }
                                }
                            }
                        }
                        try
                        {
                            SelectionnerOccurenceMot(mot);
                        }
                        catch
                        {
                            SelectionnerOccurence(0);
                        }
                    }
                    else
                    {
                        SelectionnerOccurence(0);
                    }
                }
                else
                {
                    SelectionnerOccurence(0);
                }

                info?.Invoke(
                    $"Sauvegarde du dictionnaire des mots détectés",
                    new Tuple<int, int>(8, baseStepNumber)
                );
                this.Save();
                info?.Invoke($"Document prêt pour analyse", new Tuple<int, int>(9, baseStepNumber));
                TexteEnMemoire = DocumentMainRange;
                SelectedRange.Select();
            }
            else
            {
                throw new Exception("Impossible d'analyser le document sélectionner avec word");
            }
        }

        public void ReanalyserDocument()
        {
            if (TexteEnMemoire != DocumentMainRange)
            {
                info?.Invoke(
                    $"Modification de contenu détecter : réanalyse du document ..."
                );
                DictionnaireDeTravail actuel = WorkingDictionnary.Clone();
                WorkingDictionnary = new DictionnaireDeTravail(
                    Path.GetFileNameWithoutExtension(WorkingDictionnaryPath) + ".bdic"
                );
                this.TexteEnMemoire = DocumentMainRange;
                // Reajouter tous les mots au traitement
                info?.Invoke(
                    $"Recherche des occurences des {actuel.CarteMotOccurences.Keys.Count} mots précédement identifiés ..."
                );
                AjouterListeMotsAuTraitement(actuel.CarteMotOccurences.Keys.ToList());
                List<string> motsAControler = new List<string>();
                info?.Invoke($"Rechargement des décisions ...");
                for (int i = 0; i < actuel.CarteMotOccurences.Keys.Count; i++)
                {
                    string item = actuel.CarteMotOccurences.Keys.ElementAt(i);
                    info?.Invoke("", new Tuple<int, int>(i, actuel.CarteMotOccurences.Keys.Count));
                    // Appliquer le statut de l'occurence précédente a la nouvelle occurence
                    for (
                        int indexInMap = 0;
                        indexInMap < actuel.CarteMotOccurences[item].Count;
                        indexInMap++
                    )
                    {
                        int occurence = actuel.CarteMotOccurences[item][indexInMap];
                        if (indexInMap < WorkingDictionnary.CarteMotOccurences[item].Count)
                        {
                            int newOccurence = WorkingDictionnary.CarteMotOccurences[item][
                                indexInMap
                            ];
                            if (
                                actuel.Occurences[occurence]
                                != WorkingDictionnary.Occurences[newOccurence]
                            )
                            {
                                info?.Invoke(
                                    $"Attention !! l'occurence {indexInMap + 1} du mot {item} a changer d'écriture\r\n"
                                        + $"(avant: {actuel.Occurences[occurence]} - apres {WorkingDictionnary.Occurences[newOccurence]})"
                                );
                                if (!motsAControler.Contains(item))
                                {
                                    motsAControler.Add(item);
                                }
                            }
                            WorkingDictionnary.StatutsOccurences[
                                WorkingDictionnary.CarteMotOccurences[item][indexInMap]
                            ] = actuel.StatutsOccurences[occurence];
                        }
                        else
                        {
                            info?.Invoke(
                                $"Attention !! le nombre d'occurence du mot {item} a changé"
                            );
                            if (!motsAControler.Contains(item))
                            {
                                motsAControler.Add(item);
                            }
                            break;
                        }
                    }
                }
                this.Save();
                if (motsAControler.Count > 0)
                {
                    MessageBox.Show(
                        $"Attention, les mots suivants doivent êtres recontrolés : \r\n"
                            + string.Join(", ", motsAControler),
                        "Mots à recontroler",
                        MessageBoxButton.OK
                    );
                }
            }
        }

        public List<string> GetMotsEtrangers()
        {
            info?.Invoke(
                $"Recherche des mots étrangers en utilisant la detection de word ... (peut prendre du temps)"
            );
            // NP 2014/10/14 : remplacement de la détection des erreurs d'orthographes par une détection des mots étranger
            List<Range> consecutiveForeignWord = new List<Range>();
            List<string> motsARechercher = new List<string>();
            bool isInCode = false;
            // Pour chaque mot dans le document, demandé a word ce qu'il en pense et si c'est pas du français
            try
            {
                int indicateur = 0;
                int total = _document.Words.Count;
                foreach (MSWord.Range w in _document.Words)
                {
                    indicateur++;
                    info?.Invoke("", new Tuple<int, int>(indicateur, total));
                    w.DetectLanguage();
                    string text = w.Text;
                    bool hasUppercase = Regex.IsMatch(text, $"[{MAJ}]+");
                    // Controle des mots détectés entre des ouvertures et fermetures de code dbt
                    if (text.StartsWith("[[*") || text.EndsWith("[[*"))
                    {
                        isInCode = true;
                        continue;
                    }
                    if (text.EndsWith("*]]"))
                    {
                        isInCode = false;
                        continue;
                    }
                    if (isInCode)
                    {
                        // Mot détecté dans du code dbt
                        continue;
                    }
                    // Fin de paragraphe remonté dans la liste des mots
                    if (text == "\r")
                    {
                        // Fin de paragraph, on traite les mots sauvegarder
                        if (
                            consecutiveForeignWord.Count == 1
                            && Abreviation.EstAbregeable(consecutiveForeignWord[0].Text)
                        )
                        {
                            motsARechercher.Add(consecutiveForeignWord[0].Text.Trim());
                        }
                        else if (consecutiveForeignWord.Count > 1)
                        {
                            string phrase = "";
                            foreach (Range toadd in consecutiveForeignWord)
                            {
                                if (Abreviation.EstAbregeable(toadd.Text.Trim()))
                                {
                                    motsARechercher.Add(consecutiveForeignWord[0].Text.Trim());
                                }
                                phrase += toadd.Text;
                            }
                            info?.Invoke($"- Possible phrase étrangère détecté : {phrase.Trim()}");
                        }
                        consecutiveForeignWord.Clear();
                    }
                    else if (!Regex.IsMatch(text, $"[{MIN}{MAJ}]+"))
                    {
                        // Mot ne contenant pas de lettre
                        continue;
                    }
                    else if (w.LanguageID != MSWord.WdLanguageID.wdFrench)
                    {
                        // mot étranger
                        info?.Invoke(
                            $"- Mot en langue étrangère détecté {text} (langue : {w.LanguageID})"
                        );
                        // on garde son emplacement
                        consecutiveForeignWord.Add(_document.Range(w.Start, w.End));
                    }
                    else
                    {
                        // Mot français, on traite la liste des précédents mots étranger détectés
                        if (
                            consecutiveForeignWord.Count == 1
                            && Abreviation.EstAbregeable(consecutiveForeignWord[0].Text.Trim())
                        )
                        {
                            info?.Invoke(
                                $"- {consecutiveForeignWord[0].Text} ajouté au traitement"
                            );
                            motsARechercher.Add(consecutiveForeignWord[0].Text.Trim());
                            //AjouterMotAuTraitement(
                            //    consecutiveForeignWord[0].Text.Trim(),
                            //    refreshDictionnary: false
                            //);
                        }
                        else if (consecutiveForeignWord.Count > 1)
                        {
                            string phrase = "";
                            foreach (Range toadd in consecutiveForeignWord)
                            {
                                if (Abreviation.EstAbregeable(toadd.Text))
                                {
                                    motsARechercher.Add(consecutiveForeignWord[0].Text.Trim());
                                    //AjouterMotAuTraitement(
                                    //    toadd.Text.Trim(),
                                    //    refreshDictionnary: false
                                    //);
                                }
                                phrase += toadd.Text;
                            }
                            info?.Invoke($"- Possible phrase étrangère détecté : {phrase.Trim()}");
                        }
                        consecutiveForeignWord.Clear();
                    }
                }
                return motsARechercher;
                //foreach(string mot in motsARechercher)
                //{
                //    AjouterMotAuTraitement(mot, refreshDictionnary: false);
                //}
                //WorkingDictionnary.ReorderOccurences();
                //WorkingDictionnary.ComputeWordMap();
            }
            catch (Exception e)
            {
                MessageBox.Show(
                    "Impossible de procéder a la détection des mots étranger : \r\n"
                        + e.Message
                        + "\r\n"
                        + "Veuillez installer une langue de vérification supplémentaire (Options > Langue > Langue de création et de vérification > Ajouter)"
                );
                info?.Invoke(
                    $"Impossible de procéder a la détection des mots étranger : {e.Message}"
                );
                return new List<string>();
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="mot"></param>
        /// <param name="statut"></param>
        /// <param name="selectedOccurenceIndexInMap">Occurence sur laquelle appliqué le traitement. si -1, appliqué sur toutes les occurences</param>
        public void AjouterMotAuTraitement(
            string mot,
            Statut statut = Statut.INCONNU,
            int selectedOccurenceIndexInMap = -1,
            bool refreshDictionnary = true,
            bool alerteSiNonTrouver = false
        )
        {
            string word = mot.Trim().ToLower();
            info?.Invoke(
                    $"Ajout du mots '{word}' à la recherche de mots"
                );
            // Pour un mot donné, on recherche toute ses occurence dans le texte d'origine qu'on a garder en cache
            // NP 2024/11/20 : remplacement du word boundary par la recherche de caracteres non alphanumériques et des tirets (pour éviter les mots composés)
            //new Regex($"(^|[^{ALPHANUM}-])(\\[\\[\\*i\\*\\]\\])?({word})([^{ALPHANUM}-]|$)", RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Regex toLook = SearchWord(word, RegexOptions.IgnoreCase | RegexOptions.Singleline);

            MatchCollection check = toLook.Matches(TexteEnMemoire);
            if (check.Count > 0)
            {
                info?.Invoke(
                    $"{check.Count} occurences trouvées"
                );
                foreach (Match item in check)
                {
                    bool isAlreadyProtected = item.Groups[1].Success;
                    string foundWord = item.Groups[2].Value;
                    int indexInText = item.Groups[2].Index;
                    // Récupération du contexte autour de l'occurence
                    int indexBefore = Math.Max(0, indexInText - 50);
                    int indexAfter = Math.Min(
                        TexteEnMemoire.Length - 1,
                        indexInText + foundWord.Length + 50
                    );
                    string contextBefore = TexteEnMemoire.Substring(
                        indexBefore,
                        indexInText - indexBefore
                    );
                    string contextAfter = "";

                    contextAfter = TexteEnMemoire.Substring(
                        indexInText + foundWord.Length,
                        indexAfter - foundWord.Length - indexInText
                    );
                    // On ajoute cette occurence de mot si elle n'est pas déjà dans le dictionnaire de traitement
                    if (
                        !(
                            WorkingDictionnary.CarteMotOccurences.ContainsKey(word)
                            && WorkingDictionnary.CarteMotOccurences[word].FindIndex(
                                occurence =>
                                    WorkingDictionnary.PositionsOccurences[occurence] == indexInText
                            ) >= 0
                        )
                    )
                    {
                        WorkingDictionnary.Add(
                            foundWord,
                            isAlreadyProtected ? Statut.PROTEGE : Statut.INCONNU,
                            contextBefore,
                            contextAfter,
                            indexInText
                        );
                        if (isAlreadyProtected)
                        {
                            WorkingDictionnary.StatutsAppliquer[
                                WorkingDictionnary.Occurences.Count - 1
                            ] = true;
                        }
                    }
                }
                if (refreshDictionnary)
                {
                    WorkingDictionnary.ReorderOccurences();
                    WorkingDictionnary.ComputeWordMap();
                }
                if (statut != Statut.INCONNU)
                {
                    if (selectedOccurenceIndexInMap >= 0)
                    {
                        // on a sélectionner une occurence en particulier
                        WorkingDictionnary.StatutsOccurences[
                            WorkingDictionnary.CarteMotOccurences[word][selectedOccurenceIndexInMap]
                        ] = statut;
                    }
                    else
                    {
                        // on a pas sélectionner d'occurence, on applique a toutes les occurences
                        foreach (var index in WorkingDictionnary.CarteMotOccurences[word])
                        {
                            WorkingDictionnary.StatutsOccurences[index] = statut;
                        }
                    }
                }
            }
            else if (alerteSiNonTrouver)
            {
                info?.Invoke($"Attention : {word} non retrouvé");
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="mots"></param>
        /// <param name="statut"></param>
        /// <param name="refreshDictionnary"></param>
        /// <param name="alerteSiNonTrouver"></param>
        public void AjouterListeMotsAuTraitement(
            List<string> mots,
            Statut statut = Statut.INCONNU,
            bool refreshDictionnary = true,
            bool alerteSiNonTrouver = false
        )
        {
            if (mots.Count == 0)
                return;

            string word = mots[0].Trim().ToLower();
            for (int i = 1; i < mots.Count; i++)
            {
                word += "|" + mots[i].Trim().ToLower();
            }
            Regex toLook = SearchWord(
                word,
                RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.Singleline
            );
            MatchCollection check = toLook.Matches(TexteEnMemoire);
            if (check.Count > 0)
            {
                foreach (Match item in check)
                {
                    bool isAlreadyProtected = item.Groups[1].Success;
                    string foundWord = item.Groups[2].Value;
                    int indexInText = item.Groups[2].Index;
                    // Récupération du contexte autour de l'occurence
                    int indexBefore = Math.Max(0, indexInText - 50);
                    int indexAfter = Math.Min(
                        TexteEnMemoire.Length - 1,
                        indexInText + foundWord.Length + 50
                    );
                    string contextBefore = TexteEnMemoire.Substring(
                        indexBefore,
                        indexInText - indexBefore
                    );
                    string contextAfter = "";

                    contextAfter = TexteEnMemoire.Substring(
                        indexInText + foundWord.Length,
                        indexAfter - foundWord.Length - indexInText
                    );
                    // On ajoute cette occurence de mot si elle n'est pas déjà dans le dictionnaire de traitement
                    if (
                        !(
                            WorkingDictionnary.CarteMotOccurences.ContainsKey(
                                foundWord.ToLower().Trim()
                            )
                            && WorkingDictionnary.CarteMotOccurences[
                                foundWord.ToLower().Trim()
                            ].FindIndex(
                                occurence =>
                                    WorkingDictionnary.PositionsOccurences[occurence] == indexInText
                            ) >= 0
                        )
                    )
                    {
                        WorkingDictionnary.Add(
                            foundWord,
                            isAlreadyProtected ? Statut.PROTEGE : Statut.INCONNU,
                            contextBefore,
                            contextAfter,
                            indexInText
                        );
                        if (isAlreadyProtected)
                        {
                            WorkingDictionnary.StatutsAppliquer[
                                WorkingDictionnary.Occurences.Count - 1
                            ] = true;
                        }
                    }
                }
                if (refreshDictionnary)
                {
                    WorkingDictionnary.ReorderOccurences();
                    WorkingDictionnary.ComputeWordMap();
                }
                if (statut != Statut.INCONNU)
                {
                    foreach (var index in WorkingDictionnary.CarteMotOccurences[word])
                    {
                        WorkingDictionnary.StatutsOccurences[index] = statut;
                    }
                }
            }
            else if (alerteSiNonTrouver)
            {
                info?.Invoke($"Attention : {word} non retrouvé");
            }
        }

        public void Save()
        {
            this.WorkingDictionnary.SaveJSON(
                new DirectoryInfo(Path.GetDirectoryName(WorkingDictionnaryPath))
            );
            if (File.Exists(ResumePath))
            {
                File.Delete(ResumePath);
            }
            // Fichier de reprise de traitement
            File.WriteAllLines(
                ResumePath,
                new string[] { SelectedWord, EstTerminer() ? "-- terminé" : "" }
            );
        }

        public delegate void OnProtectorProgress(
            string message = null,
            Tuple<int, int> progress = null
        );

        public void AppliquerStatutsSurDocument(OnProtectorProgress callback = null)
        {
            this.ReanalyserDocument();

            callback?.Invoke(
                $"Application des statuts sur les mots du documents...",
                new Tuple<int, int>(0, WorkingDictionnary.Occurences.Count)
            );
            for (int i = 0; i < WorkingDictionnary.Occurences.Count; i++)
            {
                AppliquerStatutSurOccurence(i, WorkingDictionnary.StatutsOccurences[i]);
                callback?.Invoke(
                    progress: new Tuple<int, int>(i + 1, WorkingDictionnary.Occurences.Count)
                );
            }
            //this.AnalyserDocument();
        }

        public void ReinitialiserDocument(OnProtectorProgress callback = null)
        {
            info?.Invoke($"Déprotection des mots contenants des chiffres ...");
            MSWord.Range protectionRange = _document.StoryRanges[WdStoryType.wdMainTextStory];
            var finder = protectionRange.Find;
            int p = 0;
            foreach (string item in occurencesToProtect)
            {
                p++;
                callback?.Invoke(
                    $"Protection automatique des mots contenants des chiffres ...",
                    new Tuple<int, int>(p, occurencesToProtect.Count)
                );
                finder.ClearFormatting();
                finder.Text = item;
                finder.Forward = true;
                finder.MatchCase = true;
                if (finder.Execute())
                {
                    Abreger(protectionRange);
                }
            }

            info?.Invoke($"Application de tous les statuts sur les mots du documents...");
            MSWord.Range docRange = _document.StoryRanges[WdStoryType.wdMainTextStory];
            for (int i = 0; i < WorkingDictionnary.Occurences.Count; i++)
            {
                callback?.Invoke(
                    $"Application de tous les statuts sur les mots du documents...",
                    new Tuple<int, int>(i + 1, WorkingDictionnary.Occurences.Count)
                );
                string searchedWord = WorkingDictionnary.Occurences[i];
                string contextAfter = WorkingDictionnary.ContextesApresOccurences[i];
                string contextBefore = WorkingDictionnary.ContextesAvantOccurences[i];
                Statut toApply = WorkingDictionnary.StatutsOccurences[i];
                docRange = _document.Range(
                    docRange.Start,
                    _document.StoryRanges[WdStoryType.wdMainTextStory].End
                );
                finder = docRange.Find;
                finder.ClearFormatting();
                // Pour eviter les mots composés
                string before = (contextBefore.Length > 0 ? "[!-]" : "");
                string after = contextAfter;
                if (after.IndexOf("\r") >= 0)
                {
                    after = after.Substring(0, after.IndexOf("\r"));
                }
                finder.Text = $"{before}{searchedWord}>{after}";
                finder.Forward = true;
                finder.MatchCase = true;
                finder.MatchWildcards = true;
                bool found = false;
                try
                {
                    found = finder.Execute();
                }
                catch
                {
                    found = false;
                }
                if (found)
                {
                    docRange = _document.Range(
                        docRange.Start + (before.Length > 0 ? 1 : 0),
                        docRange.End - after.Length
                    );
                }
                else
                {
                    // Fallback : redo the search without whole word but using the next context
                    finder.ClearFormatting();
                    finder.Text = $"{searchedWord}{after}";
                    finder.Forward = true;
                    finder.MatchCase = true;
                    finder.MatchWholeWord = false;
                    if (!finder.Execute())
                    {
                        throw new Exception(
                            $"L'occurence numéro {i} ({searchedWord}) n'a pas pu être retrouvé dans le texte"
                        );
                    }
                    else
                    {
                        docRange = _document.Range(docRange.Start, docRange.End - after.Length);
                    }
                }
                docRange.Select();
                SelectedRange = docRange;
                SelectedOccurence = i;
                Abreger(SelectedRange);
            }
            // Sauvegarde du document
            _document.Save();
            this.ReanalyserDocument();
        }

        #region Actions

        public void ProtegerOccurence()
        {
            var occurrences = WorkingDictionnary.CarteMotOccurences[SelectedWord];
            foreach (var index in occurrences)
            {
                WorkingDictionnary.StatutsOccurences[index] = Statut.PROTEGE;
            }
            SelectedRange = Proteger(SelectedRange);
            /*WorkingDictionnary.StatutsOccurences[SelectedOccurence] = Statut.PROTEGE;
            SelectedRange = Proteger(SelectedRange);*/
        }

        public void AbregerOccurence()
        {
            var occurrences = WorkingDictionnary.CarteMotOccurences[SelectedWord];
            foreach (var index in occurrences)
            {
                WorkingDictionnary.StatutsOccurences[index] = Statut.ABREGE;
            }
            SelectedRange = Abreger(SelectedRange);
            /*WorkingDictionnary.StatutsOccurences[SelectedOccurence] = Statut.ABREGE;
            SelectedRange = Proteger(SelectedRange);*/
        }

        /// <summary>
        /// Fonction utilitaire d'insertion d'un code de protection d'une sélection dans un document.
        /// </summary>
        /// <param name="current"></param>
        /// <param name="wordRange"></param>
        /// <returns></returns>
        public static MSWord.Range Proteger(MSWord.Document current, MSWord.Range wordRange)
        {
            string trimmedStart = wordRange.Text.TrimStart();
            int trimmedCount = wordRange.Text.Length - trimmedStart.Length;
            if (trimmedCount > 0)
            {
                // Reselectionner le mot sans les espaces de début
                wordRange = current.Range(wordRange.Start + trimmedCount, wordRange.End);
                wordRange.Select();
            }

            if (!EstProteger(current, wordRange))
            {
                wordRange.InsertBefore(PROTECTION_CODE);
                wordRange = current.Range(wordRange.Start + PROTECTION_CODE.Length, wordRange.End);
                wordRange.Select();
            }
            return wordRange;
        }

        /// <summary>
        /// Fonction utilitaire de protection d'un bloc de mots dans un document.
        /// </summary>
        /// <param name="current"></param>
        /// <param name="wordRange"></param>
        /// <returns></returns>
        public static MSWord.Range ProtegerBloc(MSWord.Document current, MSWord.Range wordRange)
        {
            string trimmedStart = wordRange.Text.TrimStart();
            int trimmedCount = wordRange.Text.Length - trimmedStart.Length;
            if (trimmedCount > 0)
            {
                // Reselectionner le mot sans les espaces de début
                wordRange = current.Range(
                    wordRange.Start + trimmedCount,
                    wordRange.End
                );
                wordRange.Select();
            }

            if (!EstProteger(current, wordRange))
            {
                wordRange.InsertBefore(PROTECTION_CODE_G1);
                wordRange.InsertAfter(PROTECTION_CODE_G2);
                wordRange = current.Range(
                    wordRange.Start + PROTECTION_CODE_G1.Length,
                    wordRange.End - PROTECTION_CODE_G2.Length
                );
                wordRange.Select();
            }
            return wordRange;
        }

        public void ProtegerPhrasesEtrangere()
        {
            List<string> phrasesEtrangers = new List<string>();
            List<MSWord.Range> currentPhrase = new List<MSWord.Range>();
            bool isInCode = false;

            try
            {
                foreach (MSWord.Range w in _document.Words)
                {
                    w.DetectLanguage();
                    string text = w.Text;

                    isInCode = (text.StartsWith("[[*") || text.EndsWith("[[*")) ? true : (text.EndsWith("*]]") ? false : isInCode);

                    if (w.LanguageID != MSWord.WdLanguageID.wdFrench)
                    {
                        currentPhrase.Add(w);
                    }
                    else
                    {
                        if (currentPhrase.Count > 0)
                        {
                            // Ajouter les codes de protection G1 et G2 autour de la phrase étrangère
                            MSWord.Range phraseRange = _document.Range(currentPhrase.First().Start, currentPhrase.Last().End);
                            if (!phraseRange.Text.StartsWith(PROTECTION_CODE_G1))
                            {
                                phraseRange.InsertBefore(PROTECTION_CODE_G1);
                            }
                            if (!phraseRange.Text.EndsWith(PROTECTION_CODE_G2))
                            {
                                phraseRange.InsertAfter(PROTECTION_CODE_G2);
                            }
                            phrasesEtrangers.Add(phraseRange.Text.Trim());
                            currentPhrase.Clear();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(
                    "Impossible de procéder à la détection des phrases étrangères : \r\n"
                    + e.Message
                    + "\r\n"
                    + "Veuillez installer une langue de vérification supplémentaire (Options > Langue > Langue de création et de vérification > Ajouter)"
                );
            }
        }

        public static bool EstProteger(MSWord.Document current, MSWord.Range wordRange)
        {
            string trimmedStart = wordRange.Text.TrimStart();
            int trimmedCount = wordRange.Text.Length - trimmedStart.Length;
            if (trimmedCount > 0)
            {
                // Reselectionner le mot sans les espaces de début
                wordRange = current.Range(wordRange.Start + trimmedCount, wordRange.End);
            }
            // Le mot sélectionné est au début et n'est pas précédé de suffisament de caractères pour être protégé
            if (wordRange.Start - PROTECTION_CODE.Length < 0)
            {
                return false;
            }
            // Get previous range
            MSWord.Range previous = current.Range(
                wordRange.Start - PROTECTION_CODE.Length,
                wordRange.Start
            );
            string previousText = previous.Text;
            return previousText.Equals(PROTECTION_CODE)
                || wordRange.Text.StartsWith(PROTECTION_CODE);
        }

        /// <summary>
        /// Protège un mot à une position donnée dans le document courant.
        /// (Ajoute un code de protection devant l'occurence s'il n'est pas déjà définit)
        /// </summary>
        /// <param name="wordRange">Position du mot</param>
        /// <returns>La nouvelle position du mot</returns>
        public MSWord.Range Proteger(MSWord.Range wordRange)
        {
            return Proteger(_document, wordRange);
        }

        /// <summary>
        /// Fonction utilitaire pour abréger (c.a.d. supprimer un code de protection devant) un mot dans un document
        /// </summary>
        /// <param name="current"></param>
        /// <param name="wordRange"></param>
        /// <returns></returns>
        public static MSWord.Range Abreger(Document current, MSWord.Range wordRange)
        {
            string trimmedStart = wordRange.Text.TrimStart();
            int trimmedCount = wordRange.Text.Length - trimmedStart.Length;
            if (trimmedCount > 0)
            {
                // Reselectionner le mot sans les espaces de début
                wordRange = current.Range(wordRange.Start + trimmedCount, wordRange.End);
                wordRange.Select();
            }
            if (wordRange.Start - PROTECTION_CODE.Length >= 0)
            {
                MSWord.Range previous = current.Range(
                    wordRange.Start - PROTECTION_CODE.Length,
                    wordRange.Start
                );
                string previousText = previous.Text;
                // First check if text is preceded by protection code
                if (previousText.Equals(PROTECTION_CODE))
                {
                    Range toDelete = current.Range(
                        wordRange.Start - PROTECTION_CODE.Length,
                        wordRange.Start
                    );
                    toDelete.Delete();
                    wordRange = current.Range(toDelete.End, wordRange.End);
                    wordRange.Select();
                }
            }
            // check if selection starts with protection code
            if (trimmedStart.StartsWith(PROTECTION_CODE))
            {
                Range toDelete = current.Range(
                    wordRange.Start,
                    wordRange.Start + PROTECTION_CODE.Length
                );
                toDelete.Delete();
                wordRange = current.Range(toDelete.End, wordRange.End);
                wordRange.Select();
            }
            string word = wordRange.Text.ToLower().Trim();
            return wordRange;
        }

        /// <summary>
        /// Abrège une occurence de mot a une position donnée dans le document courant.
        /// (Retire un code de protection devant le mot s'il est définit)
        /// </summary>
        /// <param name="wordRange">position du mot </param>
        /// <returns></returns>
        public MSWord.Range Abreger(MSWord.Range wordRange)
        {
            return Abreger(_document, wordRange);
        }

        public static MSWord.Range AbregerBloc(Document current, MSWord.Range wordRange)
        {
            string trimmedStart = wordRange.Text.TrimStart();
            int trimmedCount = wordRange.Text.Length - trimmedStart.Length;
            if (trimmedCount > 0)
            {
                wordRange = current.Range(
                    wordRange.Start + trimmedCount,
                    wordRange.End
                );
                wordRange.Select();
            }

            // si le mot est précédé de suffisament de caractères pour être protégé
            if (wordRange.Start - PROTECTION_CODE_G1.Length >= 0 && wordRange.End + PROTECTION_CODE_G2.Length <= current.Content.End)
            {
                MSWord.Range previous = current.Range(
                    wordRange.Start - PROTECTION_CODE_G1.Length,
                    wordRange.Start
                );
                string previousText = previous.Text;
                // si le mot est précédé du code de protection G1
                if (previousText.Equals(PROTECTION_CODE_G1))
                {
                    // Suppression du code de protection G1
                    Range toDelete = current.Range(
                        wordRange.Start - PROTECTION_CODE_G1.Length,
                        wordRange.Start
                    );

                    // Suppression du code de protection G2
                    toDelete.Delete();
                    wordRange = current.Range(toDelete.End, wordRange.End);
                    wordRange.Select();
                }
                MSWord.Range next = current.Range(
                    wordRange.End,
                    wordRange.End + PROTECTION_CODE_G2.Length
                );
                // si le mot est suivi du code de protection G2
                string nextText = next.Text;
                if (nextText.Equals(PROTECTION_CODE_G2))
                {
                    Range toDelete = current.Range(
                        wordRange.End,
                        wordRange.End + PROTECTION_CODE_G2.Length
                    );
                    toDelete.Delete();
                    wordRange = current.Range(wordRange.Start, toDelete.Start); //est ce qu'on peut lier le start et le end pour avoir moin de code
                    wordRange.Select();
                }
            }
            if (trimmedStart.StartsWith(PROTECTION_CODE_G1))
            {
                Range toDelete = current.Range(
                    wordRange.Start,
                    wordRange.Start + PROTECTION_CODE_G1.Length
                );
                toDelete.Delete();
                wordRange = current.Range(toDelete.End, wordRange.End);
                wordRange.Select();
            }
            string word = wordRange.Text.ToLower().Trim();
            return wordRange;
        }

        /// <summary>
        /// Marquer le mot sélectionné comme étant à protéger dans le traitement
        /// </summary>
        /// <param name="wordRange"></param>
        public void MarquerPourProtection(MSWord.Range wordRange)
        {
            string wordSelected = wordRange.Text.ToLower().Trim().ToLower();
            WorkingDictionnary.SetStatut(wordSelected, Statut.PROTEGE);
        }

        /// <summary>
        /// Marquer le mot sélectionné comme étant à abréger dans le traitement
        /// </summary>
        /// <param name="wordRange"></param>
        public void MarquerPourAbreviation(MSWord.Range wordRange)
        {
            string wordSelected = wordRange.Text.ToLower().Trim().ToLower();
            WorkingDictionnary.SetStatut(wordSelected, Statut.ABREGE);
        }

        /// <summary>
        /// Marquer le mot sélectionné comme ambigu, c'est à dire a traiter au cas par cas
        /// </summary>
        /// <param name="wordRange"></param>
        public void MarquerCommeAmbigu(MSWord.Range wordRange)
        {
            string wordSelected = wordRange.Text.ToLower().Trim().ToLower();
        }

        /// <summary>
        /// Vérifie si le traitement est terminé, c'est à dire si toutes les occurences ont été traité
        /// </summary>
        /// <param name="wordRange"></param>
        public bool EstTerminer()
        {
            // Si toute les occurences ont vu leur statuts attribué
            return WorkingDictionnary.StatutsOccurences.All(s => s != Statut.INCONNU);
        }

        #endregion

        #region Selection des occurences triés par mot

        // TODO : Déplacer les fonctions de navigation dans une classe a part pour simplifier la maintenance

        /// <summary>
        /// Select le mot suivant
        /// </summary>
        /// <param name="andOccurence">Si définit, sélectionne une occurence spécifique</param>
        /// <returns></returns>
        public MSWord.Range ProchainMot(int andOccurence = 0)
        {
            string occurenceSelected = WorkingDictionnary.Occurences[SelectedOccurence];
            string wordSelected = occurenceSelected.ToLower();
            List<string> keys = WorkingDictionnary.CarteMotOccurences.Keys.ToList();
            int newWordKeyIndex = (keys.Count + keys.IndexOf(wordSelected) + 1) % keys.Count;
            return SelectionnerOccurenceMot(keys[newWordKeyIndex], andOccurence);
        }

        public MSWord.Range PrecedentMot(int andOccurence = 0)
        {
            string occurenceSelected = WorkingDictionnary.Occurences[SelectedOccurence];
            string wordSelected = occurenceSelected.ToLower();
            List<string> keys = WorkingDictionnary.CarteMotOccurences.Keys.ToList();
            int newWordKeyIndex = (keys.Count + keys.IndexOf(wordSelected) - 1) % keys.Count;
            return SelectionnerOccurenceMot(keys[newWordKeyIndex], andOccurence);
        }

        /// <summary>
        /// "Supprime" le mot sélectionner et réselectionne le mot suivant
        /// </summary>
        /// <returns></returns>
        public MSWord.Range IgnorerMotEtSelectionnerSuivant()
        {
            // Récupération du mot précédent à partir de la carte actuel
            string wordSelected = SelectedWord.ToLower();

            // listes des occurences a supprimer de la liste
            List<int> wordOccurences = WorkingDictionnary.CarteMotOccurences[wordSelected];
            foreach (var occurence in WorkingDictionnary.CarteMotOccurences[wordSelected])
            {
                WorkingDictionnary.StatutsOccurences[occurence] = Statut.IGNORE;
            }

            // Sélectionner le mot suivant
            return ProchainMot();
        }

        public MSWord.Range ProchaineOccurenceDuMot()
        {
            string occurenceSelected = WorkingDictionnary.Occurences[SelectedOccurence];
            string wordSelected = occurenceSelected.ToLower();
            List<int> occurences = WorkingDictionnary.CarteMotOccurences[wordSelected];
            int wordOccurenceIndex =
                (occurences.Count + occurences.IndexOf(SelectedOccurence) + 1) % occurences.Count;
            int wordSelectedOccurence = WorkingDictionnary.CarteMotOccurences[
                occurenceSelected.ToLower()
            ][wordOccurenceIndex];
            string searchedWord = WorkingDictionnary.Occurences[wordSelectedOccurence];
            string contextBefore = WorkingDictionnary.ContextesAvantOccurences[
                wordSelectedOccurence
            ];
            try
            {
                var docRange = _document.Range(
                    SelectedRange.End,
                    _document.StoryRanges[WdStoryType.wdMainTextStory].End
                );
                var finder = docRange.Find;

                finder.ClearFormatting();
                finder.Text = (contextBefore.Length > 0 ? "[!-]<" : "<") + searchedWord + ">";
                finder.Forward = true;
                finder.MatchCase = true;
                finder.MatchWildcards = true;
                if (finder.Execute())
                {
                    docRange = _document.Range(
                        docRange.Start + (contextBefore.Length > 0 ? 1 : 0),
                        docRange.End
                    );
                    SelectedRange = docRange;
                    SelectedOccurence = wordSelectedOccurence;
                }
                else
                {
                    throw new Exception("Fin du document atteinte");
                }
            }
            catch (Exception)
            {
                var rangeStart = SelectedRange.Start;
                var rangeEnd = SelectedRange.End;
                var docEnd = _document.StoryRanges[WdStoryType.wdMainTextStory].End;
                SelectedOccurence = -1;
            }

            return SelectedRange;
        }

        public MSWord.Range PrecedenteOccurenceDuMot()
        {
            string occurenceSelected = WorkingDictionnary.Occurences[SelectedOccurence];
            string wordSelected = occurenceSelected.ToLower();
            List<int> occurences = WorkingDictionnary.CarteMotOccurences[wordSelected];
            int wordOccurenceIndex = occurences.IndexOf(SelectedOccurence);
            return wordOccurenceIndex == 0
                ? PrecedentMot(-1)
                : SelectionnerOccurence(occurences[wordOccurenceIndex - 1]);
        }

        /// <summary>
        /// Protéger un mot dans tout le document
        /// (Note: un peu lent en traitement itératif, réserver pour un traitement mot par mot séparé)
        /// </summary>
        /// <param name="mot"></param>
        public void ProtegerMot(string mot)
        {
            string wordSelected = mot.ToLower();
            foreach (var idx in WorkingDictionnary.CarteMotOccurences[wordSelected])
            {
                WorkingDictionnary.StatutsOccurences[idx] = Statut.PROTEGE;
                Range current = SelectionnerOccurence(idx);
                Proteger(current);
            }
        }

        /// <summary>
        /// Abreger un mot dans tout le document
        /// (Note: un peu lent en traitement itératif, réserver pour un traitement mot par mot séparé)
        /// </summary>
        /// <param name="mot"></param>
        public void AbregerMot(string mot)
        {
            string wordSelected = mot.ToLower();
            foreach (var idx in WorkingDictionnary.CarteMotOccurences[wordSelected])
            {
                WorkingDictionnary.StatutsOccurences[idx] = Statut.ABREGE;
                Range current = SelectionnerOccurence(idx);
                Abreger(current);
            }
        }
        #endregion


        #region Sélection des occurences non trié dans le document
        private MSWord.Range _selectedRange;

        /// <summary>
        /// Selection courante dans le document pour optimiser
        /// </summary>
        public MSWord.Range SelectedRange
        {
            get => _selectedRange;
            private set { _selectedRange = value; }
        }

        /// <summary>
        /// Selectionne une occurence identifié pour traitement dans le document
        /// </summary>
        /// <param name="newSelectedOccurenceIndex">Indice de l'occurence dans la liste des occurences identifiées pour traitement</param>
        /// <returns></returns>
        public MSWord.Range SelectionnerOccurence(int newSelectedOccurenceIndex)
        {
            if (WorkingDictionnary.Occurences.Count == 0)
            {
                throw new Exception("Aucun mot n'a été détecté pour traitement dans le document");
            }
            SelectedOccurence = newSelectedOccurenceIndex;
            int startPosition = WorkingDictionnary.PositionsOccurences[SelectedOccurence];
            SelectedRange = _document.Range(startPosition, startPosition + SelectedWord.Length);
            return SelectedRange;
        }

        /// <summary>
        /// Selectionne un mot dans le document
        /// </summary>
        /// <param name="key">Mot rechercher</param>
        /// <param name="indexCarte">numéro de l'occurence à rechercher</param>
        /// <returns></returns>
        public MSWord.Range SelectionnerOccurenceMot(string key, int indexCarte = 0)
        {
            if (WorkingDictionnary.Occurences.Count == 0)
            {
                throw new Exception("Aucun mot n'a été détecté pour traitement dans le document");
            }
            if (!WorkingDictionnary.CarteMotOccurences.ContainsKey(key.ToLower()))
            {
                throw new Exception($"Le mot {key} n'est pas détecter dans la carte des mots");
            }
            // Passage par les positions précalculés, plus efficace et plus rapide que le passage dans le finder de word (mais plus fragile si le texte est modifié)
            SelectedOccurence = WorkingDictionnary.CarteMotOccurences[key.ToLower()][indexCarte];
            int startPosition = WorkingDictionnary.PositionsOccurences[SelectedOccurence];

            SelectedRange = _document.Range(startPosition, startPosition + SelectedWord.Length);
            return SelectedRange;
        }

        /// <summary>
        /// Selectionne la prochaine occurence de mot dans le document
        /// </summary>
        /// <returns></returns>
        public MSWord.Range ProchaineOccurence()
        {
            SelectedOccurence =
                (WorkingDictionnary.Occurences.Count + SelectedOccurence + 1)
                % WorkingDictionnary.Occurences.Count;
            string occurenceSelected = WorkingDictionnary.Occurences[SelectedOccurence];
            string contextBefore = WorkingDictionnary.ContextesAvantOccurences[SelectedOccurence];
            int lastOccurenceEnd = SelectedRange.End;
            int documentEnd = _document.Content.End;
            Range toEnd =
                SelectedOccurence == 0
                    ? _document.Content
                    : _document.Range(lastOccurenceEnd, documentEnd);
            var finder = toEnd.Find;
            finder.ClearFormatting();
            finder.Text = (contextBefore.Length > 0 ? "[!-]<" : "<") + occurenceSelected + ">";
            finder.Forward = true;
            finder.MatchCase = true;
            finder.MatchWildcards = true;
            //finder.MatchWholeWord = true;
            if (finder.Execute())
            {
                toEnd = _document.Range(
                    toEnd.Start + (contextBefore.Length > 0 ? 1 : 0),
                    toEnd.End
                );
            }
            if (!finder.Execute())
            {
                // Fallback : redo the search without whole word
                toEnd =
                    SelectedOccurence == 0
                        ? _document.Content
                        : _document.Range(lastOccurenceEnd, _document.Content.End);
                finder = toEnd.Find;
                finder.ClearFormatting();
                finder.Text =
                    (
                        contextBefore.Length > 0
                            ? contextBefore.Substring(contextBefore.Length - 1)
                            : ""
                    ) + occurenceSelected;
                finder.Forward = true;
                finder.MatchCase = true;
                finder.MatchWholeWord = false;
                if (finder.Execute())
                {
                    toEnd = _document.Range(
                        toEnd.Start + (contextBefore.Length > 0 ? 1 : 0),
                        toEnd.End
                    );
                }
                else
                {
                    // Probleme a remonter
                }
            }
            else
            {
                SelectedRange = toEnd;
            }

            int selectionStart = SelectedRange.Start;
            int selectionEnd = SelectedRange.End;
            return toEnd;
        }

        /// <summary>
        /// Selectionne la prochaine occurence de mot à traiter (prochaine occurence en statut inconnu) dans le document
        /// </summary>
        /// <param name="reselectionOccurenceCourante">Autorise la reselection de l'occurence courante si son statut est inconnu</param>
        /// <returns></returns>
        public MSWord.Range ProchaineOccurenceATraiter(bool reselectionOccurenceCourante = false)
        {
            if (
                reselectionOccurenceCourante
                && (SelectedOccurenceStatut == Statut.INCONNU || !SelectedOccurenceEstTraitee) /*|| EstTerminer()*/
            )
            {
                return SelectedRange;
            }
            int counter = SelectedOccurence;
            while (counter < WorkingDictionnary.Occurences.Count)
            {
                Range selected = ProchaineOccurence();
                if (SelectedOccurenceStatut == Statut.INCONNU || !SelectedOccurenceEstTraitee)
                {
                    return selected;
                }
                counter++;
            }
            ;
            // worst case, on retourne la dernière occurence
            return SelectedRange;
        }

        /// <summary>
        /// Sélectionne l'occurence de mot précédente
        /// </summary>
        /// <returns></returns>
        public MSWord.Range PrecedenteOccurence()
        {
            SelectedOccurence =
                (WorkingDictionnary.Occurences.Count + SelectedOccurence - 1)
                % WorkingDictionnary.Occurences.Count;
            string occurenceSelected = WorkingDictionnary.Occurences[SelectedOccurence];
            string wordSelected = occurenceSelected.ToLower();
            string contextBefore = WorkingDictionnary.ContextesAvantOccurences[SelectedOccurence];
            int wordOccurenceIndex = WorkingDictionnary.CarteMotOccurences[wordSelected].IndexOf(
                SelectedOccurence
            );
            Range toBegining =
                SelectedOccurence == WorkingDictionnary.Occurences.Count - 1
                    ? _document.StoryRanges[WdStoryType.wdMainTextStory]
                    : _document.Range(
                        _document.StoryRanges[WdStoryType.wdMainTextStory].Start,
                        SelectedRange.Start
                    );
            var finder = toBegining.Find;
            finder.ClearFormatting();
            finder.Text = (contextBefore.Length > 0 ? "[!-]<" : "<") + occurenceSelected + ">";
            finder.Forward = true;
            finder.MatchCase = true;
            finder.MatchWildcards = true;
            //finder.MatchWholeWord = true;
            if (finder.Execute())
            {
                toBegining = _document.Range(
                    toBegining.Start + (contextBefore.Length > 0 ? 1 : 0),
                    toBegining.End
                );
            }
            SelectedRange = toBegining;
            return toBegining;
        }

        #endregion

        #region Interface IProtection
        /// <summary>
        /// Resélectionne si nécessaire et applique un statut sur une occurence dans le texte
        /// </summary>
        /// <param name="index">index de l'occurence dans la liste des occurences</param>
        /// <param name="statut">Statut a appliqué</param>
        public void AppliquerStatutSurOccurence(int index, Statut statut)
        {
            Range current =
                index != SelectedOccurence ? SelectionnerOccurence(index) : SelectedRange;
            SelectedOccurenceStatut = statut;
            SelectedOccurenceEstTraitee = statut != Statut.INCONNU;
            if (statut == Statut.PROTEGE)
            {
                if (!EstProteger(_document, current))
                {
                    // Ajouter ProtectionCodeLength a la position de l'occurence et  de toutes les occurences qui suivent
                    for (int i = index; i < WorkingDictionnary.Occurences.Count; i++)
                    {
                        WorkingDictionnary.PositionsOccurences[i] += PROTECTION_CODE.Length;
                    }
                    Proteger(current);
                }
            }
            else
            {
                if (EstProteger(_document, current))
                {
                    // Enlever ProtectionCodeLength a la position de l'occurence et toutes les occurences qui suivent
                    for (int i = index; i < WorkingDictionnary.Occurences.Count; i++)
                    {
                        WorkingDictionnary.PositionsOccurences[i] -= PROTECTION_CODE.Length;
                    }
                    Abreger(current);
                }
            }
        }

        /// <summary>
        /// Mise a jour du texte en mémoire pour eviter la reanalyse (par exemple apres application des statuts)
        /// </summary>
        public void RechargerTexteEnMemoire()
        {
            this.TexteEnMemoire = DocumentMainRange;
        }

        /// <summary>
        /// Sélectionner l'occurence en mémoire et remet le focus dessus dans word
        /// </summary>
        /// <param name="index"></param>
        public void AfficherOccurence(int index)
        {
            if (index != SelectedOccurence)
            {
                SelectionnerOccurence(index);
            }
            SelectedRange.Select();
        }
        #endregion
    }
}
