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
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;
using static NHibernate.Engine.TransactionHelper;
using System.Diagnostics;
using static System.Net.Mime.MediaTypeNames;
using System.Windows.Controls;
using static System.Windows.Forms.AxHost;

namespace fr.avh.braille.addin
{
    /// <summary>
    ///
    /// </summary>
    public class ProtectionWord : IProtection
    {
        private static readonly string MIN = "a-zàáâãäåæçèéêëìíîïðñòóôõöøùúûüý";
        private static readonly string MAJ = "A-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝ";
        private static readonly string NUM = "0-9";
        private static readonly string ALPHANUM = $"{MIN}{MAJ}{NUM}_"; // == a-zàáâãäåæçèéêëìíîïðñòóôõöøùúûüýA-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝ0-9_
        
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
        /// Dictionnaire de travail pour conserver les actions fait sur les mots <br/>
        /// - 0 : non-traité<br/>
        /// - 1 : abrégé globalement<br/>
        /// - 2 : protégé globalement<br/>
        /// - 3 : ambigu, traitement au cas par cas par occurence dans le titre<br/>
        /// </summary>
        public DictionnaireDeTravail DonneesTraitement = null;

        public string WorkingDictionnaryPath = null;


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
                DonneesTraitement.DerniereOccurenceSelectionnee = value;
                SelectionChanged?.Invoke(value);
            }
        }
        public Statut SelectedOccurenceStatut
        {
            get => DonneesTraitement.StatutsOccurences[_selectedOccurence];
            set => DonneesTraitement.StatutsOccurences[_selectedOccurence] = value;
        }

        public bool SelectedOccurenceEstTraitee
        {
            get => DonneesTraitement.StatutsAppliquer[_selectedOccurence];
            set => DonneesTraitement.StatutsAppliquer[_selectedOccurence] = value;
        }

        public int SelectedWordIndex
        {
            get
            {
                string wordSelected = DonneesTraitement.Occurences[SelectedOccurence].ToLower();
                return DonneesTraitement.CarteMotOccurences.Keys.ToList().IndexOf(wordSelected);
            }
        }

        public string SelectedWord
        {
            get { return DonneesTraitement.Occurences[SelectedOccurence].ToLower(); }
            set
            {
                if (DonneesTraitement.CarteMotOccurences.ContainsKey(value.ToLower().Trim())) {
                    SelectedOccurence = DonneesTraitement.CarteMotOccurences[value.ToLower().Trim()][0];
                } else {
                    throw new Exception($"Le mot {value} n'est pas dans le dictionnaire de travail");
                }
            }
        }

        public int SelectedWordOccurenceIndex
        {
            get
            {
                string wordSelected = DonneesTraitement.Occurences[SelectedOccurence].ToLower();
                return DonneesTraitement.CarteMotOccurences[wordSelected].IndexOf(
                    SelectedOccurence
                );
            }
        }

        public int SelectedWordOccurenceCount
        {
            get
            {
                string wordSelected = DonneesTraitement.Occurences[SelectedOccurence].ToLower();
                return DonneesTraitement.CarteMotOccurences[wordSelected].Count;
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
            string dicName = Path.GetFileNameWithoutExtension(filename) + ".json";
            WorkingDictionnaryPath = Path.Combine(fr.avh.braille.dictionnaire.Globals.AppData.FullName, dicName);
            try {
                if (Directory.Exists(Path.GetDirectoryName(filename))) {
            WorkingDictionnaryPath = Path.Combine(Path.GetDirectoryName(filename), dicName);
                }
            } catch(Exception e) {
               // Stocker le dictionnaire dans le dossier AppData
            }

            if (document != null)
            {
                _document = document;
                AnalyserDocument();
                
            }
            else
            {
                throw new Exception("Impossible d'analyser le document sélectionner avec word");
            }
        }

        /// <summary>
        /// Analyse le document word et enregistre les mots identifiés dans le dictionnaire de travail <br/>
        /// </summary>
        /// <exception cref="Exception"></exception>
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

                }
                DonneesTraitement = new DictionnaireDeTravail(
                    Path.GetFileNameWithoutExtension(WorkingDictionnaryPath)
                );

                info?.Invoke(
                    "Préparation du document pour analyse ...",
                    new Tuple<int, int>(2, baseStepNumber)
                );

                info?.Invoke(
                    $" Copie du texte en mémoire...",
                    new Tuple<int, int>(3, baseStepNumber)
                );
                ChargerTexteEnMemoire();

                info?.Invoke(
                    $"Analyse et extractions des mots à traiter ...",
                    new Tuple<int, int>(4, baseStepNumber)
                );
                AnalyseDuTexteEnMemoire();


                info?.Invoke(
                    $"Récupérations des statistiques des mots détectés ...",
                    new Tuple<int, int>(6, baseStepNumber)
                );
                using (var session = BaseSQlite.CreateSessionFactory(info, error).OpenSession())
                {
                    alreadyInDB = session
                        .QueryOver<Mot>()
                        .WhereRestrictionOn(m => m.Texte)
                        .IsIn(DonneesTraitement.CarteMotOccurences.Keys.ToList())
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
                                    && DonneesTraitement.StatutMot(item.Texte) == Statut.INCONNU
                                )
                                {
                                    Statut selected =
                                        item.Abreviations > item.Protections
                                            ? Statut.ABREGE
                                            : Statut.PROTEGE;
                                    // Ne mettre a jour que les mots non traité au cas ou on serait sur une reprise de traitement
                                    foreach (
                                        var index in DonneesTraitement.CarteMotOccurences[
                                            item.Texte
                                        ]
                                    )
                                    {
                                        if (
                                            DonneesTraitement.StatutsOccurences[index]
                                            == Statut.INCONNU
                                        )
                                            DonneesTraitement.StatutsOccurences[index] = selected;
                                    }
                                }
                            }
                        }
                    }
                }
                

                bool reapplyDecisions = false;
                // Rechargement du précédent dictionnaire
                if (existingDictionnary != null)
                {
                    info?.Invoke(
                        $"Rechargement des décisions enregistrés dans le précédent dictionnaire du document...",
                        new Tuple<int, int>(6, baseStepNumber)
                    );
                    DonneesTraitement.RechargerDecisionDe(existingDictionnary);
                    var res = MessageBox.Show("Souhaitez-vous que les décisions du précédent dictionnaire sur le document soient réappliquer ?" +
                        "\r\nCeci n'est pas obligatoire si le document a été sauvegarder après le précédent traitement du braille.", "Chargement réussi", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    reapplyDecisions = res == MessageBoxResult.Yes;
                }
                
                if (DonneesTraitement.DernierMotSelectionne != null 
                    && DonneesTraitement.CarteMotOccurences.ContainsKey(DonneesTraitement.DernierMotSelectionne)
                ) {
                    string mot = DonneesTraitement.DernierMotSelectionne;
                    info?.Invoke(
                            $"Relancement du traitement au mot {mot}",
                            new Tuple<int, int>(9, baseStepNumber)
                        );
                    List<string> keys = DonneesTraitement.CarteMotOccurences.Keys.ToList();
                    int wordCountToReprocess = DonneesTraitement.EstTerminer
                        ? keys.Count
                        : keys.IndexOf(mot.ToLower());
                    if (wordCountToReprocess > 0) {
                        info?.Invoke(
                            $"Reprise des décisions précédentes sur {wordCountToReprocess} mots",
                            new Tuple<int, int>(9, baseStepNumber)
                        );
                        for (int i = 0; i < wordCountToReprocess; i++) {
                            
                            string wordToCheck = keys[i];
                            List<int> toTreat = DonneesTraitement.CarteMotOccurences[
                                wordToCheck
                            ]
                                .Where(
                                    o =>
                                        DonneesTraitement.StatutsOccurences[o] == Statut.ABREGE
                                        || DonneesTraitement.StatutsOccurences[o]
                                            == Statut.PROTEGE
                                )
                                .ToList();
                            if (reapplyDecisions) {
                                info?.Invoke(
                                    $"- Reapplication des décisions sur {wordToCheck} : {DonneesTraitement.CarteMotOccurences[wordToCheck].Count} occurences, {toTreat.Count} décisions à appliquer",
                                    new Tuple<int, int>(9, baseStepNumber)
                                );

                                // Parcours des occurences avec le parcours optimisé
                                foreach (int occurence in toTreat) {
                                    try {
                                        SelectionnerOccurence(occurence);
                                        AppliquerStatutSurOccurence(
                                            SelectedOccurence,
                                            SelectedOccurenceStatut
                                        );
                                    }
                                    catch (Exception e) {
                                        MessageBox.Show(
                                            $"L'erreur suivante s'est produite lors du contrôle de {wordToCheck}\r\n"
                                                + e.Message
                                        );
                                    }
                                }
                            } else {
                                foreach (int occurence in toTreat) {
                                    try {
                                        DonneesTraitement.AppliquerStatut(occurence, DonneesTraitement.StatutsOccurences[occurence]);
                                    }
                                    catch (Exception e) {
                                        MessageBox.Show(
                                            $"L'erreur suivante s'est produite lors du contrôle de {wordToCheck}\r\n"
                                                + e.Message
                                        );
                                    }
                                }
                            }
                            
                        }
                    }
                    try {
                        SelectionnerOccurenceMot(mot);
                    }
                    catch {
                        SelectionnerOccurence(0);
                    }
                } else {
                    SelectionnerOccurence(0);
                }


                info?.Invoke(
                    $"Sauvegarde du dictionnaire des mots détectés",
                    new Tuple<int, int>(8, baseStepNumber)
                );
                this.Save();
                info?.Invoke($"Document prêt pour analyse", new Tuple<int, int>(9, baseStepNumber));
                ChargerTexteEnMemoire();
            }
            else
            {
                throw new Exception("Impossible d'analyser le document sélectionner avec word");
            }
        }

        /// <summary>
        /// Analyse du texte en mémoire
        /// </summary>
        public void AnalyseDuTexteEnMemoire()
        {
            var analyse = Abreviation.AnalyserTexte(TexteEnMemoire, info).Result.OrderBy(i => i.Key).Select(i => i.Value).ToList();
            info?.Invoke(
                $"Recherche puis protection des mots contenant des chiffres ou hors lexique français..."
            );
            int offset = 0;
            for (int i = 0; i < analyse.Count; i++) {
                if (!analyse[i].EstDejaProteger
                    && (analyse[i].ContientDesChiffres
                        || (analyse[i].EstAbregeable && !analyse[i].EstFrançaisAbregeable)
                    )
                ) {
                    Range selection = _document.Range(offset + analyse[i].Index, offset + analyse[i].Index + analyse[i].Mot.Length);
                    Range test = Proteger(selection);
                    offset += DictionnaireDeTravail.PROTECTION_CODE.Length;
                    //offset += test.Words.Count > 1 
                    //    ? DictionnaireDeTravail.PROTECTION_CODE_G1.Length + DictionnaireDeTravail.PROTECTION_CODE_G2.Length
                    //    : DictionnaireDeTravail.PROTECTION_CODE.Length;
                } else if (analyse[i].EstAbregeable
                    && (
                        (analyse[i].EstFrançaisAbregeable && analyse[i].ContientDesMajuscules)
                        || analyse[i].EstAmbigu
                    )
                ) {
                    // Reprise des contextes
                    string foundWord = analyse[i].Mot;
                    int indexInAnalysis = analyse[i].Index;

                    bool isAlreadyProtected = analyse[i].EstDejaProteger;
                    // Récupération du contexte autour de l'occurence
                    int indexBefore = Math.Max(0, indexInAnalysis - 50);
                    int indexAfter = Math.Min(
                        TexteEnMemoire.Length - 1,
                        indexInAnalysis + foundWord.Length + 50
                    );
                    string contextBefore = TexteEnMemoire.Substring(
                        indexBefore,
                        indexInAnalysis - indexBefore
                    );
                    string contextAfter = "";

                    contextAfter = TexteEnMemoire.Substring(
                        indexInAnalysis + foundWord.Length,
                        indexAfter - foundWord.Length - indexInAnalysis
                    );
                    // On ajoute cette occurence de mot si elle n'est pas déjà dans le dictionnaire de traitement
                    if (
                        !(
                            DonneesTraitement.CarteMotOccurences.ContainsKey(
                                foundWord.ToLower().Trim()
                            )
                            && DonneesTraitement.CarteMotOccurences[
                                foundWord.ToLower().Trim()
                            ].FindIndex(
                                occurence =>
                                    DonneesTraitement.PositionsOccurences[occurence] == (offset + indexInAnalysis)
                            ) >= 0
                        )
                    ) {
                        DonneesTraitement.AjouterOccurence(
                            foundWord,
                            isAlreadyProtected ? Statut.PROTEGE : Statut.INCONNU,
                            contextBefore,
                            contextAfter,
                            offset + indexInAnalysis,
                            isAlreadyProtected
                        );
                    }
                }
                info?.Invoke(
                    $"",
                    new Tuple<int, int>(i + 1, analyse.Count)
                );
            }
            ChargerTexteEnMemoire();
            DonneesTraitement.ReorderOccurences();
            DonneesTraitement.CalculCartographieMots();
        }


        /// <summary>
        /// Applique les décisions sur tout ou partie des occurences détectés
        /// </summary>
        /// <param name="surToutesOccurence">Si vrai, reapplique les décisions sur touts </param>
        /// <param name="start">Occurence a partir de laquelle appliquer les décision (incluse)</param>
        /// <param name="end">Derniere occurence a traité (incluse), -1 = pas de limite</param>
        public void AppliquerDecisions(bool surToutesOccurence = false, int start = 0, int end = -1)
        {
            int _start = Math.Min(DonneesTraitement.Occurences.Count - 1, Math.Max(0, start));
            int _end = end < 0 ? DonneesTraitement.Occurences.Count - 1 : Math.Min(DonneesTraitement.Occurences.Count - 1, Math.Max(0, start));
            if(_start > _end) {
                int t = _end;
                _end = _start;
                _start = t;
            }
            for (int i = _start; i <= _end; i++) {
                string wordKey = DonneesTraitement.Occurences[i].ToLower();
                Statut selected = DonneesTraitement.StatutsOccurences[i];
                // N'appliquer les décisions
                if (surToutesOccurence || !DonneesTraitement.StatutsAppliquer[i]) {
                    switch (selected) {
                        case Statut.INCONNU:
                            if (DonneesTraitement.CarteMotStatut.ContainsKey(wordKey)
                                && DonneesTraitement.CarteMotStatut[wordKey] != Statut.INCONNU
                                && DonneesTraitement.StatutsAppliquer[i] == false
                            ) {
                                AppliquerStatutSurOccurence(i, DonneesTraitement.CarteMotStatut[wordKey]);
                            }
                            break;
                        default:
                            AppliquerStatutSurOccurence(i, selected);
                            break;
                    }
                }
            }
        }

        public void ReanalyserDocumentSiModification()
        {
            if (TexteEnMemoire != DocumentMainRange)
            {
                info?.Invoke(
                    $"Modification de contenu détecter : réanalyse du document ..."
                );
                DictionnaireDeTravail actuel = DonneesTraitement.Clone();
                DonneesTraitement = new DictionnaireDeTravail(
                    Path.GetFileNameWithoutExtension(WorkingDictionnaryPath)
                );

                ChargerTexteEnMemoire();
                
                AnalyseDuTexteEnMemoire();
                
                info?.Invoke($"Rechargement des décisions ...");
                List<string> motsAControler = new List<string>();
                for (int i = 0; i < actuel.CarteMotOccurences.Keys.Count; i++) {
                    string item = actuel.CarteMotOccurences.Keys.ElementAt(i);
                    info?.Invoke("", new Tuple<int, int>(i, actuel.CarteMotOccurences.Keys.Count));
                    // Appliquer le statut de l'occurence précédente a la nouvelle occurence
                    for (
                        int indexInMap = 0;
                        indexInMap < actuel.CarteMotOccurences[item].Count;
                        indexInMap++
                    ) {
                        int occurence = actuel.CarteMotOccurences[item][indexInMap];
                        if (indexInMap < DonneesTraitement.CarteMotOccurences[item].Count) {
                            int newOccurence = DonneesTraitement.CarteMotOccurences[item][
                                indexInMap
                            ];
                            if (
                                actuel.Occurences[occurence]
                                != DonneesTraitement.Occurences[newOccurence]
                            ) {
                                info?.Invoke(
                                    $"Attention !! l'occurence {indexInMap + 1} du mot {item} a changer d'écriture\r\n"
                                        + $"(avant: {actuel.Occurences[occurence]} - apres {DonneesTraitement.Occurences[newOccurence]})"
                                );
                                if (!motsAControler.Contains(item)) {
                                    motsAControler.Add(item);
                                }
                            }
                            DonneesTraitement.StatutsOccurences[
                                DonneesTraitement.CarteMotOccurences[item][indexInMap]
                            ] = actuel.StatutsOccurences[occurence];
                        } else {
                            info?.Invoke(
                                $"Attention !! le nombre d'occurence du mot {item} a changé"
                            );
                            if (!motsAControler.Contains(item)) {
                                motsAControler.Add(item);
                            }
                            break;
                        }
                    }
                }
                this.Save();
                if (motsAControler.Count > 0) {
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
                            DonneesTraitement.CarteMotOccurences.ContainsKey(word)
                            && DonneesTraitement.CarteMotOccurences[word].FindIndex(
                                occurence =>
                                    DonneesTraitement.PositionsOccurences[occurence] == indexInText
                            ) >= 0
                        )
                    )
                    {
                        DonneesTraitement.AjouterOccurence(
                            foundWord,
                            isAlreadyProtected ? Statut.PROTEGE : Statut.INCONNU,
                            contextBefore,
                            contextAfter,
                            indexInText,
                            isAlreadyProtected
                        );
                    }
                }
                if (refreshDictionnary)
                {
                    DonneesTraitement.ReorderOccurences();
                    DonneesTraitement.CalculCartographieMots();
                }
                if (statut != Statut.INCONNU)
                {
                    if (selectedOccurenceIndexInMap >= 0)
                    {
                        // on a sélectionner une occurence en particulier
                        DonneesTraitement.StatutsOccurences[
                            DonneesTraitement.CarteMotOccurences[word][selectedOccurenceIndexInMap]
                        ] = statut;
                    }
                    else
                    {
                        // on a pas sélectionner d'occurence, on applique a toutes les occurences
                        foreach (var index in DonneesTraitement.CarteMotOccurences[word])
                        {
                            DonneesTraitement.StatutsOccurences[index] = statut;
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
        [Obsolete("Peut provoquer une erreur StackOverFlow sur un nombre de mot trop importants")]
        public void AjouterListeMotsAuTraitementOld(
            List<string> mots,
            Statut statut = Statut.INCONNU,
            bool refreshDictionnary = true,
            bool alerteSiNonTrouver = false
        )
        {
            if (mots.Count == 0)
                return;

            string word = mots[0].Trim().ToLower();
            for (int i = 1; i < mots.Count; i++) {
                word += "|" + mots[i].Trim().ToLower();
            }
            Regex toLook = SearchWord(
                word,
                RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.Singleline
            );
            MatchCollection check = toLook.Matches(TexteEnMemoire);
            if (check.Count > 0) {
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
                            DonneesTraitement.CarteMotOccurences.ContainsKey(
                                foundWord.ToLower().Trim()
                            )
                            && DonneesTraitement.CarteMotOccurences[
                                foundWord.ToLower().Trim()
                            ].FindIndex(
                                occurence =>
                                    DonneesTraitement.PositionsOccurences[occurence] == indexInText
                            ) >= 0
                        )
                    ) {
                        DonneesTraitement.AjouterOccurence(
                            foundWord,
                            isAlreadyProtected ? Statut.PROTEGE : Statut.INCONNU,
                            contextBefore,
                            contextAfter,
                            indexInText
                        );
                        if (isAlreadyProtected) {
                            DonneesTraitement.StatutsAppliquer[
                                DonneesTraitement.Occurences.Count - 1
                            ] = true;
                        }
                    }
                }
                
            }
            info?.Invoke($"Reconstruction de la cartographie...");
            if (refreshDictionnary) {
                DonneesTraitement.ReorderOccurences();
                DonneesTraitement.CalculCartographieMots();
            }
            if (statut != Statut.INCONNU) {
                foreach (var index in DonneesTraitement.CarteMotOccurences[word]) {
                    DonneesTraitement.StatutsOccurences[index] = statut;
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

            // Suppression de tous les hyperliens créer parfois automatiquement par word ...
            foreach (MSWord.Hyperlink link in _document.Hyperlinks) {
                link.Delete();
            }
            // Loop through fields in reverse to safely delete
            for (int i = _document.Fields.Count; i >= 1; i--) {
                MSWord.Field field = _document.Fields[i];
                if (field.Type == MSWord.WdFieldType.wdFieldHyperlink) {
                    field.Unlink(); // Replaces hyperlink with its display text
                }
            }

            List<Task<List<Tuple<string, Statut, string, string, int>>>> foundResults = new List<Task<List<Tuple<string, Statut, string, string, int>>>>();

            foreach (var mot in mots) {
                foundResults.Add(Task.Run(() =>
                {
                    string word = mot.Trim().ToLower();
                    Regex toLook = SearchWord(
                        word,
                            RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.Singleline
                        );
                    MatchCollection check = toLook.Matches(TexteEnMemoire);
                    List<Tuple<string, Statut, string, string, int>> temp = new List<Tuple<string, Statut, string, string, int>>();

                    if (check.Count > 0) {
                        foreach (Match item in check) {
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
                            temp.Add(
                                new Tuple<string, Statut, string, string, int>(foundWord,
                                    isAlreadyProtected ? Statut.PROTEGE : Statut.INCONNU,
                                    contextBefore,
                                    contextAfter,
                                    indexInText
                                )
                            );
                        }
                    }
                    return temp;
                }));
                //string word = mots[0].Trim().ToLower();
                //for (int i = 1; i < mots.Count; i++) {
                //    word += "|" + mots[i].Trim().ToLower();
                //}

            }
            List<Tuple<string, Statut, string, string, int>>[] results = null;
            try {
                results = Task.WhenAll(foundResults).Result;
            }
            catch (AggregateException ae) {
                error?.Invoke(
                    new Exception($"Erreur lors de la recherche des mots dans le document", ae)
                );
                return;
            }
            foreach (var res in results) {
                foreach (var res2 in res) {
                    // On ajoute cette occurence de mot si elle n'est pas déjà dans le dictionnaire de traitement
                    if (
                        !(
                            DonneesTraitement.CarteMotOccurences.ContainsKey(
                                res2.Item1.ToLower().Trim()
                            )
                            && DonneesTraitement.CarteMotOccurences[
                                res2.Item1.ToLower().Trim()
                            ].FindIndex(
                                occurence =>
                                    DonneesTraitement.PositionsOccurences[occurence] == res2.Item5
                            ) >= 0
                        )
                    ) {
                        DonneesTraitement.AjouterOccurence(
                            res2.Item1,
                            res2.Item2,
                            res2.Item3,
                            res2.Item4,
                            res2.Item5
                        );
                        if (res2.Item2 == Statut.PROTEGE) {
                            DonneesTraitement.StatutsAppliquer[
                                DonneesTraitement.Occurences.Count - 1
                            ] = true;
                        }
                    }

                }
            }
            info?.Invoke($"Reconstruction de la cartographie après recherche de {mots.Count} mots ...");
            if (refreshDictionnary) {
                DonneesTraitement.ReorderOccurences();
                DonneesTraitement.CalculCartographieMots();
            }
            foreach(var word in mots) {
                var _word = word.Trim().ToLower();
                if (statut != Statut.INCONNU && DonneesTraitement.CarteMotOccurences.ContainsKey(word)) {
                    foreach (var index in DonneesTraitement.CarteMotOccurences[word]) {
                        DonneesTraitement.StatutsOccurences[index] = statut;
                    }
                } else if (alerteSiNonTrouver && !DonneesTraitement.CarteMotOccurences.ContainsKey(word)) {
                    info?.Invoke($"Attention : {word} non retrouvé");
                }
            }
            info?.Invoke($"Reconstruction de la cartographie finis");


        }

        public void Save()
        {
            this.DonneesTraitement.DernierMotSelectionne = SelectedWord;
            this.DonneesTraitement.SaveJSON(
                new DirectoryInfo(Path.GetDirectoryName(WorkingDictionnaryPath))
            );
        }

        public delegate void OnProtectorProgress(
            string message = null,
            Tuple<int, int> progress = null
        );

        public void AppliquerStatutsSurDocument(OnProtectorProgress callback = null)
        {
            this.ReanalyserDocumentSiModification();

            callback?.Invoke(
                $"Application des statuts sur les mots du documents...",
                new Tuple<int, int>(0, DonneesTraitement.Occurences.Count)
            );
            for (int i = 0; i < DonneesTraitement.Occurences.Count; i++)
            {
                AppliquerStatutSurOccurence(i, DonneesTraitement.StatutsOccurences[i]);
                callback?.Invoke(
                    progress: new Tuple<int, int>(i + 1, DonneesTraitement.Occurences.Count)
                );
            }
            //this.AnalyserDocument();
        }


        #region Actions

        public void ProtegerOccurence()
        {
            var occurrences = DonneesTraitement.CarteMotOccurences[SelectedWord];
            foreach (var index in occurrences)
            {
                DonneesTraitement.StatutsOccurences[index] = Statut.PROTEGE;
            }
            SelectedRange = Proteger(SelectedRange);
            /*WorkingDictionnary.StatutsOccurences[SelectedOccurence] = Statut.PROTEGE;
            SelectedRange = Proteger(SelectedRange);*/
        }

        public void AbregerOccurence()
        {
            var occurrences = DonneesTraitement.CarteMotOccurences[SelectedWord];
            foreach (var index in occurrences)
            {
                DonneesTraitement.StatutsOccurences[index] = Statut.ABREGE;
            }
            SelectedRange = Abreger(SelectedRange);
            /*WorkingDictionnary.StatutsOccurences[SelectedOccurence] = Statut.ABREGE;
            SelectedRange = Proteger(SelectedRange);*/
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

            string textBefore = current.Range(0, wordRange.Start).Text;

            if (!DansBlockProteger(textBefore, wordRange.Text))
            {
                wordRange.InsertBefore(DictionnaireDeTravail.PROTECTION_CODE_G1);
                wordRange.InsertAfter(DictionnaireDeTravail.PROTECTION_CODE_G2);
                wordRange = current.Range(
                    wordRange.Start + DictionnaireDeTravail.PROTECTION_CODE_G1.Length,
                    wordRange.End - DictionnaireDeTravail.PROTECTION_CODE_G2.Length
                );
                wordRange.Select();
            }
            return wordRange;
        }

        public void ProtegerBloc(MSWord.Range wordRange)
        {
            
            if (DansBlockProteger(DocumentMainRange.Substring(0,wordRange.Start), wordRange.Text)
                || DansBlockProteger(DocumentMainRange.Substring(0, wordRange.End))) {

                // Si le debut est dans un block et pas la fin
                // Récupérer le départ du block et fusionner avec le 
                // Si la fin est dans un block
                // Récupérer la fin du block et fusionner avec le début du block
                
                return;
            }

            
            

            // passer en ignorer toutes les occurences situés dans l'intervale
            for (int i = 0; i < DonneesTraitement.Occurences.Count; i++) {
                if (DonneesTraitement.PositionsOccurences[i] >= wordRange.Start && DonneesTraitement.PositionsOccurences[i] <= wordRange.End) {
                    AppliquerStatutSurOccurence(i, Statut.IGNORE);
                }
            }
            ProtegerBloc(_document, wordRange);
            

        }


        /// <summary>
        /// Protege les phrases identifiées comme étrangères dans le document.
        /// Si des mots étrangers isolés sont détectés, ils sont retourné dans une liste de tuple (position, mot)
        /// Desactiver car beaucoup trop lent
        /// </summary>
        /// <returns></returns>
        //public List<Tuple<int, string>> TraiterPhrasesEtrangere()
        //{
        //    List<Tuple<int, string>> motsPossiblementEtrangers = new List<Tuple<int, string>>();
        //    List<MSWord.Range> motsEtrangesConsecutifs = new List<MSWord.Range>();
        //    bool isInCode = false;
        //    info?.Invoke(
        //        $"Protection des phrases en langues étrangères ..."
        //    );
        //    try
        //    {
        //        foreach (MSWord.Range w in _document.Words)
        //        {
        //            w.DetectLanguage();
        //            string text = w.Text;
        //            isInCode = (text.StartsWith("[[*") || text.EndsWith("[[*")) ? true : (text.EndsWith("*]]") ? false : isInCode);
        //            if (w.LanguageID != MSWord.WdLanguageID.wdFrench)
        //            {
        //                motsEtrangesConsecutifs.Add(w);
        //            }
        //            else
        //            {
        //                if (motsEtrangesConsecutifs.Count > 1)
        //                {
        //                    int startBlock = motsEtrangesConsecutifs.First().Start;
        //                    int endBlock = motsEtrangesConsecutifs.Last().End;
        //                    // Ajouter les codes de protection G1 et G2 autour de la phrase étrangère
        //                    int decalage = 0;
        //                    MSWord.Range phraseRange = _document.Range(startBlock, endBlock);
        //                    info?.Invoke(
        //                        $"Protection du block de texte {phraseRange.Text}"
        //                    );
        //                    if (!phraseRange.Text.StartsWith(DictionnaireDeTravail.PROTECTION_CODE_G1))
        //                    {
        //                        phraseRange.InsertBefore(DictionnaireDeTravail.PROTECTION_CODE_G1);
        //                        endBlock += DictionnaireDeTravail.PROTECTION_CODE_G1.Length;
        //                        decalage += DictionnaireDeTravail.PROTECTION_CODE_G1.Length;
        //                    }
        //                    if (!phraseRange.Text.EndsWith(DictionnaireDeTravail.PROTECTION_CODE_G2))
        //                    {
        //                        phraseRange.InsertAfter(DictionnaireDeTravail.PROTECTION_CODE_G2);
        //                        endBlock += DictionnaireDeTravail.PROTECTION_CODE_G2.Length;
        //                        decalage += DictionnaireDeTravail.PROTECTION_CODE_G2.Length;
        //                    }
        //                    motsEtrangesConsecutifs.Clear();
        //                    WorkingDictionnary.SauvegarderAjoutBlocIntegral(
        //                        startBlock,
        //                        endBlock,
        //                        decalage
        //                    );
        //                } else if (motsEtrangesConsecutifs.Count == 1) {
        //                    motsPossiblementEtrangers.Add(
        //                        new Tuple<int, string>(motsEtrangesConsecutifs.First().Start, motsEtrangesConsecutifs.First().Text.Trim())
        //                    );
        //                    motsEtrangesConsecutifs.Clear();
        //                }
        //            }
        //        }
        //        if (motsEtrangesConsecutifs.Count > 1) {
        //            // Cas d'une phrase en langue étrangère en fin de document
        //            // Ajouter les codes de protection G1 et G2 autour de la phrase étrangère
        //            int startBlock = motsEtrangesConsecutifs.First().Start;
        //            int endBlock = motsEtrangesConsecutifs.Last().End;
        //            int decalage = 0;
        //            MSWord.Range phraseRange = _document.Range(startBlock, endBlock);
        //            info?.Invoke(
        //                $"Protection du block de texte {phraseRange.Text}"
        //            );
        //            if (!phraseRange.Text.StartsWith(DictionnaireDeTravail.PROTECTION_CODE_G1)) {
        //                phraseRange.InsertBefore(DictionnaireDeTravail.PROTECTION_CODE_G1);
        //                endBlock += DictionnaireDeTravail.PROTECTION_CODE_G1.Length;
        //                decalage += DictionnaireDeTravail.PROTECTION_CODE_G1.Length;
        //            }
        //            if (!phraseRange.Text.EndsWith(DictionnaireDeTravail.PROTECTION_CODE_G2)) {
        //                phraseRange.InsertAfter(DictionnaireDeTravail.PROTECTION_CODE_G2);
        //                endBlock += DictionnaireDeTravail.PROTECTION_CODE_G2.Length;
        //                decalage += DictionnaireDeTravail.PROTECTION_CODE_G2.Length;
        //            }
        //            motsEtrangesConsecutifs.Clear();
        //            WorkingDictionnary.SauvegarderAjoutBlocIntegral(
        //                startBlock,
        //                endBlock,
        //                decalage
        //            ); ;
        //        } else if (motsEtrangesConsecutifs.Count == 1) {
        //            // Cas d'un mot en langue étrangère en fin de document
        //            motsPossiblementEtrangers.Add(
        //                new Tuple<int, string>(motsEtrangesConsecutifs.First().Start, motsEtrangesConsecutifs.First().Text.Trim())
        //            );
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        MessageBox.Show(
        //            "Impossible de procéder à la détection des phrases étrangères : \r\n"
        //            + e.Message
        //            + "\r\n"
        //            + "Veuillez installer une langue de vérification supplémentaire (Options > Langue > Langue de création et de vérification > Ajouter)"
        //        );
        //    }
        //    return motsPossiblementEtrangers;
        //}

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
            if (wordRange.Start - DictionnaireDeTravail.PROTECTION_CODE.Length < 0)
            {
                return false;
            }
            
            // Get previous range
            MSWord.Range previous = current.Range(
                wordRange.Start - DictionnaireDeTravail.PROTECTION_CODE.Length,
                wordRange.Start
            );
            string previousText = previous.Text;
            return previousText.Equals(DictionnaireDeTravail.PROTECTION_CODE)
                || wordRange.Text.StartsWith(DictionnaireDeTravail.PROTECTION_CODE);
        }

        public static bool DansBlockProteger(string textBefore, string textEvaluated = "")
        {
            if (textBefore.EndsWith(DictionnaireDeTravail.PROTECTION_CODE_G1) || textEvaluated.TrimStart().StartsWith(DictionnaireDeTravail.PROTECTION_CODE_G1)) {
                return true;
            }
            int lastG1 = textBefore.LastIndexOf(DictionnaireDeTravail.PROTECTION_CODE_G1);
            int lastG2 = textBefore.LastIndexOf(DictionnaireDeTravail.PROTECTION_CODE_G2);
            if (lastG1 > -1) {
                return lastG1 > lastG2;
            }

            return false;
        }

        public bool DansBlockProteger(int position)
        {
            return DansBlockProteger(DocumentMainRange.Substring(0,position));
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
            if (trimmedCount > 0) {
                // Reselectionner le mot sans les espaces de début
                wordRange = current.Range(wordRange.Start + trimmedCount, wordRange.End);
                wordRange.Select();
            }
            //if (DansBlockProteger(wordRange.Text) || DansBlockProteger(current.Range(0, wordRange.Start).Text)) {
            //    // Si le mot est déjà dans un bloc protégé, on ne le protège pas à nouveau
            //    return wordRange;
            //}

            //if (wordRange.Words.Count > 1) {
            //    wordRange.InsertBefore(DictionnaireDeTravail.PROTECTION_CODE_G1);
            //    wordRange.InsertAfter(DictionnaireDeTravail.PROTECTION_CODE_G2);
            //    wordRange = current.Range(
            //        wordRange.Start + DictionnaireDeTravail.PROTECTION_CODE_G1.Length,
            //        wordRange.End - DictionnaireDeTravail.PROTECTION_CODE_G2.Length
            //    );
            //    wordRange.Select();
            //    return wordRange;
            //}
            if (!EstProteger(current, wordRange)) {
                wordRange.InsertBefore(DictionnaireDeTravail.PROTECTION_CODE);
                wordRange = current.Range(wordRange.Start + DictionnaireDeTravail.PROTECTION_CODE.Length, wordRange.End);
                wordRange.Select();
            }
            return wordRange;
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
            if (wordRange.Start - DictionnaireDeTravail.PROTECTION_CODE.Length >= 0)
            {
                MSWord.Range previous = current.Range(
                    wordRange.Start - DictionnaireDeTravail.PROTECTION_CODE.Length,
                    wordRange.Start
                );
                string previousText = previous.Text;
                // First check if text is preceded by protection code
                if (previousText.Equals(DictionnaireDeTravail.PROTECTION_CODE))
                {
                    Range toDelete = current.Range(
                        wordRange.Start - DictionnaireDeTravail.PROTECTION_CODE.Length,
                        wordRange.Start
                    );
                    toDelete.Delete();
                    wordRange = current.Range(toDelete.End, wordRange.End);
                    wordRange.Select();
                }
            }
            // check if selection starts with protection code
            if (trimmedStart.StartsWith(DictionnaireDeTravail.PROTECTION_CODE))
            {
                Range toDelete = current.Range(
                    wordRange.Start,
                    wordRange.Start + DictionnaireDeTravail.PROTECTION_CODE.Length
                );
                toDelete.Delete();
                wordRange = current.Range(toDelete.End, wordRange.End);
                wordRange.Select();
            }
            //string word = wordRange.Text.ToLower().Trim();
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
            if (wordRange.Start - DictionnaireDeTravail.PROTECTION_CODE_G1.Length >= 0 && wordRange.End + DictionnaireDeTravail.PROTECTION_CODE_G2.Length <= current.Content.End)
            {
                MSWord.Range previous = current.Range(
                    wordRange.Start - DictionnaireDeTravail.PROTECTION_CODE_G1.Length,
                    wordRange.Start
                );
                string previousText = previous.Text;
                // si le mot est précédé du code de protection G1
                if (previousText.Equals(DictionnaireDeTravail.PROTECTION_CODE_G1))
                {
                    // Suppression du code de protection G1
                    Range toDelete = current.Range(
                        wordRange.Start - DictionnaireDeTravail.PROTECTION_CODE_G1.Length,
                        wordRange.Start
                    );

                    // Suppression du code de protection G2
                    toDelete.Delete();
                    wordRange = current.Range(toDelete.End, wordRange.End);
                    wordRange.Select();
                }
                MSWord.Range next = current.Range(
                    wordRange.End,
                    wordRange.End + DictionnaireDeTravail.PROTECTION_CODE_G2.Length
                );
                // si le mot est suivi du code de protection G2
                string nextText = next.Text;
                if (nextText.Equals(DictionnaireDeTravail.PROTECTION_CODE_G2))
                {
                    Range toDelete = current.Range(
                        wordRange.End,
                        wordRange.End + DictionnaireDeTravail.PROTECTION_CODE_G2.Length
                    );
                    toDelete.Delete();
                    wordRange = current.Range(wordRange.Start, toDelete.Start); //est ce qu'on peut lier le start et le end pour avoir moin de code
                    wordRange.Select();
                }
            }
            if (trimmedStart.StartsWith(DictionnaireDeTravail.PROTECTION_CODE_G1))
            {
                Range toDelete = current.Range(
                    wordRange.Start,
                    wordRange.Start + DictionnaireDeTravail.PROTECTION_CODE_G1.Length
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
            DonneesTraitement.SetStatut(wordSelected, Statut.PROTEGE);
        }

        /// <summary>
        /// Marquer le mot sélectionné comme étant à abréger dans le traitement
        /// </summary>
        /// <param name="wordRange"></param>
        public void MarquerPourAbreviation(MSWord.Range wordRange)
        {
            string wordSelected = wordRange.Text.ToLower().Trim().ToLower();
            DonneesTraitement.SetStatut(wordSelected, Statut.ABREGE);
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
            // Si toute les occurences ont vu leur statuts attribué et appliquer
            return DonneesTraitement.StatutsOccurences.All(s => s != Statut.INCONNU) && DonneesTraitement.StatutsAppliquer.All(s => s);
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
            string occurenceSelected = DonneesTraitement.Occurences[SelectedOccurence];
            string wordSelected = occurenceSelected.ToLower();
            List<string> keys = DonneesTraitement.CarteMotOccurences.Keys.ToList();
            int newWordKeyIndex = (keys.Count + keys.IndexOf(wordSelected) + 1) % keys.Count;
            return SelectionnerOccurenceMot(keys[newWordKeyIndex], andOccurence);
        }

        public MSWord.Range PrecedentMot(int andOccurence = 0)
        {
            string occurenceSelected = DonneesTraitement.Occurences[SelectedOccurence];
            string wordSelected = occurenceSelected.ToLower();
            List<string> keys = DonneesTraitement.CarteMotOccurences.Keys.ToList();
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

            // listes des occurences a marquer comme étant ignorer, c'est à dire a ne pas afficher dans le traitement
            List<int> wordOccurences = DonneesTraitement.CarteMotOccurences[wordSelected];
            foreach (var occurence in DonneesTraitement.CarteMotOccurences[wordSelected])
            {
                DonneesTraitement.StatutsOccurences[occurence] = Statut.IGNORE;
                DonneesTraitement.StatutsAppliquer[occurence] = true;
            }

            // Sélectionner le mot suivant
            return ProchainMot();
        }

        public MSWord.Range ProchaineOccurenceDuMot()
        {
            string occurenceSelected = DonneesTraitement.Occurences[SelectedOccurence];
            string wordSelected = occurenceSelected.ToLower();
            List<int> occurences = DonneesTraitement.CarteMotOccurences[wordSelected];
            int wordOccurenceIndex =
                (occurences.Count + occurences.IndexOf(SelectedOccurence) + 1) % occurences.Count;
            int wordSelectedOccurence = DonneesTraitement.CarteMotOccurences[
                occurenceSelected.ToLower()
            ][wordOccurenceIndex];
            string searchedWord = DonneesTraitement.Occurences[wordSelectedOccurence];
            string contextBefore = DonneesTraitement.ContextesAvantOccurences[
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
            string occurenceSelected = DonneesTraitement.Occurences[SelectedOccurence];
            string wordSelected = occurenceSelected.ToLower();
            List<int> occurences = DonneesTraitement.CarteMotOccurences[wordSelected];
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
            foreach (var idx in DonneesTraitement.CarteMotOccurences[wordSelected])
            {
                DonneesTraitement.StatutsOccurences[idx] = Statut.PROTEGE;
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
            foreach (var idx in DonneesTraitement.CarteMotOccurences[wordSelected])
            {
                DonneesTraitement.StatutsOccurences[idx] = Statut.ABREGE;
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
            private set { _selectedRange = value; _selectedRange.Select(); }
        }

        public string MotSelectionne { get => SelectedWord; }
        public int OccurenceSelectionnee { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }


        /// <summary>
        /// Selectionne une occurence identifié pour traitement dans le document
        /// </summary>
        /// <param name="newSelectedOccurenceIndex">Indice de l'occurence dans la liste des occurences identifiées pour traitement</param>
        /// <returns></returns>
        public MSWord.Range SelectionnerOccurence(int newSelectedOccurenceIndex)
        {
            if (DonneesTraitement.Occurences.Count == 0)
            {
                throw new Exception("Aucun mot n'a été détecté pour traitement dans le document");
            }
            SelectedOccurence = newSelectedOccurenceIndex;

            int startPosition = DonneesTraitement.PositionsOccurences[SelectedOccurence];
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
            if (DonneesTraitement.Occurences.Count == 0)
            {
                throw new Exception("Aucun mot n'a été détecté pour traitement dans le document");
            }
            if (!DonneesTraitement.CarteMotOccurences.ContainsKey(key.ToLower()))
            {
                throw new Exception($"Le mot {key} n'est pas détecter dans la carte des mots");
            }
            // Passage par les positions précalculés, plus efficace et plus rapide que le passage dans le finder de word (mais plus fragile si le texte est modifié)
            SelectedOccurence = DonneesTraitement.CarteMotOccurences[key.ToLower()][indexCarte];
            int startPosition = DonneesTraitement.PositionsOccurences[SelectedOccurence];
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
                (DonneesTraitement.Occurences.Count + SelectedOccurence + 1)
                % DonneesTraitement.Occurences.Count;
            string occurenceSelected = DonneesTraitement.Occurences[SelectedOccurence];
            string contextBefore = DonneesTraitement.ContextesAvantOccurences[SelectedOccurence];
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
            while (counter < DonneesTraitement.Occurences.Count)
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
                (DonneesTraitement.Occurences.Count + SelectedOccurence - 1)
                % DonneesTraitement.Occurences.Count;
            string occurenceSelected = DonneesTraitement.Occurences[SelectedOccurence];
            string wordSelected = occurenceSelected.ToLower();
            string contextBefore = DonneesTraitement.ContextesAvantOccurences[SelectedOccurence];
            int wordOccurenceIndex = DonneesTraitement.CarteMotOccurences[wordSelected].IndexOf(
                SelectedOccurence
            );
            Range toBegining =
                SelectedOccurence == DonneesTraitement.Occurences.Count - 1
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
            if (statut == Statut.PROTEGE && !EstProteger(_document, current))
            { 
                Proteger(current);
            }
            else if(EstProteger(_document, current))
            {
                Abreger(current);
            }
            DonneesTraitement.AppliquerStatut(index, statut);
        }

        public void AppliquerStatutSurBlock(int start, int end, Statut statut)
        {
            int blocIndexOfStart = DonneesTraitement.ListeBlocsIntegral.FindIndex(b => b.Item1 <= start && start <= b.Item2);
            int blocIndexOfEnd = DonneesTraitement.ListeBlocsIntegral.FindIndex(b => b.Item1 <= end && end <= b.Item2);

            int newBlockStart = blocIndexOfStart > -1 ? DonneesTraitement.ListeBlocsIntegral[blocIndexOfStart].Item1 : start;
            int newBlockEnd = blocIndexOfEnd > -1 ? DonneesTraitement.ListeBlocsIntegral[blocIndexOfEnd].Item2 : end;

            // Ajouter le nouveau block a la fin
            DonneesTraitement.DebutsBlocsIntegral.Add(newBlockStart);
            DonneesTraitement.FinsBlocsIntegral.Add(newBlockEnd);

            int decalage = 0;
            // En parcours arriere, supprimer les blocks précédents situés dans le nouveau block
            for (int i = DonneesTraitement.ListeBlocsIntegral.Count - 2; i >= 0; i--) {
                if (newBlockStart <= DonneesTraitement.ListeBlocsIntegral[i].Item1 && DonneesTraitement.ListeBlocsIntegral[i].Item2 <= newBlockEnd) {
                    DonneesTraitement.DebutsBlocsIntegral.RemoveAt(i);
                    DonneesTraitement.FinsBlocsIntegral.RemoveAt(i);
                    i--;
                }
            }

            Range toProtect = _document.Range(newBlockStart, newBlockEnd);
            if (!toProtect.Text.StartsWith(DictionnaireDeTravail.PROTECTION_CODE_G1)) {
                toProtect.InsertBefore(DictionnaireDeTravail.PROTECTION_CODE_G1);
                decalage += DictionnaireDeTravail.PROTECTION_CODE_G1.Length;
            }
            if (!toProtect.Text.EndsWith(DictionnaireDeTravail.PROTECTION_CODE_G2)) {
                toProtect.InsertAfter(DictionnaireDeTravail.PROTECTION_CODE_G2);
                decalage += DictionnaireDeTravail.PROTECTION_CODE_G2.Length;
            }

            DonneesTraitement.DebutsBlocsIntegral.Sort();
            DonneesTraitement.FinsBlocsIntegral.Sort();
        }

        /// <summary>
        /// Mise a jour du texte en mémoire pour eviter la reanalyse (par exemple apres application des statuts)
        /// </summary>
        public void ChargerTexteEnMemoire()
        {
            // Suppression de tous les hyperliens créer parfois automatiquement par word ...
            foreach (MSWord.Hyperlink link in _document.Hyperlinks) {
                link.Delete();
            }
            // Loop through fields in reverse to safely delete
            for (int i = _document.Fields.Count; i >= 1; i--) {
                MSWord.Field field = _document.Fields[i];
                if (field.Type == MSWord.WdFieldType.wdFieldHyperlink) {
                    field.Unlink(); // Replaces hyperlink with its display text
                }
            }
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
