﻿using MSWord = Microsoft.Office.Interop.Word;

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
using System.Globalization;
using System.Diagnostics;

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

        private int _selectedOccurence = -1;

        /// <summary>
        /// Occurence sélectionné (indice des listes WorkingDictionnary.occurences et occurenceSelectedAction)
        /// </summary>
        public int Occurence
        {
            get => _selectedOccurence;
            private set
            {
                _selectedOccurence = value;
                DonneesTraitement.DerniereOccurenceSelectionnee = value;
                SelectionChanged?.Invoke(value);
            }
        }


        public string MotSelectionne {
            get { return DonneesTraitement.Occurences[Occurence].ToLower(); }
            set
            {
                if (DonneesTraitement.CarteMotOccurences.ContainsKey(value.ToLower().Trim())) {
                    Occurence = DonneesTraitement.CarteMotOccurences[value.ToLower().Trim()][0];
                } else {
                    throw new Exception($"Le mot {value} n'est pas dans le dictionnaire de travail");
                }
            }
        }

        public Statut StatutOccurence
        {
            get => DonneesTraitement.StatutsOccurences[_selectedOccurence];
            set => DonneesTraitement.StatutsOccurences[_selectedOccurence] = value;
        }

        public bool OccurenceEstTraitee
        {
            get => DonneesTraitement.EstTraitee[_selectedOccurence];
            set => DonneesTraitement.EstTraitee[_selectedOccurence] = value;
        }



        public int SelectedWordIndex
        {
            get
            {
                string wordSelected = DonneesTraitement.Occurences[Occurence].ToLower();
                return DonneesTraitement.CarteMotOccurences.Keys.ToList().IndexOf(wordSelected);
            }
        }


        public int SelectedWordOccurenceIndex
        {
            get
            {
                string wordSelected = DonneesTraitement.Occurences[Occurence].ToLower();
                return DonneesTraitement.CarteMotOccurences[wordSelected].IndexOf(
                    Occurence
                );
            }
        }

        public int SelectedWordOccurenceCount
        {
            get
            {
                string wordSelected = DonneesTraitement.Occurences[Occurence].ToLower();
                return DonneesTraitement.CarteMotOccurences[wordSelected].Count;
            }
        }

        public MSWord.Document _document { get; } = null;

        public int PositionCurseur {
            get {
                if (_document != null && _document.Application != null) {
                    return _document.Application.Selection.Start;
                } else return 0;
            }
        }

        public IList<Mot> alreadyInDB { get; private set; }
        public IList<DecisionParDefaut> decisionsParDefaut { get; private set; }

        private string _texteEnMemoire = "";
        public string TexteEnMemoire {
            get { return _texteEnMemoire; }
            private set {
                _texteEnMemoire = string.Copy(value);
            }
        }

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
            Utils.OnErrorCallback error = null,
            DictionnaireDeTravail importer = null
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
                AnalyserDocument(importer);
                
            }
            else
            {
                throw new Exception("Impossible d'analyser le document sélectionner avec word");
            }
        }


        public static IEnumerable<int> AllIndexesOf(string str, string searchstring)
        {
            int minIndex = str.IndexOf(searchstring);
            while (minIndex != -1) {
                yield return minIndex;
                minIndex = str.IndexOf(searchstring, minIndex + 1);
            }
        }

        /// <summary>
        /// Protege un document pour eviter l'écriture
        /// (Pour plus tard quand on passera un task panel)
        /// </summary>
        /// <exception cref="Exception"></exception>
        public void ProtectDocument()
        {
            if (_document != null) {
                // On protège le document
                _document.Protect(
                    WdProtectionType.wdAllowOnlyReading,false, System.String.Empty, false, false
                );
            } else {
                throw new Exception("Impossible de protéger le document, il n'est pas chargé");
            }
        }

        /// <summary>
        /// Reautorise l'écriture dans le document word
        /// </summary>
        /// <exception cref="Exception"></exception>
        public void UnProtectDocument()
        {
            if (_document != null) {
                // On protège le document
                _document.Unprotect(String.Empty);
            } else {
                throw new Exception("Impossible de protéger le document, il n'est pas chargé");
            }
        }


        /// <summary>
        /// Analyse le document word et enregistre les mots identifiés dans le dictionnaire de travail <br/>
        /// </summary>
        /// <exception cref="Exception"></exception>
        public void AnalyserDocument(DictionnaireDeTravail decisionsImportees = null)
        {
            int baseStepNumber = 9;
            if (_document != null)
            {

                // Par défaut, si la preprotection automatique est activé, on reapplique les décisions du précédent dictionnaire
                // bool reapplyDecisions = OptionsComplement.Instance.ActiverPreProtectionAuto; ;
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
                        //if (!reapplyDecisions) {
                        //    var res = MessageBox.Show("Souhaitez-vous que les décisions du précédent dictionnaire sur le document soient réappliquer ?" +
                        //        "\r\nCeci n'est pas obligatoire si le document a été sauvegarder après le précédent traitement du braille.",
                        //        "Chargement réussi",
                        //        MessageBoxButton.YesNo,
                        //        MessageBoxImage.Question,
                        //        // Obligatoire pour forcer la mise en premier plan ici
                        //        MessageBoxResult.No,
                        //        MessageBoxOptions.DefaultDesktopOnly
                        //    );
                        //    reapplyDecisions = res == MessageBoxResult.Yes;
                        //}
                        
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
                    $"{DonneesTraitement.CarteMotOccurences.Count} mots abregeables distincts à traiter ({DonneesTraitement.Occurences.Count} mots dans le document) ...",
                    new Tuple<int, int>(4, baseStepNumber)
                );

                info?.Invoke(
                    $"Récupérations des statistiques pour les {DonneesTraitement.CarteMotOccurences.Count} ...",
                    new Tuple<int, int>(6, baseStepNumber)
                );
                Dictionary<string, Mot> carteMotDB = new Dictionary<string, Mot>();
                Dictionary<string, DecisionParDefaut> carteDecisionDB = new Dictionary<string, DecisionParDefaut>();
                using (var session = BaseSQlite.CreateSessionFactory(info, error).OpenSession()) {

                    alreadyInDB = session
                        .QueryOver<Mot>()
                        .WhereRestrictionOn(m => m.Texte)
                        .IsIn(DonneesTraitement.CarteMotOccurences.Keys.ToList())
                        .List();
                    decisionsParDefaut = session.QueryOver<DecisionParDefaut>()
                        .WhereRestrictionOn(d => d.Mot)
                        .IsIn(DonneesTraitement.CarteMotOccurences.Keys.ToList())
                        .List();
                    carteMotDB = alreadyInDB.ToDictionary(m => m.Texte.ToLower(), m => m);
                    carteDecisionDB = decisionsParDefaut.ToDictionary(m => m.Mot.ToLower(), m => m);
                }
                info?.Invoke(
                    $"Controle/Pré-traitement des mots du dictionnaire existant et/ou de la base de donnée des décisions ...",
                    new Tuple<int, int>(6, baseStepNumber)
                );

                int documentBarrier = 100;
                int n = 0;
                // Essai en mode par mot plutot que par occurence
                foreach (var motEtListeOccurence in DonneesTraitement.CarteMotOccurences) {
                    info?.Invoke(
                        $"",
                        new Tuple<int, int>(++n, DonneesTraitement.CarteMotOccurences.Count)
                    );
                    string wordKey = motEtListeOccurence.Key;
                    var motDBStats = carteMotDB.ContainsKey(wordKey) ? carteMotDB[wordKey] : null;
                    char decisionParDefaut = carteDecisionDB.ContainsKey(wordKey) ? carteDecisionDB[wordKey].Decision : ' ';
                    
                    bool first = true;
                    for (int i = 0; i < motEtListeOccurence.Value.Count; i++) {
                        int indexOccurence = motEtListeOccurence.Value[i];
                        // Priorité au rechargement du précédent dictionnaire (sur les occurences non protégés)
                        if (!DonneesTraitement.EstTraitee[indexOccurence] && existingDictionnary != null) {
                            if (existingDictionnary.CarteMotOccurences.ContainsKey(wordKey)
                                && i < existingDictionnary.CarteMotOccurences[wordKey].Count
                            ) {
                                int indexExistingOccurence = existingDictionnary.CarteMotOccurences[wordKey][i];
                                if (first) {
                                    info?.Invoke(
                                         $" - Application sur {wordKey} du statut {existingDictionnary.StatutsOccurences[indexExistingOccurence]} (dictionnaire précédent, par occurence)",
                                         new Tuple<int, int>(n, DonneesTraitement.CarteMotOccurences.Count)
                                     );
                                    first = false;
                                }
                                // correspondance exacte entre occurence recharge et occurence courante, on réapplique 
                                InitialiserStatutSurOccurence(indexOccurence, existingDictionnary.StatutsOccurences[indexExistingOccurence]);
                                //AppliquerStatutSurOccurence(i, existingDictionnary.StatutsOccurences[i]);
                            } else if (existingDictionnary.CarteMotStatut.ContainsKey(wordKey)) {
                                if (first) {
                                    info?.Invoke(
                                         $" - {wordKey} {existingDictionnary.CarteMotStatut[wordKey]} (dictionnaire précédent)",
                                         new Tuple<int, int>(n, DonneesTraitement.CarteMotOccurences.Count)
                                     );
                                    first = false;
                                }
                                // correspondance partielle, on réapplique le statut du mot
                                InitialiserStatutSurOccurence(indexOccurence, existingDictionnary.CarteMotStatut[wordKey]);
                                //AppliquerStatutSurOccurence(i, existingDictionnary.CarteMotStatut[DonneesTraitement.Occurences[i].ToLower()]);
                            }
                        }
                        // Si pas de dictionnaire existant mais des décisions importées, on applique les décisions importées
                        if (!DonneesTraitement.EstTraitee[indexOccurence] && decisionsImportees != null && decisionsImportees.CarteMotStatut.ContainsKey(wordKey)) {
                            if (first) {
                                info?.Invoke(
                                     $" - Application sur {wordKey} du statut {decisionsImportees.CarteMotStatut[wordKey]} (dictionnaire importé)",
                                     new Tuple<int, int>(n, DonneesTraitement.CarteMotOccurences.Count)
                                 );
                                first = false;
                            }
                            InitialiserStatutSurOccurence(indexOccurence, decisionsImportees.CarteMotStatut[wordKey]);
                        }
                        // Si aucun dictionnaire et que le mot n'est pas un mot a systématiquement évaluer
                        if (!DonneesTraitement.EstTraitee[indexOccurence] && motDBStats?.ToujoursDemander != 1) {
                            // on passe par la base de donnée des décisions par défaut
                            switch (decisionParDefaut) {
                                case 'a':
                                    if (first) {
                                        info?.Invoke(
                                             $" - Abreviation de {wordKey} (décision par défaut)",
                                             new Tuple<int, int>(n, DonneesTraitement.CarteMotOccurences.Count)
                                         );
                                        first = false;
                                    }
                                    InitialiserStatutSurOccurence(indexOccurence, Statut.ABREGE);
                                    //AppliquerStatutSurOccurence(i, Statut.ABREGE);
                                    break;
                                case 'p':
                                    if (first) {
                                        info?.Invoke(
                                             $" - Protection de {wordKey} (décision par défaut)",
                                             new Tuple<int, int>(n, DonneesTraitement.CarteMotOccurences.Count)
                                         );
                                        first = false;
                                    }
                                    InitialiserStatutSurOccurence(indexOccurence, Statut.PROTEGE);
                                    //AppliquerStatutSurOccurence(i, Statut.PROTEGE);
                                    break;
                                default:
                                    // Cas sans decisions par défaut : on donne une decision indicative
                                    if (motDBStats != null &&
                                        motDBStats.Documents > documentBarrier
                                        && Math.Max(motDBStats.Abreviations, motDBStats.Protections) > 0
                                    ) {
                                        double certainty =
                                            1.00
                                            - (
                                                (double)Math.Min(motDBStats.Abreviations, motDBStats.Protections)
                                                / (double)Math.Max(motDBStats.Abreviations, motDBStats.Protections)
                                            );
                                        if (
                                            certainty > 0.99
                                            && DonneesTraitement.StatutsOccurences[indexOccurence] == Statut.INCONNU
                                        ) {
                                            Statut selected =
                                                motDBStats.Abreviations > motDBStats.Protections
                                                    ? Statut.ABREGE
                                                    : Statut.PROTEGE;
                                            if (OptionsComplement.Instance.ActiverPreProtectionStatistique) {
                                                if (first) {
                                                    info?.Invoke(
                                                         $" - Application sur {wordKey} du statut {selected} (certitude {certainty * 100}%)",
                                                         new Tuple<int, int>(n, DonneesTraitement.CarteMotOccurences.Count)
                                                     );
                                                    first = false;
                                                }
                                                InitialiserStatutSurOccurence(indexOccurence, selected);
                                            } else {
                                                if (first) {
                                                    info?.Invoke(
                                                         $" - Prémarquage de {wordKey} en {selected} (certitude {certainty * 100}%)",
                                                         new Tuple<int, int>(n, DonneesTraitement.CarteMotOccurences.Count)
                                                     );
                                                    first = false;
                                                }
                                                DonneesTraitement.StatutsOccurences[indexOccurence] = selected;
                                            }
                                        }
                                    }
                                    break;

                            }
                        }
                    }

                    
                }

                
                //for (int i = 0; i < DonneesTraitement.Occurences.Count; i++) {
                //    info?.Invoke(
                //        $"",
                //        new Tuple<int, int>(i, DonneesTraitement.Occurences.Count)
                //    );
                //    string wordKey = DonneesTraitement.Occurences[i].ToLower();
                //    var item = carteMotDB.ContainsKey(wordKey) ? carteMotDB[wordKey] : null;
                //    char d = carteDecisionDB.ContainsKey(wordKey) ? carteDecisionDB[wordKey].Decision : ' ';
                //    if(importer != null && importer.CarteMotStatut.ContainsKey(wordKey)) {
                //        // Recharge des décisions par mot
                //        InitialiserStatutSurOccurence(i, importer.CarteMotStatut[wordKey]);
                //    }
                //    if(!DonneesTraitement.EstTraitee[i] && existingDictionnary != null) { // Recharge des décisions du précédent dictionnaire
                //        // Note : peut etre faire attention ici, pas sur que cette approche soit la bonne idée
                //        // il pourrait y avoir eu une correction changean la position de l'occurence dans la liste
                //        if (i <= existingDictionnary.Occurences.Count 
                //            && existingDictionnary.Occurences[i] == DonneesTraitement.Occurences[i]
                //        ) {
                //            // correspondance exacte entre occurence recharge et occurence courante, on réapplique 
                //            InitialiserStatutSurOccurence(i, existingDictionnary.StatutsOccurences[i]);
                //            //AppliquerStatutSurOccurence(i, existingDictionnary.StatutsOccurences[i]);
                //        } else if (existingDictionnary.CarteMotStatut.ContainsKey(DonneesTraitement.Occurences[i].ToLower())) {
                //            // correspondance partielle, on réapplique le statut du mot
                //            InitialiserStatutSurOccurence(i, existingDictionnary.CarteMotStatut[DonneesTraitement.Occurences[i].ToLower()]);
                //            //AppliquerStatutSurOccurence(i, existingDictionnary.CarteMotStatut[DonneesTraitement.Occurences[i].ToLower()]);
                //        }
                //    } 
                //    if (!DonneesTraitement.EstTraitee[i] && item?.ToujoursDemander != 1) {
                //        // Pas de docu
                //        switch (d) {
                //            case 'a':
                //                InitialiserStatutSurOccurence(i, Statut.ABREGE);
                //                //AppliquerStatutSurOccurence(i, Statut.ABREGE);
                //                break;
                //            case 'p':
                //                InitialiserStatutSurOccurence(i, Statut.PROTEGE);
                //                //AppliquerStatutSurOccurence(i, Statut.PROTEGE);
                //                break;
                //            default:
                //                // Pour les cas sans decisions par défaut, on donne une decision indicative
                //                if (item != null &&
                //                    item.Documents > documentBarrier
                //                    && Math.Max(item.Abreviations, item.Protections) > 0
                //                ) {
                //                    double certainty =
                //                        1.00
                //                        - (
                //                            (double)Math.Min(item.Abreviations, item.Protections)
                //                            / (double)Math.Max(item.Abreviations, item.Protections)
                //                        );
                //                    if (
                //                        certainty > 0.99
                //                        && DonneesTraitement.StatutsOccurences[i] == Statut.INCONNU
                //                    ) {
                //                        Statut selected =
                //                            item.Abreviations > item.Protections
                //                                ? Statut.ABREGE
                //                                : Statut.PROTEGE;
                //                        if (OptionsComplement.Instance.ActiverPreProtectionStatistique) {
                //                            InitialiserStatutSurOccurence(i, Statut.PROTEGE);
                //                        } else {
                //                            DonneesTraitement.StatutsOccurences[i] = selected;
                //                        }
                //                    }
                //                }
                //                break;
                            
                //        }
                //    }
                //}


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


        public async Task<List<int>> IndicesOccurencesHorsLexique()
        {
            List<Task<bool>> tasksHorsLexique = new List<Task<bool>>();
            foreach(var occurence in DonneesTraitement.Occurences) {
                tasksHorsLexique.Add(
                    Task.Run(() => !LexiqueFrance.EstFrancaisAbregeable(occurence))
                );
            }
            
            return await Task.WhenAll(tasksHorsLexique).ContinueWith(t =>
            {
                List<int> indices = new List<int>();
                bool[] estHorsLexique = t.Result;
                for (int i = 0; i < estHorsLexique.Length; i++) {
                    if (estHorsLexique[i]) {
                        indices.Add(i);
                    }
                }
                return indices;
            });
        }

        /// <summary>
        /// Analyse du texte en mémoire
        /// </summary>
        public void AnalyseDuTexteEnMemoire()
        {
            try {

                int offset = 0;
                List<int> debutBlocsIntegral = new List<int>();
                List<int> finBlocsIntegral = new List<int>();
                string texteAnalyser = string.Copy(TexteEnMemoire);
                
                var motsAvecChiffresAProteger = Abreviation.RechercheMotsAvecChiffres(texteAnalyser, info)
                    .OrderBy(kvp => kvp.Key)
                    .Select(kvp => kvp.Value)
                    .Where(value => !value.EstDejaProteger)
                    .ToList();
                if(motsAvecChiffresAProteger.Count > 0) {
                    info?.Invoke(
                        $"Protection de {motsAvecChiffresAProteger.Count} mots contenant des chiffres ..."
                    );
                    foreach (var occurenceATraiter in motsAvecChiffresAProteger) {
                        info?.Invoke(
                            $"- Protection de {occurenceATraiter.Mot}"
                        );
                        Range selection = _document.Range(offset + occurenceATraiter.Index, offset + occurenceATraiter.Index + occurenceATraiter.Mot.Length);
                        Range test = Proteger(selection);
                        offset += DictionnaireDeTravail.PROTECTION_CODE.Length;
                    }
                    ChargerTexteEnMemoire();
                    texteAnalyser = string.Copy(TexteEnMemoire);
                }
                info?.Invoke(
                    $"Recherche des blocs de protections existants ..."
                );
                var debutBlocs = Regex.Matches(texteAnalyser, $"\\[\\[\\*g1\\*\\]\\]");
                if (debutBlocs.Count > 0) {
                    foreach (Match res in debutBlocs) {
                        DonneesTraitement.DebutsBlocsIntegrals.Add(res.Index + DictionnaireDeTravail.PROTECTION_CODE_G1.Length);
                    }
                }
                var finBlocs = Regex.Matches(texteAnalyser, $"\\[\\[\\*g2\\*\\]\\]");
                if (finBlocs.Count > 0) {
                    foreach (Match res in finBlocs) {
                        DonneesTraitement.FinsBlocsIntegrals.Add(res.Index);
                    }
                }

                info?.Invoke($"Recherche des mots abregeables contenant des majuscules...");
                info?.Invoke($"(Lancement d'une mesure du temps d'analyse)");
                Stopwatch sw = new Stopwatch();
                sw.Start();
                var analyse = Abreviation.RechercheMotsAvecMaj(texteAnalyser).Where(o => o.EstAbregeable).ToList();
                //var analyse = Abreviation.AnalyserTexteComplet(texteAnalyser).Where(o => o.EstAbregeable && o.ContientDesMajuscules).ToList();
                info?.Invoke($"{analyse.Count} mots détectés, cartographie des mots dans le document...");
                
                foreach (var test in analyse.AsParallel()) {
                    // le logging de la progession ralenti considerablement le traitement
                    //info?.Invoke(
                    //    $"",
                    //    new Tuple<int, int>(i++, analyse.Count)
                    //);
                    // Marquer comme protéger les occurence qui sont détectés dans des blocs intégral
                    int indexBloc = DonneesTraitement.ListeBlocsIntegral.FindIndex(b => b.Item1 <= test.Index && (test.Index + test.Mot.Length) <= b.Item2);
                    if (indexBloc >= 0) {
                        test.EstDejaProteger = true;
                    }
                    DonneesTraitement.AjouterOccurence(test);
                }
                sw.Stop();
                DonneesTraitement.ReorderOccurences();
                DonneesTraitement.CalculCartographieMots();
                info?.Invoke(
                    $"Cartographie terminée en {sw.ElapsedMilliseconds} ms : {DonneesTraitement.CarteMotOccurences.Count} mots distincts"
                );

            } catch(Exception e) {
                error?.Invoke(e);
            }
            
        }


        /// <summary>
        /// Applique les décisions sur tout ou partie des occurences détectés
        /// </summary>
        /// <param name="surToutesOccurence">Si vrai, reapplique les décisions sur touts </param>
        /// <param name="start">Occurence a partir de laquelle appliquer les décision (incluse)</param>
        /// <param name="end">Derniere occurence a traité (incluse), -1 = pas de limite</param>
        public void AppliquerDecisions(bool surToutesOccurence = false, int start = 0, int end = -1)
        {
            info?.Invoke(
                $"Application des décisions sur le document ..."
            );
            int _start = Math.Min(DonneesTraitement.Occurences.Count - 1, Math.Max(0, start));
            int _end = end < 0 ? DonneesTraitement.Occurences.Count - 1 : Math.Min(DonneesTraitement.Occurences.Count - 1, Math.Max(0, start));
            if(_start > _end) {
                int t = _end;
                _end = _start;
                _start = t;
            }
            info?.Invoke("", Tuple.Create(0, _end - _start));
            for (int i = _start; i <= _end; i++) {
                string wordKey = DonneesTraitement.Occurences[i].ToLower();
                Statut selected = DonneesTraitement.StatutsOccurences[i];
                // N'appliquer les décisions
                if (surToutesOccurence || !DonneesTraitement.EstTraitee[i]) {
                    switch (selected) {
                        case Statut.INCONNU:
                            if (DonneesTraitement.CarteMotStatut.ContainsKey(wordKey)
                                && DonneesTraitement.CarteMotStatut[wordKey] != Statut.INCONNU
                                && DonneesTraitement.EstTraitee[i] == false
                            ) {
                                AppliquerStatutSurOccurence(i, DonneesTraitement.CarteMotStatut[wordKey]);
                            }
                            break;
                        default:
                            AppliquerStatutSurOccurence(i, selected);
                            break;
                    }
                }
                info?.Invoke("", Tuple.Create(_start + 1, _end - _start));
            }
            ChargerTexteEnMemoire();
        }

        /// <summary>
        /// FIXME : Cette fonction pose probleme en relance asynchrone (freeze de l'ui
        /// </summary>
        public bool ReanalyserDocumentSiModification()
        {
            if (TexteEnMemoire != DocumentMainRange)
            {
                info?.Invoke(
                    $"Modification de contenu détecter : réanalyse du document ..."
                );

                DictionnaireDeTravail actuel = DonneesTraitement != null ? DonneesTraitement.Clone() : new DictionnaireDeTravail(
                    Path.GetFileNameWithoutExtension(WorkingDictionnaryPath)
                );
                DonneesTraitement = new DictionnaireDeTravail(
                    Path.GetFileNameWithoutExtension(WorkingDictionnaryPath)
                );

                ChargerTexteEnMemoire();
                
                AnalyseDuTexteEnMemoire();
                
                info?.Invoke($"Rechargement des décisions ...");
                List<string> motsAControler = new List<string>();
                for (int i = 0; i < actuel.CarteMotOccurences.Keys.Count; i++) {
                    info?.Invoke("", new Tuple<int, int>(i, actuel.CarteMotOccurences.Keys.Count));
                    string item = actuel.CarteMotOccurences.Keys.ElementAt(i);
                    // Mot ajouté manuellement
                    if (!DonneesTraitement.CarteMotOccurences.ContainsKey(item)) {
                        info?.Invoke($"Reprise du mot ajouté manuellement {item} ...");
                        AjouterMotAuTraitement(item);
                    }
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
                return true;
            }
            return false;
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
            Regex toLook = Abreviation.SearchWord(word);

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
                            DonneesTraitement.EstTraitee[
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
                            DonneesTraitement.EstTraitee[
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
            this.DonneesTraitement.DernierMotSelectionne = MotSelectionne;
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

        /// <summary>
        /// Insère des codes G1 et G2 autour d'un bloc de texte dans un document, si ce bloc n'est pas déjà protégé.
        /// </summary>
        /// <param name="current">Documment Word</param>
        /// <param name="wordRange">Sélection dans le document word</param>
        /// <returns>La sélection protégée (sans les codes s'ils ont été rajoutés)</returns>
        public static MSWord.Range ProtegerBloc(MSWord.Document current, MSWord.Range wordRange)
        {
            // Note sur le code commenter :
            // Ne marche pas - texte parasite a cause des code duxburry qui sont pris comme des mots
            //if (wordRange.Words.Count == 0) {
            //    // Si la sélection est vide, on ne fait rien
            //    return wordRange;
            //}
            //// Un seul mot dans la sélection, on le resélectionne
            //if (wordRange.Words.Count == 1) {
            //    wordRange = wordRange.Words.First;
            //} else {
            //    wordRange = current.Range(
            //        wordRange.Words.First.Start,
            //        wordRange.Words.Last.End
            //    );
            //}
            //wordRange.Select();
            string trimmedStart = wordRange.Text.TrimStart();
            int trimmedCount = wordRange.Text.Length - trimmedStart.Length;
            if (trimmedCount > 0)
            {
                // Reselectionner le mot sans les espaces de début
                wordRange = current.Range(
                    wordRange.Start + trimmedCount,
                    wordRange.End
                );
                //wordRange.Select();
            }

            string textBefore = current.Range(0, wordRange.Start).Text;

            if (!EstProtegerBloc(textBefore, wordRange.Text))
            {
                wordRange.InsertBefore(DictionnaireDeTravail.PROTECTION_CODE_G1);
                wordRange.InsertAfter(DictionnaireDeTravail.PROTECTION_CODE_G2);
                wordRange = current.Range(
                    wordRange.Start + DictionnaireDeTravail.PROTECTION_CODE_G1.Length,
                    wordRange.End - DictionnaireDeTravail.PROTECTION_CODE_G2.Length
                );
                //wordRange.Select();
            }
            return wordRange;
        }

        /// <summary>
        /// Protéger un bloc de mots dans le document
        /// </summary>
        /// <param name="wordRange"></param>
        public void ProtegerBloc(MSWord.Range wordRange)
        {
            string trimmedStart = wordRange.Text.TrimStart();
            int trimmedCount = wordRange.Text.Length - trimmedStart.Length;
            if (trimmedCount > 0) {
                // Reselectionner le mot sans les espaces de début
                wordRange = _document.Range(wordRange.Start + trimmedCount, wordRange.End);
                //wordRange.Select();
            }

            string trimmedEnd = wordRange.Text.TrimEnd();
            if (trimmedEnd.EndsWith("\r")) trimmedEnd = trimmedEnd.Substring(0, trimmedEnd.Length - 1);
            trimmedCount = wordRange.Text.Length - trimmedEnd.Length;
            if (trimmedCount > 0) {
                // Reselectionner le mot sans les espaces de début
                wordRange = _document.Range(wordRange.Start, wordRange.End - trimmedCount);
                //wordRange.Select();
            }

            if (wordRange.Text.StartsWith(DictionnaireDeTravail.PROTECTION_CODE_G1)) {
                wordRange = _document.Range(wordRange.Start + DictionnaireDeTravail.PROTECTION_CODE_G1.Length, wordRange.End);
            }
            if (wordRange.Text.EndsWith(DictionnaireDeTravail.PROTECTION_CODE_G2)) {
                wordRange = _document.Range(wordRange.Start, wordRange.End - +DictionnaireDeTravail.PROTECTION_CODE_G2.Length);
            }

            int debutProtection = wordRange.Start;
            int finProtection = wordRange.End;

            int indexDebut = DonneesTraitement.ListeBlocsIntegral.FindIndex((t) => t.Item1 <= debutProtection && debutProtection <= t.Item2);
            int indexFin = DonneesTraitement.ListeBlocsIntegral.FindIndex((t) => t.Item1 <= finProtection && finProtection <= t.Item2);

            if (indexDebut < 0 && indexFin < 0) {

                // Supprimer tous les blocs intégrale entre wordRange.Start et wordRange.End
                for (int i = DonneesTraitement.ListeBlocsIntegral.Count - 1; i >= 0; i--) {
                    var bloc = DonneesTraitement.ListeBlocsIntegral[i];
                    if (bloc.Item1 >= debutProtection && bloc.Item2 <= finProtection) {
                        AbregerBloc(_document, _document.Range(bloc.Item1, bloc.Item2));
                        DonneesTraitement.SupprimerBlocIntegral(i, true);
                        // Décaller la fin attendu du nouveau bloc abrégé
                        finProtection -= (DictionnaireDeTravail.PROTECTION_CODE_G1.Length + DictionnaireDeTravail.PROTECTION_CODE_G2.Length);
                    }
                }

                for (int i = 0; i < DonneesTraitement.Occurences.Count; i++) {
                    if (DonneesTraitement.PositionsOccurences[i] >= debutProtection && DonneesTraitement.PositionsOccurences[i] < finProtection) {
                        AppliquerStatutSurOccurence(i, Statut.IGNORE);
                    }
                }

                // Si le bloc n'est pas dans la liste des blocs intégral, on protège la zone
                // passer en ignorer toutes les occurences situés dans l'intervale
                wordRange = ProtegerBloc(_document, wordRange);
                DonneesTraitement.AjouterBlocIntegral(wordRange.Start, wordRange.End, true);
                return;
            } else if(indexDebut == indexFin) {
                    // Deja dans un bloc protégé
                return;
            } else if(indexDebut >= 0 && indexDebut < indexFin) {
                // protection sur plusieurs bloc : fusion des blocs
                int blocStart = DonneesTraitement.ListeBlocsIntegral[indexDebut].Item1;
                int blockEnd = DonneesTraitement.ListeBlocsIntegral[indexDebut].Item2;
                int newEnd = DonneesTraitement.ListeBlocsIntegral[indexFin].Item2;
                // Supprimer les blocs y compris les bloc englobant
                for (int i = indexFin; i >= indexDebut; i--) {
                    var bloc = DonneesTraitement.ListeBlocsIntegral[i];
                    AbregerBloc(_document, _document.Range(bloc.Item1, bloc.Item2));
                    DonneesTraitement.SupprimerBlocIntegral(i);
                }
                // Creer le nouveau block
                wordRange = ProtegerBloc(_document, _document.Range(blocStart, newEnd));
                DonneesTraitement.AjouterBlocIntegral(wordRange.Start, wordRange.End, true);
                //Passer en ignorer toutes les occurence dans le nouveau block
                for (int i = DonneesTraitement.Occurences.Count - 1; i >=0 ; i--) {
                    if (DonneesTraitement.PositionsOccurences[i] >= blocStart && DonneesTraitement.PositionsOccurences[i] <= newEnd) {
                        AppliquerStatutSurOccurence(i, Statut.IGNORE);
                    }
                }

            } else if (indexDebut >= 0 && indexFin < 0) {
                // Extension d'un bloc existant vers sa fin
                int blocStart = DonneesTraitement.ListeBlocsIntegral[indexDebut].Item1;
                int blockEnd = DonneesTraitement.ListeBlocsIntegral[indexDebut].Item2;
                int newEnd = wordRange.End;
                // Supprimer les bloc intermédiaire
                for (int i = DonneesTraitement.ListeBlocsIntegral.Count - 1; i >= indexDebut; i--) {
                    var bloc = DonneesTraitement.ListeBlocsIntegral[i];
                    if (bloc.Item1 >= blocStart && bloc.Item2 <= newEnd) {
                        // On trouve un bloc qui commence après le début du bloc actuel et avant la fin du nouveau bloc
                        AbregerBloc(_document, _document.Range(bloc.Item1, bloc.Item2));
                        DonneesTraitement.SupprimerBlocIntegral(i);
                    }
                }
                wordRange = ProtegerBloc(_document, _document.Range(blocStart, newEnd));
                DonneesTraitement.AjouterBlocIntegral(blocStart, newEnd, true);
                // Passer en ignorer toutes les occurence dans le nouveau block
                for (int i = DonneesTraitement.Occurences.Count - 1; i >= 0; i--) {
                    if (DonneesTraitement.PositionsOccurences[i] >= wordRange.Start && DonneesTraitement.PositionsOccurences[i] <= wordRange.End) {
                        AppliquerStatutSurOccurence(i, Statut.IGNORE);
                    }
                }

            } else if (indexDebut < 0 && indexFin >= 0) {
                // Extension d'un bloc existant vers son début
                int blocStart = DonneesTraitement.ListeBlocsIntegral[indexFin].Item1;
                int blockEnd = DonneesTraitement.ListeBlocsIntegral[indexFin].Item2;
                int newStart = wordRange.Start;
                
                // Supprimer les bloc
                for (int i = indexFin; i >= 0; i--) {
                    var bloc = DonneesTraitement.ListeBlocsIntegral[i];
                    if (bloc.Item1 >= newStart && bloc.Item2 < blockEnd) {
                        // On trouve un bloc qui commence après le début du nouveaux bloc actuel et avant la fin du bloc actuel
                        AbregerBloc(_document, _document.Range(bloc.Item1, bloc.Item2));
                        DonneesTraitement.SupprimerBlocIntegral(i);
                    }
                }
                ProtegerBloc(_document, _document.Range(newStart, blockEnd));
                DonneesTraitement.AjouterBlocIntegral(newStart, blockEnd, true);
                //Passer en ignorer toutes les occurence dans le nouveau block
                for (int i = DonneesTraitement.Occurences.Count - 1; i >= 0; i--) {
                    if (DonneesTraitement.PositionsOccurences[i] >= newStart && DonneesTraitement.PositionsOccurences[i] <= blockEnd) {
                        AppliquerStatutSurOccurence(i, Statut.IGNORE);
                    }
                }
            }
        }

        /// <summary>
        /// Controle si une sélection est précédé ou commence par un code de protection de mot simple ([[*i*]])
        /// </summary>
        /// <param name="current"></param>
        /// <param name="wordRange"></param>
        /// <returns></returns>
        public static bool EstProtegerMot(MSWord.Document current, MSWord.Range wordRange)
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
            string previousTextAssumingICode = current.Range(
                wordRange.Start - DictionnaireDeTravail.PROTECTION_CODE.Length,
                wordRange.Start
            ).Text;

            //string previousTextAssumingG1Code = current.Range(
            //    Math.Max(0,wordRange.Start - DictionnaireDeTravail.PROTECTION_CODE.Length),
            //    wordRange.Start
            //).Text;

            //string nextTextAssumingG2Code = current.Range(
            //    wordRange.End,
            //    Math.Min(current.Content.End, wordRange.End + DictionnaireDeTravail.PROTECTION_CODE_G2.Length)
            //).Text;
            string currentText = wordRange.Text;
            return previousTextAssumingICode.Equals(DictionnaireDeTravail.PROTECTION_CODE)
                || currentText.StartsWith(DictionnaireDeTravail.PROTECTION_CODE);
               //|| ( // Mots composés entourés par un bloc
               //     (previousTextAssumingG1Code.Equals(DictionnaireDeTravail.PROTECTION_CODE_G1) || currentText.StartsWith(DictionnaireDeTravail.PROTECTION_CODE_G1)) 
               //     && (nextTextAssumingG2Code.Equals(DictionnaireDeTravail.PROTECTION_CODE_G2) || currentText.EndsWith(DictionnaireDeTravail.PROTECTION_CODE_G2))
               // );
        }


        /// <summary>
        /// Controle si un code de protection de bloc existe dans le texte avant ou dans une selection
        /// </summary>
        /// <param name="textBefore"></param>
        /// <param name="textEvaluated"></param>
        /// <returns></returns>
        public static bool EstProtegerBloc(string textBefore, string textEvaluated = "")
        {
            if (textBefore.TrimEnd().EndsWith(DictionnaireDeTravail.PROTECTION_CODE_G1) || textEvaluated.TrimStart().StartsWith(DictionnaireDeTravail.PROTECTION_CODE_G1)) {
                return true;
            }
            int lastG1 = textBefore.LastIndexOf(DictionnaireDeTravail.PROTECTION_CODE_G1);
            int lastG2 = textBefore.LastIndexOf(DictionnaireDeTravail.PROTECTION_CODE_G2);
            if (lastG1 > -1) {
                return lastG1 > lastG2;
            }

            return false;
        }

        public bool EstProtegerBloc(int position)
        {
            return EstProtegerBloc(DocumentMainRange.Substring(0,position));
        }

        public int RetrouverOccurence(Range wordRange)
        {
            string trimmedStart = wordRange.Text.TrimStart();
            int trimmedCount = wordRange.Text.Length - trimmedStart.Length;
            if (trimmedCount > 0) {
                // Reselectionner le mot sans les espaces de début
                wordRange = _document.Range(wordRange.Start + trimmedCount, wordRange.End);
                //wordRange.Select();
            }

            string trimmedEnd = wordRange.Text.TrimEnd();
            if (trimmedEnd.EndsWith("\r")) trimmedEnd = trimmedEnd.Substring(0, trimmedEnd.Length - 1);
            trimmedCount = wordRange.Text.Length - trimmedEnd.Length;
            if (trimmedCount > 0) {
                // Reselectionner le mot sans les espaces de début
                wordRange = _document.Range(wordRange.Start, wordRange.End - trimmedCount);
                //wordRange.Select();
            }

            int indexOccurence = DonneesTraitement.PositionsOccurences.FindIndex(p => p == wordRange.Start);

            return indexOccurence >= 0 && trimmedEnd == DonneesTraitement.Occurences[indexOccurence] ? indexOccurence : -1;
        }

        /// <summary>
        /// Fonction utilitaire d'insertion d'un code de protection d'une sélection dans un document.
        /// </summary>
        /// <param name="current"></param>
        /// <param name="wordRange"></param>
        /// <returns></returns>
        public static MSWord.Range Proteger(MSWord.Document current, MSWord.Range wordRange)
        {
            //if (wordRange.Words.Count == 0) {
            //    // Si la sélection est vide, on ne fait rien
            //    return wordRange;
            //}
            
            //if (wordRange.Words.Count == 1) {
            //    // Un seul mot dans la sélection, on le resélectionne
            //    wordRange = wordRange.Words.First;
            //} else {
            //    // Plusieurs mots, on sélectionne tout
            //    wordRange = current.Range(
            //        wordRange.Words.First.Start,
            //        wordRange.Words.Last.End
            //    );
            //}
            //wordRange.Select();
            string trimmedStart = wordRange.Text.TrimStart();
            int trimmedCount = wordRange.Text.Length - trimmedStart.Length;
            if (trimmedCount > 0) {
                // Reselectionner le mot sans les espaces de début
                wordRange = current.Range(wordRange.Start + trimmedCount, wordRange.End);
                //wordRange.Select();
            }

            string trimmedEnd = wordRange.Text.TrimEnd();
            if (trimmedEnd.EndsWith("\r")) trimmedEnd = trimmedEnd.Substring(0, trimmedEnd.Length - 1);
            trimmedCount = wordRange.Text.Length - trimmedEnd.Length;
            if (trimmedCount > 0) {
                // Reselectionner le mot sans les espaces de début
                wordRange = current.Range(wordRange.Start, wordRange.End - trimmedCount);
                //wordRange.Select();
            }

            if (EstProtegerBloc(current.Range(0, wordRange.Start).Text, wordRange.Text)) {
                // Si le mot est déjà dans un bloc protégé, on ne le protège pas à nouveau
                return wordRange;
            }

            if (wordRange.Words.Count > 1) {
                // Si un code de protection de mots est déjà sur un mot composé
                // Le supprimer
                int start = wordRange.Start;
                int end = wordRange.End;
                foreach(Range w in wordRange.Words) {
                    if (EstProtegerMot(current, w)) {
                        Range toDelete = current.Range(
                            w.Start - DictionnaireDeTravail.PROTECTION_CODE.Length,
                            w.Start
                        );
                        toDelete.Delete();
                        start -= DictionnaireDeTravail.PROTECTION_CODE.Length;
                        end -= DictionnaireDeTravail.PROTECTION_CODE.Length;
                    }
                }
                wordRange = current.Range(start, end);
                trimmedEnd = wordRange.Text.TrimEnd();
                trimmedCount = wordRange.Text.Length - trimmedEnd.Length;
                if (trimmedCount > 0) {
                    // Reselectionner le mot sans les espaces de début
                    wordRange = current.Range(wordRange.Start, wordRange.End - trimmedCount);
                    //wordRange.Select();
                }
                //string textTest = wordRange.Text;
                // Insérer le code de block
                wordRange.InsertBefore(DictionnaireDeTravail.PROTECTION_CODE_G1);
                wordRange.InsertAfter(DictionnaireDeTravail.PROTECTION_CODE_G2);
                wordRange = current.Range(
                    wordRange.Start + DictionnaireDeTravail.PROTECTION_CODE_G1.Length,
                    wordRange.End - DictionnaireDeTravail.PROTECTION_CODE_G2.Length
                );
                //textTest = wordRange.Text;
                //wordRange.Select();
                return wordRange;
            }
            if (!EstProtegerMot(current, wordRange)) {
                wordRange.InsertBefore(DictionnaireDeTravail.PROTECTION_CODE);
                wordRange = current.Range(wordRange.Start + DictionnaireDeTravail.PROTECTION_CODE.Length, wordRange.End);
                wordRange.Select();
            }
            //current.Save();
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
            //if (wordRange.Words.Count == 0) {
            //    // Si la sélection est vide, on ne fait rien
            //    return wordRange;
            //}

            //if (wordRange.Words.Count == 1) {
            //    // Un seul mot dans la sélection, on le resélectionne
            //    wordRange = wordRange.Words.First;
            //} else {
            //    // Plusieurs mots, on sélectionne tout
            //    wordRange = current.Range(
            //        wordRange.Words.First.Start,
            //        wordRange.Words.Last.End
            //    );
            //}
            //wordRange.Select();
            string trimmedStart = wordRange.Text.TrimStart();
            int trimmedCount = wordRange.Text.Length - trimmedStart.Length;
            if (trimmedCount > 0) {
                // Reselectionner le mot sans les espaces de début
                wordRange = current.Range(wordRange.Start + trimmedCount, wordRange.End);
                //wordRange.Select();
            }

            string trimmedEnd = wordRange.Text.TrimEnd();
            if (trimmedEnd.EndsWith("\r")) trimmedEnd = trimmedEnd.Substring(0, trimmedEnd.Length - 1);
            trimmedCount = wordRange.Text.Length - trimmedEnd.Length;
            if (trimmedCount > 0) {
                // Reselectionner le mot sans les espaces de début
                wordRange = current.Range(wordRange.Start, wordRange.End - trimmedCount);
                //wordRange.Select();
            }
            if (EstProtegerMot(current, wordRange)) {
                MSWord.Range previous = current.Range(
                    wordRange.Start - DictionnaireDeTravail.PROTECTION_CODE.Length,
                    wordRange.Start
                );
                string previousText = previous.Text;
                // First check if text is preceded by protection code
                if (previousText.Equals(DictionnaireDeTravail.PROTECTION_CODE)) {
                    Range toDelete = current.Range(
                        wordRange.Start - DictionnaireDeTravail.PROTECTION_CODE.Length,
                        wordRange.Start
                    );
                    toDelete.Delete();
                    wordRange = current.Range(toDelete.End, wordRange.End);
                    //wordRange.Select();
                    return wordRange;
                }

                if (trimmedStart.StartsWith(DictionnaireDeTravail.PROTECTION_CODE)) {
                    Range toDelete = current.Range(
                        wordRange.Start,
                        wordRange.Start + DictionnaireDeTravail.PROTECTION_CODE.Length
                    );
                    toDelete.Delete();
                    wordRange = current.Range(toDelete.End, wordRange.End);
                    //wordRange.Select();
                    return wordRange;
                }
            } else {
                string previousText = current.Range(
                    Math.Max(wordRange.Start - DictionnaireDeTravail.PROTECTION_CODE_G1.Length, 0),
                    wordRange.Start
                ).Text;
                string currentText = current.Range(
                    wordRange.Start,
                    wordRange.End
                ).Text;
                string nextText = current.Range(
                    wordRange.End,
                    Math.Min(wordRange.End + DictionnaireDeTravail.PROTECTION_CODE_G2.Length, current.Content.End)
                ).Text;
                int start = wordRange.Start;
                int end = wordRange.End;
                Range toDeleteBefore = null;
                Range toDeleteAfter = null;
                if (previousText.EndsWith(DictionnaireDeTravail.PROTECTION_CODE_G1)) {
                    toDeleteBefore = current.Range(
                        start - DictionnaireDeTravail.PROTECTION_CODE_G1.Length,
                        start
                    );
                }
                if (currentText.StartsWith(DictionnaireDeTravail.PROTECTION_CODE_G1)) {
                    toDeleteBefore = current.Range(
                        start,
                        start + DictionnaireDeTravail.PROTECTION_CODE_G1.Length
                    );
                }
                if (currentText.EndsWith(DictionnaireDeTravail.PROTECTION_CODE_G2)) {
                    toDeleteAfter = current.Range(
                        end - DictionnaireDeTravail.PROTECTION_CODE_G2.Length,
                        end
                    );
                }
                if (nextText.StartsWith(DictionnaireDeTravail.PROTECTION_CODE_G2)) {
                    toDeleteAfter = current.Range(
                        end - DictionnaireDeTravail.PROTECTION_CODE_G2.Length,
                        end
                    );
                }
                if(toDeleteBefore != null && toDeleteAfter != null) {
                    toDeleteAfter.Delete();
                    toDeleteBefore.Delete();
                    wordRange = current.Range(
                        toDeleteBefore.End,
                        toDeleteAfter.Start
                    );
                    return wordRange;
                    //string testTExt = wordRange.Text;
                }
            }
            

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


        public void AbregerBloc(MSWord.Range wordRange)
        {
            string trimmedStart = wordRange.Text.TrimStart();
            int trimmedCount = wordRange.Text.Length - trimmedStart.Length;
            if (trimmedCount > 0) {
                // Reselectionner le mot sans les espaces de début
                wordRange = _document.Range(wordRange.Start + trimmedCount, wordRange.End);
                //wordRange.Select();
            }

            string trimmedEnd = wordRange.Text.TrimEnd();
            if (trimmedEnd.EndsWith("\r")) trimmedEnd = trimmedEnd.Substring(0, trimmedEnd.Length - 1);
            trimmedCount = wordRange.Text.Length - trimmedEnd.Length;
            if (trimmedCount > 0) {
                // Reselectionner le mot sans les espaces de début
                wordRange = _document.Range(wordRange.Start, wordRange.End - trimmedCount);
                //wordRange.Select();
            }

            if (wordRange.Text.StartsWith(DictionnaireDeTravail.PROTECTION_CODE_G1)){
                wordRange = _document.Range(wordRange.Start + DictionnaireDeTravail.PROTECTION_CODE_G1.Length, wordRange.End);
            }
            if (wordRange.Text.EndsWith(DictionnaireDeTravail.PROTECTION_CODE_G2)){
                wordRange = _document.Range(wordRange.Start, wordRange.End - +DictionnaireDeTravail.PROTECTION_CODE_G2.Length);
            }

            wordRange.Select();

            int debutAbreger = wordRange.Start;
            int finAbreger = wordRange.End;
            int longueurTexte = debutAbreger - finAbreger;
            int indexDebut = DonneesTraitement.ListeBlocsIntegral.FindIndex((t) => t.Item1 <= debutAbreger && debutAbreger <= t.Item2);
            int indexFin = DonneesTraitement.ListeBlocsIntegral.FindIndex((t) => t.Item1 <= finAbreger && finAbreger <= t.Item2);
            // Texte sans plus aucun code de protection
            int LongueurTextAbreger = wordRange.Text
                .Replace(DictionnaireDeTravail.PROTECTION_CODE_G1, "")
                .Replace(DictionnaireDeTravail.PROTECTION_CODE_G2, "")
                .Replace(DictionnaireDeTravail.PROTECTION_CODE, "")
                .Length;
            if (indexDebut < 0 && indexFin < 0) {
                // Supprimer tous les blocs intégrale entre wordRange.Start et wordRange.End
                for (int i = DonneesTraitement.ListeBlocsIntegral.Count - 1; i >= 0; i--) {
                    var bloc = DonneesTraitement.ListeBlocsIntegral[i];
                    if (bloc.Item1 >= debutAbreger && bloc.Item2 <= finAbreger) {
                        AbregerBloc(_document, _document.Range(bloc.Item1, bloc.Item2));
                        DonneesTraitement.SupprimerBlocIntegral(i, true);
                        // Décaller la fin attendu du nouveau bloc abrégé
                        finAbreger -= (DictionnaireDeTravail.PROTECTION_CODE_G1.Length + DictionnaireDeTravail.PROTECTION_CODE_G2.Length);
                    }
                }

                for (int i = 0; i < DonneesTraitement.Occurences.Count; i++) {
                    if (DonneesTraitement.PositionsOccurences[i] >= debutAbreger && DonneesTraitement.PositionsOccurences[i] < finAbreger) {
                        AppliquerStatutSurOccurence(i, Statut.ABREGE);
                    }
                }

            } else if (indexDebut == indexFin) {
                // séprarer le bloc intégral en 3 partie (integral puis abrege puis integrale
                int debutProtegerAvant = DonneesTraitement.ListeBlocsIntegral[indexDebut].Item1;
                int finProtegerApres = DonneesTraitement.ListeBlocsIntegral[indexDebut].Item2;
                int tailleProtegerApres = finProtegerApres - finAbreger;
                int tailleProtegerAvant = debutAbreger - debutProtegerAvant;

                // Supprimer le bloc courant et renvoie la position du bloc "abrege"
                wordRange = AbregerBloc(_document, _document.Range(debutProtegerAvant, finProtegerApres));
                debutProtegerAvant = wordRange.Start;
                debutAbreger = debutProtegerAvant + tailleProtegerAvant;
                finProtegerApres = wordRange.End;
                finAbreger = finProtegerApres - tailleProtegerApres;
                DonneesTraitement.SupprimerBlocIntegral(indexDebut, true);
                // Abreger les mots entre les 2 blocs
                for (int i = 0; i < DonneesTraitement.Occurences.Count; i++) {
                    if (DonneesTraitement.PositionsOccurences[i] >= debutAbreger && DonneesTraitement.PositionsOccurences[i] < finAbreger) {
                        AppliquerStatutSurOccurence(i, Statut.ABREGE);
                    }
                }
                finAbreger = debutAbreger + LongueurTextAbreger;


                // Reajouter un sous block 1 avant s'il y avait d'autres mots
                if (tailleProtegerAvant > 0) {
                    wordRange = ProtegerBloc(_document, _document.Range(debutProtegerAvant, debutProtegerAvant + tailleProtegerAvant));
                    DonneesTraitement.AjouterBlocIntegral(wordRange.Start, wordRange.End, true);
                    debutAbreger = wordRange.End;
                    finAbreger = debutAbreger + LongueurTextAbreger;
                }
                // Reajouter un bloc a la fin s'il y avait d'autres mots
                if (tailleProtegerApres > 0) {
                    wordRange = ProtegerBloc(_document, _document.Range(finAbreger, finAbreger + tailleProtegerApres));
                    DonneesTraitement.AjouterBlocIntegral(finAbreger, finAbreger + tailleProtegerApres, true);
                }
                

            } else if (indexDebut >= 0 && indexDebut < indexFin) {
                // séprarer le bloc intégral en 3 partie (integral puis abrege puis integrale
                int debutProtegerAvant = DonneesTraitement.ListeBlocsIntegral[indexDebut].Item1;
                int finProtegerApres = DonneesTraitement.ListeBlocsIntegral[indexFin].Item2;
                int tailleProtegerApres = finProtegerApres - finAbreger;
                int tailleProtegerAvant = debutAbreger - debutProtegerAvant;

                // Supprimer tous les blocs dans la plage [indexDebut, indexFin]
                for (int i = indexFin; i >= indexDebut; i--) {
                    var bloc = DonneesTraitement.ListeBlocsIntegral[i];
                    wordRange = AbregerBloc(_document, _document.Range(bloc.Item1, bloc.Item2));
                    DonneesTraitement.SupprimerBlocIntegral(i, true);
                    finAbreger -= (DictionnaireDeTravail.PROTECTION_CODE_G1.Length + DictionnaireDeTravail.PROTECTION_CODE_G2.Length);
                }
                debutProtegerAvant = wordRange.Start;
                debutAbreger = debutProtegerAvant + tailleProtegerAvant;
                //Passage en abrege des occurences (et remise a jour des positions
                for (int i = 0; i < DonneesTraitement.Occurences.Count; i++) {
                    if (DonneesTraitement.PositionsOccurences[i] >= debutAbreger && DonneesTraitement.PositionsOccurences[i] < finAbreger) {
                        AppliquerStatutSurOccurence(i, Statut.ABREGE);
                    }
                }
                finAbreger = debutAbreger + LongueurTextAbreger;
                // Reajouter un sous block 1 avant s'il y avait d'autres mots
                if (tailleProtegerAvant > 0) {
                    wordRange = ProtegerBloc(_document, _document.Range(debutProtegerAvant, debutProtegerAvant + tailleProtegerAvant));
                    DonneesTraitement.AjouterBlocIntegral(wordRange.Start, wordRange.End, true);
                    debutAbreger = wordRange.End;
                    finAbreger = debutAbreger + LongueurTextAbreger;
                }
                // Reajouter un bloc a la fin s'il y avait d'autres mots
                if (tailleProtegerApres > 0) {
                    wordRange = ProtegerBloc(_document, _document.Range(finAbreger, finAbreger + tailleProtegerApres));
                    DonneesTraitement.AjouterBlocIntegral(finAbreger, finAbreger + tailleProtegerApres, true);
                }
            } else if (indexDebut >= 0 && indexFin < 0) {
                // séprarer le bloc intégral en 3 partie (integral puis abrege puis integrale
                int debutProtegerAvant = DonneesTraitement.ListeBlocsIntegral[indexDebut].Item1;
                int finProtegerApres = finAbreger;
                //int tailleProtegerApres = finProtegerApres - finAbreger;
                int tailleProtegerAvant = debutAbreger - debutProtegerAvant;

                // Supprimer tous les blocs dans la plage [indexDebut, total] mais situé avant finAbreger
                for (int i = DonneesTraitement.ListeBlocsIntegral.Count - 1; i >= indexDebut; i--) {
                    var bloc = DonneesTraitement.ListeBlocsIntegral[i];
                    if (debutAbreger <= bloc.Item1 && bloc.Item2 <= finAbreger) {
                        wordRange = AbregerBloc(_document, _document.Range(bloc.Item1, bloc.Item2));
                        DonneesTraitement.SupprimerBlocIntegral(i, true);
                        finAbreger -= (DictionnaireDeTravail.PROTECTION_CODE_G1.Length + DictionnaireDeTravail.PROTECTION_CODE_G2.Length);
                    }       
                }
                debutProtegerAvant = wordRange.Start;
                debutAbreger = debutProtegerAvant + tailleProtegerAvant;
                //Passage en abrege des occurences (et remise a jour des positions
                for (int i = 0; i < DonneesTraitement.Occurences.Count; i++) {
                    if (DonneesTraitement.PositionsOccurences[i] >= debutAbreger && DonneesTraitement.PositionsOccurences[i] < finAbreger) {
                        AppliquerStatutSurOccurence(i, Statut.ABREGE);
                    }
                }
                finAbreger = debutAbreger + LongueurTextAbreger;
                // Reajouter un sous block 1 avant s'il y avait d'autres mots
                if (tailleProtegerAvant > 0) {
                    wordRange = ProtegerBloc(_document, _document.Range(debutProtegerAvant, debutProtegerAvant + tailleProtegerAvant));
                    DonneesTraitement.AjouterBlocIntegral(wordRange.Start, wordRange.End, true);
                    debutAbreger = wordRange.End;
                    finAbreger = debutAbreger + LongueurTextAbreger;
                }
                // Reajouter un bloc a la fin s'il y avait d'autres mots
                //if (tailleProtegerApres > 0) {
                //    wordRange = ProtegerBloc(_document, _document.Range(finAbreger, finAbreger + tailleProtegerApres));
                //    DonneesTraitement.AjouterBlocIntegral(finAbreger, finAbreger + tailleProtegerApres, true);
                //}

            } else if (indexDebut < 0 && indexFin >= 0) {

                // séprarer le bloc intégral en 3 partie (integral puis abrege puis integrale
                int debutProtegerAvant = debutAbreger;
                int finProtegerApres = DonneesTraitement.ListeBlocsIntegral[indexFin].Item2;
                int tailleProtegerApres = finProtegerApres - finAbreger;
                //int tailleProtegerAvant = debutAbreger - debutProtegerAvant;

                // Supprimer tous les blocs dans la plage [indexDebut, total] mais situé avant finAbreger
                for (int i = indexFin; i >= 0; i--) {
                    var bloc = DonneesTraitement.ListeBlocsIntegral[i];
                    if (debutAbreger <= bloc.Item1 && bloc.Item2 <= finAbreger) {
                        wordRange = AbregerBloc(_document, _document.Range(bloc.Item1, bloc.Item2));
                        DonneesTraitement.SupprimerBlocIntegral(i, true);
                        finAbreger -= (DictionnaireDeTravail.PROTECTION_CODE_G1.Length + DictionnaireDeTravail.PROTECTION_CODE_G2.Length);
                    }
                }
                //debutProtegerAvant = wordRange.Start;
                debutAbreger = debutProtegerAvant;
                //Passage en abrege des occurences (et remise a jour des positions
                for (int i = 0; i < DonneesTraitement.Occurences.Count; i++) {
                    if (DonneesTraitement.PositionsOccurences[i] >= debutAbreger && DonneesTraitement.PositionsOccurences[i] < finAbreger) {
                        AppliquerStatutSurOccurence(i, Statut.ABREGE);
                    }
                }
                finAbreger = debutAbreger + LongueurTextAbreger;
                // Reajouter un sous block 1 avant s'il y avait d'autres mots
                //if (tailleProtegerAvant > 0) {
                //    wordRange = ProtegerBloc(_document, _document.Range(debutProtegerAvant, debutProtegerAvant + tailleProtegerAvant));
                //    DonneesTraitement.AjouterBlocIntegral(wordRange.Start, wordRange.End, true);
                //    debutAbreger = wordRange.End;
                //    finAbreger = debutAbreger + LongueurTextAbreger;
                //}
                // Reajouter un bloc a la fin s'il y avait d'autres mots
                if (tailleProtegerApres > 0) {
                    wordRange = ProtegerBloc(_document, _document.Range(finAbreger, finAbreger + tailleProtegerApres));
                    DonneesTraitement.AjouterBlocIntegral(finAbreger, finAbreger + tailleProtegerApres, true);
                }

            }
            ChargerTexteEnMemoire();
        }

        public static MSWord.Range AbregerBloc(Document current, MSWord.Range wordRange)
        {
            //if (wordRange.Words.Count == 0) {
            //    // Si la sélection est vide, on ne fait rien
            //    return wordRange;
            //}

            //if (wordRange.Words.Count == 1) {
            //    // Un seul mot dans la sélection, on le resélectionne
            //    wordRange = wordRange.Words.First;
            //} else {
            //    // Plusieurs mots, on sélectionne tout
            //    wordRange = current.Range(
            //        wordRange.Words.First.Start,
            //        wordRange.Words.Last.End
            //    );
            //}
            //wordRange.Select();
            string trimmedStart = wordRange.Text.TrimStart();
            int trimmedCount = wordRange.Text.Length - trimmedStart.Length;
            if (trimmedCount > 0) {
                // Reselectionner le mot sans les espaces de début
                wordRange = current.Range(wordRange.Start + trimmedCount, wordRange.End);
                //wordRange.Select();
            }

            string trimmedEnd = wordRange.Text.TrimEnd();
            if (trimmedEnd.EndsWith("\r")) trimmedEnd = trimmedEnd.Substring(0, trimmedEnd.Length - 1);
            trimmedCount = wordRange.Text.Length - trimmedEnd.Length;
            if (trimmedCount > 0) {
                // Reselectionner le mot sans les espaces de début
                wordRange = current.Range(wordRange.Start, wordRange.End - trimmedCount);
                //wordRange.Select();
            }

            string previousText = current.Range(
                    0,
                    wordRange.Start
                ).Text ?? "";
            string currentText = current.Range(
                wordRange.Start,
                wordRange.End
            ).Text ?? "";
            string nextText = current.Range(
                wordRange.End,
                current.Content.End
            ).Text ?? "";
            int start = wordRange.Start;
            int end = wordRange.End;

            Range toDeleteBefore = null;
            Range toDeleteAfter = null;
            if (previousText.EndsWith(DictionnaireDeTravail.PROTECTION_CODE_G1)) {
                toDeleteBefore = current.Range(
                    start - DictionnaireDeTravail.PROTECTION_CODE_G1.Length,
                    start
                );
            }
            if (currentText.StartsWith(DictionnaireDeTravail.PROTECTION_CODE_G1)) {
                toDeleteBefore = current.Range(
                    start,
                    start + DictionnaireDeTravail.PROTECTION_CODE_G1.Length
                );
            }
            if (currentText.EndsWith(DictionnaireDeTravail.PROTECTION_CODE_G2)) {
                toDeleteAfter = current.Range(
                    end - DictionnaireDeTravail.PROTECTION_CODE_G2.Length,
                    end
                );
            }
            if (nextText.StartsWith(DictionnaireDeTravail.PROTECTION_CODE_G2)) {
                toDeleteAfter = current.Range(
                    end,
                    end + DictionnaireDeTravail.PROTECTION_CODE_G2.Length
                );
            }
            if (toDeleteBefore != null && toDeleteAfter != null) {
                toDeleteAfter.Select();
                toDeleteAfter.Delete();
                toDeleteBefore.Select();
                toDeleteBefore.Delete();
                wordRange = current.Range(
                    toDeleteBefore.End,
                    toDeleteAfter.Start
                );
            }
            return wordRange;
        }

        

        /// <summary>
        /// Vérifie si le traitement est terminé, c'est à dire si toutes les occurences ont été traité
        /// </summary>
        /// <param name="wordRange"></param>
        public bool EstTerminer()
        {
            // Si toute les occurences ont vu leur statuts attribué et appliquer
            return DonneesTraitement.StatutsOccurences.All(s => s != Statut.INCONNU) && DonneesTraitement.EstTraitee.All(s => s);
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
            string occurenceSelected = DonneesTraitement.Occurences[Occurence];
            string wordSelected = occurenceSelected.ToLower();
            List<string> keys = DonneesTraitement.CarteMotOccurences.Keys.ToList();
            int newWordKeyIndex = (keys.Count + keys.IndexOf(wordSelected) + 1) % keys.Count;
            return SelectionnerOccurenceMot(keys[newWordKeyIndex], andOccurence);
        }

        public MSWord.Range PrecedentMot(int andOccurence = 0)
        {
            string occurenceSelected = DonneesTraitement.Occurences[Occurence];
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
            string wordSelected = MotSelectionne.ToLower();

            // listes des occurences a marquer comme étant ignorer, c'est à dire a ne pas afficher dans le traitement
            List<int> wordOccurences = DonneesTraitement.CarteMotOccurences[wordSelected];
            foreach (var occurence in DonneesTraitement.CarteMotOccurences[wordSelected])
            {
                DonneesTraitement.StatutsOccurences[occurence] = Statut.IGNORE;
                DonneesTraitement.EstTraitee[occurence] = true;
            }

            // Sélectionner le mot suivant
            return ProchainMot();
        }

        public MSWord.Range ProchaineOccurenceDuMot()
        {
            string occurenceSelected = DonneesTraitement.Occurences[Occurence];
            string wordSelected = occurenceSelected.ToLower();
            List<int> occurences = DonneesTraitement.CarteMotOccurences[wordSelected];
            int wordOccurenceIndex =
                (occurences.Count + occurences.IndexOf(Occurence) + 1) % occurences.Count;
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
                    Occurence = wordSelectedOccurence;
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
                Occurence = -1;
            }

            return SelectedRange;
        }

        public MSWord.Range PrecedenteOccurenceDuMot()
        {
            string occurenceSelected = DonneesTraitement.Occurences[Occurence];
            string wordSelected = occurenceSelected.ToLower();
            List<int> occurences = DonneesTraitement.CarteMotOccurences[wordSelected];
            int wordOccurenceIndex = occurences.IndexOf(Occurence);
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
            if (newSelectedOccurenceIndex < 0 || newSelectedOccurenceIndex >= DonneesTraitement.Occurences.Count) {
                throw new Exception($"L'occurence {newSelectedOccurenceIndex} n'est pas valide");
            }
            Occurence = newSelectedOccurenceIndex;

            int startPosition = DonneesTraitement.PositionsOccurences[Occurence];
            string mot = DonneesTraitement.Occurences[Occurence];
            SelectedRange = _document.Range(startPosition, startPosition + mot.Length);
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
            Occurence = DonneesTraitement.CarteMotOccurences[key.ToLower()][indexCarte];
            int startPosition = DonneesTraitement.PositionsOccurences[Occurence];
            SelectedRange = _document.Range(startPosition, startPosition + MotSelectionne.Length);
            
            return SelectedRange;
        }

        /// <summary>
        /// Selectionne la prochaine occurence de mot dans le document
        /// </summary>
        /// <returns></returns>
        public MSWord.Range ProchaineOccurence()
        {
            Occurence =
                (DonneesTraitement.Occurences.Count + Occurence + 1)
                % DonneesTraitement.Occurences.Count;
            string occurenceSelected = DonneesTraitement.Occurences[Occurence];
            string contextBefore = DonneesTraitement.ContextesAvantOccurences[Occurence];
            int lastOccurenceEnd = SelectedRange.End;
            int documentEnd = _document.Content.End;
            Range toEnd =
                Occurence == 0
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
                    Occurence == 0
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
                && (StatutOccurence == Statut.INCONNU || !OccurenceEstTraitee) /*|| EstTerminer()*/
            )
            {
                return SelectedRange;
            }
            int counter = Occurence;
            while (counter < DonneesTraitement.Occurences.Count)
            {
                Range selected = ProchaineOccurence();
                if (StatutOccurence == Statut.INCONNU || !OccurenceEstTraitee)
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
            Occurence =
                (DonneesTraitement.Occurences.Count + Occurence - 1)
                % DonneesTraitement.Occurences.Count;
            string occurenceSelected = DonneesTraitement.Occurences[Occurence];
            string wordSelected = occurenceSelected.ToLower();
            string contextBefore = DonneesTraitement.ContextesAvantOccurences[Occurence];
            int wordOccurenceIndex = DonneesTraitement.CarteMotOccurences[wordSelected].IndexOf(
                Occurence
            );
            Range toBegining =
                Occurence == DonneesTraitement.Occurences.Count - 1
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
            int offsetCurrent = 0;
            int offsetNext = 0;
            var previousStatut = DonneesTraitement.StatutsOccurences[index];
            if(previousStatut == statut || (previousStatut == Statut.INCONNU && statut == Statut.ABREGE)) {
                DonneesTraitement.AppliquerStatut(index, statut, offsetCurrent, offsetNext);
                // Si le statut est déjà appliqué ou pas de changement attendu dans le document, on s'arrete la
                return;
            }
            // Sinon traitement de changement
            Range current =
                index != Occurence ? SelectionnerOccurence(index) : SelectedRange;
            int start = current.Start;
            int end = current.End;
            int newStart = start;
            // Je pars du principe que le code de protection est en dehors du texte de l'occurence
            bool estProtegerBloc = EstProtegerBloc(current.Start);
            bool estProtegerMot = EstProtegerMot(_document, current);
            
            if (!estProtegerBloc) {
                // L'occurence n'est pas dans un bloc de protection, on peut appliquer le statut
                if (statut == Statut.PROTEGE && !estProtegerMot) {
                    Range t = Proteger(current);
                    newStart = t.Start;
                    int newEnd = t.End;
                    bool estMaintenantProtegerBloc = EstProtegerBloc(t.Start); // Si protection d'un mot composé
                    offsetCurrent = estMaintenantProtegerBloc ? DictionnaireDeTravail.PROTECTION_CODE_G1.Length : DictionnaireDeTravail.PROTECTION_CODE.Length;
                    offsetNext = estMaintenantProtegerBloc ? DictionnaireDeTravail.PROTECTION_CODE_G2.Length : 0;
                    if (estMaintenantProtegerBloc) { // Garder une trace du bloc integral
                        DonneesTraitement.AjouterBlocIntegral(newStart,newEnd);
                    }
                    
                } else if (estProtegerMot && (statut == Statut.ABREGE || statut == Statut.INCONNU)) {
                    Range t = Abreger(current);
                    offsetCurrent -= DictionnaireDeTravail.PROTECTION_CODE.Length;
                }
            } else if(statut == Statut.ABREGE || statut == Statut.INCONNU) {
                // Si le bloc est uniquement un bloc de protection de mot composé
                // On peut supprimer le bloc
                int indexBloc = DonneesTraitement.ListeBlocsIntegral.FindIndex(b => b.Item1 <= start && end <= b.Item2);
                if(indexBloc > -1) {
                    int blockStart = DonneesTraitement.ListeBlocsIntegral[indexBloc].Item1;
                    int blockEnd = DonneesTraitement.ListeBlocsIntegral[indexBloc].Item2;
                    if(blockStart == start && blockEnd == end) {
                        Range t = AbregerBloc(_document, _document.Range(blockStart, blockEnd));
                        DonneesTraitement.SupprimerBlocIntegral(indexBloc);
                        offsetCurrent -= DictionnaireDeTravail.PROTECTION_CODE_G1.Length;
                        offsetNext -= DictionnaireDeTravail.PROTECTION_CODE_G2.Length;
                    }
                } else if (estProtegerMot) {
                    // Cas d'une occurence modifier après protection d'un block
                    Range t = Abreger(current);
                    offsetCurrent -= DictionnaireDeTravail.PROTECTION_CODE.Length;
                }
            } // Statut IGNORER ou inconnu conserver mais non traiter
            // Sauvegarder le statut de l'occurence et les décallages de positions ajouter par le traitement
            DonneesTraitement.AppliquerStatut(index, statut, offsetCurrent, offsetNext);
        }


        /// <summary>
        /// Initialise le statut d'une occurence dans le document
        /// (Quasi identique à AppliquerStatutSurOccurence mais sans controle de bloc via le document)
        /// </summary>
        /// <param name="index"></param>
        /// <param name="statut"></param>
        public void InitialiserStatutSurOccurence(int index, Statut statut)
        {
            int offsetCurrent = 0;
            int offsetNext = 0;
            var previousStatut = DonneesTraitement.StatutsOccurences[index];
            if (previousStatut == statut || (previousStatut == Statut.INCONNU && statut == Statut.ABREGE)) {
                DonneesTraitement.AppliquerStatut(index, statut, offsetCurrent, offsetNext);
                // Si le statut est déjà appliqué ou pas de changement attendu dans le document, on s'arrete la
                return;
            }
            
            if (previousStatut == Statut.PROTEGE && (statut == Statut.ABREGE || statut == Statut.IGNORE)) {
                Range current = SelectionnerOccurence(index);
                int start = current.Start;
                int end = current.End;
                int indexBloc = DonneesTraitement.ListeBlocsIntegral.FindIndex(b => b.Item1 <= start && end <= b.Item2);
                if (indexBloc > -1) { // Mot composé
                    int blockStart = DonneesTraitement.ListeBlocsIntegral[indexBloc].Item1;
                    int blockEnd = DonneesTraitement.ListeBlocsIntegral[indexBloc].Item2;
                    if (blockStart == start && blockEnd == end) {
                        Range t = AbregerBloc(_document, _document.Range(blockStart, blockEnd));
                        DonneesTraitement.SupprimerBlocIntegral(indexBloc);
                        offsetCurrent -= DictionnaireDeTravail.PROTECTION_CODE_G1.Length;
                        offsetNext -= DictionnaireDeTravail.PROTECTION_CODE_G2.Length;
                    }
                } else { // Mot simple
                    Range t = Abreger(current);
                    int newOffet = t.Start - start;
                    offsetCurrent = newOffet;
                    //if(newOffet == DictionnaireDeTravail.PROTECTION_CODE_G1.Length) {
                    //    // Suppression de protection de bloc (mais ne devrait pas apparaitre ici
                    //    offsetNext -= DictionnaireDeTravail.PROTECTION_CODE_G2.Length;
                    //}
                }
            } else if(statut == Statut.PROTEGE && (previousStatut == Statut.INCONNU|| previousStatut == Statut.ABREGE)) {
                Range current = SelectionnerOccurence(index);
                int start = current.Start;
                int end = current.End;
                Range t = Proteger(current);
                int newStart = t.Start;
                int newEnd = t.End;
                int newOffet = t.Start - start;
                offsetCurrent = newOffet;
                if (newOffet == DictionnaireDeTravail.PROTECTION_CODE_G1.Length) {
                    // Suppression de protection de bloc (mais ne devrait pas apparaitre ici
                    offsetNext = DictionnaireDeTravail.PROTECTION_CODE_G2.Length;
                    DonneesTraitement.AjouterBlocIntegral(newStart, newEnd);
                }
            }
            DonneesTraitement.AppliquerStatut(index, statut, offsetCurrent, offsetNext);
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
            // Sauvegarde, sinon risque de corruption du texte en mémoire
            _document.Save();
            this.TexteEnMemoire = DocumentMainRange;
        }

        /// <summary>
        /// Sélectionner l'occurence en mémoire et remet le focus dessus dans word
        /// </summary>
        /// <param name="index"></param>
        public void AfficherOccurence(int index)
        {
            if (index != Occurence)
            {
                SelectionnerOccurence(index);
            }
            SelectedRange.Select();
        }

        #endregion

        #region backup code inutilisé


        //public void AppliquerStatutSurBlock(int start, int end, Statut statut)
        //{
        //    int blocIndexOfStart = DonneesTraitement.ListeBlocsIntegral.FindIndex(b => b.Item1 <= start && start <= b.Item2);
        //    int blocIndexOfEnd = DonneesTraitement.ListeBlocsIntegral.FindIndex(b => b.Item1 <= end && end <= b.Item2);

        //    int newBlockStart = blocIndexOfStart > -1 ? DonneesTraitement.ListeBlocsIntegral[blocIndexOfStart].Item1 : start;
        //    int newBlockEnd = blocIndexOfEnd > -1 ? DonneesTraitement.ListeBlocsIntegral[blocIndexOfEnd].Item2 : end;

        //    // Ajouter le nouveau block a la fin
        //    DonneesTraitement.DebutsBlocsIntegrals.Add(newBlockStart);
        //    DonneesTraitement.FinsBlocsIntegrals.Add(newBlockEnd);

        //    int decalage = 0;
        //    // En parcours arriere, supprimer les blocks précédents situés dans le nouveau block
        //    for (int i = DonneesTraitement.ListeBlocsIntegral.Count - 2; i >= 0; i--) {
        //        if (newBlockStart <= DonneesTraitement.ListeBlocsIntegral[i].Item1 && DonneesTraitement.ListeBlocsIntegral[i].Item2 <= newBlockEnd) {
        //            DonneesTraitement.DebutsBlocsIntegrals.RemoveAt(i);
        //            DonneesTraitement.FinsBlocsIntegrals.RemoveAt(i);
        //            i--;
        //        }
        //    }

        //    Range toProtect = _document.Range(newBlockStart, newBlockEnd);
        //    if (!toProtect.Text.StartsWith(DictionnaireDeTravail.PROTECTION_CODE_G1)) {
        //        toProtect.InsertBefore(DictionnaireDeTravail.PROTECTION_CODE_G1);
        //        decalage += DictionnaireDeTravail.PROTECTION_CODE_G1.Length;
        //    }
        //    if (!toProtect.Text.EndsWith(DictionnaireDeTravail.PROTECTION_CODE_G2)) {
        //        toProtect.InsertAfter(DictionnaireDeTravail.PROTECTION_CODE_G2);
        //        decalage += DictionnaireDeTravail.PROTECTION_CODE_G2.Length;
        //    }

        //    DonneesTraitement.DebutsBlocsIntegrals.Sort();
        //    DonneesTraitement.FinsBlocsIntegrals.Sort();
        //}

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
        #endregion
    }


}
