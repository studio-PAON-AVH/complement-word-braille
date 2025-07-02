using fr.avh.braille.dictionnaire;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using static fr.avh.braille.dictionnaire.Abreviation;
using Window = System.Windows.Window;

namespace fr.avh.braille.addin
{
    // TODO : Exception handling sur tous les boutons pour eviter les crash
    /// <summary>
    /// Logique d'interaction pour ProtectionIterativeDialog.xaml
    /// </summary>
    public partial class ProtectionInteractiveParOccurenceDialog : Window
    {
        ProtectionWord protecteur;

        private static readonly string dialogTitleTemplate =
            "Action pour le mot {0} dans le document";
        private static readonly string ProtegeDansXDocumentTemplate = "Protégé dans {0} documents";
        private static readonly string AbregeDansXDocumentTemplate = "Abrègé dans {0} documents";
        private static readonly string DetecteDansXDocumentTemplate = "Détecté dans {0} documents";
        private static readonly string MotSelectionneTemplate = "{0}";
        private static readonly string NbOccurenceTemplate =
            "Nombre d'occurence dans le document : {0}";

        private static readonly string RegleAbreviationTemplate = "Abreviation détecté : {0}";

        private static readonly string ProgressionXSurYTemplate = "Progression : {0} / {1}";

        ObservableCollection<MotAfficher> mots = new ObservableCollection<MotAfficher>();

        public bool IsClosed { get; private set; } = false;

        public bool UserIsWarnedAboutEnd = false;

        private int _indexDuMotSelectionner = -1;
        private int _indexDuMotSelectionnerDansLordre = -1;

        List<string> MotsSelectionnable = new List<string>();
        /// <summary>
        /// Liste des mots par ordre de premiere apparation dans le document
        /// </summary>
        List<string> MotsSelectionnablesOrdonnes = new List<string>();

        public class SelectableObject<T>
        {
            public bool IsSelected { get; set; }
            public T ObjectData { get; set; }

            public SelectableObject(T objectData)
            {
                ObjectData = objectData;
            }

            public SelectableObject(T objectData, bool isSelected)
            {
                IsSelected = isSelected;
                ObjectData = objectData;
            }
        }

        ObservableCollection<SelectableObject<Statut>> SelectionStatus = new ObservableCollection<SelectableObject<Statut>>()
        {
            {new SelectableObject<Statut>(Statut.INCONNU, true)},
            {new SelectableObject<Statut>(Statut.ABREGE, true)},
            {new SelectableObject<Statut>(Statut.PROTEGE, true)},
            {new SelectableObject<Statut>(Statut.IGNORE, false)}, // Par défaut ne pas afficher les mots en statuts ignorer
        };

        // recalcul la liste des mots a afficher a partir des filtres
        List<Statut> StatutsAfficher {
            get => SelectionStatus.Where(s => s.IsSelected)
            .Select(s => s.ObjectData)
            .ToList();
        }

        private int _occurenceATraiter = 0;

        public ProtectionInteractiveParOccurenceDialog(ProtectionWord protecteur)
        {
            InitializeComponent();
            this.protecteur = protecteur;
            StatusFilter.ItemsSource = SelectionStatus;

            //MotsSelectionnable = protecteur.DonneesTraitement.CarteMotOccurences.Keys.ToList();
            //MotsSelectionnable.Sort();
            //SelecteurMot.ItemsSource = MotsSelectionnable;

            int selectable = protecteur.DonneesTraitement.CarteMotOccurences[protecteur.MotSelectionne].FindIndex(
                o => StatutsAfficher.Contains(protecteur.DonneesTraitement.StatutsOccurences[o])
            );

            // Sélectionner la premiere occurence affiché, ou la premiere occurence du mot si aucune occurrence 
            // ne correspond aux filtres
            _occurenceATraiter = protecteur.DonneesTraitement.PositionsOccurences.FindIndex(
                    p => p >= protecteur.PositionCurseur
                );
            _occurenceATraiter = _occurenceATraiter >= 0 ? _occurenceATraiter : 0;
            Range next = protecteur.SelectionnerOccurence(_occurenceATraiter);

            //Range next = protecteur.SelectionnerOccurenceMot(protecteur.MotSelectionne, Math.Max(0, selectable));
            RechargerFenetre();
            next.Select();
        }

        /// <summary>
        /// Sélection de la prochaine occurence a traiter
        /// </summary>
        /// <param name="reselectionMot">Si vrai, permet la reselection de l'occurence courante (pour reprise/relancement du traitement après un arret)</param>
        private void SelectionProchainMotATraiter(bool reselectionMot = false)
        {
            int safety = 0;
            Range next = reselectionMot ? protecteur.SelectedRange : protecteur.ProchainMot();
            bool hasStatutNonAppliquer = false;
            do {
                List<int> occurenceMot = protecteur.DonneesTraitement.CarteMotOccurences[
                    protecteur.MotSelectionne
                ];
                for(int i = 0; i < occurenceMot.Count && !hasStatutNonAppliquer; i++) {
                    if(!protecteur.DonneesTraitement.StatutsAppliquer[occurenceMot[i]]) {
                        hasStatutNonAppliquer = true;
                    }
                }
                if (!hasStatutNonAppliquer)
                {
                    next = protecteur.ProchainMot();
                    hasStatutNonAppliquer = false;
                }
                safety++;
            } while (!hasStatutNonAppliquer && safety < protecteur.DonneesTraitement.CarteMotOccurences.Keys.Count);
            RechargerFenetre();
            next.Select();
        }

        

        private void SelectionProchainMot()
        {
            Range next;
            int safety = 0;
            do { // NP : continuer tant qu'on est sur un statut de mot ignoré
                next = protecteur.ProchainMot();
            } while(protecteur.StatutOccurence == Statut.IGNORE
                    && safety < protecteur.DonneesTraitement.CarteMotOccurences.Count);
            RechargerFenetre();
            next.Select();
        }

        private void ProtegerIci_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(protecteur.MotSelectionne))
            {
                foreach (var mot in mots)
                {
                    protecteur.DonneesTraitement.StatutsOccurences[mot.Index] = Statut.PROTEGE;
                }
                VueDictionnaire_Refresh();
            }
        }

        private void AbregerIci_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(protecteur.MotSelectionne))
            {
                foreach (var mot in mots)
                {
                    protecteur.DonneesTraitement.StatutsOccurences[mot.Index] = Statut.ABREGE;
                }
                VueDictionnaire_Refresh();
            }
        }

        private bool _hasFinishedReview = false;

        //private void Next_Click(object sender, RoutedEventArgs e)
        //{
        //    System.Threading.Tasks.Task.Run(() =>
        //    {
        //        AppliquerStatuts();
                
        //        Dispatcher.Invoke(() =>
        //        {
                    
        //            if (
        //                protecteur.SelectedWord.ToLower()
        //                == protecteur.DonneesTraitement.CarteMotOccurences.Keys.Last()
        //            ) { // Si on est sur le dernier mot
        //                if (!protecteur.EstTerminer()) { // S'il reste des éléments à traiter (pour lesquels un statut n'a pas été choisi)
        //                    if (!_hasFinishedReview)
        //                    { // Si on est pas déja en mode sélectif
        //                        // on averti l'utilisateur qu'il reste des éléments à traiter et que le parcours va repartir de la premiere occurence non traité
        //                        MessageBox.Show(
        //                            "Des occurences sans décision sont toujours présentes, le parcours va maintenant se concentrer sur les mots ayant des occurences sans statuts.",
        //                            "Des occurences en statut inconnus sont toujours présentes",
        //                            MessageBoxButton.OK,
        //                            MessageBoxImage.Warning
        //                        );
        //                        // On change de mode
        //                        _hasFinishedReview = true;
        //                    }
        //                    // Repartir de la premiere occurence et selectionner la première occurence en statut Ignorer
        //                    protecteur.SelectionnerOccurence(0);
        //                    SelectionProchainMotATraiter(true);
        //                }
        //                else // Si on est sur le dernier mot et qu'il n'y a plus d'occurence à traiter
        //                {
        //                    // on averti l'utilisateur qu'il n'y a plus d'occurence à traiter et qu'il peut fermer la fenetre pour terminer le traitement
        //                    var message = MessageBox.Show(
        //                        "Tous les mots du documents identifiés ont été traités. Voulez-vous fermer la fenêtre de traitement ?",
        //                        "Fin de l'analyse",
        //                        MessageBoxButton.YesNo,
        //                        MessageBoxImage.Question
        //                    );
        //                    // Désactivation du mode passage en revue sélectif
        //                    _hasFinishedReview = false;
        //                    if (message == MessageBoxResult.Yes) { // Si l'utilisateur a confirmer vouloir a fermer la fenettre
        //                        // On revient au début et on ferme la fenetre
        //                        protecteur.SelectionnerOccurence(0);
        //                        Dispatcher.Invoke(() => Close());
        //                    }
        //                    else { // Sinon on passe au mot suivant
        //                        SelectionProchainMot();
        //                    }
        //                }
        //            }
        //            else
        //            { // Cas général
        //                if (_hasFinishedReview)
        //                { // si on est en mode parcours d'occurences non traité
        //                    if (!protecteur.EstTerminer())
        //                    {   // s'il reste des occurences non traitées
        //                        // Selectionner la prochaine occurence en statut Inconnu
        //                        SelectionProchainMotATraiter();
        //                    }
        //                    else
        //                    { // Sinon on a terminé le traitement complet en mode review selected
        //                        // on averti l'utilisateur qu'il n'y a plus d'occurence à traiter et qu'il peut fermer la fenetre pour terminer le traitement
        //                        var message = MessageBox.Show(
        //                            "Tous les mots du documents identifiés ont été traités. Voulez-vous fermer la fenêtre ?",
        //                            "Fin de l'analyse",
        //                            MessageBoxButton.YesNo,
        //                            MessageBoxImage.Question
        //                        );
        //                        // Désactivation du mode passage en revue sélectif
        //                        _hasFinishedReview = false;
        //                        // on revient a la première occurence détectée
        //                        var range = protecteur.SelectionnerOccurence(0);
        //                        // Si l'utilisateur a demander a fermer la fenettre
        //                        if (message == MessageBoxResult.Yes)
        //                        {
        //                            Dispatcher.Invoke(() => Close());
        //                        }
        //                        else
        //                        {
        //                            // Sinon on remet en avant le mot et on recharge la fenetre
        //                            range?.Select();
        //                            RechargerFenetre();
        //                        }
        //                    }
        //                }
        //                else
        //                { // mode normal de traitement sur la première passe, on passe au mot suivant
        //                    SelectionProchainMot();
        //                }
        //            }
        //        });
        //    });
        //}

        /// <summary>
        /// Appliquer les statuts sur les occurences du mot sélectionné
        /// </summary>
        //private void AppliquerStatuts()
        //{
        //    string mot = protecteur.SelectedWord;
        //    List<int> occurences = protecteur.DonneesTraitement.CarteMotOccurences[
        //            protecteur.SelectedWord
        //        ].Where(
        //            i => protecteur.DonneesTraitement.StatutsOccurences[i] == Statut.ABREGE 
        //                || protecteur.DonneesTraitement.StatutsOccurences[i] == Statut.PROTEGE
        //                || protecteur.DonneesTraitement.StatutsOccurences[i] == Statut.IGNORE
        //        ).ToList();
        //    Dispatcher.Invoke(() =>
        //    {
        //        ProgressAnalyse.Maximum = occurences.Count;
        //        ProgressAnalyse.Value = 0;
        //    });


        //    for(int i = 0; i < occurences.Count; i++) {
        //        protecteur.SelectionnerOccurenceMot(mot, i).Select();
        //        protecteur.AppliquerStatutSurOccurence(protecteur.SelectedOccurence, protecteur.SelectedOccurenceStatut);

        //        Dispatcher.Invoke(() =>
        //        {
        //            ProgressAnalyse.Value = i+1;
        //            ProgressIndicator.Content =
        //                $"{ProgressAnalyse.Value} / {ProgressAnalyse.Maximum}";
        //            this.UpdateLayout();
        //        });
        //    }
        //    protecteur.ChargerTexteEnMemoire();
        //}

        private void Previous_Click(object sender, RoutedEventArgs e)
        {
            Range previous;
            int safety = 0;
            do {
                // NP : continuer tant qu'on est sur un statut de mot ignoré
                previous = protecteur.PrecedentMot();
                safety++;
            } while(protecteur.StatutOccurence == Statut.IGNORE 
                    && safety < protecteur.DonneesTraitement.CarteMotOccurences.Count);
            RechargerFenetre();
            previous.Select();
        }

        private void ProtectionDialog_Load(object sender, EventArgs e)
        {
            SelectionProchainMotATraiter();
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            protecteur.Save();
            // Ask users if he wants to apply all decision on the document before closing
            //if(System.Windows.MessageBox.Show(
            //    "Voulez-vous appliquer les décisions prises sur le document avant de fermer la fenêtre ?",
            //    "Appliquer les décisions",
            //    MessageBoxButton.YesNo,
            //    MessageBoxImage.Question
            //) == MessageBoxResult.Yes) {
            //    Dispatcher.Invoke(() =>
            //    {
            //        ProgressAnalyse.Maximum = protecteur.DonneesTraitement.Occurences.Count;
            //        ProgressAnalyse.Value = 0;
            //    });
            //    protecteur.AppliquerStatutsSurDocument((m,p) => {
            //        Dispatcher.Invoke(() =>
            //        {
            //            ProgressAnalyse.Value = p.Item1;
            //            ProgressAnalyse.Maximum = p.Item2;
            //        });
            //    });
            //}
        }

        private void VueOccurences_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var dataGrid = sender as DataGrid;
            var selectedItem = dataGrid.SelectedItem as MotAfficher;
            var selectionIndex = dataGrid.SelectedIndex;
            if (selectedItem != null)
            {
                protecteur.SelectionnerOccurence(selectedItem.Index).Select();
            }
        }

        private bool _hasChanged = false;

        /// <summary>
        /// Rechargement des informations de la fenêtre à partir du mot sélectionné dans le protecteur
        /// </summary>
        private void RechargerFenetre()
        {
            protecteur.SelectedRange.Select();
            Title = string.Format(dialogTitleTemplate, protecteur.MotSelectionne);
            //Previous.IsEnabled = protecteur.SelectedOccurence > 0;
            
            MotSelectionne.Content = string.Format(MotSelectionneTemplate, protecteur.MotSelectionne);
            int motsUniquesTraites = protecteur.DonneesTraitement.CarteMotOccurences.Keys
                .ToList()
                .IndexOf(protecteur.MotSelectionne);

            MotsSelectionnable = protecteur.DonneesTraitement.CarteMotOccurences.Keys.OrderBy(k => k).ToList();
            //SelecteurMot.ItemsSource = MotsSelectionnable;
            _indexDuMotSelectionner = MotsSelectionnable.IndexOf(protecteur.MotSelectionne);
            //SelecteurMot.SelectedIndex = _indexDuMotSelectionner;

            
           MotsSelectionnablesOrdonnes = protecteur.DonneesTraitement.CarteMotOccurences.Keys.OrderBy(
                k => {
                    List<int> validOccurence = protecteur.DonneesTraitement.CarteMotOccurences[k]
                        //.Where(i => protecteur.DonneesTraitement.StatutsOccurences[i] != Statut.IGNORE)
                        .ToList();
                    if(validOccurence.Count == 0) {
                        return int.MaxValue; // mot ignorer en fin de tri
                    }
                    return protecteur.DonneesTraitement.PositionsOccurences[
                            validOccurence[0]
                        ];
                    }
            ).ToList();
            //MotDansOrdreDocument.ItemsSource = MotsSelectionnablesOrdonnes.Select((s, i) => $"{s} - {i+1}");
            _indexDuMotSelectionnerDansLordre = MotsSelectionnablesOrdonnes.IndexOf(protecteur.MotSelectionne);
            //MotDansOrdreDocument.SelectedIndex = _indexDuMotSelectionnerDansLordre;

            ProgressionXSurY.Content = string.Format(
                ProgressionXSurYTemplate, protecteur.Occurence + 1, protecteur.DonneesTraitement.Occurences.Count
            );

            NbOccurence.Content = string.Format(
                NbOccurenceTemplate,
                protecteur.DonneesTraitement.CarteMotOccurences[protecteur.MotSelectionne].Count
            );
            RegleAbreviation.Content = string.Format(
                RegleAbreviationTemplate,
                Abreviation.RegleAppliquerSur(protecteur.MotSelectionne) ?? "aucune"
            );
            var listMot = protecteur.alreadyInDB
                .Where(x => x.Texte == protecteur.MotSelectionne.ToLower())
                .ToList();
            if (listMot.Count > 0)
            {
                // Le mot existe dans la base,
                // on met a jour le texte
                ProtegeDansXDocument.Content = string.Format(
                    ProtegeDansXDocumentTemplate,
                    listMot[0].Protections.ToString()
                );
                AbregeDansXDocument.Content = string.Format(
                    AbregeDansXDocumentTemplate,
                    listMot[0].Abreviations.ToString()
                );
                DetecteDansXDocument.Content = string.Format(
                    DetecteDansXDocumentTemplate,
                    listMot[0].Documents.ToString()
                );
                CommentairesMot.Content =
                    listMot[0].Commentaires.Length > 0
                        ? listMot[0].Commentaires
                        : "Pas de commentaires";
            }
            else
            {
                ProtegeDansXDocument.Content = string.Format(
                    ProtegeDansXDocumentTemplate,
                    0
                );
                AbregeDansXDocument.Content = string.Format(AbregeDansXDocumentTemplate, 0);
                DetecteDansXDocument.Content = string.Format(
                   DetecteDansXDocumentTemplate,
                   0
               );
                CommentairesMot.Content = "Le mot n'existe pas dans la base de donnée";
            }
            VueDictionnaire_Refresh();
            //ProgressAnalyse.Value = 0;
            //ProgressIndicator.Content = null;
            _hasChanged = false;
        }

        private void VueDictionnaire_Refresh()
        {
            
            mots = new ObservableCollection<MotAfficher>(
                protecteur.DonneesTraitement
                    .OccurencesAsListOfTuples()
                    .Where(
                        m => m.Item2.ToLower().Trim() == protecteur.MotSelectionne.ToLower().Trim()
                        && StatutsAfficher.Contains(m.Item3)
                    )
                    .Select(
                        (tuple) =>
                            new MotAfficher()
                            {
                                Index = tuple.Item1,
                                Texte = tuple.Item2,
                                Statut = tuple.Item3,
                                ContexteAvant = tuple.Item4,
                                ContexteApres = tuple.Item5
                            }
                    )
            );
            VueOccurences.DataContext = mots;
            // Selection de l'occurence in
        }

        private void Statut_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox _statut = sender as ComboBox;
            MotAfficher selected = (MotAfficher)_statut.DataContext;
            if (selected != null && _statut.SelectedItem != null)
            {
                selected.StatutChoisi = _statut.SelectedItem.ToString();
                protecteur.DonneesTraitement.StatutsOccurences[selected.Index] = selected.Statut;
                protecteur.Save();
                _hasChanged = true;
                //protecteur.AppliquerStatutSurOccurence(selected.Index, selected.Statut);
            }
        }

        private void StatusFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Ne pas afficher la sélection pour ne garder qu'une combobox de filtre
            ComboBox comboBox = (ComboBox)sender;
            comboBox.SelectedItem = null;
        }

        private void StatusFilter_StatusCheckChange(object sender, RoutedEventArgs e)
        {
            VueDictionnaire_Refresh();
        }


        private void Supprimer_Click(object sender, RoutedEventArgs e)
        {
            // Fenetre d'alerte
            // si OK, supprimer le mot du dictionnaire de travail
            // puis recalculer le cache
            // Note NP 2024 10 07 : remplacer la suppression par un marquage en statut ignoré
            var result = MessageBox.Show(
                "Voulez-vous vraiment ignorer ce mot ?\r\n" +
                "Il ne sera plus présenté dans la fenêtre de traitement",
                "Suppression du mot",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question
            );
            if (result == MessageBoxResult.Yes)
            {

                var range = protecteur.IgnorerMotEtSelectionnerSuivant();
                range?.Select();
                RechargerFenetre();
            }
        }


        private void AppliquerDecisionEtPasserAOccurenceSuivante(Statut selectionne, bool surtoutesLesOccurences)
        {
            bool hasNext = _occurenceATraiter < protecteur.DonneesTraitement.Occurences.Count - 1;
            Range selection = protecteur.SelectionnerOccurence(_occurenceATraiter);
            //selection.Select();
            if (surtoutesLesOccurences) {
                int currentOccurence = _occurenceATraiter;
                string wordKey = protecteur.DonneesTraitement.Occurences[currentOccurence].ToLower().Trim();
                foreach (int occurence in protecteur.DonneesTraitement.CarteMotOccurences[wordKey]) {
                    if(occurence <= currentOccurence) {
                        protecteur.AppliquerStatutSurOccurence(occurence, selectionne);
                    } else {
                        // Premarquer les occurence suivantes sans appliquer le statut
                        protecteur.DonneesTraitement.StatutsOccurences[occurence] = selectionne;
                    }
                }
            } else { 
                protecteur.AppliquerStatutSurOccurence(_occurenceATraiter, selectionne);
            }
            selection = protecteur.SelectionnerOccurence(_occurenceATraiter);
            //selection.Select();
            if (hasNext) {
                do {
                    _occurenceATraiter = protecteur.Occurence + 1;
                    selection = protecteur.SelectionnerOccurence(_occurenceATraiter);
                    if( new Statut[] { Statut.PROTEGE, Statut.ABREGE }.Contains(protecteur.StatutOccurence)) {
                        // l'occurence a été marqué mais n'a pas été préalablement traité
                        if (!protecteur.OccurenceEstTraitee) {
                            protecteur.AppliquerStatutSurOccurence(protecteur.Occurence, protecteur.StatutOccurence);
                        }
                        hasNext = _occurenceATraiter < protecteur.DonneesTraitement.Occurences.Count - 1;
                    } else {
                        // l'occurence n'a pas été pré-marqué ou a été précédement ignorer, on passe au traitement de l'occurence sélectionné
                        hasNext = false;
                    }
                } while (hasNext);
                protecteur.ChargerTexteEnMemoire();
                selection = protecteur.SelectionnerOccurence(_occurenceATraiter);
                RechargerFenetre();
                selection.Select();
            } else {
                int firstIgnore = protecteur.DonneesTraitement.StatutsOccurences.IndexOf(
                    Statut.IGNORE
                );
                if(firstIgnore >= 0) {
                    // Si on a des occurences ignorées, on repart de la première occurence ignorée
                    selection = protecteur.SelectionnerOccurence(firstIgnore);
                    _occurenceATraiter = firstIgnore;
                    RechargerFenetre();
                    selection.Select();
                } else {
                    MessageBox.Show(
                        "Tous les mots identifiés a partir de votre sélection de départ ont été traités",
                        "Fin du traitement",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information
                    );
                    this.Close();
                }
            }

        }

        private void IgnorerOccurence_Click(object sender, RoutedEventArgs e)
        {
            AppliquerDecisionEtPasserAOccurenceSuivante(Statut.IGNORE, false);
        }

        private void ProtegerOccurence_Click(object sender, RoutedEventArgs e)
        {
            AppliquerDecisionEtPasserAOccurenceSuivante(Statut.PROTEGE, false);
        }

        private void ProtegerMot_Click(object sender, RoutedEventArgs e)
        {
            AppliquerDecisionEtPasserAOccurenceSuivante(Statut.PROTEGE, true);
        }

        private void AbregerMot_Click(object sender, RoutedEventArgs e)
        {
            AppliquerDecisionEtPasserAOccurenceSuivante(Statut.ABREGE, true);
        }

        private void AbregerOccurence_Click(object sender, RoutedEventArgs e)
        {
            AppliquerDecisionEtPasserAOccurenceSuivante(Statut.ABREGE, false);
        }

        private void Reselectionner_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
