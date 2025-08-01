﻿using fr.avh.braille.dictionnaire;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Window = System.Windows.Window;

namespace fr.avh.braille.addin
{
    // TODO : Exception handling sur tous les boutons pour eviter les crash
    /// <summary>
    /// Logique d'interaction pour ProtectionIterativeDialog.xaml
    /// </summary>
    public partial class ProtectionInteractiveParMotsDialog : Window
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


        List<MotAfficher> OccurencesTraitees = new List<MotAfficher>();
        ObservableCollection<MotAfficher> OccurencesAffichees = new ObservableCollection<MotAfficher>();

        public bool IsClosed { get; private set; } = false;

        public bool UserIsWarnedAboutEnd = false;

        private int _indexDuMotSelectionner = -1;
        private int _indexDuMotSelectionnerDansLordre = -1;


        private bool peutRetraiterMot = false;

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

        public ProtectionInteractiveParMotsDialog(ProtectionWord protecteur, bool peutRetraiter = false)
        {
            InitializeComponent();
            this.protecteur = protecteur;

            
            StatusFilter.ItemsSource = SelectionStatus;
            peutRetraiterMot = peutRetraiter;

            //MotsSelectionnable = protecteur.DonneesTraitement.CarteMotOccurences.Keys.ToList();
            //MotsSelectionnable.Sort();
            //SelecteurMot.ItemsSource = MotsSelectionnable;

            int selectable = protecteur.DonneesTraitement.CarteMotOccurences[protecteur.MotSelectionne].FindIndex(
                o => StatutsAfficher.Contains(protecteur.DonneesTraitement.StatutsOccurences[o])
            );

            if (!peutRetraiter) {
                // reselectionner le premier mot a traiter
                SelectionProchainMotATraiter(true);
            } else {
                // Sélectionner la premiere occurence affiché, ou la premiere occurence du mot si aucune occurrence 
                // ne correspond aux filtres
                Range next = protecteur.SelectionnerOccurenceMot(protecteur.MotSelectionne, Math.Max(0, selectable));
                RechargerFenetre();
                next.Select();
            }
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
                    if(!protecteur.DonneesTraitement.EstTraitee[occurenceMot[i]]) {
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

        private void ProtegerMot_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(protecteur.MotSelectionne))
            {
                foreach (var mot in OccurencesAffichees)
                {
                    mot.Statut = Statut.PROTEGE;
                }
                VueDictionnaire_Refresh();
            }
        }

        private void AbregerMot_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(protecteur.MotSelectionne))
            {
                foreach (var mot in OccurencesAffichees)
                {
                    mot.Statut = Statut.ABREGE;
                }
                VueDictionnaire_Refresh();
            }
        }

        private bool _hasFinishedReview = false;

        private void Next_Click(object sender, RoutedEventArgs e)
        {
            System.Threading.Tasks.Task.Run(() =>
            {
                AppliquerStatuts();
                
                Dispatcher.Invoke(() =>
                {
                    
                    if (
                        protecteur.MotSelectionne.ToLower()
                        == protecteur.DonneesTraitement.CarteMotOccurences.Keys.Last()
                    ) { // Si on est sur le dernier mot
                        if (!protecteur.EstTerminer()) { // S'il reste des éléments à traiter (pour lesquels un statut n'a pas été choisi)
                            if (!_hasFinishedReview)
                            { // Si on est pas déja en mode sélectif
                                // on averti l'utilisateur qu'il reste des éléments à traiter et que le parcours va repartir de la premiere occurence non traité
                                MessageBox.Show(
                                    "Des occurences sans décision sont toujours présentes, le parcours va maintenant se concentrer sur les mots ayant des occurences sans statuts.",
                                    "Des occurences en statut inconnus sont toujours présentes",
                                    MessageBoxButton.OK,
                                    MessageBoxImage.Warning
                                );
                                // On change de mode
                                _hasFinishedReview = true;
                                peutRetraiterMot = true;
                            }
                            // Repartir de la premiere occurence et selectionner la première occurence en statut Ignorer
                            protecteur.SelectionnerOccurence(0);
                            SelectionProchainMotATraiter(true);
                        }
                        else // Si on est sur le dernier mot et qu'il n'y a plus d'occurence à traiter
                        {
                            // on averti l'utilisateur qu'il n'y a plus d'occurence à traiter et qu'il peut fermer la fenetre pour terminer le traitement
                            var message = MessageBox.Show(
                                "Tous les mots du documents identifiés ont été traités. Voulez-vous fermer la fenêtre de traitement ?",
                                "Fin de l'analyse",
                                MessageBoxButton.YesNo,
                                MessageBoxImage.Question
                            );
                            // Désactivation du mode passage en revue sélectif
                            _hasFinishedReview = false;
                            if (message == MessageBoxResult.Yes) { // Si l'utilisateur a confirmer vouloir a fermer la fenettre
                                // On revient au début et on ferme la fenetre
                                protecteur.SelectionnerOccurence(0);
                                Dispatcher.Invoke(() => Close());
                            }
                            else { // Sinon on passe au mot suivant
                                SelectionProchainMot();
                            }
                        }
                    }
                    else
                    { // Cas général
                        if (_hasFinishedReview)
                        { // si on est en mode parcours d'occurences non traité
                            if (!protecteur.EstTerminer())
                            {   // s'il reste des occurences non traitées
                                // Selectionner la prochaine occurence en statut Inconnu
                                SelectionProchainMotATraiter();
                            }
                            else
                            { // Sinon on a terminé le traitement complet en mode review selected
                                // on averti l'utilisateur qu'il n'y a plus d'occurence à traiter et qu'il peut fermer la fenetre pour terminer le traitement
                                var message = MessageBox.Show(
                                    "Tous les mots du documents identifiés ont été traités. Voulez-vous fermer la fenêtre ?",
                                    "Fin de l'analyse",
                                    MessageBoxButton.YesNo,
                                    MessageBoxImage.Question
                                );
                                // Désactivation du mode passage en revue sélectif
                                _hasFinishedReview = false;
                                // on revient a la première occurence détectée
                                var range = protecteur.SelectionnerOccurence(0);
                                // Si l'utilisateur a demander a fermer la fenettre
                                if (message == MessageBoxResult.Yes)
                                {
                                    Dispatcher.Invoke(() => Close());
                                }
                                else
                                {
                                    // Sinon on remet en avant le mot et on recharge la fenetre
                                    range?.Select();
                                    RechargerFenetre();
                                }
                            }
                        }
                        else
                        { // mode normal de traitement sur la première passe, on passe au mot suivant
                            if (!protecteur.EstTerminer()) {   // s'il reste des occurences non traitées
                                // Selectionner la prochaine occurence en statut Inconnu
                                SelectionProchainMotATraiter();
                            } else SelectionProchainMot();
                        }
                    }
                });
            });
        }

        private void ProgressHandler(string message, Tuple<int, int> progress)
        {
            Dispatcher.Invoke(() =>
            {
                ProgressIndicator.Content = message;
                ProgressAnalyse.Maximum = progress.Item2;
                ProgressAnalyse.Value = progress.Item1;
            });
        }

        /// <summary>
        /// Appliquer les statuts sur les occurences du mot sélectionné
        /// </summary>
        private void AppliquerStatuts()
        {
            
            Dispatcher.Invoke(() =>
            {
                ProgressAnalyse.Maximum = OccurencesTraitees.Count;
                ProgressAnalyse.Value = 0;
            });
            int p = 0;
            foreach(var occurence in OccurencesTraitees) {
                // NP : comme certaines décisions peuvent être "prémarquées" sans être appliquer,
                // si l'occurence n'etait pas traitée auparavant, on la repasse en inconnu avant
                // d'appeler la fonction d'application de statut
                if (!protecteur.DonneesTraitement.EstTraitee[occurence.Index]) {
                    protecteur.DonneesTraitement.StatutsOccurences[occurence.Index] = Statut.INCONNU;
                }
                // NP : essai avec la fonction plus rapide qui ne controle pas les blocs existants
                protecteur.InitialiserStatutSurOccurence(occurence.Index, occurence.Statut);
                Dispatcher.Invoke(() =>
                {
                    ProgressAnalyse.Value = ++p;
                    ProgressIndicator.Content =
                        $"{ProgressAnalyse.Value} / {ProgressAnalyse.Maximum}";
                    this.UpdateLayout();
                });
            }
            //for(int i = 0; i < occurences.Count; i++) {
            //    MotAfficher motAfficher = OccurencesAffichees.FirstOrDefault(m => m.Index == occurences[i]);
            //    protecteur.AppliquerStatutSurOccurence(occurences[i], motAfficher.Statut);
            //    //protecteur.SelectionnerOccurenceMot(mot, i).Select();
            //    //protecteur.AppliquerStatutSurOccurence(protecteur.Occurence, protecteur.StatutOccurence)   
            //}
            protecteur.ChargerTexteEnMemoire();
        }

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
            Previous.IsEnabled = protecteur.Occurence > 0;
            
            MotSelectionne.Content = string.Format(MotSelectionneTemplate, protecteur.MotSelectionne);

            // NP 2025 07 08 : les transcripteurs ne veulent pas voir/repasser sur les mots pré traité finalement
            // idee : ne plus rendre selectionnable les mots traités
            MotsSelectionnable = protecteur.DonneesTraitement.CarteMotOccurences
                .Where(kv => peutRetraiterMot ? true : kv.Value.Any(
                    i => !protecteur.DonneesTraitement.EstTraitee[i]
                )).ToDictionary(kv => kv.Key, kv => kv.Value)
                .Keys.OrderBy(k => k).ToList();
            SelecteurMot.ItemsSource = MotsSelectionnable;
            _indexDuMotSelectionner = MotsSelectionnable.IndexOf(protecteur.MotSelectionne);
            SelecteurMot.SelectedIndex = _indexDuMotSelectionner;

            
           MotsSelectionnablesOrdonnes = protecteur.DonneesTraitement.CarteMotOccurences
                .Where(kv => peutRetraiterMot ? true : kv.Value.Any(
                    i => !protecteur.DonneesTraitement.EstTraitee[i]
                )).ToDictionary(kv => kv.Key, kv => kv.Value)
                .Keys.OrderBy(
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
            MotDansOrdreDocument.ItemsSource = MotsSelectionnablesOrdonnes.Select((s, i) => $"{s} - {i+1}");
            _indexDuMotSelectionnerDansLordre = MotsSelectionnablesOrdonnes.IndexOf(protecteur.MotSelectionne);
            MotDansOrdreDocument.SelectedIndex = _indexDuMotSelectionnerDansLordre;

            Total.Content = string.Format(
                "/ {0}", MotsSelectionnablesOrdonnes.Count
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
            OccurencesTraitees = protecteur.DonneesTraitement
                    .OccurencesAsListOfTuples()
                    .Where(
                        m => m.Item2.ToLower().Trim() == protecteur.MotSelectionne.ToLower().Trim()
                    ).Select(
                        (tuple) =>
                            new MotAfficher()
                            {
                                Index = tuple.Item1,
                                Texte = tuple.Item2,
                                Statut = tuple.Item3,
                                ContexteAvant = tuple.Item4,
                                ContexteApres = tuple.Item5
                            }
                    ).ToList();
            VueDictionnaire_Refresh();
            ProgressAnalyse.Value = 0;
            ProgressIndicator.Content = null;
            _hasChanged = false;
        }

        private void VueDictionnaire_Refresh()
        {

            OccurencesAffichees = new ObservableCollection<MotAfficher>(
                OccurencesTraitees.Where(m => StatutsAfficher.Contains(m.Statut))
            );
            
            VueOccurences.DataContext = OccurencesAffichees;
            // Selection de l'occurence in
        }

        private void Statut_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox _statut = sender as ComboBox;
            MotAfficher selected = (MotAfficher)_statut.DataContext;
            if (selected != null && _statut.SelectedItem != null)
            {
                selected.StatutChoisi = _statut.SelectedItem.ToString();
                //protecteur.DonneesTraitement.StatutsOccurences[selected.Index] = selected.Statut;
                //protecteur.Save();
                //_hasChanged = true;
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

        private void SelecteurMot_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (
                SelecteurMot.SelectedIndex >= 0
                && _indexDuMotSelectionner != SelecteurMot.SelectedIndex
            )
            {
                _indexDuMotSelectionner = SelecteurMot.SelectedIndex;
                // TODO : charge le mot sélectionné ici dans la fenêtre
                string mot = SelecteurMot.SelectedValue.ToString();
                var range = protecteur.SelectionnerOccurenceMot(mot);
                range?.Select();
                RechargerFenetre();
            }
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

        private void Progression_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (
                MotDansOrdreDocument.SelectedIndex >= 0
                && _indexDuMotSelectionnerDansLordre != MotDansOrdreDocument.SelectedIndex
            ) {
                _indexDuMotSelectionnerDansLordre = MotDansOrdreDocument.SelectedIndex;
                // TODO : charge le mot sélectionné ici dans la fenêtre
                string mot = MotsSelectionnablesOrdonnes[MotDansOrdreDocument.SelectedIndex];
                var range = protecteur.SelectionnerOccurenceMot(mot);
                range?.Select();
                RechargerFenetre();
            }
        }
    }
}
