using Antlr.Runtime;
using fr.avh.braille.dictionnaire.Entities;
using NHibernate.Mapping;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace fr.avh.braille.dictionnaire
{
    /// <summary>
    /// Logique d'interaction pour Consultation.xaml
    /// </summary>
    public partial class Consultation : Window
    {

        private string _motRechercher = "";

        private Mot _motSelectionner = null;


        public Mot MotSelectionner { 
            get {
                return _motSelectionner;
            } 
            
            set { 
                _motSelectionner = value;
                if (_motSelectionner != null) {
                    ProtectionsValue.IsEnabled = true;
                    ProtectionsValue.Text = _motSelectionner.Protections.ToString();
                    AbreviationsValue.IsEnabled = true;
                    AbreviationsValue.Text = _motSelectionner.Abreviations.ToString();
                    ToujoursDemanderValue.IsEnabled = true;
                    ToujoursDemanderValue.IsChecked = _motSelectionner.ToujoursDemander == 1;
                    CommentairesValue.IsEnabled = true;
                    CommentairesValue.Text = _motSelectionner.Commentaires;
                    SaveButton.IsEnabled = false;
                } else {
                    ProtectionsValue.IsEnabled = false;
                    ProtectionsValue.Text = "0";
                    AbreviationsValue.IsEnabled = false;
                    AbreviationsValue.Text = "0";
                    ToujoursDemanderValue.IsEnabled = false;
                    ToujoursDemanderValue.IsChecked = false;
                    CommentairesValue.IsEnabled = false;
                    CommentairesValue.Text = "";
                }
            }
        }


        public Consultation()
        {
            InitializeComponent();
            ProtectionsValue.IsEnabled = false;
            ProtectionsValue.Text = "";
            AbreviationsValue.IsEnabled = false;
            AbreviationsValue.Text = "";
            ToujoursDemanderValue.IsEnabled = false;
            ToujoursDemanderValue.IsChecked = false;
            CommentairesValue.IsEnabled = false;
            CommentairesValue.Text = "";
            SaveButton.IsEnabled = false;
        }
        public Consultation(string motConsulter) : this()
        {
            SelecteurMot.SelectedIndex = 0;
            InfosMotsTrouver.Dispatcher.Invoke(new Action(() =>
            {
                InfosMotsTrouver.Content = "Recherche en cours ...";
            }));
            SelecteurMot.Dispatcher.Invoke(new Action(() =>
            {
                SelecteurMot.Items.Clear();
            }));
            using (var session = BaseSQlite.CreateSessionFactory().OpenSession()) {
                List<Mot> trouver = session.Query<Mot>().Where(
                    m => m.Texte == motConsulter
                ).ToList();
                InfosMotsTrouver.Dispatcher.Invoke(new Action(() =>
                {
                    InfosMotsTrouver.Content = trouver.Count == 0 ? "Aucun mot trouvé" : $"{trouver.Count} mots trouvés";
                }));
                SelecteurMot.Dispatcher.Invoke(new Action(() =>
                {
                    SelecteurMot.Items.Clear();
                    SelecteurMot.SelectedIndex = -1;
                    int i = -1;
                    foreach (var mot in trouver) {
                        
                        i++;
                        SelecteurMot.Items.Add(mot.Texte);
                        if (_motRechercher.ToLower() == mot.Texte) {
                            SelecteurMot.SelectedIndex = i;
                        }
                    }
                    // Pas de correspondance exacte dans la base de donnée
                    if (SelecteurMot.SelectedIndex == -1) {
                        // On ajoute une entrée supplémentaire pour ajouter une action supplémentaire
                        SelecteurMot.Items.Add("Ajouter le mot dans la base...");
                        SelecteurMot.SelectedIndex = 0;
                    }
                }));

            }

        }

        private Regex PositiveNumberOnly = new Regex("[0-9]+", RegexOptions.Compiled);
        private bool IsTextAllowed(string text)
        {
            return PositiveNumberOnly.IsMatch(text);
        }

        private void ValidateNumbers(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !IsTextAllowed(e.Text);
        }

        private void CompteurProtection_ValueChanged(object sender, EventArgs e)
        {

            if (MotSelectionner == null && _motRechercher.Length > 0) {
                MotSelectionner = new Mot()
                {
                    Texte = _motRechercher,
                    Protections = 0,
                    Abreviations = 0,
                    ToujoursDemander = 0,
                    Commentaires = "",
                    DateAjout = DateTime.Now.ToString(),
                };
            }
            if (MotSelectionner != null) {
                MotSelectionner.Protections = ProtectionsValue.Text.Length > 0 ? long.Parse(ProtectionsValue.Text) : MotSelectionner.Protections;
                SaveButton.IsEnabled = true;
            }
             
            
        }

        private void CompteurAbreviation_ValueChanged(object sender, EventArgs e)
        {
            if (MotSelectionner == null && _motRechercher.Length > 0) {
                MotSelectionner = new Mot() { 
                    Texte = _motRechercher,
                    Protections = 0,
                    Abreviations = 0,
                    ToujoursDemander = 0,
                    Commentaires = "",
                    DateAjout = DateTime.Now.ToString(),
                };
            }
            if (MotSelectionner != null) {
                MotSelectionner.Abreviations = AbreviationsValue.Text.Length > 0 ? long.Parse(AbreviationsValue.Text) : MotSelectionner.Abreviations;
                SaveButton.IsEnabled = true;
            }
            
        }

        private void ToujoursDemanderValue_CheckedChanged(object sender, EventArgs e)
        {
            
            if (MotSelectionner == null && _motRechercher.Length > 0) {
                MotSelectionner = new Mot()
                {
                    Texte = _motRechercher,
                    Protections = 0,
                    Abreviations = 0,
                    ToujoursDemander = 0,
                    Commentaires = "",
                    DateAjout = DateTime.Now.ToString(),
                };
            }
            if(MotSelectionner != null) {
                MotSelectionner.ToujoursDemander = ToujoursDemanderValue.IsChecked == true ? 1 : 0;
                SaveButton.IsEnabled = true;
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (MotSelectionner == null && _motRechercher.Length > 0) {
                MotSelectionner = new Mot()
                {
                    Texte = _motRechercher,
                    Protections = 0,
                    Abreviations = 0,
                    ToujoursDemander = 0,
                    Commentaires = "",
                    DateAjout = DateTime.Now.ToString(),
                };
            }
            if (MotSelectionner != null) {
                if( MotSelectionner.Id == 0 &&
                    MessageBox.Show(
                        $"Voulez-vous enregistrer {MotSelectionner.Texte} dans la base de donnée ?",
                        "Nouveau mot",
                        MessageBoxButton.YesNo
                    ) != MessageBoxResult.Yes
                ) {
                    return;
                }   
                using (var session = BaseSQlite.CreateSessionFactory().OpenSession()) {
                    using (var transaction = session.BeginTransaction()) {
                        //if (MotSelectionner != null) {
                        //    session.SaveOrUpdate(MotSelectionner);
                        //    transaction.Commit();
                        //}
                    }
                }
            }
            
        }

        private void SelecteurMot_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(SelecteurMot.SelectedItem != null) {
                using (var session = BaseSQlite.CreateSessionFactory().OpenSession()) {
                    List<Mot> motsSelectionner = session.Query<Mot>().Where(
                        m => m.Texte == SelecteurMot.SelectedItem.ToString()
                    ).ToList();

                    if (motsSelectionner.Count > 0) {
                        MotSelectionner = motsSelectionner[0];
                        _motRechercher = MotSelectionner.Texte;
                        SaveButton.IsEnabled = false;
                    } else if(_motRechercher.Length > 0) {
                        // Création du mot sélectionner s'il n'existe pas dans la base de donnée
                        MotSelectionner = new Mot()
                        {
                            Texte = _motRechercher,
                            Protections = 0,
                            Abreviations = 0,
                            ToujoursDemander = 0,
                            Commentaires = "",
                            DateAjout = DateTime.Now.ToString(),
                        };
                        SaveButton.IsEnabled = true;
                    }
                }
            } else {
                MotSelectionner = null;
            }
            
        }

        CancellationTokenSource tokenSource = new CancellationTokenSource();
        Task awaitTask = null;

        private async void SearchText_TextChanged(object sender, TextChangedEventArgs e)
        {
            _motRechercher = SearchText.Text;
            SelecteurMot.Items.Clear();
            MotSelectionner = null;
            SaveButton.IsEnabled = false;

            if(_motRechercher.Length == 0) {
                InfosMotsTrouver.Content = "(Saisissez un mot dans le champ Texte pour commencer)";
                return;
            }
            if (_motRechercher.Length > 0) {
                SelecteurMot.SelectedIndex = 0;
                InfosMotsTrouver.Dispatcher.Invoke(new Action(() =>
                {
                    InfosMotsTrouver.Content = "Recherche en cours ...";
                }));
                SelecteurMot.Dispatcher.Invoke(new Action(() =>
                {
                    SelecteurMot.Items.Clear();
                }));
                // Cancel previous tasks
                tokenSource.Cancel();
                awaitTask?.Wait();
                tokenSource = new CancellationTokenSource();
                CancellationToken token = tokenSource.Token;
                // Attendre une seconde pour lancer la recherche
                awaitTask = Task.Run(() =>
                    {
                        int awaitTime = 1000;
                        for (int i = 0; i < awaitTime; i += 1) {
                            if (token.IsCancellationRequested) {
                                return false;
                            }
                            Thread.Sleep(1);
                        }
                        return true;
                    },
                    tokenSource.Token
                ).ContinueWith((t) => 
                    {
                        if (t.Result == true) {
                            // La tache d'attente async a fini, on peut lancer une recherche en arrière plan
                            Task search = Task.Run(() =>
                            {
                                using (var session = BaseSQlite.CreateSessionFactory().OpenSession()) {

                                    List<Mot> trouver = session.Query<Mot>().Where(
                                        m => m.Texte.Contains(_motRechercher.ToLower())
                                    ).ToList();
                                    InfosMotsTrouver.Dispatcher.Invoke(new Action(() =>
                                    {
                                        InfosMotsTrouver.Content = trouver.Count == 0 ? "Aucun mot trouvé" : $"{trouver.Count} mots trouvés";
                                    }));
                                    SelecteurMot.Dispatcher.Invoke(new Action(() =>
                                    {
                                        SelecteurMot.Items.Clear();
                                        SelecteurMot.SelectedIndex = -1;
                                        int i = -1;
                                        foreach (var mot in trouver) {
                                            if (token.IsCancellationRequested) {
                                                return;
                                            }
                                            i++;
                                            SelecteurMot.Items.Add(mot.Texte);
                                            if (_motRechercher.ToLower() == mot.Texte) {
                                                SelecteurMot.SelectedIndex = i;
                                            }
                                        }
                                        // Pas de correspondance exacte dans la base de donnée
                                        if (SelecteurMot.SelectedIndex == -1) {
                                            // On ajoute une entrée supplémentaire pour ajouter une action supplémentaire
                                            SelecteurMot.Items.Add("Ajouter le mot dans la base...");
                                            SelecteurMot.SelectedIndex = 0;
                                        }
                                    }));

                                }
                            }, tokenSource.Token);
                        }
                    }
                );
            }

        }

        private void CommentairesValue_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (MotSelectionner == null && _motRechercher.Length > 0) {
                MotSelectionner = new Mot()
                {
                    Texte = _motRechercher,
                    Protections = 0,
                    Abreviations = 0,
                    ToujoursDemander = 0,
                    Commentaires = "",
                    DateAjout = DateTime.Now.ToString(),
                };
            }
            if (MotSelectionner != null) {
                MotSelectionner.Commentaires = CommentairesValue.Text;
                SaveButton.IsEnabled = true;
            }
        }
    }
}
