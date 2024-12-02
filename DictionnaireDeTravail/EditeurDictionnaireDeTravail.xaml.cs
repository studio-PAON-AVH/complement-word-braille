
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace fr.avh.braille.dictionnaire
{

    public class MotAfficher
    {
        private static Dictionary<string, Statut> _statuts = new Dictionary<string, Statut> {
            { "Inconnu", Statut.INCONNU },
            { "Abréger", Statut.ABREGE },
            { "Protéger", Statut.PROTEGE },
            { "Ignorer", Statut.IGNORE }
        };
        private static Dictionary<Statut, string> _revert = _statuts.ToDictionary((kvp) => kvp.Value, (kvp) => kvp.Key);
        public string[] StatutsPossible { get => _statuts.Keys.ToArray(); }

        public string StatutChoisi { 
            get => _revert[Statut]; 
            set { 
                Statut = _statuts[value];
            }
        }

        public string Texte { get; set; }

        public string ContexteAvant { get; set; } = null;
        public string ContexteApres { get; set; } = null;
        public string Contexte { get; set; } = null;
        public Statut Statut { get; set; }

        public int Index { get; set; }
    }
    /// <summary>
    /// Logique d'interaction pour EditeurDictionnaireDeTravail.xaml
    /// </summary>
    public partial class EditeurDictionnaireDeTravail : Window
    {

        private DictionnaireDeTravail dictionnaire = new DictionnaireDeTravail("Placeholder");

        private string dictionnairePath = "";

        ObservableCollection<MotAfficher> mots = new ObservableCollection<MotAfficher>();

        private Dictionary<Statut, CheckBox> filteredStatuts = new Dictionary<Statut, CheckBox>()
        {
            { Statut.INCONNU, new CheckBox(){ IsChecked = true, Content = "Inconnu" } },
            { Statut.ABREGE, new CheckBox(){ IsChecked = true, Content = "Abréger" } },
            { Statut.PROTEGE, new CheckBox(){ IsChecked = true, Content = "Protéger" } },
            { Statut.IGNORE, new CheckBox(){ IsChecked = true, Content = "Ignorer" } }
        };

        IProtection protecteur = null;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dictionnaire"></param>
        /// <param name="dictionnairePath"></param>
        /// <param name="protecteur">Protecteur intéragissant avec un document</param>
        /// <param name="contextes"></param>
        public EditeurDictionnaireDeTravail(DictionnaireDeTravail dictionnaire, string dictionnairePath, IProtection protecteur = null)
        {
            InitializeComponent();
            this.dictionnaire = dictionnaire;
            this.protecteur = protecteur;
            
            Title = $"Édition du dictionnaire {dictionnaire.NomDictionnaire}";
            NameDictionnaire.Content = dictionnaire.NomDictionnaire;
            CompteurMots.Content = $"{dictionnaire.Occurences.Count} mots répertoriés";
            
            this.dictionnairePath = dictionnairePath;

            FiltreStatuts.Children.Clear();
            foreach (var item in filteredStatuts) {
                item.Value.Checked += FiltreStatuts_Changed;
                item.Value.Unchecked += FiltreStatuts_Changed;
                FiltreStatuts.Children.Add(item.Value);
            }
            SelecteurMot.ItemsSource = dictionnaire.CarteMotOccurences.Keys;
            VueDictionnaire_Refresh();
        }

        public EditeurDictionnaireDeTravail(string dictionnairePath)
        {
            if(!File.Exists(dictionnairePath)) {
                throw new FileNotFoundException($"Le fichier {dictionnairePath} n'existe pas");
            }
            InitializeComponent();
            Title = $"Édition du dictionnaire {Path.GetFileName(dictionnairePath)}";
            this.dictionnairePath = dictionnairePath;
            NameDictionnaire.Content = "Chargement ...";
            DictionnaireDeTravail.FromDictionnaryFileJSON(dictionnairePath)
                .ContinueWith((task) => {
                    dictionnaire = task.Result;
                    Dispatcher.Invoke(() =>
                    {
                        Title = $"Édition du dictionnaire {dictionnaire.NomDictionnaire}";
                        NameDictionnaire.Content = dictionnaire.NomDictionnaire;
                        CompteurMots.Content = $"{dictionnaire.Occurences.Count} mots répertoriés";
                        VueDictionnaire_Refresh();
                    });
                });

            FiltreStatuts.Children.Clear();
            foreach (var item in filteredStatuts) {
                item.Value.Checked += FiltreStatuts_Changed;
                item.Value.Unchecked += FiltreStatuts_Changed;
                FiltreStatuts.Children.Add(item.Value);
            }
            SelecteurMot.ItemsSource = dictionnaire.CarteMotOccurences.Keys;

        }

        private void VueDictionnaire_Refresh()
        {
            // recalcul la liste des mots a afficher a partir des filtres
            List<Statut> activeStatuts = filteredStatuts
                .Where((kvp) => kvp.Value.IsChecked == true)
                .Select((kvp) => kvp.Key)
                .ToList();

            mots = new ObservableCollection<MotAfficher>(
                dictionnaire.OccurencesAsListOfTuples()
                .Where(m =>
                    activeStatuts.Contains(m.Item3)
                    && (
                        FiltreTexte.Text.Length == 0
                        ? (SelecteurMot.SelectedIndex >= 0 
                            ? m.Item2.ToLower().Trim() == SelecteurMot.SelectedValue.ToString() 
                            : true
                        )
                        : m.Item2.ToLower().Trim().Contains(FiltreTexte.Text.ToLower().Trim())
                    )
                )
                .Select(
                        (tuple) => new MotAfficher()
                        {
                            Index = tuple.Item1,
                            Texte = tuple.Item2,
                            Statut = tuple.Item3,
                            ContexteAvant = tuple.Item4,
                            ContexteApres = tuple.Item5
                        }
                    )
                ); 
            VueDictionnaire.DataContext = mots;
            CompteurAfficher.Content = $"{mots.Count} mots affichés";

        }

        private void FiltreStatuts_Changed(object sender, RoutedEventArgs e)
        {
            VueDictionnaire_Refresh();
        }

        

        private void Mot_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2) {
                TextBlock _mot = sender as TextBlock;
                new Consultation(_mot.Text).ShowDialog();
            }
            
        }

        private void Statut_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox _statut = sender as ComboBox;
            MotAfficher selected = (MotAfficher)_statut.DataContext;
            if (selected != null && _statut.SelectedValue != null) {
                selected.StatutChoisi = _statut.SelectedValue.ToString();
                dictionnaire.StatutsOccurences[selected.Index] = selected.Statut;
                protecteur?.AppliquerStatutSurOccurence(selected.Index, selected.Statut);
                
                if (this.dictionnairePath.Length > 0 && Directory.Exists(Path.GetDirectoryName(dictionnairePath))) {
                    dictionnaire.SaveJSON(new DirectoryInfo(Path.GetDirectoryName(dictionnairePath)));
                }
            }
        }

        private bool _blockRefresh = false;

        private void FiltreTexte_TextChanged(object sender, TextChangedEventArgs e)
        {
            if(_blockRefresh) {
                return;
            }
            _blockRefresh = true;
            SelecteurMot.SelectedIndex = -1;
            VueDictionnaire_Refresh();
            _blockRefresh = false;
        }

        private void SelecteurMot_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_blockRefresh) {
                return;
            }
            _blockRefresh = true;
            FiltreTexte.Text = "";
            VueDictionnaire_Refresh();
            _blockRefresh = false;
        }

        private void ProtegerLaSelection_Click(object sender, RoutedEventArgs e)
        {
            foreach(MotAfficher mot in mots) {
                dictionnaire.StatutsOccurences[mot.Index] = Statut.PROTEGE;
                mot.Statut = Statut.PROTEGE;
                protecteur?.AppliquerStatutSurOccurence(mot.Index, Statut.PROTEGE);
            }
            VueDictionnaire_Refresh();
        }

        private void AbregerLaSelection_Click(object sender, RoutedEventArgs e)
        {
            foreach (MotAfficher mot in mots) {
                dictionnaire.StatutsOccurences[mot.Index] = Statut.ABREGE;
                mot.Statut = Statut.ABREGE;
                protecteur?.AppliquerStatutSurOccurence(mot.Index, Statut.ABREGE);
            }
            VueDictionnaire_Refresh();
        }

        private void ReinitialiserLaSelection_Click(object sender, RoutedEventArgs e)
        {
            foreach (MotAfficher mot in mots) {
                dictionnaire.StatutsOccurences[mot.Index] = Statut.INCONNU;
                mot.Statut = Statut.INCONNU;
                protecteur?.AppliquerStatutSurOccurence(mot.Index, Statut.INCONNU);
            }
            VueDictionnaire_Refresh();
        }

        private void VueDictionnaire_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var dataGrid = sender as DataGrid;
            var selectedItem = dataGrid.SelectedItem as MotAfficher;
            if (selectedItem != null) {
                protecteur?.AfficherOccurence(selectedItem.Index);
            }
        }
    }
}
