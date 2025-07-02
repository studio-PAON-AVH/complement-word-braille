using fr.avh.braille.dictionnaire;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
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
using System.Windows.Threading;

namespace fr.avh.braille.addin
{
    /// <summary>
    /// Logique d'interaction pour ListeMotsHorsLexique.xaml
    /// </summary>
    public partial class ListeMotsHorsLexique : Window
    {

        private ProtectionWord protecteur;
        ObservableCollection<MotAfficher> mots = new ObservableCollection<MotAfficher>();


        public ListeMotsHorsLexique()
        {
            InitializeComponent();
        }


        public ListeMotsHorsLexique(ProtectionWord p)
        {
            InitializeComponent();
            protecteur = p;
            VueDictionnaire_Refresh();

        }

        private void VueOccurences_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var dataGrid = sender as DataGrid;
            var selectedItem = dataGrid.SelectedItem as MotAfficher;
            var selectionIndex = dataGrid.SelectedIndex;
            if (selectedItem != null) {
                protecteur.SelectionnerOccurence(selectedItem.Index).Select();
            }
        }

        private void VueDictionnaire_Refresh()
        {
            Dispatcher dispatcher = this.Dispatcher;
            dispatcher.Invoke(async () =>
            {
                protecteur.ReanalyserDocumentSiModification();
                var indices = await protecteur.IndicesOccurencesHorsLexique();
                mots = new ObservableCollection<MotAfficher>(
                        protecteur.DonneesTraitement
                            .OccurencesAsListOfTuples()
                            .Where((o, index) => indices.Contains(index))
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
                this.UpdateLayout();
            });
        }

        private void Window_GotFocus(object sender, RoutedEventArgs e)
        {
            if (protecteur != null) {
                VueDictionnaire_Refresh();
            }
        }
    }
}
