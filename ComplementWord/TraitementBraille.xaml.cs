using fr.avh.braille.dictionnaire;
using Microsoft.Win32;
using System;
using System.IO;
using System.Windows;
using System.Windows.Threading;

namespace fr.avh.braille.addin
{
    /// <summary>
    /// Logique d'interaction pour TraitementBraille.xaml
    /// </summary>
    public partial class TraitementBraille : Window
    {

        public void SetProgress(int progress, int maxValue)
        {
            ProgressIndicator.Value = progress;
            ProgressIndicator.Maximum = maxValue;
        }

        public void AddMessage(string message)
        {
            ProgressMessages.AppendText(message + "\r\n");
            ProgressMessages.ScrollToEnd();
        }
        
        public TraitementBraille(string title = "")
        {
            InitializeComponent();
            if(!string.IsNullOrWhiteSpace(title)) {
                Title = "Traitement du braille pour " + title;
            }
        }

        // TODO : déléguer le traitement du document (ProtectionWord) a cette fenêtre

        private ProtectionWord _protectionWord;

        /// <summary>
        /// Liaison pour actions sur le document associés
        /// </summary>
        /// <param name="protection"></param>
        public void BindToProtection(ProtectionWord protection)
        {
            _protectionWord = protection;
            if (_protectionWord != null) {
                LancerAnalysePhrases.IsEnabled = true;
                LancerParcourMots.IsEnabled = true;
                LancerChargementDictionnaire.IsEnabled = true;
            } else {
                LancerAnalysePhrases.IsEnabled = false;
                LancerParcourMots.IsEnabled = false;
                LancerChargementDictionnaire.IsEnabled = false;
            }

        }

        private void LancerChargementDictionnaire_Click(object sender, RoutedEventArgs e)
        {

            DictionnaireDeTravail importer = null;
            // Rechercher un fichier .bdic ou .json ou .ddic
            OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Fichiers de dictionnaire (*.bdic, *.json, *.ddic)|*.bdic;*.json;*.ddic|Tous les fichiers (*.*)|*.*",
                DefaultExt = ".json",
                Title = "Sélectionner un fichier de dictionnaire"
            };
            if (openFileDialog.ShowDialog() == true) {
                string filePath = openFileDialog.FileName;
                try {
                    switch(Path.GetExtension(filePath).ToLower()) {
                        case ".json":
                            importer = DictionnaireDeTravail.FromDictionnaryFileJSON(filePath).Result;
                            break;
                        case ".bdic":
                            importer = DictionnaireDeTravail.FromDictionnaryFileBDIC(filePath).Result;
                            break;
                        case ".ddic":
                            importer = DictionnaireDeTravail.FromDictionnaryFileDDIC(filePath).Result;
                            break;
                        default:
                            throw new NotSupportedException("Format de fichier non supporté : " + Path.GetExtension(filePath));
                    }
                } catch (Exception ex) {
                    MessageBox.Show(
                        "Impossible de charger ce dictionnaire :\r\n" + ex.Message,
                        "Erreur de chargement",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error
                    );
                }
            }
            if(importer != null) {
                _protectionWord.DonneesTraitement.RechargerDecisionDe(importer);
                var res = MessageBox.Show("Dictionnaire chargé avec succès : " + _protectionWord.DonneesTraitement.NomDictionnaire +
                    "\r\nSouhaitez-vous appliquer les décisions sur le document immédiatement ?", "Chargement réussi", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if(res == MessageBoxResult.Yes) {
                    try {
                        _protectionWord.AppliquerDecisions(false);
                        MessageBox.Show("Décisions appliquées avec succès.", "Application réussie", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex) {
                        //MessageBox.Show("Erreur lors de l'application des décisions :\r\n" + ex.Message, "Erreur d'application", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        private void LancerAnalysePhrases_Click(object sender, RoutedEventArgs e)
        {

        }

        private void LancerParcourMots_Click(object sender, RoutedEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                LancerAnalysePhrases.IsEnabled = false;
                LancerParcourMots.IsEnabled = false;
                LancerChargementDictionnaire.IsEnabled = false;
                _protectionWord.ReanalyserDocumentSiModification();
                LancerAnalysePhrases.IsEnabled = true;
                LancerParcourMots.IsEnabled = true;
                LancerChargementDictionnaire.IsEnabled = true;
                // Empecher la modification du document pendant que la fenêtre est ouverte pour éviter de casser les positions
                new ProtectionInteractiveParMotsDialog(_protectionWord).ShowDialog();
            });
        }
    }
}
