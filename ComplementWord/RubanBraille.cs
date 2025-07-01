
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows;
using fr.avh.braille.dictionnaire;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Threading;
using System.Windows.Forms;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.Threading.Tasks;

namespace fr.avh.braille.addin
{
    public partial class RubanBraille
    {
        ProtectionWord protectionTool;
        ProtectionInteractiveParMotsDialog actions = null;

        private void RubanAVH_Load(object sender, RibbonUIEventArgs e)
        {
            Globals.ThisAddIn.Application.WindowActivate += Application_WindowActivate;
            AfficherDictionnaire.Enabled = false;
            Refocus.Enabled = false;
            PremierMot.Enabled = false;
            OccurencePrecedente.Enabled = false;
            OccurenceSuivante.Enabled = false;
            //AnalyseMotsEtranger.Enabled = false;
        }
        

        // Mise a jour du ruban en cas de changement de document
        private void Application_WindowActivate(Document newActiveDocument, Word.Window Wn)
        {
            if (newActiveDocument != null) {
                if (Globals.ThisAddIn.documentProtection.ContainsKey(newActiveDocument.FullName)) {
                    protectionTool = Globals.ThisAddIn.documentProtection[newActiveDocument.FullName];
                    AfficherDictionnaire.Enabled = true;
                    Refocus.Enabled = true;
                    PremierMot.Enabled = true;
                    OccurencePrecedente.Enabled = true;
                    OccurenceSuivante.Enabled = true;
                    //AnalyseMotsEtranger.Enabled = true;
                } else {
                    protectionTool = null;
                    AfficherDictionnaire.Enabled = false;
                    Refocus.Enabled = false;
                    PremierMot.Enabled = false;
                    OccurencePrecedente.Enabled = false;
                    OccurenceSuivante.Enabled = false;
                    //AnalyseMotsEtranger.Enabled = false;
                }
            } else {
                AfficherDictionnaire.Enabled = false;
                Refocus.Enabled = false;
                PremierMot.Enabled = false;
                OccurencePrecedente.Enabled = false;
                OccurenceSuivante.Enabled = false;
                //AnalyseMotsEtranger.Enabled = false;
            }
        }
        TraitementBraille progressDialog = null;

        Dispatcher _dispatcher = Dispatcher.CurrentDispatcher;
        private void InfoCallbak(string message, Tuple<int, int> progessTuple = null)
        {
            _dispatcher.Invoke(() =>
            {
                if (progressDialog == null || !progressDialog.IsLoaded) {
                    progressDialog = new TraitementBraille(Path.GetFileNameWithoutExtension(Globals.ThisAddIn.Application.ActiveDocument.FullName));
                }
                progressDialog.Show();
                progressDialog.Focus();
                progressDialog.Dispatcher.Invoke(() =>
                {
                    try {
                        if (progessTuple != null) {
                            progressDialog.SetProgress(progessTuple.Item1, progessTuple.Item2);
                        }
                        if (message.Length > 0) {
                            progressDialog.AddMessage(message);
                            fr.avh.braille.dictionnaire.Globals.log(message);
                        }
                    }
                    catch (Exception e) {

                    }
                });
            });
            
        }

        private void ErrorCallbak(Exception ex)
        {
            _dispatcher.Invoke(() =>
            {
                if (progressDialog == null || !progressDialog.IsLoaded) {
                    progressDialog = new TraitementBraille(Path.GetFileNameWithoutExtension(Globals.ThisAddIn.Application.ActiveDocument.FullName));
                }
                progressDialog.Show();
                progressDialog.Focus();
                fr.avh.braille.dictionnaire.Globals.log(ex);
                StringBuilder message = new StringBuilder($"L'erreur suivante a été remonté lors de l'action et doit être remonté à l'équipe de développement du complément : \r\n" +
                    $"{ex.Message}\r\n");
                string stack = ex.StackTrace;
                while (ex.InnerException != null) {
                    ex = ex.InnerException;
                    message.Append($"{ex.Message}\r\n");
                }
                message.Append($"{stack}\r\n");
                progressDialog.Dispatcher.Invoke(() =>
                {
                    try {
                        progressDialog.AddMessage(message.ToString());
                        fr.avh.braille.dictionnaire.Globals.log(message.ToString());
                    }
                    catch (Exception e) {

                    }
                });
                System.Windows.MessageBox.Show(
                    message.ToString(),
                    "Erreur lors de l'analyse du document",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );
            });
        }


        private async Task<ProtectionWord> GetProtectionToolAsync()
        {
            Dispatcher _d = Dispatcher.CurrentDispatcher;
            if (Globals.ThisAddIn.documentProtection.ContainsKey(Globals.ThisAddIn.Application.ActiveDocument.FullName)) {
                return Globals.ThisAddIn.documentProtection[Globals.ThisAddIn.Application.ActiveDocument.FullName];
            } else {
                return await Globals.ThisAddIn.AnalyzeCurrentDocument(InfoCallbak, ErrorCallbak).ContinueWith(t => { 
                    ProtectionWord protection = t.Result;
                    protection.SelectionChanged += (index) =>
                    {
                        _d.Invoke(() =>
                        {
                            MotSelectionner.Label = $"Mot : {protection.MotSelectionne}";
                            Progression.Label = $"Occurence : {protection.SelectedWordOccurenceIndex + 1} / {protection.SelectedWordOccurenceCount}";
                        });
                    };
                    return protection;
                });
             }
        }


        private void AfficherDictionnaire_Click(object sender, RibbonControlEventArgs e)
        {
            EditeurDictionnaireDeTravail editeur = new EditeurDictionnaireDeTravail(protectionTool.DonneesTraitement, protectionTool.WorkingDictionnaryPath, protectionTool);
            editeur.ShowDialog();
        }

        private void Refocus_Click(object sender, RibbonControlEventArgs e)
        {
            protectionTool?.SelectedRange.Select();
        }

        private void OccurencePrecedente_Click(object sender, RibbonControlEventArgs e)
        {
            protectionTool?.PrecedenteOccurence().Select();
        }

        private void OccurenceSuivante_Click(object sender, RibbonControlEventArgs e)
        {
            protectionTool?.ProchaineOccurence().Select();
        }

        private void ProtectionDB_Click(object sender, RibbonControlEventArgs e)
        {
            if (!BaseSQlite.dbExists()) {
                TraitementBraille progressDialog = new TraitementBraille();
                progressDialog.Show();
                BaseSQlite.CheckForUpdates((message, progessTuple) =>
                {
                    if (progessTuple != null) {
                        progressDialog.Dispatcher.Invoke(() =>
                        {
                            progressDialog.SetProgress(progessTuple.Item1, progessTuple.Item2);
                        });
                    }
                    progressDialog.Dispatcher.Invoke(() =>
                    {
                        progressDialog.AddMessage(message);
                    });
                });
                progressDialog.Close();
            }
            new Consultation().ShowDialog();
        }

        private void PremierMot_Click(object sender, RibbonControlEventArgs e)
        {
            protectionTool?.AfficherOccurence(0);
        }

        private void ProtegerSelection_Click(object sender, RibbonControlEventArgs e)
        {
            Range selection = Globals.ThisAddIn.Application.Selection.Range;
            
            if(selection != null && selection.Text != null && selection.Text.Trim().Length > 0) {
                // mode manuel, pas de test d'abrégeable
                
                if (Globals.ThisAddIn.documentProtection.ContainsKey(Globals.ThisAddIn.Application.ActiveDocument.FullName)) {
                    // si un dictionnaire de traitement est déja présent
                    ProtectionWord protector = Globals.ThisAddIn.documentProtection[Globals.ThisAddIn.Application.ActiveDocument.FullName];
                    protector.ReanalyserDocumentSiModification();
                    if(selection.Words.Count > 1) {
                        // Si la selection contient plusieurs mots
                        protector.ProtegerBloc(selection);
                        //ProtectionWord.ProtegerBloc(Globals.ThisAddIn.Application.ActiveDocument, selection);
                    } else {
                        protector.Proteger(selection);
                        // proposer d'appliquer cette protection a toutes les occurences 
                        int selectedIndex = -1;
                        if (System.Windows.MessageBox.Show(
                            "Voulez-vous protéger toutes les occurences de ce mot ?", "Protection de mot", MessageBoxButton.YesNo
                            ) == MessageBoxResult.No
                        ) {
                            // Compter le nombre d'occurence avant le mot sélectionné
                            Range toBegin = Globals.ThisAddIn.Application.ActiveDocument.Range(
                               Globals.ThisAddIn.Application.ActiveDocument.Content.Start,
                               selection.Start
                            );
                            string text = toBegin.Text;
                            string word = selection.Text.Trim();
                            Regex search = ProtectionWord.SearchWord(word, RegexOptions.IgnoreCase | RegexOptions.Singleline);
                            MatchCollection matches = search.Matches(text);
                            selectedIndex = matches.Count;
                        }
                        // Ajouter le mot au traitement
                        protector.AjouterMotAuTraitement(selection.Text.Trim().ToLower(), Statut.PROTEGE, selectedIndex);
                        // resélectionner la premiere occurence du mot
                        protector.SelectionnerOccurenceMot(selection.Text.Trim().ToLower(), selectedIndex >= 0 ? selectedIndex : 0);
                        if (actions == null || !actions.IsLoaded) {
                            actions = new ProtectionInteractiveParMotsDialog(protectionTool);
                            actions.ShowDialog();
                        } else {
                            actions.Activate();
                        }
                    }  
                } else {
                    // Pas de dictionnaire de traitement, si la selection de mots est composé de plusieurs mots, protéger le texte entier avec ProtegerBloc
                    if(selection.Text.Split(' ').Length > 1)
                    {
                        ProtectionWord.ProtegerBloc(Globals.ThisAddIn.Application.ActiveDocument, selection);
                        return;
                    }
                    else
                    {
                        // Pas de dictionnaire de traitement, on protège juste la selection
                        ProtectionWord.Proteger(Globals.ThisAddIn.Application.ActiveDocument, selection);
                    }
                }
            }
        }

        private void AbregerSelection_Click(object sender, RibbonControlEventArgs e)
        {
            Range selection = Globals.ThisAddIn.Application.Selection.Range;

            if (selection != null && selection.Text != null && selection.Text.Trim().Length > 0) {
                // mode manuel, pas de test d'abrégeable

                if (Globals.ThisAddIn.documentProtection.ContainsKey(Globals.ThisAddIn.Application.ActiveDocument.FullName)) {
                    // si un dictionnaire de traitement est déja présent
                    ProtectionWord protector = Globals.ThisAddIn.documentProtection[Globals.ThisAddIn.Application.ActiveDocument.FullName];
                    protector.ReanalyserDocumentSiModification();
                    // proposer d'appliquer cette protection a toutes les occurences 
                    int selectedIndex = -1;
                    if (System.Windows.MessageBox.Show(
                        "Voulez-vous abréger toutes les occurences de ce mot ?", "Protection de mot", MessageBoxButton.YesNo
                        ) == MessageBoxResult.No
                    ) {
                        Range toBegin = Globals.ThisAddIn.Application.ActiveDocument.Range(
                           Globals.ThisAddIn.Application.ActiveDocument.Content.Start,
                           selection.Start
                        );
                        string text = toBegin.Text;
                        string word = selection.Text.Trim();
                        Regex search = ProtectionWord.SearchWord(word, RegexOptions.IgnoreCase | RegexOptions.Singleline);
                        MatchCollection matches = search.Matches(text);
                        selectedIndex = matches.Count;
                    }
                    // Ajouter le mot au traitement
                    protector.AjouterMotAuTraitement(selection.Text.Trim().ToLower(), Statut.ABREGE, selectedIndex);
                    // resélectionner la premiere occurence du mot
                    protector.SelectionnerOccurenceMot(selection.Text.Trim().ToLower(), selectedIndex >= 0 ? selectedIndex : 0);
                    if (actions == null || !actions.IsLoaded) {
                        actions = new ProtectionInteractiveParMotsDialog(protectionTool);
                        actions.ShowDialog();
                    } else {
                        actions.Activate();
                    }
                } else {

                    if (selection.Text.Split(' ').Length > 1)
                    {
                        ProtectionWord.AbregerBloc(Globals.ThisAddIn.Application.ActiveDocument, selection);
                        return;
                    }
                    else
                    {
                        // Pas de dictionnaire de traitement, on protège juste la selection
                        ProtectionWord.Abreger(Globals.ThisAddIn.Application.ActiveDocument, selection);
                    }
                    
                }
            }
        }

        private void OuvrirOptions_Click(object sender, RibbonControlEventArgs e)
        {
            // Afficher un formulaire d'option
            FormulaireOptionsComplement options = new FormulaireOptionsComplement();
            options.ShowDialog();
        }

        private void ChargerDecisions_Click(object sender, RibbonControlEventArgs e)
        {
            Dispatcher _d = Dispatcher.CurrentDispatcher;
            _d.Invoke(async () =>
            {
                try {
                    InfoCallbak("Chargement d'un autre dictionnaire de décisions...");
                    protectionTool = await GetProtectionToolAsync();
                    protectionTool.ReanalyserDocumentSiModification();
                    // Rechercher un fichier .bdic ou .json ou .ddic
                    OpenFileDialog openFileDialog = new OpenFileDialog
                    {
                        Filter = "Fichiers de décisions (*.bdic, *.json, *.ddic)|*.bdic;*.json;*.ddic",
                        DefaultExt = ".json",
                        Title = "Sélectionner un fichier de décisions"
                    };
                    if (openFileDialog.ShowDialog() == DialogResult.OK) {
                        string filePath = openFileDialog.FileName;
                        protectionTool.ImporterUnDictionnaire(filePath);
                        InfoCallbak(filePath + " chargé avec succès, vous pouvez commencez le traitement par mots");
                    }
                } catch (Exception ex) {
                    ErrorCallbak(ex);
                    return;
                }
                
            });
        }

        private void LancerTraitementMots_Click(object sender, RibbonControlEventArgs e)
        {
            Dispatcher _d = Dispatcher.CurrentDispatcher;
            _d.Invoke(async () => {
                try {
                    InfoCallbak("Lancement du traitement par mots ...");
                    protectionTool = await GetProtectionToolAsync();
                    protectionTool.ReanalyserDocumentSiModification();
                    new ProtectionInteractiveParMotsDialog(protectionTool).ShowDialog();
                } catch (AggregateException ex) {
                    ErrorCallbak(ex);
                    return;
                }
                
            });
        }

        private ListeMotsHorsLexique _listeMotsHorsLexique = null;



        private void MotsHorsLexique_Click(object sender, RibbonControlEventArgs e)
        {
            Dispatcher _d = Dispatcher.CurrentDispatcher;
            _d.Invoke(async () =>
            {
                try {
                    InfoCallbak("Récupération des mots hors lexique ...");
                    protectionTool = await GetProtectionToolAsync();
                    protectionTool.ReanalyserDocumentSiModification();
                    if (_listeMotsHorsLexique != null && _listeMotsHorsLexique.IsLoaded) {
                        _listeMotsHorsLexique.Activate();
                        return;
                    }
                    _listeMotsHorsLexique = new ListeMotsHorsLexique(protectionTool);
                    _listeMotsHorsLexique.Show();
                }
                catch (AggregateException ex) {
                    ErrorCallbak(ex);
                    return;
                }
            });
            
        }
    }
}
