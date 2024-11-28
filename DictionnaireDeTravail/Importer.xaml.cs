using fr.avh.archivage;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;

using DialogResultOld = System.Windows.Forms.DialogResult;
using Path = System.IO.Path;
using DragEventArgs = System.Windows.DragEventArgs;
using DataFormats = System.Windows.DataFormats;
using DragDropEffects = System.Windows.DragDropEffects;
using FolderBrowserDialog = System.Windows.Forms.FolderBrowserDialog;
using OpenFileDialog = System.Windows.Forms.OpenFileDialog;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using System.Linq;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;


namespace fr.avh.braille.dictionnaire
{
    /// <summary>
    /// Logique d'interaction pour Importer.xaml
    /// </summary>
    public partial class Importer : Window
    {
        private List<string> entriesDataList = new List<string>();
        private Dictionary<string, string> logsPerEntries = new Dictionary<string, string>();

        public Importer()
        {
            InitializeComponent();
        }

        private static readonly Regex expectedStructure = new Regex(
            @".*\.(d?dic|csv)",
            RegexOptions.Compiled | RegexOptions.IgnoreCase
        );

        private void selectEntriesInFolder(string selectedPath)
        {
            if (Directory.Exists(Utils.LongPath(selectedPath)) || File.Exists(Utils.LongPath(selectedPath))) {
                List<FileInfo> entries = Utils.LazyRecusiveSearch(
                    selectedPath,
                    expectedStructure,
                    Utils.SearchType.Files,
                    (string log, Tuple<int, int> progressIndicator) =>
                    {
                        ProgressionText.Dispatcher.Invoke(
                            new Action(() =>
                            {
                                ProgressionText.AppendText(log);
                            })
                        );
                        if (!logsPerEntries.ContainsKey("")) {
                            logsPerEntries.Add("", log);
                        } else {
                            logsPerEntries[""] += log;
                        }
                    }
                );
                foreach (FileInfo archive in entries) {
                    if (!entriesDataList.Contains(archive.FullName)) {
                        entriesDataList.Add(archive.FullName);
                        string folderName = Path.GetFileName(archive.FullName);
                        // Ajoutez le chemin sélectionné à la ListBox
                        SelectedEntries.Items.Add(folderName);
                        // Activez le bouton de lancement si au moins un élément est dans la ListBox
                        Launch.IsEnabled = SelectedEntries.Items.Count > 0;
                    } else {

                    }
                }
            }
        }

        private void SelectedEntries_OnDragEnter(object sender, DragEventArgs e)
        {
            // Vérifie si l'utilisateur a glissé et déposé des fichiers.
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                // Si des fichiers ont été glissés et déposés, définissez les effets d'opération de glisser-déposer comme "Copy".
                e.Effects = DragDropEffects.Copy;
                e.Handled = true;
            } else {
                // Signifie que l'opération de glisser-déposer n'est pas autorisée et n'aura aucun effet.
                e.Effects = DragDropEffects.None;
            }
        }

        private void SelectedEntries_OnDragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files != null && files.Length > 0) {
                List<string> existingFolders = new List<string>();
                foreach (string file in files) {
                    selectEntriesInFolder(file);
                }
            }
        }

        private void SelectedEntries_OnDragOver(object sender, DragEventArgs e)
        {
            e.Effects |= DragDropEffects.None;

            if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                string target = (string)e.Data.GetData(DataFormats.FileDrop);

                if (target != null) {
                    e.Effects = DragDropEffects.Copy;
                }
            }
        }

        private void SelectedEntries_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }

        CancellationTokenSource cancellation = new CancellationTokenSource();
        Task runner;
        private void Launch_Click(object sender, RoutedEventArgs e)
        {
            if(runner != null && runner.Status == TaskStatus.Running) {
                cancellation.Cancel();
                runner.Wait();
            }
            if(entriesDataList.Count > 0) {
                List<string> dicEntries = entriesDataList.FindAll(f => f.ToLower().EndsWith("dic"));
                List<string> csvEntries = entriesDataList.FindAll(f => f.ToLower().EndsWith("csv"));
                runner = Task.Run(() =>
                {
                    BaseSQlite.AnalyserDictionnaires(
                        dicEntries,
                        (string info, Tuple<int, int> progress) => {
                            ProgressionText.Dispatcher.Invoke(new Action(() =>
                            {
                                ProgressionText.AppendText(info + "\r\n");
                                if (progress != null) {
                                    Progress.Minimum = 0;
                                    Progress.Maximum = progress.Item2;
                                    Progress.Value = progress.Item1;
                                    
                                }
                            }));
                        },
                        cancellation
                    );
                    foreach (var item in csvEntries) {
                        BaseSQlite.updateFromCSV(item);
                    }
                }, cancellation.Token);
                
            }
            
            
        }

        private void Consultation_Click(object sender, RoutedEventArgs e)
        {
            Consultation consultation = new Consultation();
            consultation.Show();
        }

        private void AddEntry_Click(object sender, RoutedEventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog()) {
                dialog.Filter = "Fichiers de dictionnaire (*.dic, *.ddic, *.csv)|*.dic;*.ddic;*.csv";
                dialog.Multiselect = true;
                DialogResultOld result = dialog.ShowDialog();

                if (result == DialogResultOld.OK) {
                    foreach (FileInfo entry in dialog.FileNames.Select(f => new FileInfo(f))) {
                        if (!entriesDataList.Contains(entry.FullName)) {
                            entriesDataList.Add(entry.FullName);
                            string folderName = Path.GetFileName(entry.FullName);
                            // Ajoutez le chemin sélectionné à la ListBox
                            SelectedEntries.Items.Add(folderName);
                            // Activez le bouton de lancement si au moins un élément est dans la ListBox
                            Launch.IsEnabled = SelectedEntries.Items.Count > 0;
                        } else {

                        }
                    }
                }
            }
        }

        private void AddEntries_Click(object sender, RoutedEventArgs e)
        {
            using (FolderBrowserDialog browserDialog = new FolderBrowserDialog()) {
                DialogResultOld result = browserDialog.ShowDialog();

                if (result == DialogResultOld.OK) {
                    selectEntriesInFolder(browserDialog.SelectedPath);
                }
            }
        }

        private void ClearEntries_Click(object sender, RoutedEventArgs e)
        {
            SelectedEntries.Items.Clear();
            ProgressionText.Clear();
            logsPerEntries.Clear();
            entriesDataList.Clear();
            Progress.Value = 0;
        }

        private void SuppEntry_Click(object sender, RoutedEventArgs e)
        {
            var selection = SelectedEntries.SelectedItems;
            for (int i = selection.Count - 1; i >= 0; i--) {
                string found = entriesDataList.Find((f) => Path.GetFileName(f) == selection[i].ToString());
                if (found != "") {
                    entriesDataList.Remove(found);
                }
                SelectedEntries.Items.Remove(selection[i]);
            }
        }

        private void SelectedEntries_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            this.EditionDictionnaire.IsEnabled = SelectedEntries.SelectedItems.Count == 1;
        }

        private void OnKeyDownHandler(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete) {
                var selection = SelectedEntries.SelectedItems;
                for (int i = selection.Count - 1; i >= 0; i--) {
                    string found = entriesDataList.Find((f) => Path.GetFileName(f) == selection[i].ToString());
                    if (found != "") {
                        entriesDataList.Remove(found);
                    }
                    SelectedEntries.Items.Remove(selection[i]);
                }
            }
        }

        private void EditionDictionnaire_Click(object sender, RoutedEventArgs e)
        {
            try {
                var selection = SelectedEntries.SelectedItems[0];
                string found = entriesDataList.Find((f) => Path.GetFileName(f) == selection.ToString());
                
                EditeurDictionnaireDeTravail editeur = new EditeurDictionnaireDeTravail(found);
                editeur.Show();
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
            
        }
    }
}
