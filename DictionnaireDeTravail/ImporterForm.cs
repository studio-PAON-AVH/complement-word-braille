using NHibernate;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace fr.avh.braille.dictionnary
{
    public partial class ImporterForm : Form
    {
        ISessionFactory sessionFactory;
        private string inputPath = "";


        private delegate void AppendText(string text);
        private AppendText logging;
        public ImporterForm(ISessionFactory sessionFactory)
        {
            this.InitializeComponent();
            this.sessionFactory = sessionFactory;
            logging = new AppendText(journal.AppendText);
        }


        private void Launcher_OnDragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files != null && files.Length > 0)
            {
                selectedFilePath.Text = inputPath = files[0];
            }
        }

        private void Launcher_OnDragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else e.Effect = DragDropEffects.None;
        }

        CancellationTokenSource cancellation = new CancellationTokenSource();
        private void launchScript_Click(object sender, EventArgs e)
        {
            if (inputPath == "")
            {
                MessageBox.Show("Veuillez séléectionner un dossier ou un fichier dictionnaire à charger.", "Aucun fichier sélectionné", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            if (!importProcess.IsBusy)
            {
                importProcess.RunWorkerAsync();
                launchScript.Text = "Annuler";
                cancellation = new CancellationTokenSource();
            } else
            {
                importProcess.CancelAsync();
                journal.AppendText("Annulation du traitement ... \r\n");
                cancellation.Cancel();
                launchScript.Text = "Lancer l\'import";
            }
            
              
        }

        private void browseFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Fichier dic|*.dic;*.ddic|Tous les fichiers (*.*)|*.*";
                openFileDialog.Multiselect = true;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    inputPath = openFileDialog.FileName;
                    selectedFilePath.Text = inputPath;
                    launchScript.Enabled = true;
                }
            }
        }

        private void selectedFilePath_TextChanged(object sender, EventArgs e)
        {
            inputPath = selectedFilePath.Text;
            if (inputPath.StartsWith("\"") && inputPath.EndsWith("\""))
            {
                inputPath = inputPath.Substring(1, inputPath.Length - 2);
            }
            launchScript.Enabled = (inputPath != "");
        }

        private void browseFolder_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog selectFolderDialog = new FolderBrowserDialog())
            {
                if (selectFolderDialog.ShowDialog() == DialogResult.OK)
                {
                    inputPath = selectFolderDialog.SelectedPath;
                    selectedFilePath.Text = inputPath;
                    launchScript.Enabled = true;
                }
            }
        }



        private void importProcess_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            if (File.Exists(inputPath))
            {
                if (inputPath.EndsWith(".ddic") || inputPath.EndsWith(".dic")) {
                    bool start = true;
                    DictionnaireSQlite.AnalyserDictionnaires(
                        new List<string>()
                        {
                        inputPath
                        },
                        (string info, Tuple<int, int> progress) =>
                        {
                            this.Invoke(new Action(() =>
                            {
                                logging(info + "\r\n");
                                if (progress != null) {
                                    if (start) {
                                        ImportProgress.Minimum = 0;
                                        ImportProgress.Maximum = progress.Item2;
                                        ImportProgress.Step = 1;
                                        start = false;
                                    }
                                    ImportProgress.PerformStep();
                                }

                            }));
                        },
                        cancellation
                        );
                } else if (inputPath.EndsWith(".csv")) {
                    bool start = true;
                    DictionnaireSQlite.updateFromCSV(
                        inputPath,
                        (string info, Tuple<int, int> progress) =>
                        {
                            this.Invoke(new Action(() =>
                            {
                                logging(info + "\r\n");
                                if (progress != null) {
                                    if (start) {
                                        ImportProgress.Minimum = 0;
                                        ImportProgress.Maximum = progress.Item2;
                                        ImportProgress.Step = 1;
                                        start = false;
                                    }
                                    ImportProgress.PerformStep();
                                }

                            }));
                        },
                        cancellation
                    );
                } else { 
                }
                    
            }
            else if (Directory.Exists(inputPath))
            {
                IEnumerable<string> dictionnaries = Directory.EnumerateFiles(inputPath, "*.dic", SearchOption.AllDirectories);
                dictionnaries = dictionnaries.Concat(Directory.EnumerateFiles(inputPath, "*.ddic", SearchOption.AllDirectories));
                //AnalyseurDeDictionnaires metaDictionnaire = DictionnaireSQlite.Load();//new MetaDictionnaire();

                bool start = true;
                DictionnaireSQlite.AnalyserDictionnaires(
                    dictionnaries.ToList(),
                    (string info, Tuple<int,int> progress) => {
                        this.Invoke(new Action(() =>
                        {
                            logging(info + "\r\n");
                            if(progress != null) {
                                if (start) {
                                    ImportProgress.Minimum = 0;
                                    ImportProgress.Maximum = progress.Item2;
                                    ImportProgress.Step = 1;
                                    start = false;
                                }
                                ImportProgress.PerformStep();
                            }
                            
                        }));
                    },
                    cancellation
                    );
            }
            else
            {
                MessageBox.Show(
                    "Le chemin \r\n"
                        + inputPath + "\r\n" +
                        "est introuvable ou de pointe vers un dictionnaire ou un dossier valide.\r\n" +
                        "Veuillez le déplacer sur votre poste ou sélectionnez un autre fichier",
                    "Chemin de fichier ou dossier invalide",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                
            }
        }

        private void importProcess_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                this.Invoke(logging, "Traitement annulé\r\n");
                journal.AppendText("Traitement annulé");
            }
            else if (e.Error != null)
            {
                this.Invoke(logging, "Traitement en erreur : " + e.Error.Message + "\r\n");
            }
            else
            {
                this.Invoke(logging, "Traitement terminé\r\n");
            }
            launchScript.Text = "Lancer l\'import";
        }

        private void Consulter_Click(object sender, EventArgs e)
        {
            var consult = new ConsultationMot(this.sessionFactory);
            consult.Show();
        }
    }

}
