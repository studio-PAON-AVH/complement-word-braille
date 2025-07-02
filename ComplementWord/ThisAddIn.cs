using System;
using System.Collections.Generic;
using Task = System.Threading.Tasks.Task;
using fr.avh.archivage;
using System.Threading.Tasks;
using System.Windows.Threading;
using System.IO;
using System.Text;
using System.Windows;

namespace fr.avh.braille.addin
{
    public partial class ThisAddIn
    {
        public Dictionary<string, ProtectionWord> documentProtection = new Dictionary<string, ProtectionWord>();

        public Dictionary<string, TraitementBraille> journal = new Dictionary<string, TraitementBraille>();

        TraitementBraille progressDialog = null;

        Dispatcher _dispatcher = Dispatcher.CurrentDispatcher;
        public void InfoCallback(string file, string message, Tuple<int, int> progessTuple = null)
        {
            _dispatcher.Invoke(() =>
            {
                TraitementBraille progressDialog;
                if (!journal.ContainsKey(file) || !journal[file].IsLoaded) {
                    progressDialog = new TraitementBraille(Path.GetFileNameWithoutExtension(file));
                    journal[file] =  progressDialog;
                } else {
                    progressDialog = journal[file];
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

        public void ErrorCallback(string file, Exception ex)
        {
            _dispatcher.Invoke(() =>
            {
                TraitementBraille progressDialog;
                if (!journal.ContainsKey(file) || !journal[file].IsLoaded) {
                    progressDialog = new TraitementBraille(Path.GetFileNameWithoutExtension(file));
                    journal[file] = progressDialog;
                } else {
                    progressDialog = journal[file];
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


        public Task<ProtectionWord> AnalyzeCurrentDocument()
        {
            return Task.Run(() =>
            {
                try {
                    if (!documentProtection.ContainsKey(Globals.ThisAddIn.Application.ActiveDocument.FullName)) {
                        documentProtection.Add(
                            Globals.ThisAddIn.Application.ActiveDocument.FullName,
                            new ProtectionWord(Globals.ThisAddIn.Application.ActiveDocument, (m, t) =>
                            {
                                this.InfoCallback(file: Globals.ThisAddIn.Application.ActiveDocument.FullName, m, t);
                            }, (ex) =>
                            {
                                this.ErrorCallback(file: Globals.ThisAddIn.Application.ActiveDocument.FullName, ex);
                            })
                        );
                    }
                    return documentProtection[Globals.ThisAddIn.Application.ActiveDocument.FullName];
                } catch (Exception e) {
                    InfoCallback(Globals.ThisAddIn.Application.ActiveDocument.FullName, "Erreur lors de l'analyse du document : " + e.Message+"\r\n"+e.StackTrace);
                    ErrorCallback(file: Globals.ThisAddIn.Application.ActiveDocument.FullName, e);
                    return null;
                }
                
            });
            
        }
       
        //private BrailleTaskPaneHolder _brailleTaskPane;
        //private Microsoft.Office.Tools.CustomTaskPane _myCustomTaskPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            AddinUpdater.CheckForUpdate();
            // Pour plus tard : remplacer la fenêtre du journal pas un task pane qui hébergera le calcul pour le document courant
            // 
            //_brailleTaskPane = new BrailleTaskPaneHolder();
            //_myCustomTaskPane = this.CustomTaskPanes.Add(
            //    _brailleTaskPane,
            //    "Traitement du braille"
            //);
            //_myCustomTaskPane.Visible = true;
            // Sauvegarder le fichier DBTCodes.dic sur le disque de l'utilisateur
            //string dbtCodes = Properties.Resources.DBTCodes;
            //DirectoryInfo appData = fr.avh.braille.dictionnaire.Globals.AppData;
            //string dbtCodesDicPath = Path.Combine(appData.FullName, "DBTCodes.dic");
            //using (StreamWriter text = File.CreateText(dbtCodesDicPath)) {
            //    text.Write(dbtCodes);
            //}
            //this.Application.CustomDictionaries.Add(dbtCodesDicPath);

        }

        private void ThisAddin_DocumentChange()
        {
            // TODO : lors d'un changement de document, charger les données d'analyse correspondante dans le taskpane
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }

        #region Code généré par VSTO

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
