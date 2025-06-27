using System;
using System.Collections.Generic;
using Task = System.Threading.Tasks.Task;
using fr.avh.archivage;
using System.Threading.Tasks;

namespace fr.avh.braille.addin
{
    public partial class ThisAddIn
    {
        public Dictionary<string, ProtectionWord> documentProtection = new Dictionary<string, ProtectionWord>();

        public Task<ProtectionWord> AnalyzeCurrentDocument(Utils.OnInfoCallback info, Utils.OnErrorCallback error)
        {
            return Task.Run(() =>
            {
                try {
                    if (!documentProtection.ContainsKey(Globals.ThisAddIn.Application.ActiveDocument.FullName)) {
                        documentProtection.Add(
                            Globals.ThisAddIn.Application.ActiveDocument.FullName,
                            new ProtectionWord(Globals.ThisAddIn.Application.ActiveDocument, info, error)
                        );
                    }
                    return documentProtection[Globals.ThisAddIn.Application.ActiveDocument.FullName];
                } catch (Exception e) {
                    info("Erreur lors de l'analyse du document : " + e.Message);
                    info(e.StackTrace);
                    error(e);
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
