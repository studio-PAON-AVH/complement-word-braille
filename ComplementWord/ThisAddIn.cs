using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using fr.avh.braille.addin;
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


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Sauvegarder le fichier DBTCodes.dic sur le disque de l'utilisateur
            //string dbtCodes = Properties.Resources.DBTCodes;
            //DirectoryInfo appData = fr.avh.braille.dictionnaire.Globals.AppData;
            //string dbtCodesDicPath = Path.Combine(appData.FullName, "DBTCodes.dic");
            //using (StreamWriter text = File.CreateText(dbtCodesDicPath)) {
            //    text.Write(dbtCodes);
            //}
            //this.Application.CustomDictionaries.Add(dbtCodesDicPath);
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
