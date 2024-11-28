namespace fr.avh.braille.dictionnary
{
    partial class ImporterForm
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur Windows Form

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            this.selectedFilePath = new System.Windows.Forms.TextBox();
            this.browseFile = new System.Windows.Forms.Button();
            this.launchScript = new System.Windows.Forms.Button();
            this.browseFolder = new System.Windows.Forms.Button();
            this.journal = new System.Windows.Forms.TextBox();
            this.importProcess = new System.ComponentModel.BackgroundWorker();
            this.reanalyzeExistingDictionnaries = new System.Windows.Forms.CheckBox();
            this.Consulter = new System.Windows.Forms.Button();
            this.ImportProgress = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // selectedFilePath
            // 
            this.selectedFilePath.AllowDrop = true;
            this.selectedFilePath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.selectedFilePath.Location = new System.Drawing.Point(166, 20);
            this.selectedFilePath.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.selectedFilePath.Name = "selectedFilePath";
            this.selectedFilePath.Size = new System.Drawing.Size(577, 22);
            this.selectedFilePath.TabIndex = 3;
            this.selectedFilePath.TextChanged += new System.EventHandler(this.selectedFilePath_TextChanged);
            // 
            // browseFile
            // 
            this.browseFile.Location = new System.Drawing.Point(12, 11);
            this.browseFile.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.browseFile.Name = "browseFile";
            this.browseFile.Size = new System.Drawing.Size(148, 41);
            this.browseFile.TabIndex = 4;
            this.browseFile.Text = "Fichiers des dictionnaires";
            this.browseFile.UseVisualStyleBackColor = true;
            this.browseFile.Click += new System.EventHandler(this.browseFile_Click);
            // 
            // launchScript
            // 
            this.launchScript.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.launchScript.Location = new System.Drawing.Point(749, 11);
            this.launchScript.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.launchScript.Name = "launchScript";
            this.launchScript.Size = new System.Drawing.Size(142, 41);
            this.launchScript.TabIndex = 5;
            this.launchScript.Text = "Lancer l\'import";
            this.launchScript.UseVisualStyleBackColor = true;
            this.launchScript.Click += new System.EventHandler(this.launchScript_Click);
            // 
            // browseFolder
            // 
            this.browseFolder.Location = new System.Drawing.Point(12, 56);
            this.browseFolder.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.browseFolder.Name = "browseFolder";
            this.browseFolder.Size = new System.Drawing.Size(148, 41);
            this.browseFolder.TabIndex = 6;
            this.browseFolder.Text = "Dossier de dictionnaire";
            this.browseFolder.UseVisualStyleBackColor = true;
            this.browseFolder.Click += new System.EventHandler(this.browseFolder_Click);
            // 
            // journal
            // 
            this.journal.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.journal.Location = new System.Drawing.Point(12, 105);
            this.journal.Multiline = true;
            this.journal.Name = "journal";
            this.journal.Size = new System.Drawing.Size(879, 278);
            this.journal.TabIndex = 7;
            // 
            // importProcess
            // 
            this.importProcess.WorkerSupportsCancellation = true;
            this.importProcess.DoWork += new System.ComponentModel.DoWorkEventHandler(this.importProcess_DoWork);
            this.importProcess.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.importProcess_RunWorkerCompleted);
            // 
            // reanalyzeExistingDictionnaries
            // 
            this.reanalyzeExistingDictionnaries.AutoSize = true;
            this.reanalyzeExistingDictionnaries.Location = new System.Drawing.Point(166, 47);
            this.reanalyzeExistingDictionnaries.Name = "reanalyzeExistingDictionnaries";
            this.reanalyzeExistingDictionnaries.Size = new System.Drawing.Size(268, 20);
            this.reanalyzeExistingDictionnaries.TabIndex = 8;
            this.reanalyzeExistingDictionnaries.Text = "Réanalyser les dictionnaires déjà traités";
            this.reanalyzeExistingDictionnaries.UseVisualStyleBackColor = true;
            // 
            // Consulter
            // 
            this.Consulter.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.Consulter.Location = new System.Drawing.Point(749, 56);
            this.Consulter.Name = "Consulter";
            this.Consulter.Size = new System.Drawing.Size(142, 43);
            this.Consulter.TabIndex = 9;
            this.Consulter.Text = "Consulter la base";
            this.Consulter.UseVisualStyleBackColor = true;
            this.Consulter.Click += new System.EventHandler(this.Consulter_Click);
            // 
            // ImportProgress
            // 
            this.ImportProgress.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ImportProgress.Location = new System.Drawing.Point(12, 389);
            this.ImportProgress.Name = "ImportProgress";
            this.ImportProgress.Size = new System.Drawing.Size(879, 31);
            this.ImportProgress.TabIndex = 10;
            // 
            // Importer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(903, 432);
            this.Controls.Add(this.ImportProgress);
            this.Controls.Add(this.Consulter);
            this.Controls.Add(this.reanalyzeExistingDictionnaries);
            this.Controls.Add(this.journal);
            this.Controls.Add(this.browseFolder);
            this.Controls.Add(this.launchScript);
            this.Controls.Add(this.browseFile);
            this.Controls.Add(this.selectedFilePath);
            this.Name = "Importer";
            this.Text = "Charger un nouveau dictionnaire";
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.Launcher_OnDragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.Launcher_OnDragEnter);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox selectedFilePath;
        private System.Windows.Forms.Button browseFile;
        private System.Windows.Forms.Button launchScript;
        private System.Windows.Forms.Button browseFolder;
        private System.Windows.Forms.TextBox journal;
        private System.ComponentModel.BackgroundWorker importProcess;
        private System.Windows.Forms.CheckBox reanalyzeExistingDictionnaries;
        private System.Windows.Forms.Button Consulter;
        private System.Windows.Forms.ProgressBar ImportProgress;
    }
}

