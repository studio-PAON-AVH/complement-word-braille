namespace fr.avh.braille.addin
{
    partial class BrailleTaskPaneControls
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

        #region Code généré par le Concepteur de composants

        /// <summary> 
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas 
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.WordTree = new System.Windows.Forms.TreeView();
            this.StatusAnalyse = new System.Windows.Forms.Label();
            this.VueMotsStatus = new System.Windows.Forms.DataGridView();
            this.documentProtectionToolBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.MotDocument = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.StatusMotDocument = new System.Windows.Forms.DataGridViewComboBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.VueMotsStatus)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.documentProtectionToolBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // WordTree
            // 
            this.WordTree.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.WordTree.Location = new System.Drawing.Point(3, 41);
            this.WordTree.Name = "WordTree";
            this.WordTree.Size = new System.Drawing.Size(375, 150);
            this.WordTree.TabIndex = 0;
            this.WordTree.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.WordTree_AfterSelect);
            // 
            // StatusAnalyse
            // 
            this.StatusAnalyse.AutoSize = true;
            this.StatusAnalyse.Location = new System.Drawing.Point(7, 12);
            this.StatusAnalyse.Name = "StatusAnalyse";
            this.StatusAnalyse.Size = new System.Drawing.Size(44, 16);
            this.StatusAnalyse.TabIndex = 1;
            this.StatusAnalyse.Text = "label1";
            // 
            // VueMotsStatus
            // 
            this.VueMotsStatus.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.VueMotsStatus.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.VueMotsStatus.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.MotDocument,
            this.StatusMotDocument});
            this.VueMotsStatus.Location = new System.Drawing.Point(4, 197);
            this.VueMotsStatus.MultiSelect = false;
            this.VueMotsStatus.Name = "VueMotsStatus";
            this.VueMotsStatus.RowHeadersWidth = 51;
            this.VueMotsStatus.RowTemplate.Height = 24;
            this.VueMotsStatus.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.VueMotsStatus.Size = new System.Drawing.Size(374, 260);
            this.VueMotsStatus.TabIndex = 2;
            this.VueMotsStatus.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.VueMotsStatus_CellEnter);
            // 
            // documentProtectionToolBindingSource
            // 
            this.documentProtectionToolBindingSource.DataSource = typeof(AVH.braille.protection.dictionnary.DocumentProtectionTool);
            // 
            // MotDocument
            // 
            this.MotDocument.HeaderText = "Mot";
            this.MotDocument.MinimumWidth = 6;
            this.MotDocument.Name = "MotDocument";
            this.MotDocument.ReadOnly = true;
            this.MotDocument.Width = 150;
            // 
            // StatusMotDocument
            // 
            this.StatusMotDocument.HeaderText = "Status";
            this.StatusMotDocument.MinimumWidth = 6;
            this.StatusMotDocument.Name = "StatusMotDocument";
            this.StatusMotDocument.Width = 125;
            // 
            // BrailleTaskPaneControls
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.VueMotsStatus);
            this.Controls.Add(this.StatusAnalyse);
            this.Controls.Add(this.WordTree);
            this.Name = "BrailleTaskPaneControls";
            this.Size = new System.Drawing.Size(381, 745);
            ((System.ComponentModel.ISupportInitialize)(this.VueMotsStatus)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.documentProtectionToolBindingSource)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TreeView WordTree;
        private System.Windows.Forms.Label StatusAnalyse;
        private System.Windows.Forms.DataGridView VueMotsStatus;
        private System.Windows.Forms.BindingSource documentProtectionToolBindingSource;
        private System.Windows.Forms.DataGridViewTextBoxColumn MotDocument;
        private System.Windows.Forms.DataGridViewComboBoxColumn StatusMotDocument;
    }
}
