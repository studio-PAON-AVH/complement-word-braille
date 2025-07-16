namespace fr.avh.braille.addin
{
    partial class RubanBraille : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RubanBraille()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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
            this.protectorRibbonAVH = this.Factory.CreateRibbonTab();
            this.BrailleActions = this.Factory.CreateRibbonGroup();
            this.LancerTraitement = this.Factory.CreateRibbonButton();
            this.MotsHorsLexique = this.Factory.CreateRibbonButton();
            this.ChargerDecisions = this.Factory.CreateRibbonButton();
            this.InfoSelection = this.Factory.CreateRibbonGroup();
            this.MotSelectionner = this.Factory.CreateRibbonLabel();
            this.Progression = this.Factory.CreateRibbonLabel();
            this.Refocus = this.Factory.CreateRibbonButton();
            this.Navigation = this.Factory.CreateRibbonGroup();
            this.PremierMot = this.Factory.CreateRibbonButton();
            this.OccurencePrecedente = this.Factory.CreateRibbonButton();
            this.OccurenceSuivante = this.Factory.CreateRibbonButton();
            this.Data = this.Factory.CreateRibbonGroup();
            this.AfficherDictionnaire = this.Factory.CreateRibbonButton();
            this.ProtectionDB = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.ProtegerSelection = this.Factory.CreateRibbonButton();
            this.AbregerSelection = this.Factory.CreateRibbonButton();
            this.Options = this.Factory.CreateRibbonGroup();
            this.OuvrirOptions = this.Factory.CreateRibbonButton();
            this.protectorRibbonAVH.SuspendLayout();
            this.BrailleActions.SuspendLayout();
            this.InfoSelection.SuspendLayout();
            this.Navigation.SuspendLayout();
            this.Data.SuspendLayout();
            this.group1.SuspendLayout();
            this.Options.SuspendLayout();
            this.SuspendLayout();
            // 
            // protectorRibbonAVH
            // 
            this.protectorRibbonAVH.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.protectorRibbonAVH.Groups.Add(this.BrailleActions);
            this.protectorRibbonAVH.Groups.Add(this.InfoSelection);
            this.protectorRibbonAVH.Groups.Add(this.Navigation);
            this.protectorRibbonAVH.Groups.Add(this.Data);
            this.protectorRibbonAVH.Groups.Add(this.group1);
            this.protectorRibbonAVH.Groups.Add(this.Options);
            this.protectorRibbonAVH.Label = "Braille";
            this.protectorRibbonAVH.Name = "protectorRibbonAVH";
            // 
            // BrailleActions
            // 
            this.BrailleActions.Items.Add(this.LancerTraitement);
            this.BrailleActions.Items.Add(this.MotsHorsLexique);
            this.BrailleActions.Items.Add(this.ChargerDecisions);
            this.BrailleActions.Label = "Actions";
            this.BrailleActions.Name = "BrailleActions";
            // 
            // LancerTraitement
            // 
            this.LancerTraitement.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.LancerTraitement.Label = "&Lancer le traitement";
            this.LancerTraitement.Name = "LancerTraitement";
            this.LancerTraitement.ShowImage = true;
            this.LancerTraitement.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LancerTraitementMots_Click);
            // 
            // MotsHorsLexique
            // 
            this.MotsHorsLexique.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.MotsHorsLexique.Enabled = false;
            this.MotsHorsLexique.Label = "&Mots hors lexique";
            this.MotsHorsLexique.Name = "MotsHorsLexique";
            this.MotsHorsLexique.ShowImage = true;
            this.MotsHorsLexique.Visible = false;
            this.MotsHorsLexique.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.MotsHorsLexique_Click);
            // 
            // ChargerDecisions
            // 
            this.ChargerDecisions.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ChargerDecisions.Label = "&Charger des décisions";
            this.ChargerDecisions.Name = "ChargerDecisions";
            this.ChargerDecisions.ShowImage = true;
            this.ChargerDecisions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ChargerDecisions_Click);
            // 
            // InfoSelection
            // 
            this.InfoSelection.Items.Add(this.MotSelectionner);
            this.InfoSelection.Items.Add(this.Progression);
            this.InfoSelection.Items.Add(this.Refocus);
            this.InfoSelection.Label = "Statut";
            this.InfoSelection.Name = "InfoSelection";
            // 
            // MotSelectionner
            // 
            this.MotSelectionner.Label = "Mot : ";
            this.MotSelectionner.Name = "MotSelectionner";
            // 
            // Progression
            // 
            this.Progression.Label = "Progression :";
            this.Progression.Name = "Progression";
            // 
            // Refocus
            // 
            this.Refocus.Label = "Resélectionner";
            this.Refocus.Name = "Refocus";
            this.Refocus.ShowImage = true;
            this.Refocus.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Refocus_Click);
            // 
            // Navigation
            // 
            this.Navigation.Items.Add(this.PremierMot);
            this.Navigation.Items.Add(this.OccurencePrecedente);
            this.Navigation.Items.Add(this.OccurenceSuivante);
            this.Navigation.Label = "Navigation";
            this.Navigation.Name = "Navigation";
            // 
            // PremierMot
            // 
            this.PremierMot.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.PremierMot.Label = "Revenir au début";
            this.PremierMot.Name = "PremierMot";
            this.PremierMot.ShowImage = true;
            this.PremierMot.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PremierMot_Click);
            // 
            // OccurencePrecedente
            // 
            this.OccurencePrecedente.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.OccurencePrecedente.Label = "Mot précédent";
            this.OccurencePrecedente.Name = "OccurencePrecedente";
            this.OccurencePrecedente.ShowImage = true;
            this.OccurencePrecedente.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OccurencePrecedente_Click);
            // 
            // OccurenceSuivante
            // 
            this.OccurenceSuivante.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.OccurenceSuivante.Label = "Mot suivant";
            this.OccurenceSuivante.Name = "OccurenceSuivante";
            this.OccurenceSuivante.ShowImage = true;
            this.OccurenceSuivante.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OccurenceSuivante_Click);
            // 
            // Data
            // 
            this.Data.Items.Add(this.AfficherDictionnaire);
            this.Data.Items.Add(this.ProtectionDB);
            this.Data.Label = "Données";
            this.Data.Name = "Data";
            // 
            // AfficherDictionnaire
            // 
            this.AfficherDictionnaire.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AfficherDictionnaire.Label = "Dictionnaire de travail";
            this.AfficherDictionnaire.Name = "AfficherDictionnaire";
            this.AfficherDictionnaire.ShowImage = true;
            this.AfficherDictionnaire.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AfficherDictionnaire_Click);
            // 
            // ProtectionDB
            // 
            this.ProtectionDB.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ProtectionDB.Label = "Base de protection";
            this.ProtectionDB.Name = "ProtectionDB";
            this.ProtectionDB.ShowImage = true;
            this.ProtectionDB.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProtectionDB_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.ProtegerSelection);
            this.group1.Items.Add(this.AbregerSelection);
            this.group1.Label = "Sélection";
            this.group1.Name = "group1";
            // 
            // ProtegerSelection
            // 
            this.ProtegerSelection.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ProtegerSelection.Label = "&Proteger la sélection";
            this.ProtegerSelection.Name = "ProtegerSelection";
            this.ProtegerSelection.ShowImage = true;
            this.ProtegerSelection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProtegerSelection_Click);
            // 
            // AbregerSelection
            // 
            this.AbregerSelection.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AbregerSelection.Label = "&Abréger la sélection";
            this.AbregerSelection.Name = "AbregerSelection";
            this.AbregerSelection.ShowImage = true;
            this.AbregerSelection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AbregerSelection_Click);
            // 
            // Options
            // 
            this.Options.Items.Add(this.OuvrirOptions);
            this.Options.Label = "Options";
            this.Options.Name = "Options";
            // 
            // OuvrirOptions
            // 
            this.OuvrirOptions.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.OuvrirOptions.Label = "Options";
            this.OuvrirOptions.Name = "OuvrirOptions";
            this.OuvrirOptions.ShowImage = true;
            this.OuvrirOptions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OuvrirOptions_Click);
            // 
            // RubanBraille
            // 
            this.Name = "RubanBraille";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.protectorRibbonAVH);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RubanAVH_Load);
            this.protectorRibbonAVH.ResumeLayout(false);
            this.protectorRibbonAVH.PerformLayout();
            this.BrailleActions.ResumeLayout(false);
            this.BrailleActions.PerformLayout();
            this.InfoSelection.ResumeLayout(false);
            this.InfoSelection.PerformLayout();
            this.Navigation.ResumeLayout(false);
            this.Navigation.PerformLayout();
            this.Data.ResumeLayout(false);
            this.Data.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.Options.ResumeLayout(false);
            this.Options.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab protectorRibbonAVH;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup BrailleActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AfficherDictionnaire;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Navigation;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton OccurencePrecedente;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton OccurenceSuivante;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Data;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ProtectionDB;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton PremierMot;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup InfoSelection;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel MotSelectionner;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel Progression;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Refocus;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ProtegerSelection;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AbregerSelection;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Options;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton OuvrirOptions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ChargerDecisions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton LancerTraitement;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton MotsHorsLexique;
    }

    partial class ThisRibbonCollection
    {
        internal RubanBraille RubanAVH
        {
            get { return this.GetRibbon<RubanBraille>(); }
        }
    }
}
