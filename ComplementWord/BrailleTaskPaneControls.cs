using fr.avh.braille.addin;
using fr.avh.braille.dictionnary.Entities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace fr.avh.braille.addin
{
    public partial class BrailleTaskPaneControls : UserControl
    {
        TreeNode currentWordNode;
        TreeNode currentOccurenceNode;

        List<Status> statusList = new List<Status>()
        {
            Status.INCONNU,
            Status.ABREGER,
            Status.PROTEGER,
            Status.AMBIGU
        };

        DocumentProtector boundProtector = null;
        public BrailleTaskPaneControls(DocumentProtector protector = null)
        {
            InitializeComponent();
            if(protector != null)
            {
                bindProtectionTool(protector);
            }
            this.StatusMotDocument.Items.AddRange((object[])statusList.Select(w => w.ToString()).ToArray());
        }

        public void bindProtectionTool(DocumentProtector protector)
        {
            
            VueMotsStatus.Rows.Clear();
            foreach(var wordFound in protector.wordOccurences.OrderBy((k) => k.Key))
            {
                var action = protector.wordSelectedAction.ContainsKey(wordFound.Key) ?
                    protector.wordSelectedAction[wordFound.Key] :
                    0;
                VueMotsStatus.Rows.Add(wordFound.Key, statusList.Find(w => w.Code == action).ToString());
            }

            WordTree.Nodes.Clear();
            this.boundProtector = protector;
            foreach (var wordAndOccurences in protector.wordOccurences.OrderBy((k) => k.Key))
            {
                WordTree.Nodes.Add(wordAndOccurences.Key,wordAndOccurences.Key);
                foreach (var occurence in wordAndOccurences.Value)
                {
                    WordTree.Nodes[wordAndOccurences.Key].Nodes.Add(occurence.ToString(), protector.occurencesRequiringAction[occurence]);
                }

            }
            currentWordNode = WordTree.Nodes[protector.SelectedWord];
            currentWordNode.ExpandAll();
            currentOccurenceNode = currentWordNode.Nodes[protector.SelectedOccurenceIndex.ToString()];
            WordTree.Select();
            WordTree.SelectedNode = currentOccurenceNode;

            boundProtector.addSelectionChangeCallBack(onSelectedOccurence);
            boundProtector.addOnWordActionSelectedCallBack(onWordActionChanged);
            boundProtector.addOnOccurenceActionSelectedCallBack(onOccurenceActionChanged);
        }

        public void onSelectedOccurence(int newOccurence)
        {
            if(this.boundProtector != null)
            {
                currentWordNode = WordTree.Nodes[boundProtector.SelectedWord];
                currentWordNode.ExpandAll();
                currentOccurenceNode = currentWordNode.Nodes[boundProtector.SelectedOccurenceIndex.ToString()];
                //WordTree.Select();
                WordTree.SelectedNode = currentOccurenceNode;

            }
        }

        public void onWordActionChanged(string mot, Status newStatus)
        {
            WordTree.Nodes[mot.ToLower()].Text = mot.ToLower() + " - " + newStatus.Nom;
        }

        public void onOccurenceActionChanged(int occurenceIndex, Status newStatus)
        {
            string mot = boundProtector.occurencesRequiringAction[occurenceIndex];
            TreeNode nodeMot = WordTree.Nodes[mot.ToLower()];
            TreeNode nodeOccurence = nodeMot.Nodes[occurenceIndex.ToString()];
            WordTree.Nodes[mot.ToLower()].Nodes[occurenceIndex.ToString()].Text = mot + " - " + newStatus.Nom;
        }

        private void WordTree_AfterSelect(object sender, TreeViewEventArgs e)
        {
            TreeNode selected = e.Node;
            if(selected != null)
            {
                if(selected.Parent != null)
                {
                    // Trouver la clé et sélectionné l'occurence
                }
            }
        }

        private void VueMotsStatus_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            // Charger la liste des occurences
        }
    }
}
