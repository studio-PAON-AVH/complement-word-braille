
using System.Collections.Generic;
using System.Linq;

namespace fr.avh.braille.dictionnaire
{
    public class MotAfficher
    {
        private static Dictionary<string, Statut> _statuts = new Dictionary<string, Statut> {
            { "Inconnu", Statut.INCONNU },
            { "Abréger", Statut.ABREGE },
            { "Protéger", Statut.PROTEGE },
            { "Ignorer", Statut.IGNORE }
        };
        private static Dictionary<Statut, string> _revert = _statuts.ToDictionary((kvp) => kvp.Value, (kvp) => kvp.Key);
        public string[] StatutsPossible { get => _statuts.Keys.ToArray(); }

        public string StatutChoisi { 
            get => _revert[Statut]; 
            set { 
                Statut = _statuts[value];
            }
        }

        public string Texte { get; set; }

        public string ContexteAvant { get; set; } = null;
        public string ContexteApres { get; set; } = null;
        public string Contexte { get; set; } = null;
        public Statut Statut { get; set; }

        public int Index { get; set; }
    }
}
