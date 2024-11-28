using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace fr.avh.braille.dictionnaire
{
    /// <summary>
    /// Interface devant être implémenter par un programme de protection de document
    /// pour intéragir avec le document
    /// 
    /// Pour le moment, pour le dictionnaire, nous n'avons besoin que de la méthode AppliquerStatutSurOccurence pour appliquer
    /// un statut particulier sur une occurence particulière
    /// </summary>
    public interface IProtection
    {
        /// <summary>
        /// Applique un statut choisi sur une occurence donnée
        /// </summary>
        /// <param name="index"></param>
        /// <param name="statut"></param>
        void AppliquerStatutSurOccurence(int index, Statut statut);

        /// <summary>
        /// Remet l'occurence en avant dans le contenu
        /// </summary>
        /// <param name="index"></param>
        void AfficherOccurence(int index);
    }
}
