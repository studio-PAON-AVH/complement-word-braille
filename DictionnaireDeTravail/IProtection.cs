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
        /// Applique un statut choisi sur une plage du texte
        /// </summary>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <param name="statut"></param>
        void AppliquerStatutSurBlock(int start, int end, Statut statut);

        void AfficherOccurence(int index);


        // Pour refactorisation, generalisation du traitement
        // L'idée est que le traitement soit portable a d'autre programme ultérieurement

        //DictionnaireDeTravail DonneesTraitement { get; }

        //string MotSelectionne { get; }

        //object ProchainMot();

        //object PrecedentMot();

        //int OccurenceSelectionnee { get; set; }

        //object ProchaineOccurence();

        //object PrecedenteOccurence();

        // Notes pour refactorisation ultérieur
        // Renommer l'interface en IDocument ou créer une classe abstraite DocumentBrailleBase
        // Ajouter les méthodes de navigation suivantes
        // Constructeur => créer un dictionnaire de travail vide
        // - Analyser document => détecter les occurences de mots et
        // - Selectionner Occurence
        // - Selectionner Mot
        // - Prochaine Occurence
        // - Prochain Mot
        // - Appliquer statut sur occurence
        // - Proteger emplacement (debut et fin) et renvoi le nombre de caractère ajouter


    }
}
