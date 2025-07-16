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
        /// Selectionne et met en avant une occurence donnée
        /// </summary>
        /// <param name="index"></param>
        void AfficherOccurence(int index);

        //bool EstProtegerMot(int positionMotDansTexte);

        //bool EstProtegerBloc(int start, int end);

        /// <summary>
        /// Inject des codes de protection avant ou autour d'un emplacement dans le texte
        /// et renvoi le nombre de caractères ajouter avant et après l'emplacement.
        /// </summary>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <returns></returns>
        //Tuple<int, int> ProtegerEmplacementEtRenvoyerDecalages(int start, int end);

        /// <summary>
        /// Supprime les codes de protection avant ou autour d'un emplacement dans le texte
        /// et renvoi le nombre de caractères enlever avant et après l'emplacement.
        /// </summary>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <returns></returns>
        //Tuple<int, int> AbregerEmplacementEtRenvoyerDecalages(int start, int end);

        

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
