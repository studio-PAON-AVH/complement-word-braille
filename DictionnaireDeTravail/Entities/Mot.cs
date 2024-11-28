
using System.Collections.Generic;

namespace fr.avh.braille.dictionnaire.Entities
{
    /// <summary>
    /// Mot identifié lors du traitement de documents pour le braille
    /// </summary>
    public class Mot
    {
        /// <summary>
        /// Identifiant du mot dans la base de données
        /// </summary>
        public virtual int Id { get; protected set; }
        /// <summary>
        /// Texte du mot
        /// </summary>
        public virtual string Texte { get; set; }

        /// <summary>
        /// Date d'ajout du mot dans la base de données
        /// </summary>
        public virtual string DateAjout { get; set; }

        /// <summary>
        /// Compteur des protections du mot
        /// </summary>
        public virtual long Protections { get; set; }

        /// <summary>
        /// Compteur des abréviations du mot
        /// </summary>
        public virtual long Abreviations { get; set; }

        /// <summary>
        /// Compteur du nombre de documents analysé et dans lesquels le mot a été trouvé
        /// </summary>
        public virtual long Documents { get; set; }

        /// <summary>
        /// Indique si le mot requiert systématiquement une action du transcripteur.
        /// <br/>
        /// Certains mots peuvent être ambigu et nécessiter systématiquement 
        /// une analyse contextuel : <br/>
        /// Le mot peut être un prénom dans certains contexte, ou bien appartenir a une autre langue
        /// ou si le mot peut aussi être un prénom dans certain cas.<br/>
        /// (Note : sqlite ne gère pas les booleans, on utilise 0 ou null pour la valeur "faux" 
        /// et 1 pour la valeur "vrai")
        /// </summary>
        public virtual int ToujoursDemander { get; set; }

        /// <summary>
        /// Commentaires ou notes sur le mot, tels que les ambiguités pouvant être rencontré
        ///
        /// </summary>
        public virtual string Commentaires { get; set; }

        /// <summary>
        /// Liaison entre le mot et les documents, avec un statut associé a cette liaison.
        /// </summary>
        //public virtual IList<StatutMotDocument> StatutsDocuments { get; set; }

        public Mot()
        {
            //StatutsDocuments = new List<StatutMotDocument>();
        }

        public Mot(Mot copy)
        {
            Id = copy.Id;
            Texte = copy.Texte;
            DateAjout = copy.DateAjout;
            Protections = copy.Protections;
            Abreviations = copy.Abreviations;
            Documents = copy.Documents;
            ToujoursDemander = copy.ToujoursDemander;
            Commentaires = copy.Commentaires;
        }
    }
}
