using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace fr.avh.braille.dictionnaire.Entities
{
    public class Dictionnaire
    {
        /// <summary>
        /// Identifiant du dictionnaire dans la base de données
        /// </summary>
        public virtual int Id { get; protected set; }

        /// <summary>
        /// Nom du dictionnaire dans la base de données
        /// </summary>
        public virtual string Nom { get; set; }

        /// <summary>
        /// Date d'ajout du dictionnaire dans la base de données
        /// </summary>
        public virtual string DateAjout { get; set; }
    }
}
