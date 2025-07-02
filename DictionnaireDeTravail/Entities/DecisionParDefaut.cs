using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace fr.avh.braille.dictionnaire.Entities
{
    public class DecisionParDefaut
    {
        /// <summary>
        /// 
        /// </summary>
        public virtual int Id { get; protected set; }

        /// <summary>
        /// 
        /// </summary>
        public virtual string Mot { get; set; }

        /// <summary>
        /// decision retourné par les transcripteurs : a p ou c (abrege, proteger ou context)
        /// </summary>
        public virtual char Decision { get; set; }

        /// <summary>
        /// commentaire retourné par les transcripteurs : a p ou c (abrege, proteger ou context)
        /// </summary>
        public virtual string Commentaires { get; set; }
    }
}
