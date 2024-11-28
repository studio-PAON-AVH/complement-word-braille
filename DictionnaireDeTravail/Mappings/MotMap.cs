using FluentNHibernate.Mapping;

namespace fr.avh.braille.dictionnaire.Mappings
{
    public class MotMap : ClassMap<Entities.Mot>
    {
        public MotMap()
        {
            Id(x => x.Id);
            Map(x => x.Texte).Unique().Index("idx__Texte");
            Map(x => x.DateAjout);
            Map(x => x.Protections);
            Map(x => x.Abreviations);
            Map(x => x.Documents);
            Map(x => x.ToujoursDemander);
            Map(x => x.Commentaires);
        }
    }
}
