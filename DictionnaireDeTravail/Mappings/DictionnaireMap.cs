using FluentNHibernate.Mapping;

namespace fr.avh.braille.dictionnaire.Mappings
{
    public class DictionnaireMap : ClassMap<Entities.Dictionnaire>
    {
        public DictionnaireMap()
        {
            Id(x => x.Id);
            Map(x => x.Nom).Unique().Index("idx__NomDictionnaire");
            Map(x => x.DateAjout);
        }
    }
}
