using FluentNHibernate.Mapping;

namespace fr.avh.braille.dictionnaire.Mappings
{
    public class DecisionParDefautMap : ClassMap<Entities.DecisionParDefaut>
    {
        public DecisionParDefautMap()
        {
            Id(x => x.Id);
            Map(x => x.Mot);
            Map(x => x.Decision);
            Map(x => x.Commentaires);
        }
    }
}
