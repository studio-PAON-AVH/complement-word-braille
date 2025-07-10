
using fr.avh.archivage;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace fr.avh.braille.dictionnaire
{
    /// <summary>
    /// Controle d'abreviabilité de mots pour le braille abrégé.<br/>
    /// Se base sur le manuel de braille abrégé de l'AVH :<br/>
    /// <see href="https://www.avh.asso.fr/sites/default/files/manuel_abrege2013_complet_noir_0.pdf"/>
    /// </summary>
    public static class Abreviation
    {
        /// <summary>
        /// Liste des mots abréger extrait du manuel de braille abrégé de l'AVH :
        /// https://www.avh.asso.fr/sites/default/files/manuel_abrege2013_complet_noir_0.pdf
        /// </summary>
        public static readonly HashSet<string> MotsAbreger = new HashSet<string>
        {
            "absolu",
            "absolument",
            "à cause",
            "action",
            "actionnaire",
            "affaire",
            "afin",
            "ai",
            "ailleurs",
            "ainsi",
            "alors",
            "à mesure",
            "amour",
            "amoureuse",
            "amoureusement",
            "amoureux",
            "à peine",
            "à peu près",
            "apparemment",
            "apparence",
            "apparent",
            "après",
            "à présent",
            "assez",
            "à travers",
            "attentif",
            "attention",
            "attentive",
            "attentivement",
            "au",
            "au contraire",
            "aucun",
            "aucune",
            "aucunement",
            "au-dessous",
            "au-dessus",
            "aujourd'hui",
            "auparavant",
            "auprès",
            "auquel",
            "aussi",
            "aussitôt",
            "autant que",
            "autour",
            "autre",
            "autre chose",
            "autrefois",
            "autrement",
            "autre part",
            "auxquelles",
            "auxquels",
            "avance",
            "avancement",
            "avant",
            "avantage",
            "avantageuse",
            "avantageusement",
            "avantageux",
            "avec",
            "avoir",
            "ayant",
            "beaucoup",
            "besogne",
            "besogneuse",
            "besogneux",
            "besoin",
            "bête",
            "bêtement",
            "bien",
            "bienfaisance",
            "bienfait",
            "bienfaiteur",
            "bientôt",
            "bienveillance",
            "bienveillant",
            "bizarre",
            "bizarrement",
            "bonheur",
            "bonjour",
            "bonne",
            "bonnement",
            "bonté",
            "boulevard",
            "braille",
            "branche",
            "branchement",
            "brave",
            "bravement",
            "bruit",
            "brusque",
            "brusquement",
            "budget",
            "budgétaire",
            "caractère",
            "caractéristique",
            "ce",
            "ceci",
            "cela",
            "celle",
            "celui",
            "cependant",
            "certain",
            "certaine",
            "certainement",
            "certes",
            "certitude",
            "ces",
            "c'est-à-dire",
            "cet  î",
            "cette",
            "ceux",
            "chacun",
            "chacune",
            "chagrin",
            "chaleur",
            "chaleureuse",
            "chaleureusement",
            "chaleureux",
            "champ",
            "change",
            "changement",
            "chaque",
            "charitable",
            "charitablement",
            "charité",
            "chaud",
            "chaudement",
            "chemin",
            "chère",
            "chèrement",
            "chez",
            "chiffrage",
            "chiffre",
            "choeur",
            "choix",
            "chose",
            "circonstance",
            "circonstanciel",
            "circonstancielle",
            "civil",
            "civilement",
            "civilisation",
            "civilité",
            "coeur",
            "combien",
            "comme",
            "commencement",
            "comment",
            "commentaire",
            "commun",
            "communal",
            "communautaire",
            "communauté",
            "communaux",
            "communément",
            "communion",
            "complément",
            "complémentaire",
            "complet",
            "complète",
            "complètement",
            "conclusion",
            "condition",
            "conditionnel",
            "conditionnelle",
            "conditionnellement",
            "confiance",
            "confiant",
            "congrès",
            "connaissance",
            "connaître",
            "consciemment",
            "conscience",
            "consciencieuse",
            "consciencieusement",
            "consciencieux",
            "conscient",
            "conséquemment",
            "conséquence",
            "conséquent",
            "considérable",
            "considérablement",
            "considération",
            "contraire",
            "contrairement",
            "contre",
            "conversation",
            "côté",
            "couple",
            "courage",
            "courageuse",
            "courageusement",
            "courageux",
            "d'abord",
            "danger",
            "dangereuse",
            "dangereusement",
            "dangereux",
            "dans",
            "davantage",
            "de",
            "debout",
            "dedans",
            "degré",
            "dehors",
            "déjà",
            "demain",
            "depuis",
            "dernier",
            "dernière",
            "dernièrement",
            "derrière",
            "des",
            "dès",
            "désormais",
            "desquelles",
            "desquels",
            "destin",
            "destinataire",
            "destination",
            "de suite",
            "devant",
            "différemment",
            "différence",
            "différent",
            "difficile",
            "difficilement",
            "difficulté",
            "digne",
            "dignement",
            "dignitaire",
            "dignité",
            "discours",
            "dispositif",
            "disposition",
            "distance",
            "distant",
            "donc",
            "dont",
            "douleur",
            "douloureuse",
            "douloureusement",
            "douloureux",
            "doute",
            "du",
            "duquel",
            "effectif",
            "effective",
            "effectivement",
            "effet",
            "égal",
            "également",
            "égalitaire",
            "égalité",
            "égaux",
            "élément",
            "élémentaire",
            "elle",
            "en",
            "encore",
            "endroit",
            "énergie",
            "énergique",
            "énergiquement",
            "enfin",
            "en mesure",
            "ennui",
            "ennuyeuse",
            "ennuyeux",
            "enquête",
            "enquêteur",
            "enquêteuse",
            "en réalité",
            "ensemble",
            "ensuite",
            "entier",
            "entière",
            "entièrement",
            "environ",
            "espèce",
            "espérance",
            "espoir",
            "esprit",
            "essentiel",
            "essentielle",
            "essentiellement",
            "est",
            "et",
            "étant",
            "et caetera",
            "été",
            "être",
            "événement",
            "éventualité",
            "éventuel",
            "éventuelle",
            "éventuellement",
            "excellence",
            "excellent",
            "excès",
            "excessif",
            "excessive",
            "excessivement",
            "exercice",
            "expérience",
            "expérimental",
            "expérimentalement",
            "expérimentateur",
            "expérimentation",
            "expérimentaux",
            "explicatif",
            "explication",
            "explicative",
            "expressif",
            "expression",
            "expressive",
            "expressivement",
            "extérieur",
            "extérieurement",
            "extrême",
            "extrêmement",
            "extrémité",
            "facile",
            "facilement",
            "facilité",
            "faire",
            "faut",
            "faute",
            "faveur",
            "favorable",
            "favorablement",
            "féminin",
            "féminine",
            "femme",
            "fête",
            "fidèle",
            "fidèlement",
            "fidélité",
            "figuration",
            "figure",
            "fille",
            "fils",
            "fonction",
            "fonctionnaire",
            "fonctionnel",
            "fonctionnement",
            "force",
            "forcément",
            "fortune",
            "fraternel",
            "fraternelle",
            "fraternellement",
            "fraternité",
            "fréquemment",
            "fréquence",
            "fréquent",
            "fréquentation",
            "frère",
            "garde",
            "général",
            "généralement",
            "généralisation",
            "généralité",
            "généraux",
            "généreuse",
            "généreusement",
            "généreux",
            "générosité",
            "gloire",
            "glorieuse",
            "glorieusement",
            "glorieux",
            "gouvernement",
            "gouvernemental",
            "gouvernementaux",
            "gouverneur",
            "grâce",
            "gracieuse",
            "gracieusement",
            "gracieux",
            "grand",
            "grande",
            "grandement",
            "grandeur",
            "grave",
            "gravement",
            "gravité",
            "groupe",
            "groupement",
            "guère",
            "guerre",
            "habitude",
            "habituel",
            "habituelle",
            "habituellement",
            "hasard",
            "hasardeuse",
            "hasardeux",
            "hélas",
            "heure",
            "heureuse",
            "heureusement",
            "heureux",
            "hier",
            "histoire",
            "historique",
            "historiquement",
            "hiver",
            "hommage",
            "homme",
            "honnête",
            "honnêtement",
            "honnêteté",
            "honneur",
            "honorable",
            "honorablement",
            "honoraire",
            "horaire",
            "horizon",
            "horizontal",
            "horizontalement",
            "horizontalité",
            "horizontaux",
            "humain",
            "humaine",
            "humainement",
            "humanitaire",
            "humanité",
            "hypothèse",
            "hypothétique",
            "hypothétiquement",
            "idéal",
            "idéalement",
            "idéaux",
            "idée",
            "il",
            "ils",
            "image",
            "imaginable",
            "imaginaire",
            "imagination",
            "immédiat",
            "immédiatement",
            "impression",
            "impressionnable",
            "inférieur",
            "inférieurement",
            "infériorité",
            "inquiet",
            "inquiète",
            "inquiétude",
            "intelligemment",
            "intelligence",
            "intelligent",
            "intérieur",
            "intérieurement",
            "jadis",
            "jamais",
            "je",
            "jeune",
            "jour",
            "joyeuse",
            "joyeusement",
            "joyeux",
            "juge",
            "jugement",
            "jusque",
            "juste",
            "justement",
            "justice",
            "la",
            "la plupart",
            "laquelle",
            "le",
            "lecture",
            "lequel",
            "les",
            "lesquelles",
            "lesquels",
            "lettre",
            "libéral",
            "libéralité",
            "libération",
            "libéraux",
            "liberté",
            "libre",
            "librement",
            "ligne",
            "livre",
            "logique",
            "logiquement",
            "loin",
            "lointain",
            "lointaine",
            "longtemps",
            "lorsque",
            "lourd",
            "lourdement",
            "lourdeur",
            "lui",
            "lumière",
            "lumineuse",
            "lumineusement",
            "lumineux",
            "luminosité",
            "madame",
            "mademoiselle",
            "magnificence",
            "magnifique",
            "magnifiquement",
            "maintenant",
            "mais",
            "malgré",
            "malheur",
            "malheureuse",
            "malheureusement",
            "malheureux",
            "manière",
            "mauvais",
            "me",
            "meilleur",
            "même",
            "merci",
            "mère",
            "mes",
            "mesdames",
            "mesdemoiselles",
            "messieurs",
            "mettre",
            "mieux",
            "mission",
            "missionnaire",
            "mobile",
            "mobilisation",
            "mobilité",
            "moins",
            "moment",
            "momentanément",
            "monsieur",
            "multiple",
            "multiplicateur",
            "multiplication",
            "multiplicité",
            "musique",
            "mystère",
            "mystérieuse",
            "mystérieusement",
            "mystérieux",
            "naguère",
            "nation",
            "national",
            "nationalité",
            "nationaux",
            "nature",
            "naturel",
            "naturelle",
            "naturellement",
            "ne",
            "néanmoins",
            "nécessaire",
            "nécessairement",
            "nécessité",
            "nécessiteuse",
            "nécessiteux",
            "nombre",
            "nombreuse",
            "nombreux",
            "non seulement",
            "nos",
            "notre",
            "nôtre",
            "nous",
            "nouveau",
            "nouveauté",
            "nouvel",
            "nouvelle",
            "nouvellement",
            "objectif",
            "objection",
            "objective",
            "objectivement",
            "objectivité",
            "objet",
            "observateur",
            "observation",
            "occasion",
            "occasionnel",
            "occasionnelle",
            "occasionnellement",
            "oeuvre",
            "office",
            "officiel",
            "officielle",
            "officiellement",
            "officieuse",
            "officieusement",
            "officieux",
            "on",
            "opinion",
            "originaire",
            "originairement",
            "original",
            "originalement",
            "originalité",
            "originaux",
            "origine",
            "ou",
            "outrage",
            "outrageuse",
            "outrageusement",
            "outrageux",
            "outre",
            "ouvrage",
            "ouvrier",
            "ouvrière",
            "par",
            "parce que",
            "par conséquent",
            "par dessous",
            "par dessus",
            "par exemple",
            "parfois",
            "parmi",
            "parole",
            "par suite",
            "particularité",
            "particulier",
            "particulière",
            "particulièrement",
            "partout",
            "pas",
            "pauvre",
            "pauvrement",
            "pauvreté",
            "pendant",
            "pensée",
            "pensif",
            "pensive",
            "pensivement",
            "père",
            "personnage",
            "personnalité",
            "personne",
            "personnel",
            "personnelle",
            "personnellement",
            "petit",
            "peu à peu",
            "peuple",
            "peuplement",
            "peut-être",
            "place",
            "placement",
            "plaisir",
            "plus",
            "plusieurs",
            "plus tard",
            "plus tôt",
            "plutôt",
            "point",
            "pointe",
            "populaire",
            "populairement",
            "popularité",
            "population",
            "populeuse",
            "populeux",
            "possibilité",
            "possible",
            "pour",
            "pour ainsi dire",
            "pourquoi",
            "pourtant",
            "praticable",
            "pratique",
            "pratiquement",
            "premier",
            "première",
            "premièrement",
            "près",
            "presque",
            "preuve",
            "primitif",
            "primitive",
            "primitivement",
            "principal",
            "principalement",
            "principaux",
            "principe",
            "prix",
            "probabilité",
            "probable",
            "probablement",
            "prochain",
            "prochaine",
            "prochainement",
            "producteur",
            "productif",
            "production",
            "productive",
            "productivement",
            "productivité",
            "produit",
            "profit",
            "profitable",
            "profiteur",
            "profiteuse",
            "progrès",
            "progressif",
            "progression",
            "progressive",
            "progressivement",
            "projecteur",
            "projection",
            "projet",
            "proportion",
            "proportionnalité",
            "proportionnel",
            "proportionnelle",
            "proportionnellement",
            "proposition",
            "puis",
            "puisque",
            "puissance",
            "qualité",
            "quand",
            "quant",
            "quantité",
            "que",
            "quel",
            "quelconque",
            "quelle",
            "quelque",
            "quelque chose",
            "quelquefois",
            "quelque part",
            "quelque temps",
            "question",
            "questionnaire",
            "qui",
            "quiconque",
            "quoi",
            "quoique",
            "raison",
            "raisonnable",
            "raisonnablement",
            "raisonnement",
            "rapport",
            "rapporteur",
            "rare",
            "rarement",
            "rareté",
            "réalisable",
            "réalisateur",
            "réalisation",
            "réalité",
            "réel",
            "réelle",
            "réellement",
            "réflexion",
            "regard",
            "regret",
            "regrettable",
            "relatif",
            "relation",
            "relative",
            "relativement",
            "relativité",
            "remarquable",
            "remarquablement",
            "remarque",
            "remerciement",
            "renseignement",
            "rêve",
            "rêveur",
            "rêveuse",
            "rien",
            "rôle",
            "route",
            "rythme",
            "rythmique",
            "sans",
            "sans cesse",
            "sans doute",
            "se",
            "séculaire",
            "séculairement",
            "seigneur",
            "semblable",
            "semblablement",
            "sentiment",
            "sentimental",
            "sentimentalité",
            "sentimentaux",
            "ses",
            "seul",
            "seulement",
            "si",
            "siècle",
            "simple",
            "simplement",
            "simplicité",
            "simplification",
            "soeur",
            "soin",
            "solitaire",
            "solitairement",
            "solitude",
            "sommaire",
            "sommairement",
            "somme",
            "son",
            "sont",
            "sorte",
            "soudain",
            "soudaine",
            "soudainement",
            "soudaineté",
            "souffrance",
            "souffrant",
            "sous",
            "souvent",
            "subjectif",
            "subjective",
            "subjectivement",
            "subjectivité",
            "sujet",
            "sujétion",
            "supérieur",
            "supérieurement",
            "supériorité",
            "sur",
            "surtout",
            "systématique",
            "systématiquement",
            "système",
            "tandis que",
            "te",
            "tel",
            "telle",
            "tellement",
            "temporaire",
            "temporairement",
            "temporel",
            "temporelle",
            "temps",
            "tenir",
            "terre",
            "tes",
            "tête",
            "théorie",
            "théorique",
            "théoriquement",
            "titre",
            "toujours",
            "tour à tour",
            "tous ' w",
            "tout â",
            "tout à coup",
            "tout à fait",
            "toute",
            "toutefois",
            "tragique",
            "tragiquement",
            "trajet",
            "tranquille",
            "tranquillement",
            "tranquillité",
            "travail",
            "travailleur",
            "travailleuse",
            "travaux",
            "travers",
            "très",
            "très bien",
            "trop",
            "type",
            "typique",
            "typiquement",
            "un",
            "une",
            "unique",
            "uniquement",
            "unité",
            "univers",
            "universalité",
            "universel",
            "universelle",
            "universellement",
            "universitaire",
            "université",
            "usage",
            "utile",
            "utilement",
            "utilisable",
            "utilisateur",
            "utilisation",
            "utilitaire",
            "utilité",
            "valeur",
            "venir",
            "véritable",
            "véritablement",
            "vérité",
            "vieux",
            "vif",
            "vis-à-vis",
            "vive",
            "vivement",
            "voici",
            "voilà",
            "volontaire",
            "volontairement",
            "volonté",
            "volontiers",
            "vos",
            "votre",
            "vôtre",
            "vous",
            "voyage",
            "voyageur",
            "voyageuse",
            "vraiment",
        };

        // règles d'abréviation
        // Abréviations de groupes de lettres
        // Notes :
        // c. = consonne
        // v. = voyelle
        // d. = début
        // t. = terminaison (fin de mot)
        // Si rien, abréviation employable n'importe quand

        public static readonly string VOYELLES = "aeiouyàáâãäåæèéêëęìíîïðòóôõöøùúûüýÿœ";
        public static readonly string CONSONNES = "bcdfghjklmnpqrstvwxzçčłñṇšṣṭžẓ";
        // Note : a l'exception de in, les signes abrégeant des groupes de lettres ne sont utilisables que sur des lettres appartenant
        // a une meme syllabe
        // possiblement besoin de découper le mot en syllabes pour certaines abbréviation
        public static readonly string[] TERMINAISON = { "e", "s", "es", };

        public static string RegleAppliquerSur(string mot)
        {
            string cleanedup = mot.ToLower().Trim();
            if (cleanedup.Length < 2)
                return "aucune"; // au cas ou, pas d'abbréviation des lettres isolées je crois
            foreach (var abbr in MotsAbreger)
            {
                string temp = cleanedup;

                int ending = TERMINAISON.ToList().FindIndex(
                    t => temp.EndsWith(t) 
                    && !abbr.EndsWith(t)
                    && abbr.Length == (temp.Length - t.Length)
                );
                if (ending >= 0)
                    temp = temp.Substring(0, mot.Length - TERMINAISON[ending].Length);
                if (abbr == temp)
                {
                    return temp;
                }
            }

            List<Tuple<int, string>> syllabes = DecoupageSyllabes(cleanedup);
            foreach (var ambiguite in AmbiguiterAbbreviation)
            {
                if (ambiguite.Value(cleanedup))
                {
                    //Console.WriteLine($"{cleanedup} abrégable par la règle de '{ambiguite.Key}'");
                    return "Non-abrégeable sur "
                        + ambiguite.Key
                        + " ("
                        + string.Join("|", syllabes.Select(t => t.Item2).ToArray())
                        + ")";
                }
            }
            foreach (var testAbbr in AbbreviationGroupeLettres)
            {
                if (testAbbr.Value(cleanedup))
                {
                    //Console.WriteLine($"{cleanedup} abrégable par la règle de '{testAbbr.Key}'");
                    return testAbbr.Key
                        + " ("
                        + string.Join("|", syllabes.Select(t => t.Item2).ToArray())
                        + ")";
                }
            }
            return "aucune (" + string.Join("|", syllabes.Select(t => t.Item2).ToArray()) + ")";
        }

        public static bool EstAbregeable(string mot)
        {
            string cleanedup = mot.ToLower().Trim();
            if (cleanedup.Length < 2)
                return false; // au cas ou, pas d'abbréviation des lettres isolées je crois
            foreach (var abbr in MotsAbreger)
            {
                string temp = cleanedup;
                int ending = TERMINAISON.ToList().FindIndex(
                    t => temp.EndsWith(t)
                    && !abbr.EndsWith(t)
                    && abbr.Length == (temp.Length - t.Length)
                );
                if (ending >= 0)
                    temp = temp.Substring(0, mot.Length - TERMINAISON[ending].Length);
                if (abbr == temp)
                {
                    //Console.WriteLine("Mot abrégeable trouvé depuis la liste : " + abbr);
                    return true;
                }
            }
            foreach (var ambiguite in AmbiguiterAbbreviation)
            {
                if (ambiguite.Value(cleanedup))
                {
                    //Console.WriteLine($"{cleanedup} abrégable par la règle de '{ambiguite.Key}'");
                    return false;
                }
            }
            foreach (var testAbbr in AbbreviationGroupeLettres)
            {
                if (testAbbr.Value(cleanedup))
                {
                    //Console.WriteLine($"{cleanedup} abrégable par la règle de '{testAbbr.Key}'");
                    return true;
                }
            }
            return false;
        }

        private static readonly HashSet<string> CONSONNESINCASSABLE = new HashSet<string>
        {
            "ch",
            "ph",
            "th",
            "gh",
            "gn",
            "br",
            "cr",
            "dr",
            "fr",
            "gr",
            "kr",
            "pr",
            "tr",
            "vr",
            "bl",
            "cl",
            "fl",
            "gl",
            "pl"
        };

        private static readonly HashSet<string> PREFIXES = new HashSet<string>
        {
            "acantho", //épine	acanthacées, acanthe
            "lomb", //région lombaire	lombalgie
            "acou", //entendre	acoustique, acouphène
            "lum", //lumière, partie creuse d’un tube	luminaire, luminance, luminisme, lumitype
            "acro",
            "acrie", //extrémité	acrobate, acrostiche
            "macro", //grand	macrocosme
            "actino", //rayon	actinique, actinomètre
            "mal",
            "malé",
            "mau", //mauvais	malodorant, maléfique
            "ad", //vers, ajouté à	administré
            "mé",
            "més", //mauvais	médisance, mésalliance
            "adén", //glande, ganglion lymphatique	adénome, adénoïde
            "médull", //moelle	médullaire
            "aéro", //air	aéronaute, aéronef, aérophagie, aérostat
            "méga",
            "mégalo", //grand, gros	mégalithe, mégalomane
            "agro", //champ	agronome
            "melo", //chant	mélodique, mélodrame
            "all",
            "allo", //étranger	allopathie, allophone
            "més",
            "méso", //milieu	mésopotamien
            "ambi", //deux, autour, doublement	ambidextre, ambivalent
            "meta", //après, changement	métamorphose, métaphysique
            "amphi", //autour, doublement	amphithéâtre, amphibie
            "météor",
            "météoro", //élevé dans les airs	météore, météorologie
            "an", //sans	analphabète, anarchie
            "métr",
            "métro", //mesure	métrique, métronome
            "ana", //de bas en haut, à l'inverse	anagramme, anachronisme, anastrophe
            "mi-", //milieu	midi, mi"figue, "mi"raisin
            "andro", //homme (mâle)	androgyne
            "micro", //petit	microbe, microbiologie
            "anémo", //vent	anémomètre
            "mis",
            "miso", //haine	misanthrope, misogyne
            "angio", //vaisseau	angioplastie
            "mném",
            "mnémo", //mémoire	mnémotechnique
            "anté", //avant, précédent	antérieur, antédiluvien
            "mono", //seul	monogramme, monolithe
            "anth",
            "antho", //fleur, meilleur	anthémis, anthologie
            "morpho", //forme	morphologie
            "anthrac", //charbon	anthracite
            "multi", //nombreux	multicolore, multiforme, multiple
            "anthropo", //homme (espèce)	anthropologie
            "myco", //champignon	mycologie
            "anti", //contre	antipathie, antireligieux
            "myél", //moelle	myélopathie
            "apo", //éloignement	apogée
            "myo", //muscle	myocarde
            "apo", //hors de, à partir de, loin de	apostasie, apostrophe, apothéose
            "myri",
            "myria", //dix mille	myriade
            "arch", //qui commande, au"dessus	archevêque
            "mythe", //légende	mythologie
            "archéo", //ancien	archéologie
            "nas", //nez	nasalisation, nasique
            "archi", //supériorité, au plus haut degré	archiprêtre, archimillionnaire
            "natr", //sodium	natrémie
            "arithm",
            "arithmo", //nombre	arithmétique
            "nécro", //mort	nécrologie, nécropole
            "artério", //artère	artériosclérose
            "néo", //nouveau	néologisme, néophyte
            "arthr",
            "arthro", //articulation	arthrite, arthropodes
            "néphr",
            "néphro", //rein	néphrite
            "astér",
            "astéro",
            "astr",
            "astro", //astre, étoile	astérisque, astronaute
            "neuro",
            "névr", //nerf	neurologie, névrose
            "audi", //audition	audimat
            "non-", //négation	nonchalant
            "auto", //de soi"même	autobiographie, autodidacte, automobile
            "noso", //maladie	nosologie
            "bactéri",
            "bactério", //bâton	bactéricide, bactériologie
            "nuclé", //noyau	nucléaire
            "bar", // baro	pression	baromètre
            "ob",
            "oc",
            "of",
            "op", //devant, en opposition	obnubiler
            "béné",
            "bien", //bien	bienfaiteur, bénéfique
            "octa",
            "octo", //huit	octaèdre, octogone
            "bi",
            "bis",
            "bes", //deux fois	bipède
            "ocul", //oeil	occulter
            "biblio", //livre	bibliographie, bibliothèque
            "odont",
            "odonto", //dent	odontologie
            "bio", //vivant	biographie, biologie
            "olfact", //odorat	olfactif
            "blasto", //germe	blastoderme
            "olig",
            "oligo", //peu nombreux	oligarchie
            "bléphar",
            "blépharo", //paupière	blépharite
            "omni", //tout	omniscient, omnivore
            "brachy", //court	brachycéphale
            "onco", //tumeur	oncologie
            "brady", //lent	bradycardie, bradypsychie
            "oniro", //songé	oniromancie, onirique
            "brom",
            "bromo", //puanteur	brome, bromure
            "ophtalm",
            "ophtalmo", //oeil	ophtalmologie
            "bronch",
            "broncho", //gorge, bronche	bronche, bronchique
            "orchi", //testicule	orchidée
            "bryo", //mousse	bryophile
            "ornitho", //oiseau	ornithologiste
            "bucc", //bouche	buccal
            "oro", //montagne	orographie
            "butyr",
            "butyro", //beurre	butyrique
            "ortho", //droit	orthographe, orthopédie
            "caco",
            "cach", //mauvais	cacographie, cacophonie
            "osm", //odeur	osmium
            "calc", //calcium	calcification
            "osté",
            "ostéo", //os	ostéite, ostéomyélite
            "calli", //beau	calligraphie, callipyge
            "ot",
            "oto", //oreille	oto"rhino-laryngologie
            "cardi",
            "cardio", //coeur	cardiaque, cardiogramme, cardiographie
            "outre", //au"delà de	outrepasser
            "caryo", //noyau cellulaire	caryopse
            "ovari", //ovaire	ovarien, ovarite
            "cata", //de haut en bas, complètement	cataracte, catastrophe
            "oxy", //aigu, acide	oxyton, oxygène
            "cata", //en bas	catacombes
            "pachy", //épais	pachyderme
            "cén",
            "céno", //commun	cenobite, cénesthésie
            "paléo", //ancien	paléographie, paléolithique
            "céno", //vide	cénotaphe
            "pan",
            "pant",
            "pan",
            "panto", //tout	panthéisme, pantographe
            "céphal",
            "céphalo", //tête	céphalalgie, céphalopodes
            "par",
            "per", //à travers, achèvement	parcourir
            "cérébell", //cervelet	cérébelleux
            "para", //contre, auprès	parasite
            "cervic", //cou, col	cervical
            "path",
            "patho", //maladie, souffrance	pathogène, pathologie
            "chalco", //cuivre	chalcographie
            "péd", //enfant	pédagogie, pédiatrie
            "cheir",
            "chir", //main	chiromancie, chiropratique
            "péni", //pauvreté, diminution	pénitence, pénitencier
            "chimi", //substance chimique	chimiothérapie
            "penta", //cinq	pentagone
            "chloro", //vert	chlorate, chlorhydrique, chlorophyle
            "per", //à travers	percolateur, perforer
            "chol",
            "cholé", //bile	cholagogue, cholémie
            "peri", //autour	périoste, périphrase, périphérique
            "chromat",
            "chrom",
            "chromat",
            "chromo", //couleur	chromatique, chromosome
            "phago", //manger	phagocyte
            "chron",
            "chrono", //temps	chronique, chronographie, chronologie, chronomètre
            "pharmac",
            "pharmaco", //médicament	pharmaceutique, pharmacopée
            "chrys",
            "chryso", //or	chrysostome, chrysolithe
            "pharyng",
            "pharyngo", //gosier	pharyngite
            "cinémat",
            "cinémato",
            "ciné",
            "cinét",
            "ciné",
            "cinéto", //mouvement	cinématographe, cinétique
            "phén",
            "phéno", //apparaître	phénomène
            "circum",
            "circon", //autour	circonvenir, circumpolaire, cironférence
            "phil",
            "philo", //qui aime philanthrope, philatélie, philosophie
            "cis", //en deçà de	cisalpin
            "phléb", //veine	phlébite
            "co",
            "com",
            "con",
            "cor", //avec	cohabiter
            "phon",
            "phono", //voix	phonographe
            "col", //côlon (gros intestin)	colique
            "photo", //lumière	photographe
            "colp", //vagin	colpocèle
            "phréno", //diaphragme	phrénique
            "conch",
            "concho", //coquille	conchylien, conchyliologie
            "phyllo", //feuille	phylloxéra
            "contra",
            "contre", //contre, en face de	contresens, contradiction
            "phys",
            "physio", //nature	physiocrate, physique
            "cosm",
            "cosmo", //monde	cosmique, cosmogonie, cosmopolite
            "phyt",
            "phyto", //plante	phytophage
            "cox", //hanche	coxalgie
            "plast", //façonné	plasticité, plastique
            "crâni", //crâne	crâniopharyngiome
            "pleur", //plèvre (membrane du thorax)	pleurodynie
            "cry", //froid	cryogénique
            "pleur",
            "pleuro", //côté	pleurite
            "crypt",
            "crypto", //caché	crypte, cryptogame
            "plouto", //richesse	ploutocratie
            "cyan",
            "cyano", //bleu	cyanure
            "pneum",
            "pneumat", //air, respiration	pneumatique
            "cycl",
            "cyclo", //cercle	cyclique, cyclone, cyclotourisme
            "pneumo", //poumon	pneumonie
            "cyst", //vessie, poche	cystite, cystique
            "pod",
            "podo", //pied	podomètre
            "cyto", //cellule	cytologie
            "polio", //substance grise	poliomyélite
            "dactyl",
            "dactylo", //doigt	dactylographie
            "poly", //nombreux	polyèdre, polygone
            "dé",
            //"d",
            "des", //cessation	désunion
            "post", //après	postdater, postscolaire
            "déca",
            "déci", //dix	décamètre, décimètre
            "pré", //devant	préétabli, préhistoire, préliminaire
            "dém",
            "démo", //peuple	démocrate, démographie
            "pro", //en avant	proposer, projeter, prolonger
            "derm",
            "dermo",
            "dermato", //peau	derme, dermique, dermatologie
            "proct", //anus	proctologie
            "deut", //second	deutéron
            "prosop", //visage	prosopopée
            "di", //deux fois	diptyque, disyllabe
            "prosta", //prostate	prostatique
            "dia", //à travers, séparé de	diagonal, diaphane, diorama
            "prot",
            "proto", //premier	prototype
            "didact", //enseigner	didactique
            "proté", //protéine, forme changeante	protéolyse
            "dis",
            "dif",
            "dis", //séparation	diverger
            "pseud",
            "pseudo", //faux	pseudonyme
            "disc", //disque intervertébral	hernie discale
            "​psych",
            "​psycho", //âme	psychologue
            "dodéca", //douze	dodécagone
            "ptéro", //aile	ptérodactyle
            "dolicho", //long	dolichocéphale
            "pulm", //poumon	pulmonaire
            "dors", //dos	dorsal
            "pyél", //bassinet du rein	py​élite
            "dory", //lance	doryphore
            "pyo", //pus, suppuration	pyogène
            "dynam",
            "dynamo", //force	dynamite, dynamomètre
            "pyr",
            "pyro", //feu	pyrotechnie
            "dys", //difficulté	dyspepsie, dyslexie
            "quadr",
            "quadri",
            "quadru", //quatre	quadrijumeaux, quadrupède
            "échin",
            "échino", //épine, hérisson	échinoderme
            "quasi", //presque	quasi"contrat, "quasi"délit
            "électr",
            "électro", //ambre jaune	électrochoc
            "quinqu", //cinq	quinquagénaire, quinquennal
            "embryo", //foetus	embryologie
            //"r",
            "re", //de nouveau	rouvrir, réargenter
            "en",
            "em", //dans	encéphale, endémie, enfermer
            "rachi", //colonne vertébrale	rachidien
            "endo", //en dedans	endoderme, endocarde, endocrine
            "radio", //rayon	radiographie, radiologie
            "entér",
            "entéro", //entrailles	entérite
            "rect", //rectum	rectoscopie
            "entomo", //insecte	entomologiste
            "rétro", //en retour	rétroactif, rétrograder
            "entre",
            "inter", //Entre, réciproquement	entreposer, entrecôte
            "rhino", //nez	rhinocéros
            "éo", //aurore	éocène
            "rhizo", //racine	rhizome, rhizopodes
            "epi", //sur, au"dessus	épiderme, épizootie
            "rhodo", //rose	rhododendron
            "erg", //travail	ergonomie
            "rub", //rouge	rubéole
            "érythr", //rouge	érythème, érythrine
            "sarco", //chair	sarcophage
            "eu", //agréable, bien, bon	euphorie, euphémisme, euphonie
            "saur", //lézard	sauriens
            "ex-", //à l’extérieur, hors, qui a cessé d'être	expatrié, ex"employé
            "scaph", //barque	scaphandrier
            "exo", //au dehors	exotisme, exonérer
            "schizo", //qui fend	schizophrénie
            "extra", //superlatif, hors de	extra"fin, "extraordinaire, extra"territorialité
            "extra-",
            "séma",
            "séméio",
            "sémio", //signe	sémantique, sémaphore,  sémiologie
            "galact",
            "galacto", //lait	galactose, galaxie
            "semi-",
            "demi-", //	semi"circulaire
            "gam",
            "gamo", //mariage	gamète
            "sidér",
            "sidéro", //fer	sidérurgique
            "gastro", //ventre	gastropodes, gastronome
            "simili", //semblable	similigravure, simili marbre
            "gé",
            "géo", //terre	géographie, géologie
            "solén",
            "soléno", //tuyau	solénoïde
            "genu", //genou	génuflexion
            "somat",
            "somato", //corps	somatique
            "géront",
            "géronto", //vieillard	gérontocratie
            "sou",
            "sous-",
            "suc",
            "suf",
            "sug",
            "sup", //sous, presque	soucoupe
            "gingiv", //gencive	gingivite
            "spélé",
            "spéléo", //caverne	spéléologie
            "gloss",
            "glosso", //langue	glossaire
            "sphéno", //coin	sphénoïde
            "gluc",
            "gluco", //doux	glucose, glycogène
            "sphér",
            "sphéro", //globe	sphérique, sphénoïde
            "glyc",
            "glyco",
            "glycér",
            "glycéro", //doux	glycérine
            "spin", //épine, moelle épinière	spinal
            "granul", //granulation	granuleux
            "splén", //rate	splénite
            "graph",
            "grapho", //écrire	graphologie, graphème
            "spondyl", //vertèbre	spondylite
            "gyn",
            "gynéco", //femme	gynécée, gynécologie
            "stat", //stable	statique, statistique
            "gyro", //cercle	gyroscope
            "stéa", //graisse	stéarine
            "hagi",
            "hagio", //sacré	hagiographie
            "stéré",
            "stéréo", //solide	stéréoscope
            "halo", //sel	halogène
            "stomat",
            "stomato", //bouche	stomatologie
            "hecto", //cent	hectomètre
            "styl ",
            "stylo", //colonne	stylite
            "héli",
            "hélio", //soleil	héliothérapie
            "sub", //sous	subalterne, subdélégué, subdiviser
            "hémat",
            "hémato",
            "hémo", //sang	hématose, hémorragie
            "super",
            "supra", //au"dessus	superstructure, supranational
            "hémi", //demi	hémicycle, hémisphère
            "sus", //au dessus, plus	sus"mentionné
            "hépat",
            "hépato", //foie	hépatique, hépatite
            "sy",
            "syn",
            "sym", //avec	sympathie, synonyme
            "hept",
            "hepta", //sept	heptasyllabe
            "tachy", //rapide	tachymètre
            "hétéro", //autre	hétérogène
            "tauto", //le même	tautologie
            "hexa", //six	hexagone
            "taxi", //taxe	taximètre
            "hiér",
            "hiéro", //sacré	hiéroglyphe
            "techn",
            "techno", //art	technique, technologie
            "hipp",
            "hippo", //cheval	hippodrome
            "télé", //loin	télépathie, téléphone
            "hist",
            "histo", //tissu	histologie
            "térat", //monstre	tératologie
            "homéo",
            "hom",
            "homo", // semblable   homéopathie, homologue
            "tétra", //quatre	tétragone
            "hor",
            "horo", //heure	horoscope, horodateur
            "thalasso", //mer	thalassothérapie
            "hydr",
            "hydro", //eau, (fluide)	hydraulique, hydre, hydrologie, hydrothérapie
            "théo", //dieu	théocratie, théologie
            "hygro", //humide	hygromètre, hygroscope
            "thérapeut", //qui soigne	thérapeutique
            "hyper", //plus, au dessus	hypermétrope, hypertension, hypertrophie
            "therm",
            "thermo", //chaleur	thermomètre
            "hypn",
            "hypno", //sommeil	hypnose, hypnotisme
            "thorac", //thorax	thoracique
            "hypo", //moins, en dessous	hypophyse, hypodermique
            "thromb", //coagulation, caillot	thrombose
            "hystér",
            "hystéro", //utérus	hystérographie
            "top",
            "topo", //lieu	topographie, toponymie
            "iatr", // "iâtre", //	médecin	pédiatre pediatrie
            "trans", //au"delà de, à travers	transformer, transhumant
            "icon",
            "icono", //image	icône, iconoclaste
            "trauma",
            "traumat", //blessure, choc violent	traumatisé
            "idé",
            "idéo", //idée	idéogramme, idéologie
            "tré", //au"delà	trépasser
            "idi",
            "idio", //particulier	idiome, idiotisme
            "tri", //trois	tripartite, trisaieul, tricolore
            "in",
            "im",
            "il",
            "ir", //entrer, privé de, négation	infiltrer, insinuer, illettré, impropre, inexact
            "trich", //poil	trichogramme
            "inter", //entre	interallié, interligne
            "typo", //caractère	typographie, typologie
            "intra", //au"dedans	intramusculaire
            "ultra", //au"delà de	ultrason, ultraviolet
            "isch", //suppression, arrêt	ischémique
            "uni", //un	uniforme
            "iso", //égal	isomorphe, isotherme
            "urano", //ciel	uranographie
            "juxta", //auprès de	juxtalinéaire, juxtaposer
            "uré", //urine	urémie
            "kali", //potassium	kaliémie
            "urétr", //urètre	urétral
            "kilo", //mille	kilogramme
            "vas", //vaisseau	vasomoteur
            "kinés",
            "kinét", //mouvement	kinestésie
            "vascul", //vaisseau sanguin	vasculaire
            "lapar", //paroi abdominale	laparoscopie
            "vésic", //vessie	vésicule
            "laryng",
            "laryngo", //gorge	laryngologie
            "vi",
            "vice-", //suppléance	vice"président, "vice"amiral
            "leuc",
            "leuco", //blanc	leucocyte, leucémie
            "viscér", //viscère	viscéral
            "lipo", //lipide	liposuccion
            "xanth", //jaune	xanthine
            "litho", //pierre	lithographique
            "xén",
            "xéno", //étranger	xénophobe
            "loco", //mettre en mouvement	locomotion
            "xér",
            "xéro", //sec	xérophagie
            "log",
            "logo", //discours, science	logomachie
            "xylo", //bois	xylophone
            "zoo", //animal	zoologie
        }
            .OrderBy(p => p)
            .OrderBy(p => p.Length)
            .Reverse()
            .ToHashSet();

        /// <summary>
        /// Test si un préfixe est présent dans un mot avant un indice donné
        /// </summary>
        /// <param name="mot"></param>
        /// <param name="indice"></param>
        /// <returns></returns>
        private static string GetMatchingPrefix(string mot, int indice)
        {
            string substring = mot.Substring(0, indice);
            foreach (var prefix in PREFIXES.Where(p => p.Length <= indice))
            {
                if (substring.Substring(indice - prefix.Length) == prefix)
                {
                    return prefix;
                }
            }
            return null;
        }

        /// <summary>
        /// Découpage d'un mot en syllabes (requis pour tous les groupes sauf exceptions mentionné)
        /// </summary>
        /// <param name="mot"></param>
        /// <returns></returns>
        private static List<Tuple<int, string>> DecoupageSyllabes(string mot)
        {
            /*
            La règle générale est de séparer les syllabes entre une voyelle et une consonne.
            Exemples : Cou–pant / sa–pin / be–nêt
            
            Lorsque deux consonnes se suivent, la césure s’effectue entre les deux, ce qui est toujours le cas dès lors qu’elles sont doublées.
            Exemples : Mar–teau / pel–le / fem–me
            
            Lorsque la première consonne est suivie de la lettre « r » ou de la lettre « l« , ces deux consonnes ne peuvent être séparées dans le cas d’un mot monosyllabe.
            Exemples : Clan / prix
            
            Lorsque trois consonnes se suivent la coupure doit s’effectuer après la deuxième sauf si on a deux consonnes identiques.
            Exemples : Domp–ter / Ap–prendre
            
            Si les lettres « l » et « r » sont accolées à la deuxième consonne, la coupure doit se faire après la première consonne.
            Exemples : Pren–dre / câ–ble
            
            On ne sépare jamais les groupes de consonnes « ch« , « ph« , « th« , « gn » .
            Exemples : Pê–cher / cam–phre / a–po–thé–ose / cam–pa–gne
            
            On ne sépare pas deux voyelles ou les mots contenant un « x » .
            Exemples : A–vion / rayon (correspond à rai-ion) / exem–ple (correspond à eg-zem-ple)
            
            On ne découpe pas un mot après une apostrophe.
            Exemples : L’arbre / l’élève
            */
            List<Tuple<int, string>> syllabes = new List<Tuple<int, string>>();
            int syllabeStart = 0;
            int voyellePrecedente = -1;
            bool separator = false;
            for (int i = 0; i < mot.Length; i++)
            {
                char lettre = mot[i];
                // Si la précédente chaine d'évaluation était une chaine de séparateurs
                if (separator && (VOYELLES.Contains(lettre) || CONSONNES.Contains(lettre)))
                {
                    separator = false;
                    // Je vais garder les séparateurs dans la syllabe (sinon, ajouter -1 a la longueur)
                    syllabes.Add(
                        Tuple.Create(syllabeStart, mot.Substring(syllabeStart, i - syllabeStart))
                    );
                    // on reset la syllabe
                    syllabeStart = i;
                    voyellePrecedente = -1;
                }
                if (VOYELLES.Contains(lettre))
                {
                    if (voyellePrecedente >= 0 && (i - 1) != voyellePrecedente)
                    {
                        // Doute sur la voyelle "y" (exemple de rayon = rai-ion a l'oral)
                        // nouveau noyau de sillabe trouvé
                        // on regarde le groupe de consonne entre les deux noyaux
                        string between = mot.Substring(
                            voyellePrecedente + 1,
                            i - voyellePrecedente - 1
                        );
                        if (between.Length == 1)
                        {
                            if (between.ToLower() == "x")
                            {
                                // cas particulier, pas de coupure sylabique sur X
                                // (Pas sur de la découpe ici, je sais pas si on suit les regles de l'oral ou de l'écrit

                                // on continue
                            }
                            else
                            {
                                // on casse avant la consonne
                                syllabes.Add(
                                    Tuple.Create(
                                        syllabeStart,
                                        mot.Substring(
                                            syllabeStart,
                                            voyellePrecedente + 1 - syllabeStart
                                        )
                                    )
                                );
                                syllabeStart = voyellePrecedente + 1;
                            }
                        }
                        else if (between.Length == 2)
                        {
                            // On ne sépare jamais les groupes de consonnes « ch« , « ph« , « th« , « gn »
                            // et on ne sépare pas les groupes de consonnes si elles sont différentes et suivi d'un r ou d'un l
                            if (CONSONNESINCASSABLE.Contains(between.ToLower()))
                            {
                                // on casse avant les consonnes
                                syllabes.Add(
                                    Tuple.Create(
                                        syllabeStart,
                                        mot.Substring(
                                            syllabeStart,
                                            voyellePrecedente + 1 - syllabeStart
                                        )
                                    )
                                );
                                syllabeStart = voyellePrecedente + 1;
                            }
                            else
                            {
                                // Lorsque deux consonnes se suivent, la césure s’effectue entre les deux, ce qui est toujours le cas dès lors qu’elles sont doublées.
                                syllabes.Add(
                                    Tuple.Create(
                                        syllabeStart,
                                        mot.Substring(
                                            syllabeStart,
                                            voyellePrecedente + 2 - syllabeStart
                                        )
                                    )
                                );
                                syllabeStart = voyellePrecedente + 2;
                            }
                        }
                        else if (between.Length >= 3)
                        {
                            bool a2consonneIdentique = false;
                            for (int j = 0; j < between.Length - 1; j++)
                            {
                                if (between[j] == between[j + 1])
                                {
                                    a2consonneIdentique = true;
                                    // 2 consonnes identiques trouvé, on sépare après la première
                                    syllabes.Add(
                                        Tuple.Create(
                                            syllabeStart,
                                            mot.Substring(
                                                syllabeStart,
                                                voyellePrecedente + j + 2 - syllabeStart
                                            )
                                        )
                                    );
                                    syllabeStart = voyellePrecedente + j + 1;
                                    break;
                                }
                            }
                            if (!a2consonneIdentique)
                            {
                                // On regarde s'il y a un groupe de consonnes incassable (duo sonore)
                                int indexOfIncass = -1;
                                foreach (var test in CONSONNESINCASSABLE)
                                {
                                    if (between.ToLower().Contains(test))
                                    {
                                        indexOfIncass = between.ToLower().IndexOf(test);
                                        break;
                                    }
                                }
                                if (indexOfIncass > -1)
                                {
                                    // on sépare avant la double consonne
                                    syllabes.Add(
                                        Tuple.Create(
                                            syllabeStart,
                                            mot.Substring(
                                                syllabeStart,
                                                voyellePrecedente + indexOfIncass + 1 - syllabeStart
                                            )
                                        )
                                    );
                                    syllabeStart = voyellePrecedente + indexOfIncass + 1;
                                }
                                else if (between.ToLower()[2] == 'r' || between.ToLower()[2] == 'l')
                                {
                                    // si r ou l accolé a la deuxieme consonne, on sépare apres la première
                                    syllabes.Add(
                                        Tuple.Create(
                                            syllabeStart,
                                            mot.Substring(
                                                syllabeStart,
                                                voyellePrecedente + 2 - syllabeStart
                                            )
                                        )
                                    );
                                    syllabeStart = voyellePrecedente + 2;
                                }
                                else
                                {
                                    // Lorsque trois consonnes se suivent la coupure doit s’effectuer après la deuxième
                                    syllabes.Add(
                                        Tuple.Create(
                                            syllabeStart,
                                            mot.Substring(
                                                syllabeStart,
                                                voyellePrecedente + 3 - syllabeStart
                                            )
                                        )
                                    );
                                    syllabeStart = voyellePrecedente + 3;
                                }
                            }
                        }
                        else
                        {
                            // normalement pas possible
                        }
                    }
                    voyellePrecedente = i;
                }
                else if (!CONSONNES.Contains(lettre))
                {
                    separator = true;
                }
            }
            if (syllabeStart < mot.Length - 1)
                syllabes.Add(Tuple.Create(syllabeStart, mot.Substring(syllabeStart)));

            return syllabes;
        }

        private static bool DevantVoyelle(string mot, string abbr)
        {
            mot = mot.ToLower().Trim();
            int index = mot.IndexOf(abbr);
            return index > -1
                && index < mot.Length - abbr.Length
                && VOYELLES.Contains(mot[index + abbr.Length]);
        }

        private static bool DevantConsonne(string mot, string abbr)
        {
            mot = mot.ToLower().Trim();
            int index = mot.IndexOf(abbr);
            return index > -1
                && index < mot.Length - abbr.Length
                && CONSONNES.Contains(mot[index + abbr.Length]);
        }

        private static bool Terminaison(string mot, string abbr)
        {
            mot = mot.ToLower().Trim();
            // suppression du pluriel si le terme rechercher n'est pas possiblement une forme de pluriel
            if (mot.ToLower().EndsWith("s") && !abbr.EndsWith("s"))
                mot = mot.Substring(0, mot.Length - 1);
            int index = mot.ToLower().IndexOf(abbr);
            return mot.Length > abbr.Length && index == mot.Length - abbr.Length;
        }

        public static Dictionary<string, Predicate<string>> AmbiguiterAbbreviation = new Dictionary<
            string,
            Predicate<string>
        >
        {
            { "z", (mot) => Terminaison(mot, "z") && !Terminaison(mot, "ez") }, // Présence d'un z isolé en fin du mot d'origine
        };

        // Regles extraites du tableau du manuel
        public static Dictionary<string, Predicate<string>> AbbreviationGroupeLettres =
            new Dictionary<string, Predicate<string>>
            {
                { "logiquement", (mot) => Terminaison(mot, "logiquement") },
                { "ablement", (mot) => Terminaison(mot, "ablement") },
                { "ellement", (mot) => Terminaison(mot, "ellement") },
                { "logique", (mot) => Terminaison(mot, "logique") },
                { "quement", (mot) => Terminaison(mot, "quement") },
                { "bilité", (mot) => Terminaison(mot, "bilité") },
                { "tement", (mot) => Terminaison(mot, "tement") }, // TODO a faire : (le t ne se contracte pas avec celui d'une syllabe précédente)
                { "vement", (mot) => Terminaison(mot, "vement") },
                { "ation", (mot) => Terminaison(mot, "ation") },
                { "ition", (mot) => Terminaison(mot, "ition") },
                { "logie", (mot) => Terminaison(mot, "logie") },
                {
                    "trans",
                    (mot) =>
                    {
                        // Le dis s'abrège si  (i = position du com dans le mot)
                        // - début de mot ou de ligne apres coupure, devant consonne (je vais garder que début de mot, donc si i == 0)
                        // - apres prefixe (i > 1) devant consonne (i < taille - abbr et consonne en i+abbr)
                        int index = mot.ToLower().IndexOf("trans");
                        return (
                                index == 0
                                && mot.ToLower().Length > 3
                                && CONSONNES.Contains(mot.ToLower()[index + 3])
                            )
                            || (
                                index > 1
                                && index < mot.ToLower().Length - 3
                                && CONSONNES.Contains(mot.ToLower()[index + 3])
                            );
                    }
                },
                { "able", (mot) => Terminaison(mot, "able") },
                { "elle", (mot) => Terminaison(mot, "elle") },
                {
                    "ain",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower()).FindIndex((s) => s.Item2.Contains("ain"))
                        >= 0
                },
                {
                    "oin",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower()).FindIndex((s) => s.Item2.Contains("oin"))
                        >= 0
                },
                { "ait", (mot) => Terminaison(mot, "ait") },
                { "ant", (mot) => Terminaison(mot, "ant") },
                {
                    "com",
                    (mot) =>
                    {
                        // Le com s'abrège si (i = position du com dans le mot)
                        // - début de mot ou de ligne apres coupure (je vais garder que début de mot, donc si i == 0, et tester que les mots plus grand ?)
                        // - apres prefixe (i > 0) devant consonne (i < taille - 1 et consonne après i)
                        List<Tuple<int, string>> syllabes = DecoupageSyllabes(mot.ToLower());
                        int index = syllabes.FindIndex(s => s.Item2.Contains("com"));
                        int indexInSyllable =
                            index >= 0 ? syllabes[index].Item2.IndexOf("com") : -1;
                        return (index == 0 && indexInSyllable == 0)
                            || (
                                index > 1
                                && index < (mot.ToLower().Length - 3)
                                && CONSONNES.Contains(mot.ToLower()[index + 3])
                            );
                    }
                },
                {
                    "con",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower())
                            .FindIndex(
                                (s) =>
                                    DevantConsonne(s.Item2, "con")
                                    || s.Item1 + s.Item2.Length < mot.Length
                                        && DevantConsonne(
                                            mot.ToLower().Substring(s.Item1, s.Item2.Length + 1),
                                            "con"
                                        )
                            ) >= 0
                },
                {
                    "dis",
                    (mot) =>
                    {
                        // Le dis s'abrège si  (i = position du com dans le mot)
                        // - début de mot ou de ligne apres coupure, devant consonne (je vais garder que début de mot, donc si i == 0)
                        // - apres prefixe (i > 1) devant consonne (i < taille - abbr et consonne en i+abbr)
                        int index = mot.ToLower().IndexOf("dis");
                        return (
                                index == 0
                                && mot.ToLower().Length > 3
                                && CONSONNES.Contains(mot.ToLower()[index + 3])
                            )
                            || (
                                index > 1
                                && index < mot.ToLower().Length - 3
                                && CONSONNES.Contains(mot.ToLower()[index + 3])
                            );
                    }
                },
                { "ent", (mot) => Terminaison(mot, "ent") },
                { "ess", (mot) => mot.ToLower().StartsWith("ess") },
                {
                    "eur",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower()).FindIndex((s) => s.Item2.Contains("eur"))
                        >= 0
                },
                { "ien", (mot) => mot.ToLower().Contains("ien") }, // s'abrège dans tous les cas (d'apres marius : Le groupe "ien" est abrégé quand il est en fin de mot et précédé d'une consonne.)
                { "ieu", (mot) => mot.ToLower().Contains("ieu") }, // s'abrège dans tous les cas
                { "ion", (mot) =>
                        DecoupageSyllabes(mot.ToLower())
                            .FindIndex(
                                (s) =>
                                    DevantConsonne(s.Item2, "ion")
                                    || s.Item1 + s.Item2.Length < mot.Length
                                        && DevantConsonne(
                                            mot.ToLower().Substring(s.Item1, s.Item2.Length + 1),
                                            "ion"
                                        )
                            ) >= 0
                        || Terminaison(mot, "ion")
                }, // d'apres marius : Le groupe "ion" est abrégé quand il est en fin de mot ou suivi d'une consonne (ce n'est pas le cas de Giono).
                //{ "ion", (mot) => mot.ToLower().Contains("ion") }, // s'abrège dans tous les cas
                {
                    "our",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower())
                            .FindIndex(
                                (s) =>
                                    DevantConsonne(s.Item2, "our")
                                    || s.Item1 + s.Item2.Length < mot.Length
                                        && DevantConsonne(
                                            mot.ToLower().Substring(s.Item1, s.Item2.Length + 1),
                                            "our"
                                        )
                            ) >= 0
                        || Terminaison(mot, "our")
                }, // Note : devant c. et t. (pas sur de ma regle ici)
                {
                    "pro",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower())
                            .FindIndex(
                                (s) =>
                                    DevantConsonne(s.Item2, "pro")
                                    || s.Item1 + s.Item2.Length < mot.Length
                                        && DevantConsonne(
                                            mot.ToLower().Substring(s.Item1, s.Item2.Length + 1),
                                            "pro"
                                        )
                            ) >= 0
                },
                { "que", (mot) => Terminaison(mot, "que") },
                {
                    "ai",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower()).FindIndex((s) => s.Item2.Contains("ai"))
                        >= 0
                },
                {
                    "an",
                    (mot) =>
                    {
                        // Les signes pour an, eu, et or ne s'emploient pas isolément
                        // le signe an ne s'emploie pas à la fin des mots.
                        List<Tuple<int, string>> syllabes = DecoupageSyllabes(mot.ToLower());
                        int anIndex = syllabes.FindIndex((s) => s.Item2.Contains("an"));
                        int anIndexInSyllabe =
                            anIndex >= 0 ? syllabes[anIndex].Item2.IndexOf("an") : -1;
                        return anIndex > -1 // le mot contient une syllabe qui contien an
                            && !(
                                anIndex == syllabes.Count - 1
                                && anIndexInSyllabe == syllabes[anIndex].Item2.Length - 2
                            ); // mais pas a la fin de la derniere syllabe
                    }
                },
                {
                    "ar",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower()).FindIndex((s) => s.Item2.Contains("ar"))
                        >= 0
                },
                {
                    "au",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower()).FindIndex((s) => s.Item2.Contains("au"))
                        >= 0
                },
                {
                    "bl",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower())
                            .FindIndex(
                                (s) =>
                                    DevantVoyelle(s.Item2, "bl")
                                    || s.Item1 + s.Item2.Length < mot.Length
                                        && DevantVoyelle(
                                            mot.ToLower().Substring(s.Item1, s.Item2.Length + 1),
                                            "bl"
                                        )
                            ) >= 0
                },
                {
                    "br",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower())
                            .FindIndex(
                                (s) =>
                                    DevantVoyelle(s.Item2, "br")
                                    || s.Item1 + s.Item2.Length < mot.Length
                                        && DevantVoyelle(
                                            mot.ToLower().Substring(s.Item1, s.Item2.Length + 1),
                                            "br"
                                        )
                            ) >= 0
                },
                {
                    "ch",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower()).FindIndex((s) => s.Item2.Contains("ch"))
                        >= 0
                },
                {
                    "cl",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower())
                            .FindIndex(
                                (s) =>
                                    DevantVoyelle(s.Item2, "cl")
                                    || s.Item1 + s.Item2.Length < mot.Length
                                        && DevantVoyelle(
                                            mot.ToLower().Substring(s.Item1, s.Item2.Length + 1),
                                            "cl"
                                        )
                            ) >= 0
                },
                {
                    "cr",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower())
                            .FindIndex(
                                (s) =>
                                    DevantVoyelle(s.Item2, "cr")
                                    || s.Item1 + s.Item2.Length < mot.Length
                                        && DevantVoyelle(
                                            mot.ToLower().Substring(s.Item1, s.Item2.Length + 1),
                                            "cr"
                                        )
                            ) >= 0
                },
                {
                    "dr",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower())
                            .FindIndex(
                                (s) =>
                                    DevantVoyelle(s.Item2, "dr")
                                    || s.Item1 + s.Item2.Length < mot.Length
                                        && DevantVoyelle(
                                            mot.ToLower().Substring(s.Item1, s.Item2.Length + 1),
                                            "dr"
                                        )
                            ) >= 0
                },
                {
                    "em",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower())
                            .FindIndex(
                                (s) =>
                                    DevantConsonne(s.Item2, "em")
                                    || s.Item1 + s.Item2.Length < mot.Length
                                        && DevantConsonne(
                                            mot.ToLower().Substring(s.Item1, s.Item2.Length + 1),
                                            "em"
                                        )
                            ) >= 0
                },
                {
                    "en",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower()).FindIndex((s) => s.Item2.Contains("en"))
                        >= 0
                },
                {
                    "er",
                    (mot) =>
                        mot.ToLower().IndexOf("er") > 0
                        && DecoupageSyllabes(mot.ToLower()).FindIndex((s) => s.Item2.Contains("er"))
                            >= 0
                }, // Le signe pour er ne s'emploie pas au début des mots.
                {
                    "es",
                    (mot) =>
                    {
                        List<Tuple<int, string>> syllabes = DecoupageSyllabes(mot.ToLower());
                        int index = syllabes.FindIndex(s => s.Item2.Contains("es"));
                        int indexInSyllabe = index >= 0 ? syllabes[index].Item2.IndexOf("es") : -1;
                        int indexInMot = mot.ToLower().IndexOf("es");
                        // Regle pour ES
                        //  d. d'un mot ou d'une ligne après coupure
                        // - après préfixe
                        // - t.
                        return mot.ToLower().Length > 2
                            && (
                                (index == 0 && indexInSyllabe == 0)
                                || (
                                    indexInMot > 0
                                    && GetMatchingPrefix(mot.ToLower(), indexInMot) != null
                                )
                                || (
                                    index == syllabes.Count - 1
                                    && indexInSyllabe == syllabes[index].Item2.Length - 2
                                )
                            );
                    }
                },
                {
                    "eu",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower()).FindIndex((s) => s.Item2.Contains("eu"))
                        >= 0
                }, // Les signes pour an, eu, et or ne s'emploient pas isolément
                {
                    "ex",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower())
                            .FindIndex(
                                (s) =>
                                    DevantConsonne(s.Item2, "ex")
                                    || s.Item1 + s.Item2.Length < mot.Length
                                        && DevantConsonne(
                                            mot.ToLower().Substring(s.Item1, s.Item2.Length + 1),
                                            "ex"
                                        )
                            ) >= 0
                },
                { "ez", (mot) => Terminaison(mot, "ez") },
                {
                    "fl",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower())
                            .FindIndex(
                                (s) =>
                                    DevantVoyelle(s.Item2, "fl")
                                    || s.Item1 + s.Item2.Length < mot.Length
                                        && DevantVoyelle(
                                            mot.ToLower().Substring(s.Item1, s.Item2.Length + 1),
                                            "fl"
                                        )
                            ) >= 0
                },
                {
                    "fr",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower())
                            .FindIndex(
                                (s) =>
                                    DevantVoyelle(s.Item2, "fr")
                                    || s.Item1 + s.Item2.Length < mot.Length
                                        && DevantVoyelle(
                                            mot.ToLower().Substring(s.Item1, s.Item2.Length + 1),
                                            "fr"
                                        )
                            ) >= 0
                },
                {
                    "gl",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower())
                            .FindIndex(
                                (s) =>
                                    DevantVoyelle(s.Item2, "gl")
                                    || s.Item1 + s.Item2.Length < mot.Length
                                        && DevantVoyelle(
                                            mot.ToLower().Substring(s.Item1, s.Item2.Length + 1),
                                            "gl"
                                        )
                            ) >= 0
                },
                { "gn", (mot) => mot.ToLower().IndexOf("gn") > 0 && !Terminaison(mot, "gn") },
                {
                    "gr",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower())
                            .FindIndex(
                                (s) =>
                                    DevantVoyelle(s.Item2, "gr")
                                    || s.Item1 + s.Item2.Length < mot.Length
                                        && DevantVoyelle(
                                            mot.ToLower().Substring(s.Item1, s.Item2.Length + 1),
                                            "gr"
                                        )
                            ) >= 0
                },
                {
                    "im",
                    (mot) =>
                    {
                        string lowMot = mot.ToLower();
                        int index = lowMot.IndexOf("im");
                        return index == 0
                            && lowMot.Length > 2
                            && new List<char>() { 'm', 'b', 'p', }.Contains(lowMot[2]);
                    }
                },
                {
                    "in",
                    (mot) =>
                        mot.ToLower().StartsWith("in")
                        || DecoupageSyllabes(mot.ToLower()).FindIndex((s) => s.Item2.Contains("in"))
                            >= 0
                },
                {
                    "ll",
                    (mot) =>
                    {
                        int index = mot.ToLower().IndexOf("ll");
                        return index > 0
                            && mot.ToLower().Length >= 4
                            && index < mot.ToLower().Length - 2
                            && VOYELLES.Contains(mot.ToLower()[index - 1]) // voyelle avant
                            && VOYELLES.Contains(mot.ToLower()[index + 2]); // voyelle apres
                    }
                },
                {
                    "oi",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower()).FindIndex((s) => s.Item2.Contains("oi"))
                        >= 0
                },
                {
                    "om",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower())
                            .FindIndex(
                                (s) =>
                                    DevantConsonne(s.Item2, "om")
                                    || s.Item1 + s.Item2.Length < mot.Length
                                        && DevantConsonne(
                                            mot.ToLower().Substring(s.Item1, s.Item2.Length + 1),
                                            "om"
                                        )
                            ) >= 0
                        || Terminaison(mot, "om")
                }, // Note : devant c. et t. (pas sur de ma regle ici)
                {
                    "on",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower()).FindIndex((s) => s.Item2.Contains("on"))
                        >= 0
                },
                {
                    "or",
                    (mot) =>
                        mot.ToLower().Length > 2
                        && mot.ToLower().IndexOf("or") > 0
                        && DecoupageSyllabes(mot.ToLower()).FindIndex((s) => s.Item2.Contains("or"))
                            >= 0
                }, // Les signes pour an, eu, et or ne s'emploient pas isolément
                {
                    "ou",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower()).FindIndex((s) => s.Item2.Contains("ou"))
                        >= 0
                },
                {
                    "pl",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower())
                            .FindIndex(
                                (s) =>
                                    DevantVoyelle(s.Item2, "pl")
                                    || s.Item1 + s.Item2.Length < mot.Length
                                        && DevantVoyelle(
                                            mot.ToLower().Substring(s.Item1, s.Item2.Length + 1),
                                            "pl"
                                        )
                            ) >= 0
                },
                {
                    "pr",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower())
                            .FindIndex(
                                (s) =>
                                    DevantVoyelle(s.Item2, "pr")
                                    || s.Item1 + s.Item2.Length < mot.Length
                                        && DevantVoyelle(
                                            mot.ToLower().Substring(s.Item1, s.Item2.Length + 1),
                                            "pr"
                                        )
                            ) >= 0
                },
                {
                    "qu",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower()).FindIndex((s) => s.Item2.Contains("qu"))
                        >= 0
                },
                { "re", (mot) => mot.ToLower().IndexOf("re") == 0 && DevantConsonne(mot, "re") }, // Le signe pour re ne s'emploie qu'au début des mots, devant consonne (qu'il soit ou non une syllable entière)
                {
                    "ss",
                    (mot) =>
                    {
                        int index = mot.ToLower().IndexOf("ss");
                        return index > 0
                            && mot.ToLower().Length >= 4
                            && index < mot.ToLower().Length - 2
                            && VOYELLES.Contains(mot.ToLower()[index - 1]) // voyelle avant
                            && VOYELLES.Contains(mot.ToLower()[index + 2]); // voyelle apres
                    }
                },
                {
                    "tr",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower())
                            .FindIndex(
                                (s) =>
                                    DevantVoyelle(s.Item2, "tr")
                                    || s.Item1 + s.Item2.Length < mot.Length
                                        && DevantVoyelle(
                                            mot.ToLower().Substring(s.Item1, s.Item2.Length + 1),
                                            "tr"
                                        )
                            ) >= 0
                },
                {
                    "tt",
                    (mot) =>
                    {
                        int index = mot.ToLower().IndexOf("tt");
                        return index > 0
                            && mot.ToLower().Length >= 4
                            && index < mot.ToLower().Length - 2
                            && VOYELLES.Contains(mot.ToLower()[index - 1]) // voyelle avant
                            && VOYELLES.Contains(mot.ToLower()[index + 2]); // voyelle apres
                    }
                },
                {
                    "ui",
                    (mot) =>
                        DecoupageSyllabes(mot.ToLower())
                            .FindIndex(
                                (s) =>
                                    DevantConsonne(s.Item2, "ui")
                                    || s.Item1 + s.Item2.Length < mot.Length
                                        && DevantConsonne(
                                            mot.ToLower().Substring(s.Item1, s.Item2.Length + 1),
                                            "ui"
                                        )
                            ) >= 0
                },
            };

        // Notes : regles ecrites apres le tableau du manuel
        // 1. le signe pour ER ne s'emploie pas au début des mots.
        // 2. Le signe pour RE ne s'emploie qu'au début des mots, devant consonne qu'ils forme ou non une syllabe
        // 3. Le signe pour IN s'emploie en début de mot dans tous les cas
        // 4. Les signes pour AN, EU et OR ne s'emploient pas isolément
        // 4.1. le signe AN ne s'emploie pas à la fin des mots.
        // 5. Il est important de remarquer que les signes abrégeant des groupes de lettres peuvent être utilisés
        //    quelle qu'en soit la prononciation.
        // 6. Les signes abrégeant des groupes de lettres, à l'exception de IN, ne peuvent être utilisés
        //    que pour des lettres appartenant à une même syllabe
        // 6.1 Il y a exception pour les doubles consonnes LL , SS  et TT  qui ne pourraient s'employer autrement,
        //     et pour certains signes de finales spécifiés au tableau.
        // 7. Le t de la finale tement ne se contracte pas avec celui qui le précède dans certains mots.
        // 8. Les groupes AIN et OIN, dans une même syllabe, s'abrègent (voir manuel pour les signes si besoin)
        // 8.1 Par contre, le signe IN ne peut être utilisé si le groupe de lettres appartient à 2 syllabes.
        // 9. Pour éviter une interprétation souvent difficile de la division syllabique, le groupe IEUR s'abrège
        //    (voir manuel pour signe)
        // 10. Les groupes IEN, ION et IEU s'abrègent dans tous les cas
        // 11. Le groupe IENT s'abrège quelle que soit sa prononciation sauf forme verbale
        //     (exemple, "ils oublient" s'abregege différement)
        // 12. Le groupe ESS au début d'un mot s'écrit (voir symbole abrégé dans le manuel)
        // 13. Les signes pour COM, DIS, ES et TRANS qui s'emploient au début des mots ou d'une ligne après coupure,
        //     peuvent être précédés des préfixes RE et IN.
        // 14. N'importe quel nombre de signes inférieurs et de ponctuations peuvent être écrits
        //     successivement à la condition que la séquence comprenne au moins un signe supérieur.
        //     Cette règle s'applique également à un mot coupé en fin de ligne, ainsi qu'au rejet à la ligne
        //     suivante de la fin du mot
        // 15. Lorsque l'abréviation amène à employer deux fois de suite le même signe, celui-ci n'est utilisé que la seconde fois.
        // 16. regle sur le mot bleu, voir manuel
        // 17. Les signes (majuscule, mise en evidence) placés au début d'un mot et (modificateur mathematique) placé devant un nombre,
        //     gardent leur signification de "indicateur de majuscule", "indicateur général de mise en évidence"
        //     et "modificateur mathématique".
        // 18. Le signe dit "indicateur de valeur de base", placé au début d'un mot, indique que dans ce mot,
        //     tous les signes conservent leur valeur alphabétique.
        //     Il doit être employé chaque fois qu'il peut y avoir confusion sur la valeur alphabétique
        //     ou abréviative à attribuer à un ou plusieurs signes du mot

        //TODO : Les mots abregeable qui ne sont pas dans le "glaff abregeable" peuvent être considérés comme "potentiellement étrangés au français"

        

        #region Analyseur de texte

        private static readonly string MIN = "a-zàáâãäåæçèéêëìíîïðñòóôõöøùúûüý";
        private static readonly string MAJ = "A-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝ";
        private static readonly string NUM = "0-9";
        private static readonly string ALPHANUM = $"{MIN}{MAJ}{NUM}_"; // == a-zàáâãäåæçèéêëìíîïðñòóôõöøùúûüýA-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝ0-9_
        private static readonly string REG_CS = "\\[\\[\\*"; // Debut de code duxburry
        private static readonly string REG_CE = "\\*\\]\\]"; // Fin de code duxburry
        //private static readonly string WORD_WITH_NUMBERS_PATTERN =
        //    $"({REG_CS}i{REG_CE})?"
        //    + // Code duxburry de protection optionnel (récupérer dans le groupe 1) (0 ou 1)
        //    $"(?<!{REG_CS})"
        //    + // Negative lookbehind (ne pas match le texte commençant par un indicateur de code duxburry)
        //    $"("
        //    + $"\\b"
        //    + // Debut de mot (word boundary)
        //    $"[{NUM}{MIN}{MAJ}]*"
        //    + // Chiffre ou lettre (0 ou n)
        //    $"("
        //    + $"[{MIN}{MAJ}][{NUM}]"
        //    + // Lettre suivie d'un chiffre
        //    $"|"
        //    + $"[{NUM}][{MIN}{MAJ}]"
        //    + // Chiffre suivie d'une lettre
        //    $")+"
        //    + // Motif Chiffre+Lettre ou lettre+chiffre (1 ou n)
        //    $"[{NUM}{MIN}{MAJ}]*"
        //    + // suivi de chiffres ou de lettres (0 ou n)
        //    $"(?!{REG_CE})"
        //    + // Negative lookahead (ne pas match de texte suivi d'une fin de code duxburry)
        //    $"\\b"
        //    + // Fin de mot (word boundary)
        //    $")";


        ///// <summary>
        ///// Motif de recherche de mots contenant au moins 1 majuscule : <br/>
        ///// = début de ligne ou caractère non alphanumérique,<br/>
        ///// puis optionnellement suivi d'un code de protection,<br/>
        ///// puis optionnellement suivi d'un groupe contenant <br/>
        /////    0 ou N caractère majuscule ou minuscule,<br/>
        /////    suivi d'un groupe contenant au choix<br/>
        /////    - soit une ou plusieurs (minuscules ou majuscule ou tiret) suivie d'une majuscule<br/>
        /////    - soit une majuscule suivi d'une ou plusieurs(minuscules ou majuscule ou tiret) <br/>
        /////    suivi de 0 ou N caractère minuscule, majuscule, underscore ou tiret<br/>
        /////    et suivi d'un caractère (non alphanumérique y compris sans apostrophe pour éviter les prefix en Qu' et C') ou de la fin de ligne
        ///// </summary>
        //private static readonly string WORD_WITH_CAPITAL_PATTERN =
        //    $"(?<=[^{ALPHANUM}-]|^)" // group 1 linestart or non alphanum character
        //    + $"({REG_CS}i{REG_CE})?" // group 2 Optional protection code
        //    + $"(" // group 3 : searched word
        //    + $"[{MIN}{MAJ}_-]*" // optionnal prefix
        //    + $"(" // Group 4: one or more capital letters and at least one other letter
        //    + $"[{MIN}{MAJ}-]+[{MAJ}]"
        //    + $"|[{MAJ}][{MIN}{MAJ}-]+"
        //    + $")" // end group 4
        //    + $"[{MIN}{MAJ}_-]*" // optionnal suffix
        //    + $")+" // end group 3
        //    + $"(?=[^{ALPHANUM}'’-]|$)"; // group 5 separator (non alphanum character, no apostrophes, or end of line)

        /// <summary>
        /// Motif de recherche de mot(s)
        /// </summary>
        /// <param name="pattern"></param>
        /// <param name="opts"></param>
        /// <returns>
        /// Regex with
        /// - Group 1 : Code de protection optionnel
        /// - Group 2 : Motif détecté
        /// </returns>
        public static Regex SearchWord(string pattern)
        {
            return new Regex(
                $"(?<=[^{ALPHANUM}-]|^)({REG_CS}i{REG_CE}|{REG_CS}g1{REG_CE})?(?<!{REG_CS})({pattern})({REG_CS}g2{REG_CE})?(?=[^{ALPHANUM}'’-]|$)",
                RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.Singleline
            );
        }

        ///// <summary>
        ///// Recherche des mots contenant un chiffre <br/>
        ///// (lettre minuscule ou majuscul précédé ou suivi d'un chiffre)<br/>
        ///// Goupes :<br/>
        ///// - [1] : recup du code de protection s'il existe<br/>
        ///// - [2] : le mot contenant un ou plusieurs nombre
        /////
        ///// </summary>
        //private static readonly Regex WORD_WITH_NUMBERS = new Regex(
        //    WORD_WITH_NUMBERS_PATTERN,
        //    RegexOptions.Compiled | RegexOptions.Singleline
        //);

        ///// <summary>
        ///// Recherche tous les mots contenant au moins une majuscule
        ///// Goupes :<br/>
        ///// - [1] : recup du code de protection s'il existe<br/>
        ///// - [2] : le mot contenant une ou plusieurs majuscules
        ///// </summary>
        //private static readonly Regex WORD_WITH_CAPITAL = new Regex(
        //    WORD_WITH_CAPITAL_PATTERN,
        //    RegexOptions.Compiled | RegexOptions.Singleline
        //);


        /// <summary>
        /// Expression de recherche de mots (contenant au moins une majuscule ou une minuscule) hors code duxburry
        /// </summary>
        private static readonly Regex WORDS = new Regex(
                $"(?<=[^{ALPHANUM}-]|^)({REG_CS}i{REG_CE}|{REG_CS}g1{REG_CE})?(?<!{REG_CS})([{ALPHANUM}_-]*[{MIN}{MAJ}][{ALPHANUM}_-]*)({REG_CS}g2{REG_CE})?(?=[^{ALPHANUM}'’-]|$)",
                RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.Singleline
            );
        public static readonly Regex WORDSMAJNUM = new Regex(
            $"(?<=[^{ALPHANUM}-]|^)" +
            $"({REG_CS}i{REG_CE}|{REG_CS}g1{REG_CE})?" + // prefix de protection (mot ou bloc)
            $"(?<!{REG_CS})" +
            $"([{ALPHANUM}_-]*" +
                $"([{MIN}{MAJ}]+[{MAJ}{NUM}]|[{MAJ}{NUM}][{MIN}{MAJ}]+)" +
            $"[{ALPHANUM}_-]*)" +
            $"({REG_CS}g2{REG_CE})?" + // suffix de protection de bloc
            $"(?=[^{ALPHANUM}'’-]|$)",
                RegexOptions.Compiled | RegexOptions.Singleline
            );

        public static readonly Regex WORDSNUM = new Regex(
            $"(?<=[^{ALPHANUM}-]|^)" +
            $"({REG_CS}i{REG_CE}|{REG_CS}g1{REG_CE})?" + // prefix de protection (mot ou bloc)
            $"(?<!{REG_CS})" +
            $"([{ALPHANUM}_-]*" +
                $"([{MIN}{MAJ}]+[{NUM}]|[{NUM}][{MIN}{MAJ}]+)" +
            $"[{ALPHANUM}_-]*)" +
            $"({REG_CS}g2{REG_CE})?" + // suffix de protection de bloc
            $"(?=[^{ALPHANUM}'’-]|$)",
                RegexOptions.Compiled | RegexOptions.Singleline
            );

        public static readonly Regex WORDSMAJ = new Regex(
           $"(?<=[^{ALPHANUM}-]|^)" +
           $"({REG_CS}i{REG_CE}|{REG_CS}g1{REG_CE})?" + // prefix de protection (mot ou bloc)
           $"(?<!{REG_CS})" +
           $"([{ALPHANUM}_-]*" +
               $"([{MIN}{MAJ}]+[{MAJ}]|[{MAJ}][{MIN}{MAJ}]+)" +
           $"[{ALPHANUM}_-]*)" +
           $"({REG_CS}g2{REG_CE})?" + // suffix de protection de bloc
           $"(?=[^{ALPHANUM}'’-]|$)",
               RegexOptions.Compiled | RegexOptions.Singleline
           );


        public class OccurenceATraiter
        {
            public string Mot { get; set; }
            public int Index { get; set; }
            public bool EstDejaProteger { get; set; }
            public bool CommenceUnBlocIntegral { get; set; }
            public bool TermineUnBlocIntegral { get; set; }
            public bool ContientDesChiffres { get; set; }
            public bool ContientDesMajuscules { get; set; }
            public bool EstAbregeable { get; set; }

            public string ContexteAvant { get; set; } = "";

            public string ContexteApres { get; set; } = "";

            public bool EstFrançaisAbregeable { get; set; } = false;

            public bool EstAmbigu { get; set; } = false;
            public OccurenceATraiter(string mot, int index, bool estDejaProteger, bool contientDesChiffres = false, bool estAbregeable = false, bool estFrancais = false, bool estAmbigu = false, bool commenceUnBlocIntegral = false , bool termineUnBlocIntegral = false)
            {
                Mot = mot;
                Index = index;
                ContientDesChiffres = contientDesChiffres;
                ContientDesMajuscules = contientDesChiffres;
                EstDejaProteger = estDejaProteger;
                EstAbregeable = estAbregeable;
                //EstFrançaisAbregeable = estFrancais;
                //EstAmbigu = estAmbigu;
                TermineUnBlocIntegral = termineUnBlocIntegral;
                CommenceUnBlocIntegral = commenceUnBlocIntegral;
            }


        }

        /// <summary>
        /// Analyse complete du texte du document <br/>
        /// Juger trop long par les transcripteurs, doit revenir a la détection des seuls mots avec chiffres ou majuscules <br/>
        /// 
        /// </summary>
        /// <param name="texteAAnalyser"></param>
        /// <param name="info"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static Dictionary<int,OccurenceATraiter> AnalyserTexteComplet(string texteAAnalyser, Utils.OnInfoCallback info = null)
        {
            try {
                Dictionary<int, OccurenceATraiter> motsAnalyses = new Dictionary<int, OccurenceATraiter>();

                MatchCollection result = WORDS.Matches(texteAAnalyser);
                if (result.Count > 0) {
                    info?.Invoke($"Analyse des mots {result.Count} mots du document", new Tuple<int, int>(0, result.Count));
                    List<Task<OccurenceATraiter>> tasks = new List<Task<OccurenceATraiter>>();
                    int i = 1;
                    foreach (Match item in result) {
                        info?.Invoke(
                            "",
                            new Tuple<int, int>(i++, result.Count)
                        );
                        bool isAlreadyProtected = item.Groups[1].Success;
                        bool commenceUnBlocIntegral = item.Groups[1].Success && item.Groups[1].Value == "[[*g1*]]";
                        string foundWord = item.Groups[2].Value.Trim();
                        int pos = item.Groups[2].Index;
                        bool termineUnBlocIntegral = item.Groups[3].Success && item.Groups[3].Value == "[[*g2*]]";
                        string wordKey = foundWord.ToLower();
                        tasks.Add(Task.Run(() => new OccurenceATraiter(foundWord, pos, isAlreadyProtected)
                        {
                            ContientDesChiffres = Regex.Match(foundWord, $"[{NUM}]").Success,
                            ContientDesMajuscules = Regex.Match(foundWord, $"[{MAJ}]").Success,
                            EstAbregeable = EstAbregeable(wordKey),
                            EstFrançaisAbregeable = LexiqueFrance.EstFrancaisAbregeable(wordKey),
                            EstAmbigu = LexiqueFrance.EstAmbigu(wordKey)
                        }));
                    }

                    try {
                        i = 1;
                        info?.Invoke(
                            "Compilation des résultat de l'analyse"
                        );
                        
                        OccurenceATraiter[] tests = Task.WhenAll(tasks).Result;
                        foreach (OccurenceATraiter testMot in tests) {
                            info?.Invoke(
                                "",
                                new Tuple<int, int>(i++, result.Count)
                            );
                            motsAnalyses.Add(testMot.Index, testMot);
                        }
                    }
                    catch (Exception e) {
                        throw new Exception("Erreur lors de l'analyse des mots", e);
                    }
                }
                return motsAnalyses;
            }
            catch (AggregateException e) {
                throw new Exception("Erreur lors de l'analyse du texte", e);
            }
            
        }

        public static Dictionary<int, OccurenceATraiter> RechercheMotsAvecChiffres(string texteAAnalyser, Utils.OnInfoCallback info = null)
        {
            try {
                Dictionary<int, OccurenceATraiter> motsAnalyses = new Dictionary<int, OccurenceATraiter>();
                info?.Invoke($"Recherche des mots contenant des chiffres", new Tuple<int, int>(0, 0));

                MatchCollection result = WORDSNUM.Matches(texteAAnalyser);
                if (result.Count > 0) {
                    info?.Invoke($"Récupérations des {result.Count} mots détectés", new Tuple<int, int>(0, result.Count));
                    int i = 1;
                    foreach (Match item in result) {
                        info?.Invoke(
                            "",
                            new Tuple<int, int>(i++, result.Count)
                        );
                        bool isAlreadyProtected = item.Groups[1].Success;
                        bool commenceUnBlocIntegral = item.Groups[1].Success && item.Groups[1].Value == "[[*g1*]]";
                        string foundWord = item.Groups[2].Value.Trim();
                        int pos = item.Groups[2].Index;
                        bool termineUnBlocIntegral = item.Groups[3].Success && item.Groups[3].Value == "[[*g2*]]";
                        string wordKey = foundWord.ToLower();

                        motsAnalyses[pos] = new OccurenceATraiter(foundWord, pos, isAlreadyProtected)
                        {
                            ContientDesChiffres = true,
                            ContientDesMajuscules = false,
                            EstAbregeable = false,
                            //EstFrançaisAbregeable = LexiqueFrance.EstFrancaisAbregeable(wordKey),
                            //EstAmbigu = LexiqueFrance.EstAmbigu(wordKey)
                            CommenceUnBlocIntegral = commenceUnBlocIntegral,
                            TermineUnBlocIntegral = termineUnBlocIntegral
                        };
                    }
                }

                return motsAnalyses;
            }
            catch (AggregateException e) {
                throw new Exception("Erreur lors de l'analyse du texte", e);
            }
        }
        public static List<Task<OccurenceATraiter>> RechercheMotsAvecMaj(string texteAAnalyser, Utils.OnInfoCallback info = null)
        {
            try {
                List<Task<OccurenceATraiter>> motsAnalyses = new List<Task<OccurenceATraiter>>();
                info?.Invoke($"Recherche des mots contenant des majuscules...", new Tuple<int, int>(0, 0));

                MatchCollection result = WORDSMAJ.Matches(texteAAnalyser);
                if (result.Count > 0) {
                    info?.Invoke($"Lancement des taches d'analyse des {result.Count} mots détectés", new Tuple<int, int>(0, result.Count));
                    int i = 1;
                    foreach (Match item in result) {
                        info?.Invoke(
                            "",
                            new Tuple<int, int>(i++, result.Count)
                        );
                        bool isAlreadyProtected = item.Groups[1].Success;
                        bool commenceUnBlocIntegral = item.Groups[1].Success && item.Groups[1].Value == "[[*g1*]]";
                        string foundWord = item.Groups[2].Value.Trim();
                        int pos = item.Groups[2].Index;
                        bool termineUnBlocIntegral = item.Groups[3].Success && item.Groups[3].Value == "[[*g2*]]";
                        string wordKey = foundWord.ToLower();
                        // Récupération du contexte autour de l'occurence
                        int indexBefore = Math.Max(0, pos - 50);
                        int indexAfter = Math.Min(
                            texteAAnalyser.Length - 1,
                            pos + foundWord.Length + 50
                        );
                        string contextBefore = texteAAnalyser.Substring(
                            indexBefore,
                            pos - indexBefore
                        );
                        string contextAfter = "";

                        contextAfter = texteAAnalyser.Substring(
                            pos + foundWord.Length,
                            indexAfter - foundWord.Length - pos
                        );

                        motsAnalyses.Add(Task.Run(() => new OccurenceATraiter(foundWord, pos, isAlreadyProtected)
                        {
                            ContientDesChiffres = false,
                            ContientDesMajuscules = true,
                            EstAbregeable = EstAbregeable(wordKey),
                            //EstFrançaisAbregeable = LexiqueFrance.EstFrancaisAbregeable(wordKey),
                            //EstAmbigu = LexiqueFrance.EstAmbigu(wordKey)
                            CommenceUnBlocIntegral = commenceUnBlocIntegral,
                            TermineUnBlocIntegral = termineUnBlocIntegral,
                            ContexteAvant = contextBefore,
                            ContexteApres = contextAfter
                        }));
                    }
                }

                return motsAnalyses;
            }
            catch (AggregateException e) {
                throw new Exception("Erreur lors de l'analyse du texte", e);
            }
        }

       
        #endregion


    }
}
