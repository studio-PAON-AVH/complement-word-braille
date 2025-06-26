using fr.avh.braille.dictionnaire.Entities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace fr.avh.braille.dictionnaire
{
    public class LexiqueFrance
    {
        /// <summary>
        /// Lexique français construit a l'aide du GLAFF et du dictionnaire DELA.
        /// Pour chaque mot, il est indiqué s'il est abrégeable ou non.
        /// </summary>
        public static Lazy<Dictionary<string, bool>> LexiqueComplet = new Lazy<Dictionary<string, bool>>(() =>
        {
            Dictionary<string, bool> _lexique = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            Assembly.GetExecutingAssembly()
                        .GetManifestResourceNames()
                        .ToList()
                        .ForEach(name =>
                        {
                            if (name.EndsWith("megalexique-braille.txt")) {
                                using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(name))
                                using (StreamReader reader = new StreamReader(stream)) {
                                    string line;
                                    while ((line = reader.ReadLine()) != null) {
                                        string motEtStatut = line.Trim().ToLower();
                                        // 1 = abregeable, 0 = non abregeable
                                        int code = motEtStatut[0] - '0';
                                        string mot = motEtStatut.Substring(1).Trim();
                                        if (!string.IsNullOrWhiteSpace(mot)) {
                                            _lexique.Add(mot, code == 1);
                                        } else {
                                            // Si le mot est vide, on ne l'ajoute pas
                                            Console.WriteLine("Mot vide trouvé dans le lexique : " + motEtStatut);
                                        }
                                    }
                                }
                            }
                        });
           return _lexique;
        });
        /// <summary>
        /// Lexique ne contenant que les mots français abregeables.
        /// </summary>
        public static Lazy<HashSet<string>> LexiqueAbreger = new Lazy<HashSet<string>>(
            () => LexiqueComplet.Value.Where(kv => kv.Value).Select(kv => kv.Key).ToHashSet(StringComparer.OrdinalIgnoreCase)
        );

        public static Lazy<Dictionary<string, string>> LexiqueAmbigu = new Lazy<Dictionary<string, string>>(() =>
        {
            Dictionary<string, string> _lexique = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            Assembly.GetExecutingAssembly()
                .GetManifestResourceNames()
                .ToList()
                .ForEach(name =>
                {
                    if (name.EndsWith("lexique_ambigus.csv")) {
                        using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(name))
                        using (StreamReader reader = new StreamReader(stream)) {
                            string line = reader.ReadLine();
                            // skip la premiere ligne
                            while ((line = reader.ReadLine()) != null) {
                                string motEtCommentaire = line.Trim().ToLower();
                                if (!string.IsNullOrWhiteSpace(motEtCommentaire)) {
                                    string[] donnees = motEtCommentaire.Split(new[] { ';' });
                                    donnees[0] = donnees[0].Replace("\"", ""); // enlever les guillemets
                                    donnees[1] = donnees[1].Replace("\"", ""); // enlever les guillemets
                                    _lexique.Add(donnees[0], donnees[1]);
                                }
                            }
                        }
                    }
                });
            return _lexique;
        });

        /// <summary>
        /// 
        /// </summary>
        /// <param name="mot"></param>
        /// <returns></returns>
        public static bool EstFrancais(string mot)
        {
            string m = mot.ToLower().Trim();
            List<string> decomposed = m.Split(new[] { '-' }).ToList();

            // Le mot appartient au lexique si 
            // 1. il est présent dans le lexique
            // 2. ou s'il est au pluriel et que le singulier est présent dans le lexique
            // 3. ou si c'est un mot composé, si tous les sous-mots sont présents dans le lexique
            return LexiqueComplet.Value.Keys.Contains(m)
                || (
                    m.EndsWith("s") ? LexiqueComplet.Value.Keys.Contains(m.Substring(0, m.Length - 1)) : false
                )
                || decomposed.All(p =>
                    LexiqueComplet.Value.Keys.Contains(p)
                    || (p.EndsWith("s") ? LexiqueComplet.Value.Keys.Contains(p.Substring(0, p.Length - 1)) : false)
                 );
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="mot"></param>
        /// <returns></returns>
        public static bool EstFrancaisAbregeable(string mot)
        {
            string m = mot.ToLower().Trim();
            List<string> decomposed = m.Split(new[] { '-' }).ToList();

            // Le mot appartient au lexique si 
            // 1. il est présent dans le lexique
            // 2. ou s'il est au pluriel et que le singulier est présent dans le lexique
            // 3. ou si c'est un mot composé, si tous les sous-mots sont présents dans le lexique
            return LexiqueAbreger.Value.Contains(m)
                || (
                    m.EndsWith("s") ? LexiqueAbreger.Value.Contains(m.Substring(0, m.Length - 1)) : false
                )
                || decomposed.All(p =>
                    LexiqueAbreger.Value.Contains(p)
                    || (p.EndsWith("s") ? LexiqueAbreger.Value.Contains(p.Substring(0, p.Length - 1)) : false)
                 );
        }

        /// <summary>
        /// Test si un mot est "ambigu" (s'il existe en français ou dans une autre langue,
        /// ou si le mot est un nom ou un acronyme selon le contexte)
        /// </summary>
        /// <param name="mot"></param>
        /// <returns></returns>
        public static bool EstAmbigu(string mot)
        {
            string m = mot.ToLower().Trim();
            List<string> decomposed = m.Split(new[] { '-' }).ToList();

            // Le mot appartient au lexique si 
            // 1. il est présent dans le lexique
            // 2. ou s'il est au pluriel et que le singulier est présent dans le lexique
            // 3. ou si c'est un mot composé, si tous les sous-mots sont présents dans le lexique
            return LexiqueAmbigu.Value.Keys.Contains(m)
                || (
                    m.EndsWith("s") ? LexiqueAmbigu.Value.Keys.Contains(m.Substring(0, m.Length - 1)) : false
                )
                || decomposed.All(p =>
                    LexiqueAmbigu.Value.Keys.Contains(p)
                    || (p.EndsWith("s") ? LexiqueAmbigu.Value.Keys.Contains(p.Substring(0, p.Length - 1)) : false)
                 );
        }

    }
}
