using fr.avh.braille.dictionnaire;
using fr.avh.braille.dictionnaire.Entities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ScriptsBaseDeDonnée
{
    internal class Program
    {

        static readonly Dictionary<string, string> Reencodeur = new Dictionary<string, string>()
        {
            { "ã©", "é" },
            { "ãª", "ê" },
            { "ã´", "ô" },
            { "ã¨", "è"  },

            
        };

        static readonly List<string> PrefixeAEnlever = new List<string>()
        {
            "j'l'",
            "j't'",
            "j'm'",
            "v'z'",
            "v'",
            "z'",
            "l'",
            "s'",
            "d'",
            "qu'",
            "m'",
            "n'",
            "t'",
            "j'",
        };
        static void CleanupDatabase()
        {
            List<Mot> toRemove = new List<Mot>();
            List<Mot> contenuDelaBase = new List<Mot>();
            // Pour chaque mot de la base de donnée
            using (var session = BaseSQlite.CreateSessionFactory().OpenSession()) {
                Console.WriteLine($"Requete à la base ...");
                contenuDelaBase = session
                        .QueryOver<Mot>()
                        .List()
                        .ToList();
                Console.WriteLine($" -- Controle de {contenuDelaBase.Count} mots ...");
                foreach (var mot in contenuDelaBase) {
                    try {
                        if(PrefixeAEnlever.Any((p) => mot.Texte.ToLower().StartsWith(p))) {
                            Console.WriteLine($"Supression de {mot.Texte}");
                            toRemove.Add(mot);
                        }
                        //if (!Abbreviation.EstAbregeable(mot.Texte)) {
                        //    toRemove.Add(mot);
                        //}
                    } catch (Exception e) {
                        Console.WriteLine("Erreur lors du controle du mot " + mot.Texte + " : " + e.Message);
                        if(mot.Texte.Length > 5)
                            Console.WriteLine("Controler " + mot.Texte);
                        toRemove.Add(mot);
                    }
                }
                Console.WriteLine($" -- Suppression de {toRemove.Count} mots ...");
                using (var t = session.BeginTransaction()) {
                    foreach (var mot in toRemove) {
                        int test = PrefixeAEnlever.FindIndex((p) => mot.Texte.ToLower().StartsWith(p));
                        if(test >= 0) {
                            string cleanedup = mot.Texte.ToLower().Substring(PrefixeAEnlever[test].Length);
                            int found = -1;
                            for (int i = 0; i < contenuDelaBase.Count && found < 0; i++) {
                                if (contenuDelaBase[i].Texte == cleanedup) {
                                    found = i;
                                }
                            }
                            if (found >= 0) {
                                Mot toUpdate = contenuDelaBase[found];
                                Console.WriteLine($"Fusion de {mot.Texte} avec {toUpdate.Texte}");
                                toUpdate.Protections += mot.Protections;
                                toUpdate.Abreviations += mot.Abreviations;
                                toUpdate.Documents += mot.Documents;
                                session.Update(toUpdate);
                                session.Delete(mot);
                            } else {
                                Console.WriteLine($"remplacement de {mot.Texte} avec {cleanedup}");
                                mot.Texte = cleanedup;
                                session.Update(mot);
                            }
                        } else {
                            Console.WriteLine($"Probleme avec {mot.Texte} !!!!!!!!");
                        }
                        
                    }
                    t.Commit();
                }
                //using (var t = session.BeginTransaction()) {
                //    foreach (var mot in toRemove) {
                //        Console.WriteLine($"Supression de {mot.Texte}");
                //        session.Delete(mot);
                //        t.Commit();
                //    }    
                //}
            }


        }

        static void exportAsCSV()
        {
            FileStream output = File.Open("resultat.csv", FileMode.Create);

            using (StreamWriter writer = new StreamWriter(output)) {
                writer.WriteLine($"Mot;Protections;Abreviations;Commentaire;Recommandation");
                using (var session = BaseSQlite.CreateSessionFactory().OpenSession()) {
                    List<Mot> mots = session.QueryOver<Mot>().List().ToList();
                    foreach (var mot in mots) {
                        writer.WriteLine($"{mot.Texte};{mot.Protections};{mot.Abreviations};{mot.Commentaires};");
                    }
                }
            }
            output.Close();
        }

        static void Main(string[] args)
        {
            exportAsCSV();
            //CleanupDatabase();
            return;
        }
    }
}
