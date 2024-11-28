using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace fr.avh.braille.dictionnaire
{
    /// <summary>
    /// Représentation mémoire d'un fichier "dictionnaire"
    /// obtenu apres application du traitement de protection
    /// sur un fichier word (voir macro VBA)
    /// </summary>
    public class AncienDictionnaireDeTravail
    {
        public string NomDictionnaire { get; }

        public Dictionary<string, Statut> Mots { get; }

        public AncienDictionnaireDeTravail(string nom)
        {
            NomDictionnaire = nom;
            Mots = new Dictionary<string, Statut>();
        }

        public async static Task<AncienDictionnaireDeTravail> FromDictionnaryFile(string filePath, CancellationTokenSource canceler = null)
        {
            if (canceler != null && canceler.Token.IsCancellationRequested) {
                return null;
            }
            AncienDictionnaireDeTravail result = new AncienDictionnaireDeTravail(Path.GetFileName(filePath));
            Globals.logAsync($"Parsing {result.NomDictionnaire}");
            // Récupération du contenu
            Encoding wd1252 = Encoding.GetEncoding(1252);
            // Note : le contenu est en wd1252 et pas en UTF8
            // Conversion du contenu en de la codepage 1252 en UTF8
            // Note: la première ligne contient le compteur de mots
            using (StreamReader fileReader = new StreamReader(filePath, wd1252)) {
                string line;
                while ((line = await fileReader.ReadLineAsync()) != null) {
                    if (canceler != null && canceler.Token.IsCancellationRequested) {
                        return null;
                    }
                    // Only parse lines that have more than 1 caractere and that are not pure numbers
                    if (line.Length > 1 && !line.All((c) => char.IsDigit(c))) {
                        int code = line[0] - '0';
                        string word = line.Substring(1);
                        try {
                            // get the code as first char
                            Statut wordStatus = Enum.IsDefined(typeof(Statut), code)
                                ? (Statut)code
                                : Statut.INCONNU;
                            result.Mots.Add(word.ToLower(), wordStatus);
                        }
                        catch (Exception e) {
                            // Abort transaction
                            Globals.logAsync($"Skipping line {line} due to error while parsing : {e.Message}");
                            //throw new Exception($"Une erreur s'est produite en analysant '{line}'", e);
                        }

                    } else {
                        Globals.logAsync($"Skipping line {line} in {result.NomDictionnaire}");
                    }
                }
            }
            return result;
        }
        /// <summary>
        /// Sauvegarder sur le disque
        /// </summary>
        /// <param name="path"></param>
        public void Save(DirectoryInfo folder)
        {
            Encoding wd1252 = Encoding.GetEncoding(1252);
            using (StreamWriter fileWriter = new StreamWriter(Path.Combine(folder.FullName,NomDictionnaire), false, wd1252)) {
                fileWriter.WriteLine(Mots.Count);
                foreach(var mot in Mots) {
                    fileWriter.WriteLine($"{mot.Value.ToString("d")}{mot.Key.ToLower()}");
                }
            }
        }


    }
}
