using fr.avh.braille.dictionnaire.Entities;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace fr.avh.braille.dictionnaire
{
    public class DictionnaireDeTravail
    {
        public string NomDictionnaire { get; }

        /// <summary>
        /// Liste des occurences de mots détectés dans un texte
        /// (La casse est conservée)
        /// </summary>
        public List<string> Occurences { get; private set; }

        /// <summary>
        /// Position des occurences dans le texte analysé
        /// Par défaut, la position est mise à 0 si l'information n'est pas disponible
        /// </summary>
        public List<int> PositionsOccurences { get; private set; }

        /// <summary>
        /// Liste des statuts d'occurences
        /// </summary>
        public List<Statut> StatutsOccurences { get; private set; }

        public List<bool> StatutsAppliquer { get; private set; }

        /// <summary>
        /// Liste des contextes avant les occurences
        /// </summary>
        public List<string> ContextesAvantOccurences { get; private set; }

        /// <summary>
        /// Liste des contextes apres les occurences
        /// </summary>
        public List<string> ContextesApresOccurences { get; private set; }

        /// <summary>
        /// Dictionnaires de listes des indices d'occurences trouvés pour un mot donnée.
        /// La clé est le mot en minuscule
        /// </summary>
        public Dictionary<string, List<int>> CarteMotOccurences { get; private set; }

        public bool ContientLeMot(string mot)
        {
            return CarteMotOccurences.ContainsKey(mot.ToLower().Trim());
        }

        /// <summary>
        /// Liste d'occurences condensées en tuples
        /// </summary>
        /// <param name="mot">si définit, renvoie uniquement les occurences du mot spécifié dans le texte</param>
        /// <returns>Une liste de tuple (index, mot, statut, contextAvant, contextApres)</returns>
        public List<Tuple<int, string, Statut, string, string>> OccurencesAsListOfTuples(
            string mot = ""
        )
        {
            List<Tuple<int, string, Statut, string, string>> result =
                new List<Tuple<int, string, Statut, string, string>>();
            if (mot.Length > 0)
            {
                if (CarteMotOccurences.ContainsKey(mot.ToLower().Trim()))
                {
                    foreach (int indexOccurence in CarteMotOccurences[mot.ToLower().Trim()])
                    {
                        result.Add(
                            new Tuple<int, string, Statut, string, string>(
                                indexOccurence,
                                Occurences[indexOccurence],
                                StatutsOccurences[indexOccurence],
                                ContextesAvantOccurences[indexOccurence],
                                ContextesApresOccurences[indexOccurence]
                            )
                        );
                    }
                }
            }
            else
            {
                for (int indexOccurence = 0; indexOccurence < Occurences.Count; indexOccurence++)
                {
                    result.Add(
                        new Tuple<int, string, Statut, string, string>(
                            indexOccurence,
                            Occurences[indexOccurence],
                            StatutsOccurences[indexOccurence],
                            ContextesAvantOccurences[indexOccurence],
                            ContextesApresOccurences[indexOccurence]
                        )
                    );
                }
            }

            return result;
        }

        public DictionnaireDeTravail Clone()
        {
            DictionnaireDeTravail clone = new DictionnaireDeTravail(NomDictionnaire);
            clone.Occurences = new List<string>(Occurences);
            clone.PositionsOccurences = new List<int>(PositionsOccurences);
            clone.StatutsOccurences = new List<Statut>(StatutsOccurences);
            clone.StatutsAppliquer = new List<bool>(StatutsAppliquer);
            clone.CarteMotOccurences = new Dictionary<string, List<int>>(CarteMotOccurences);
            clone.ContextesAvantOccurences = new List<string>(ContextesAvantOccurences);
            clone.ContextesApresOccurences = new List<string>(ContextesApresOccurences);

            return clone;
        }

        /// <summary>
        /// Ajouter une nouvelle occurence de mot dans le dictionnaire
        /// </summary>
        /// <param name="mot"></param>
        /// <param name="statut"></param>
        /// <param name="contexte"></param>
        public void Add(
            string mot,
            Statut statut,
            string contexteAvant = "",
            string contextApres = "",
            int position = 0
        )
        {
            lock (this) {
                Occurences.Add(mot);
                PositionsOccurences.Add(position);
                StatutsOccurences.Add(statut);
                StatutsAppliquer.Add(false);
                ContextesAvantOccurences.Add(contexteAvant);
                ContextesApresOccurences.Add(contextApres);
            }

        }

        public void ReorderOccurences()
        {
            // Réordonner les listes d'informations des occurences en fonction de la liste des positions
            List<int> reorderedIndexes = new List<int>();
            for (int i = 0; i < PositionsOccurences.Count; i++)
            {
                reorderedIndexes.Add(i);
            }
            reorderedIndexes.Sort(
                (a, b) => PositionsOccurences[a].CompareTo(PositionsOccurences[b])
            );

            List<string> newOccurences = new List<string>(Occurences.Count);
            List<int> newPositionsOccurences = new List<int>(PositionsOccurences.Count);
            List<Statut> newStatutsOccurences = new List<Statut>(StatutsOccurences.Count);
            List<bool> newStatutsAppliquer = new List<bool>(StatutsAppliquer.Count);
            List<string> newContextesAvantOccurences = new List<string>(
                ContextesAvantOccurences.Count
            );
            List<string> newContextesApresOccurences = new List<string>(
                ContextesApresOccurences.Count
            );

            // NP 2024 10 04 : supprimer les détections en doublons
            // Reconstruirer les listes d'occurences basé sur cette liste d'index
            int previousPosition = -1;
            foreach (int newIndex in reorderedIndexes)
            {
                if (PositionsOccurences[newIndex] == previousPosition)
                {
                    // On ignore les doublons en nous basant sur les positions dans le texte analysé
                    continue;
                }
                else
                {
                    previousPosition = PositionsOccurences[newIndex];
                }
                newOccurences.Add(Occurences[newIndex]);
                newPositionsOccurences.Add(PositionsOccurences[newIndex]);
                newStatutsOccurences.Add(StatutsOccurences[newIndex]);
                newStatutsAppliquer.Add(StatutsAppliquer[newIndex]);
                newContextesAvantOccurences.Add(ContextesAvantOccurences[newIndex]);
                newContextesApresOccurences.Add(ContextesApresOccurences[newIndex]);
            }
            // Remplacement des listes précédentes
            Occurences = newOccurences;
            PositionsOccurences = newPositionsOccurences;
            StatutsOccurences = newStatutsOccurences;
            StatutsAppliquer = newStatutsAppliquer;
            ContextesAvantOccurences = newContextesAvantOccurences;
            ContextesApresOccurences = newContextesApresOccurences;
        }

        /// <summary>
        /// Recalcul la carte des mots et de leurs indices dans la liste des occurences
        /// </summary>
        public void ComputeWordMap()
        {
            CarteMotOccurences.Clear();
            for (int i = 0; i < Occurences.Count; i++)
            {
                string mot = Occurences[i].ToLower().Trim();
                if (!CarteMotOccurences.ContainsKey(mot))
                {
                    CarteMotOccurences.Add(mot, new List<int>());
                }
                CarteMotOccurences[mot].Add(i);
            }
        }

        public void SetStatut(string mot, Statut selected)
        {
            if (CarteMotOccurences.ContainsKey(mot.ToLower().Trim()))
            {
                foreach (int index in CarteMotOccurences[mot.ToLower().Trim()])
                {
                    StatutsOccurences[index] = selected;
                }
            }
        }

        /// <summary>
        /// Récupérer le statut global d'un mot dans le dictionnaire
        /// </summary>
        /// <param name="mot"></param>
        /// <returns></returns>
        public Statut StatutMot(string mot)
        {
            Statut statut = Statut.INCONNU;
            if (CarteMotOccurences.ContainsKey(mot.ToLower().Trim()))
            {
                foreach (int index in CarteMotOccurences[mot.ToLower().Trim()])
                {
                    if (StatutsOccurences[index] == Statut.INCONNU)
                    {
                        // S'il reste une occurence inconnu, le statut est inconnu
                        return Statut.INCONNU;
                    }
                    else if (
                        StatutsOccurences[index] == Statut.IGNORE
                        || statut == Statut.PROTEGE && StatutsOccurences[index] == Statut.ABREGE
                        || statut == Statut.ABREGE && StatutsOccurences[index] == Statut.PROTEGE
                    )
                    {
                        // Si une seul occurence est marqué ambigu,
                        // ou si le mot a été protégé ou abrégé puis a ensuite été respectivement abrégé ou protégé
                        // On considere qu'on ignore le statut général (detournement de sens du statut ignorer, mais c'est pour simplifier)
                        // (normalement ce statut c'est pour dire que le transcripteur a ignorer le mot lors du traitement)
                        return Statut.IGNORE;
                    }
                    else
                    {
                        statut = StatutsOccurences[index];
                    }
                }
            }
            return statut;
        }

        public DictionnaireDeTravail(string nom)
        {
            NomDictionnaire = nom;
            Occurences = new List<string>();
            PositionsOccurences = new List<int>();
            StatutsOccurences = new List<Statut>();
            StatutsAppliquer = new List<bool>();
            CarteMotOccurences = new Dictionary<string, List<int>>();
            ContextesAvantOccurences = new List<string>();
            ContextesApresOccurences = new List<string>();
        }

        public static async Task<DictionnaireDeTravail> FromDictionnaryFile(
            string filePath,
            CancellationTokenSource canceler = null
        )
        {
            if (canceler != null && canceler.Token.IsCancellationRequested)
            {
                return null;
            }
            DictionnaireDeTravail result = new DictionnaireDeTravail(Path.GetFileName(filePath));
            Globals.logAsync($"Analyse de {result.NomDictionnaire}");
            // Récupération du contenu
            Encoding utf8 = Encoding.UTF8;
            // NP 20240909
            // possible probleme de chargement des dictionnaires a confirmer
            // prevoire le chargement de tous le fichier en amont si ça se confirme
            // et remplacer le writeline par un write + '\n' (pour virer le \r de microsoft)

            // plus de compteur de mot (le nb de ligne est le nb de mot)
            // 1 ligne = 1 code + 1 occurence d'un mot identifié dans le texte (avec sa casse)
            // A noter que la ligne du mot est supposément l'identifiant de son occurence dans le texte
            using (StreamReader fileReader = new StreamReader(filePath, utf8))
            {
                string line;
                while ((line = await fileReader.ReadLineAsync()) != null)
                {
                    if (line.Length == 0)
                    { // avoid empty lines
                        continue;
                    }
                    if (canceler != null && canceler.Token.IsCancellationRequested)
                    {
                        return null;
                    }

                    int code = line[0] - '0';
                    string word = line.Substring(1);
                    try
                    {
                        // get the code as first char
                        Statut wordStatus = Enum.IsDefined(typeof(Statut), code)
                            ? (Statut)code
                            : Statut.INCONNU;
                        result.Add(word, wordStatus);
                    }
                    catch (Exception e)
                    {
                        // Abort transaction
                        Globals.logAsync(
                            $"Skipping line {line} due to error while parsing : {e.Message}"
                        );
                        //throw new Exception($"Une erreur s'est produite en analysant '{line}'", e);
                    }
                }
            }
            result.ComputeWordMap();
            return result;
        }

        /// <summary>
        /// Sauvegarder sur le disque
        /// </summary>
        /// <param name="path"></param>
        public void Save(DirectoryInfo folder)
        {
            Encoding utf8 = Encoding.UTF8;
            using (
                StreamWriter fileWriter = new StreamWriter(
                    Path.Combine(folder.FullName, NomDictionnaire),
                    false,
                    utf8
                )
            )
            {
                for (int i = 0; i < Occurences.Count; i++)
                {
                    fileWriter.WriteLine($"{(int)StatutsOccurences[i]}{Occurences[i]}");
                }
            }
        }

        public static async Task<DictionnaireDeTravail> FromDictionnaryFileJSON(
           string filePath,
           CancellationTokenSource canceler = null
       )
        {
            // Vérifier si l'opération a été annulée
            if (canceler != null && canceler.Token.IsCancellationRequested)
            {
                return null;
            }

            // Créer un nouveau dictionnaire de travail à partir du nom du fichier
            DictionnaireDeTravail result = new DictionnaireDeTravail(Path.GetFileName(filePath));
            Globals.logAsync($"Analyse de {result.NomDictionnaire}");
            Encoding utf8 = Encoding.UTF8;

            // Ouvre le fichier en lecture avec l'encodage UTF-8
            using (StreamReader fileReader = new StreamReader(filePath, utf8))
            {
                string jsonContent = await fileReader.ReadToEndAsync();
                dynamic jsonObject = JsonConvert.DeserializeObject(jsonContent);

                // Parcours de la liste des mots
                foreach (var mot in jsonObject.mots)
                {
                    string word = mot.Name;
                    Statut statut = (Statut)mot.Value.statut;
                    List<int> occurences = mot.Value.occurences.ToObject<List<int>>();
                    foreach (int occurence in occurences)
                    {
                        result.Add(word, statut, position: occurence);
                    }
                }
            }
            result.ComputeWordMap();
            return result;
        }

        /// <summary>
        /// Sauvegarder sur le disque
        /// </summary>
        /// <param name="path"></param>
        public void SaveJSON(DirectoryInfo folder)
        {
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var jsonObject = new
            {
                version = "1.0",
                nom = NomDictionnaire,
                mots = CarteMotOccurences.ToDictionary(
                    kvp => kvp.Key,
                    kvp => new
                    {
                        statut = StatutMot(kvp.Key),
                        occurences = kvp.Value
                    }
                ),
            };
            string jsonContent = JsonConvert.SerializeObject(jsonObject, Newtonsoft.Json.Formatting.Indented);
            string cleanedNomDictionnaire = NomDictionnaire.Replace(".bdic", "");
            File.WriteAllText(Path.Combine(appDataPath, cleanedNomDictionnaire + ".json"), jsonContent);
        }
    }
}
