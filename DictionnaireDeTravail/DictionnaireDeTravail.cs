using fr.avh.braille.dictionnaire.Entities;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NHibernate.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using static fr.avh.braille.dictionnaire.Abreviation;

namespace fr.avh.braille.dictionnaire
{
    public class DictionnaireDeTravail
    {
        #region Constantes et Champs
        public static readonly string PROTECTION_CODE = "[[*i*]]"; // prochain mot en integral
        public static readonly string PROTECTION_CODE_G1 = "[[*g1*]]"; // debut bloc integral
        public static readonly string PROTECTION_CODE_G2 = "[[*g2*]]"; // debut bloc abregé

        public string NomDictionnaire { get; private set; }

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
        /// Liste des statuts d'occurences demander
        /// </summary>
        public List<Statut> StatutsOccurences { get; private set; }

        /// <summary>
        /// Liste de bool indiquant si le statut d'une occurence est appliqué ou non dans le document
        /// (   vrai si statut protégé avec code de protection devant l'occurence,
        ///     ou si statut abrégé ou ignorer sans code devant l'occurence
        /// )
        /// </summary>
        public List<bool> EstTraitee { get; private set; }

        /// <summary>
        /// Liste des contextes avant les occurences
        /// </summary>
        public List<string> ContextesAvantOccurences { get; private set; }

        /// <summary>
        /// Liste des contextes apres les occurences
        /// </summary>
        public List<string> ContextesApresOccurences { get; private set; }

        /// <summary>
        /// Position de debut d'un bloc integral (encadré par des codes g1 et g2)
        /// (exclure le code G2 si présent)
        /// </summary>
        public SortedSet<int> DebutsBlocsIntegrals { get; private set; }

        /// <summary>
        /// Position de fin d'un bloc integral (encadré par des codes g1 et g2)
        /// (Exclure le code G1 si présent)
        /// </summary>
        public SortedSet<int> FinsBlocsIntegrals { get; private set; }

        /// <summary>
        /// Liste des blocs de mots en intégral sous la forme d'un tuple de positions (debut, fin) dans le texte
        /// Si le bloc n'est pas terminé (bloc allant jusqu'a la fin du document), la fin est -1
        /// </summary>
        public List<Tuple<int,int>> ListeBlocsIntegral { 
            get {
                return DebutsBlocsIntegrals.Select(
                    (debut, index) => new Tuple<int, int>(
                        debut,
                        index < FinsBlocsIntegrals.Count 
                            ? FinsBlocsIntegrals.ElementAt(index)
                            : -1
                        )
                    ).OrderBy(t => t.Item1).ToList();
            }
        }

        public enum FORMAT
        {
            DDIC, // Dictionnaire de travail
            BDIC, // Dictionnaire de base
            JSON  // Dictionnaire en JSON
        };

        public FORMAT Format { get; private set; } = FORMAT.JSON; // Par défaut, on utilise le format BDIC (pour les dictionnaires de base)


        // TODO : ajouter des objets de gestion des phrases étrangères pour faire une interface dédié:
        // Le traitement des phrases étrangères par Word est extrèmement long ...


        /// <summary>
        /// Dictionnaires de listes des indices d'occurences trouvés pour un mot donnée.
        /// La clé est le mot en minuscule
        /// </summary>
        public Dictionary<string, List<int>> CarteMotOccurences { get; private set; }

        /// <summary>
        /// Statut global pour un mot
        /// INCONNU par défaut, mais changer quand un transcripteur choisi de protéger, abréger ou ignorer un mot
        /// </summary>
        public Dictionary<string, Statut> CarteMotStatut { get; private set; }

        public string DernierMotSelectionne { get; set; } = null;

        public int DerniereOccurenceSelectionnee { get; set; } = -1;

        public bool EstTerminer { get; set; } = false;

        #endregion
        public DictionnaireDeTravail(string nom)
        {
            NomDictionnaire = nom;
            Occurences = new List<string>();
            PositionsOccurences = new List<int>();
            StatutsOccurences = new List<Statut>();
            EstTraitee = new List<bool>();
            CarteMotOccurences = new Dictionary<string, List<int>>();
            CarteMotStatut = new Dictionary<string, Statut>();
            ContextesAvantOccurences = new List<string>();
            ContextesApresOccurences = new List<string>();
            DebutsBlocsIntegrals = new SortedSet<int>();
            FinsBlocsIntegrals = new SortedSet<int>();
        }

        public static async Task<DictionnaireDeTravail> DepuisFichier(string filePath)
        {
            if (!File.Exists(filePath)) {
                throw new FileNotFoundException(filePath);
            }
            DictionnaireDeTravail importer = null;
            try {
                switch (Path.GetExtension(filePath).ToLower()) {
                    case ".json":
                        importer = await FromDictionnaryFileJSON(filePath);
                        break;
                    case ".bdic":
                        importer = await FromDictionnaryFileBDIC(filePath);
                        break;
                    case ".ddic":
                        importer = await FromDictionnaryFileDDIC(filePath);
                        break;
                    default:
                        throw new NotSupportedException("Format de fichier non supporté : " + Path.GetExtension(filePath));
                }
            }
            catch (Exception ex) {
                throw new Exception($"Impossible de charger le dictionnaire {filePath}", ex);
            }
            return importer;
        }

        #region format DDIC
        /// <summary>
        /// Dictionnaire ne contenant qu'une carte de statut a appliquer
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="canceler"></param>
        /// <returns></returns>
        public static async Task<DictionnaireDeTravail> FromDictionnaryFileDDIC(
            string filePath,
            CancellationTokenSource canceler = null
        )
        {
            if (canceler != null && canceler.Token.IsCancellationRequested) {
                return null;
            }
            DictionnaireDeTravail result = new DictionnaireDeTravail(Path.GetFileNameWithoutExtension(filePath));
            Globals.logAsync($"Analyse de {result.NomDictionnaire}");

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
                            result.CarteMotStatut[word.ToLower().Trim()] = wordStatus;
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
            result.Format = FORMAT.DDIC; // On indique que le format est DDIC
            return result;
        }


        public void SaveDDIC(DirectoryInfo folder)
        {
            Encoding wd1252 = Encoding.GetEncoding(1252);
            using (StreamWriter fileWriter = new StreamWriter(Path.Combine(folder.FullName, NomDictionnaire), false, wd1252)) {
                fileWriter.WriteLine(CarteMotStatut.Count);
                foreach (var mot in CarteMotStatut) {
                    fileWriter.WriteLine($"{mot.Value.ToString("d")}{mot.Key.ToLower()}");
                }
            }
        }

        #endregion

        #region format BDIC

        public static async Task<DictionnaireDeTravail> FromDictionnaryFileBDIC(
            string filePath,
            CancellationTokenSource canceler = null
        )
        {
            if (canceler != null && canceler.Token.IsCancellationRequested) {
                return null;
            }
            DictionnaireDeTravail result = new DictionnaireDeTravail(Path.GetFileNameWithoutExtension(filePath));
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
            using (StreamReader fileReader = new StreamReader(filePath, utf8)) {
                string line;
                while ((line = await fileReader.ReadLineAsync()) != null) {
                    if (line.Length == 0) { // avoid empty lines
                        continue;
                    }
                    if (canceler != null && canceler.Token.IsCancellationRequested) {
                        return null;
                    }

                    int code = line[0] - '0';
                    string word = line.Substring(1);
                    try {
                        // get the code as first char
                        Statut wordStatus = Enum.IsDefined(typeof(Statut), code)
                            ? (Statut)code
                            : Statut.INCONNU;
                        result.AjouterOccurence(word.Trim(), wordStatus);
                    }
                    catch (Exception e) {
                        // Abort transaction
                        Globals.logAsync(
                            $"Skipping line {line} due to error while parsing : {e.Message}"
                        );
                        //throw new Exception($"Une erreur s'est produite en analysant '{line}'", e);
                    }
                }
            }
            result.CalculCartographieMots();
            foreach(var motEtOccurences in result.CarteMotOccurences) {
                string mot = motEtOccurences.Key;
                List<int> occurences = motEtOccurences.Value;
                result.CarteMotStatut[mot] = Statut.INCONNU;
                if (occurences.Count > 0) {
                    // recalcul du statut global stricte (ne reprendre que les mots pour lesquels une décision systématique a été prise)
                    result.CarteMotStatut[mot] = result.CalculStatutMot(mot, true);
                }
            }

            result.Format = FORMAT.BDIC; // On indique que le format est BDIC
            return result;
        }




        /// <summary>
        /// Sauvegarder sur le disque
        /// </summary>
        /// <param name="path"></param>
        public void SaveBDIC(DirectoryInfo folder)
        {
            Encoding utf8 = Encoding.UTF8;
            using (
                StreamWriter fileWriter = new StreamWriter(
                    Path.Combine(folder.FullName, NomDictionnaire + ".bdic"),
                    false,
                    utf8
                )
            ) {
                for (int i = 0; i < Occurences.Count; i++) {
                    fileWriter.WriteLine($"{(int)StatutsOccurences[i]}{Occurences[i]}");
                }
            }
        }

        #endregion

        #region format JSON

        internal class  MotJson
        {
            // Pour chaque mot, on veut garder un statut par défault, et une liste d'occurence spécifique
            /// <summary>
            /// Statut par défaut pour un mot
            /// </summary>
            public int statut;

            public int nb_occurences;

            /// <summary>
            /// Dictionnaire des occurences (indexé par leur indice dans la liste des occurences globales) avec un statut différent
            /// </summary>
            public Dictionary<string, int> occurences;
        }

        internal class DictionnaireEnJson
        {
            /// <summary>
            /// Version du dictionnaire, pour la compatibilité future
            /// </summary>
            public string version { get; set; }
            /// <summary>
            /// Nom du document associé, sans extension
            /// </summary>
            public string document { get; set; }

            public bool est_terminer { get; set; }

            /// <summary>
            /// Dernier mot sélectionné dans le dictionnaire, pour la reprise de la saisie
            /// </summary>
            public string mot_courant { get; set; }

            /// <summary>
            /// Derniere occurence sélectionné dans le dictionnaire, pour la reprise de la saisie
            /// </summary>
            public int occurence_courante { get; set; }


            /// <summary>
            /// Données sauvegardées pour chaque mot du dictionnaire
            /// </summary>
            public Dictionary<string, MotJson> mots { get; set; }
        }

        /// <summary>
        /// Charger un dictionnaire depuis un fichier JSON
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="canceler"></param>
        /// <returns></returns>
        public static async Task<DictionnaireDeTravail> FromDictionnaryFileJSON(
            string filePath,
            CancellationTokenSource canceler = null
        )
        {
            if (canceler != null && canceler.Token.IsCancellationRequested) {
                return null;
            }
            DictionnaireDeTravail result = new DictionnaireDeTravail(Path.GetFileNameWithoutExtension(filePath));
            Globals.logAsync($"Analyse de {result.NomDictionnaire}");
            Encoding utf8 = Encoding.UTF8;

            using (StreamReader fileReader = new StreamReader(filePath, utf8)) {
                string jsonContent = await fileReader.ReadToEndAsync();
                
                DictionnaireEnJson jsonObject = JsonConvert.DeserializeObject<DictionnaireEnJson>(jsonContent);
                result.NomDictionnaire = jsonObject.document;
                result.DernierMotSelectionne = jsonObject.mot_courant;
                result.DerniereOccurenceSelectionnee = jsonObject.occurence_courante;
                result.EstTerminer = jsonObject.est_terminer;
                foreach (var mot in jsonObject.mots) {
                    string word = mot.Key;
                    MotJson motJson = mot.Value;
                    Statut statut = (Statut)motJson.statut;
                    int nbOccurences = motJson.nb_occurences;
                    Dictionary<string, int> occurences = motJson.occurences;
                    for (int i = 0; i < nbOccurences; i++) {
                        // On ajoute l'occurence avec le statut par défaut
                        result.AjouterOccurence(
                            word,
                            (occurences != null && occurences.Count > 0 && occurences.ContainsKey(i.ToString())) 
                                ? (Statut)occurences[i.ToString()] 
                                : statut,
                            position: -1 // Position dans le contenu indéterminée pour le moment
                        );
                    }
                    if(occurences.Count > 0) {
                        // Garder un statut "Inconnu" au global pour forcer la vérification du mot
                        result.CarteMotStatut[word.ToLower().Trim()] = Statut.INCONNU;
                    }
                }
            }
            result.CalculCartographieMots();
            foreach(var motEtOccurences in result.CarteMotOccurences) {
                string mot = motEtOccurences.Key;
                List<int> occurences = motEtOccurences.Value;
                result.CarteMotStatut[mot] = Statut.INCONNU;
                if (occurences.Count > 0) {
                    // recalcul du statut global en mode souple (on reprend le statut général mais on a quand même les décisions par occurence
                    result.CarteMotStatut[mot] = result.CalculStatutMot(mot);
                }
            }
            return result;
        }



        /// <summary>
        /// Sauvegarder sur le disque en format JSON
        /// </summary>
        /// <param name="path"></param>
        public void SaveJSON(DirectoryInfo folder)
        {
            var jsonObject = new DictionnaireEnJson
            {
                version = "1.0",
                document = Path.GetFileNameWithoutExtension(NomDictionnaire),
                mot_courant = DernierMotSelectionne ?? Occurences.FirstOrDefault(),
                occurence_courante = DerniereOccurenceSelectionnee,
                mots = CarteMotOccurences.ToDictionary(
                    kvp => kvp.Key,
                    kvp =>
                    {
                        // Calculer le statut le plus fréquent
                        var statutFrequent = kvp.Value
                            .GroupBy(index => EstTraitee[index] ? StatutsOccurences[index] : Statut.INCONNU)
                            .OrderByDescending(group => group.Count())
                            .First()
                            .Key;

                        // Créer un dictionnaire des occurences en excluant les occurences avec le statut le plus fréquent
                        var occurences = kvp.Value
                        .Where(index => (EstTraitee[index] ? StatutsOccurences[index] : Statut.INCONNU) != statutFrequent)
                        .ToDictionary(
                            index => kvp.Value.IndexOf(index).ToString(),
                            index => EstTraitee[index] ? (int)StatutsOccurences[index] : (int)Statut.INCONNU
                        );

                        return new MotJson { statut = (int)statutFrequent, nb_occurences = kvp.Value.Count, occurences = occurences };
                    }
                ),
            };
            string jsonContent = JsonConvert.SerializeObject(
                jsonObject,
                Newtonsoft.Json.Formatting.Indented
            );
            string cleanedNomDictionnaire = Path.GetFileNameWithoutExtension(NomDictionnaire);
            File.WriteAllText(
                Path.Combine(folder.FullName, cleanedNomDictionnaire + ".json"),
                jsonContent
            );
        }

        #endregion

        public void RechargerDecisionDe(DictionnaireDeTravail importer)
        {
            switch (importer.Format) {
                case FORMAT.DDIC:
                    // Contient uniquement une cartographie des mots et statuts
                    foreach(var motEtStatut in importer.CarteMotStatut) {
                        string mot = motEtStatut.Key;
                        Statut statut = motEtStatut.Value;
                        CarteMotStatut[mot] = statut;
                        if (CarteMotOccurences.ContainsKey(mot)) {
                            Statut newStat = CarteMotStatut[mot];
                            // On rappatrie le statut sur toutes les occurences du mot
                            foreach (var occurenceMot in CarteMotOccurences[mot]) {
                                // changement de statut au rechargement : occurence a retraité si elle l'etait deja
                                if (newStat != StatutsOccurences[occurenceMot]) {
                                    EstTraitee[occurenceMot] = false;
                                }
                                StatutsOccurences[occurenceMot] = newStat;
                            }
                        }
                    }
                    break;
                case FORMAT.BDIC:
                    // Bdic = version simplifié avec un code par occurence
                    if (importer.NomDictionnaire == this.NomDictionnaire) {
                        // Si même nom de dictionnaire, on recharge
                        // - La carte des statut par mot (version stricte)
                        // - Pour chaque occurence de chaque mot dans l'importer, le statut de l'occurence
                        //DernierMotSelectionne = importer.DernierMotSelectionne;
                        foreach (var motEtOccurences in importer.CarteMotOccurences) {
                            string mot = motEtOccurences.Key;
                            List<int> srcOccurences = motEtOccurences.Value;
                            CarteMotStatut[mot] = importer.CarteMotStatut[mot];
                            if (CarteMotOccurences.ContainsKey(mot)) {
                                // Si statut stricte trouvé, on recharge le statut sur toutes les occurences du mot
                                if (CarteMotStatut[mot] != Statut.INCONNU) {
                                    // On rappatrie le statut sur toutes les occurences du mot
                                    foreach (var occurenceMot in CarteMotOccurences[mot]) {
                                        Statut newStat = CarteMotStatut[mot];
                                        // changement de statut au rechargement : occurence a retraité si elle l'etait deja
                                        if (newStat != StatutsOccurences[occurenceMot]) {
                                            EstTraitee[occurenceMot] = false;
                                        }
                                        StatutsOccurences[occurenceMot] = newStat;
                                    }
                                } else {
                                    // On rappatrie les décisions spécifiques de chaque occurence 
                                    for (int i = 0; i < srcOccurences.Count; i++) {
                                        List<int> target = CarteMotOccurences[mot];
                                        if (i < target.Count) {
                                            Statut newStat = importer.StatutsOccurences[srcOccurences[i]];
                                            // changement de statut au rechargement : occurence a retraité
                                            if (newStat != StatutsOccurences[target[i]]) {
                                                EstTraitee[target[i]] = false;
                                            }
                                            StatutsOccurences[target[i]] = newStat;
                                        } else {
                                            // Plus d'occurence dans le nouveau dictionnaire
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    } else {
                        // Rechargement uniquement des décisions "par défaut" et certaines (pour lesquels il n'y pas de contre décision dans un document)
                        foreach (var temp in importer.CarteMotStatut) {
                            string mot = temp.Key;
                            CarteMotStatut[mot] = temp.Value;
                            if (CarteMotOccurences.ContainsKey(mot)) {
                                if (temp.Value != Statut.INCONNU) {
                                    // On rappatrie le statut sur toutes les occurences du mot
                                    foreach (var occurenceMot in CarteMotOccurences[mot]) {
                                        // changement de statut au rechargement : occurence a retraité
                                        if (temp.Value != StatutsOccurences[occurenceMot]) {
                                            EstTraitee[occurenceMot] = false;
                                        }
                                        StatutsOccurences[occurenceMot] = temp.Value;
                                    }
                                }
                            }
                        }
                    }
                    break;
                case FORMAT.JSON:
                    if (importer.NomDictionnaire == this.NomDictionnaire) {
                        DernierMotSelectionne = importer.DernierMotSelectionne;
                        // Rechargement des décisions sur les mots détectés dans le texte
                        foreach (var motEtOccurences in importer.CarteMotOccurences) {
                            string mot = motEtOccurences.Key;
                            List<int> srcOccurences = motEtOccurences.Value;
                            CarteMotStatut[mot] = importer.CarteMotStatut[mot];
                            if (CarteMotOccurences.ContainsKey(mot)) {
                                if (CarteMotStatut[mot] != Statut.INCONNU) {
                                    // On rappatrie le statut sur toutes les occurences du mot
                                    foreach (var occurenceMot in CarteMotOccurences[mot]) {
                                        Statut newStat = CarteMotStatut[mot];
                                        // changement de statut au rechargement : occurence a retraité
                                        if(newStat != StatutsOccurences[occurenceMot]) {
                                            EstTraitee[occurenceMot] = false;
                                        }
                                        StatutsOccurences[occurenceMot] = newStat;
                                    }
                                } else {
                                    // On rappatrie les décisions spécifiques de chaque occurence 
                                    for (int i = 0; i < srcOccurences.Count; i++) {
                                        List<int> target = CarteMotOccurences[mot];
                                        if (i < target.Count) {
                                            Statut newStat = importer.StatutsOccurences[srcOccurences[i]];
                                            // changement de statut au rechargement : occurence a retraité
                                            if (newStat != StatutsOccurences[target[i]]) {
                                                EstTraitee[target[i]] = false;
                                            }
                                            StatutsOccurences[target[i]] = importer.StatutsOccurences[srcOccurences[i]];
                                        }
                                    }
                                }
                            }
                        }
                    } else {
                        // Rechargement uniquement des décisions "par défaut" et certaines (pour lesquels il n'y pas de contre décision dans un document)
                        foreach (var temp in importer.CarteMotStatut) {
                            string mot = temp.Key;
                            CarteMotStatut[mot] = temp.Value;
                            if (CarteMotOccurences.ContainsKey(mot)) {
                                if (temp.Value != Statut.INCONNU) {
                                    // On rappatrie le statut sur toutes les occurences du mot
                                    foreach (var occurenceMot in CarteMotOccurences[mot]) {
                                        // changement de statut au rechargement : occurence a retraité
                                        if (temp.Value != StatutsOccurences[occurenceMot]) {
                                            EstTraitee[occurenceMot] = false;
                                        }
                                        StatutsOccurences[occurenceMot] = temp.Value;
                                    }
                                }
                            }
                        }
                    }
                    break;
                default:
                    throw new NotSupportedException($"Format de dictionnaire {importer.Format} non supporté.");
            }

        }


        /// <summary>
        /// Appliquer un statut sur une occurence dans le dictionnaire de decision
        /// </summary>
        /// <param name="occurence">indice de l'occurence</param>
        /// <param name="statut">Statut a appliquer</param>
        /// <param name="decalageAvant">Decalage a ajouter a l'occurence</param>
        /// <param name="decalageApres">Decalage supplementaire a ajouter aux occurences suivante (decalageFinale = decalageAvant + decalageApres)</param>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        public void AppliquerStatut(int occurence, Statut statut, int decalageAvant = 0, int decalageApres = 0)
        {
            if (occurence < 0 || occurence >= StatutsOccurences.Count) {
                throw new ArgumentOutOfRangeException(nameof(occurence), "L'index d'occurence est hors des limites de la liste.");
            }
            var previous = EstTraitee[occurence] ? StatutsOccurences[occurence] : Statut.INCONNU;
            StatutsOccurences[occurence] = statut;
            EstTraitee[occurence] = statut != Statut.INCONNU;
            if (decalageAvant != 0 || decalageApres != 0) {
                // Si un offset est spécifié, on l'applique à la position de l'occurence
                PositionsOccurences[occurence] += decalageAvant;
                int decalageOccurencesSuivantes = decalageAvant + decalageApres;
                // Si un offset positif ou negatif (mais pas 0) est spécifié, on décale les occurences suivantes
                if (decalageOccurencesSuivantes != 0) {
                    // On décalle les occurences suivantes
                    for (int i = occurence + 1; i < PositionsOccurences.Count; i++) {
                        PositionsOccurences[i] += decalageOccurencesSuivantes;
                    }
                    // Et les blocs en intéral
                    for (int i = 0; i < DebutsBlocsIntegrals.Count; i++) {
                        if (DebutsBlocsIntegrals.ElementAt(i) > PositionsOccurences[occurence]) {
                            // Si le debut du bloc est après l'occurence, on décale le bloc
                            int oldStart = DebutsBlocsIntegrals.ElementAt(i);
                            DebutsBlocsIntegrals.Remove(oldStart);
                            DebutsBlocsIntegrals.Add(oldStart + decalageOccurencesSuivantes);

                            int oldEnd = FinsBlocsIntegrals.ElementAt(i);
                            if (oldEnd != int.MaxValue) {
                                FinsBlocsIntegrals.Remove(oldEnd);
                                FinsBlocsIntegrals.Add(oldEnd + decalageOccurencesSuivantes);
                            }
                        }
                    }
                }
            }
        }

        public Tuple<int,int> AjouterBlocIntegral(int debut, int fin = int.MaxValue, bool recalculerOffset = false)
        {
            DebutsBlocsIntegrals.Add(debut);
            FinsBlocsIntegrals.Add(fin);
            // Tri des blocs par ordre croissant de début
            
            if (recalculerOffset) {
                for(int i = 0; i < Occurences.Count; i++) {
                    PositionsOccurences[i] +=
                        (PositionsOccurences[i] >= debut ? PROTECTION_CODE_G1.Length : 0) +
                        (PositionsOccurences[i] > fin ? PROTECTION_CODE_G2.Length : 0);
                }
                for (int i = 0; i < DebutsBlocsIntegrals.Count; i++) {
                    int oldStart = DebutsBlocsIntegrals.ElementAt(i);
                    if(oldStart > fin) {
                        DebutsBlocsIntegrals.Remove(oldStart);
                        DebutsBlocsIntegrals.Add(oldStart + PROTECTION_CODE_G1.Length + PROTECTION_CODE_G2.Length);
                    }

                    int oldEnd = FinsBlocsIntegrals.ElementAt(i);
                    if (oldEnd != int.MaxValue && oldEnd > fin) {
                        FinsBlocsIntegrals.Remove(oldEnd);
                        FinsBlocsIntegrals.Add(oldEnd + PROTECTION_CODE_G1.Length + PROTECTION_CODE_G2.Length);
                    }
                }
            }
            return new Tuple<int, int>(debut, fin);
        }
        public Tuple<int, int> SupprimerBlocIntegral(int index, bool recalculerOffset = false)
        {
            if (index < 0 || index >= DebutsBlocsIntegrals.Count) {
                throw new ArgumentOutOfRangeException(nameof(index), "L'index du bloc est hors des limites de la liste.");
            }
            int debut = DebutsBlocsIntegrals.ElementAt(index);
            int fin = FinsBlocsIntegrals.ElementAt(index);
            if (recalculerOffset) {
                for (int i = 0; i < Occurences.Count; i++) {
                    PositionsOccurences[i] -=
                        (PositionsOccurences[i] >= debut ? PROTECTION_CODE_G1.Length : 0) +
                        (PositionsOccurences[i] > fin ? PROTECTION_CODE_G2.Length : 0);
                }
                for (int i = 0; i < DebutsBlocsIntegrals.Count; i++) {
                    if(i != index) {
                        int oldStart = DebutsBlocsIntegrals.ElementAt(i);
                        if (oldStart > fin) {
                            DebutsBlocsIntegrals.Remove(oldStart);
                            DebutsBlocsIntegrals.Add(oldStart - PROTECTION_CODE_G1.Length - PROTECTION_CODE_G2.Length);
                        }

                        int oldEnd = FinsBlocsIntegrals.ElementAt(i);
                        if (oldEnd != int.MaxValue && oldEnd > fin) {
                            FinsBlocsIntegrals.Remove(oldEnd);
                            FinsBlocsIntegrals.Add(oldEnd - PROTECTION_CODE_G1.Length - PROTECTION_CODE_G2.Length);
                        }
                    }
                }
            }
            DebutsBlocsIntegrals.Remove(debut);
            FinsBlocsIntegrals.Remove(fin);
            // Renvoi le tuple de la nouvelle position du bloc déprotégé
            return new Tuple<int, int>(debut - PROTECTION_CODE_G1.Length, fin);

        }


        /// <summary>
        /// Calcul du statut majoritaire d'un mot dans le dictionnaire de travail
        /// </summary>
        /// <param name="mot"></param>
        /// <param name="strict">Ne garder le statut que s'il est unique</param>
        /// <returns></returns>
        public Statut CalculStatutMot(string mot, bool strict = false)
        {
            string key = mot.ToLower().Trim();
            if (CarteMotOccurences.ContainsKey(key)) {
                List<Statut> statuts = CarteMotOccurences[key].Select(i => StatutsOccurences[i]).ToList();
                Dictionary<Statut, int> statutCounts = new Dictionary<Statut, int>();
                foreach (Statut s in statuts) {
                    if (!statutCounts.ContainsKey(s)) {
                        statutCounts[s] = 0;
                    }
                    statutCounts[s]++;
                }
                CarteMotStatut[key] = 
                    statutCounts.Count == 0 ? Statut.IGNORE 
                    : strict ? (statutCounts.Count == 1 ? statutCounts.First().Key : Statut.INCONNU)
                        : statutCounts.OrderByDescending(kvp => kvp.Value).First().Key;
                return CarteMotStatut[key];
            }
            return Statut.IGNORE;
        }

        public void AppliquerStatut(string mot, Statut statut, int offsetBefore = int.MinValue, int offsetAfter = int.MinValue)
        {
            if (CarteMotOccurences.ContainsKey(mot.ToLower().Trim())) {
                foreach (int index in CarteMotOccurences[mot.ToLower().Trim()]) {
                    AppliquerStatut(index, statut, offsetBefore, offsetAfter);
                }
            }
            CarteMotStatut[mot.ToLower().Trim()] = statut;
        }


        public bool Contient(string mot)
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
            clone.EstTraitee = new List<bool>(EstTraitee);
            clone.CarteMotOccurences = new Dictionary<string, List<int>>(CarteMotOccurences);
            clone.CarteMotOccurences = new Dictionary<string, List<int>>(CarteMotOccurences);
            clone.ContextesAvantOccurences = new List<string>(ContextesAvantOccurences);
            clone.ContextesApresOccurences = new List<string>(ContextesApresOccurences);
            clone.DebutsBlocsIntegrals = new SortedSet<int>(DebutsBlocsIntegrals);
            clone.FinsBlocsIntegrals = new SortedSet<int>(FinsBlocsIntegrals);
            return clone;
        }
        /// <summary>
        /// Ajouter une occurence de mot dans le dictionnaire
        /// </summary>
        /// <param name="mot"></param>
        /// <param name="statut"></param>
        /// <param name="contexteAvant"></param>
        /// <param name="contextApres"></param>
        /// <param name="position">-1 pour nouvelle occurence indéterminé, sinon position de l'occurence dans le document</param>
        public void AjouterOccurence(
            string mot,
            Statut statut,
            string contexteAvant = "",
            string contextApres = "",
            int position = -1,
            bool statutDejaAppliquer = false
        )
        {
            lock (this) {
                string wordKey = mot.ToLower().Trim();
                // Si l'occurence n'existe pas déjà dans le dictionnaire
                if (position == -1 || !(
                    CarteMotOccurences.ContainsKey(wordKey)
                    && CarteMotOccurences[wordKey]
                        .FindIndex(
                            occurence => PositionsOccurences[occurence] == position
                        ) >= 0
                )) {
                
                        Occurences.Add(mot);
                        PositionsOccurences.Add(position);
                        StatutsOccurences.Add(statut);
                        EstTraitee.Add(statutDejaAppliquer);
                        ContextesAvantOccurences.Add(contexteAvant);
                        ContextesApresOccurences.Add(contextApres);
                }
            }
        }

        public void AjouterOccurence(Abreviation.OccurenceATraiter toAdd)
        {
            lock (this) {
                string wordKey = toAdd.Mot.ToLower().Trim();
                // Si l'occurence n'existe pas déjà dans le dictionnaire
                if ( !(
                    CarteMotOccurences.ContainsKey(wordKey)
                    && CarteMotOccurences[wordKey]
                        .FindIndex(
                            occurence => PositionsOccurences[occurence] == toAdd.Index
                        ) >= 0
                )) {
                
                    Occurences.Add(toAdd.Mot);
                    PositionsOccurences.Add(toAdd.Index);
                    StatutsOccurences.Add((toAdd.EstDejaProteger || (toAdd.CommenceUnBlocIntegral && toAdd.TermineUnBlocIntegral)) ? Statut.PROTEGE : Statut.INCONNU);
                    EstTraitee.Add((toAdd.EstDejaProteger || (toAdd.CommenceUnBlocIntegral && toAdd.TermineUnBlocIntegral)));
                    ContextesAvantOccurences.Add(toAdd.ContexteAvant);
                    ContextesApresOccurences.Add(toAdd.ContexteApres);
                }
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
            List<bool> newStatutsAppliquer = new List<bool>(EstTraitee.Count);
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
                newStatutsAppliquer.Add(EstTraitee[newIndex]);
                newContextesAvantOccurences.Add(ContextesAvantOccurences[newIndex]);
                newContextesApresOccurences.Add(ContextesApresOccurences[newIndex]);
            }
            // Remplacement des listes précédentes
            Occurences = newOccurences;
            PositionsOccurences = newPositionsOccurences;
            StatutsOccurences = newStatutsOccurences;
            EstTraitee = newStatutsAppliquer;
            ContextesAvantOccurences = newContextesAvantOccurences;
            ContextesApresOccurences = newContextesApresOccurences;
        }

        /// <summary>
        /// Recalcul la carte des mots et de leurs indices dans la liste des occurences
        /// </summary>
        public void CalculCartographieMots()
        {
            CarteMotOccurences.Clear();
            for (int i = 0; i < Occurences.Count; i++)
            {
                string wordKey = Occurences[i].ToLower().Trim();
                if (!CarteMotOccurences.ContainsKey(wordKey))
                {
                    CarteMotOccurences.Add(wordKey, new List<int>());
                }
                CarteMotOccurences[wordKey].Add(i);
            }
        }

        public void SetStatut(string mot, Statut selected)
        {
            string wordKey = mot.ToLower().Trim();
            if (CarteMotOccurences.ContainsKey(wordKey))
            {
                foreach (int index in CarteMotOccurences[wordKey])
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
            string wordKey = mot.ToLower().Trim();
            if (CarteMotOccurences.ContainsKey(wordKey))
            {
                foreach (int index in CarteMotOccurences[wordKey])
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

        
    }
}
