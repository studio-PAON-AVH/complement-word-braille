using FluentNHibernate.Cfg.Db;
using FluentNHibernate.Cfg;
using NHibernate;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NHibernate.Tool.hbm2ddl;
using fr.avh.braille.dictionnaire.Entities;
using FluentNHibernate.Conventions;
using fr.avh.archivage;
using System.Threading;
using static fr.avh.archivage.Utils;

namespace fr.avh.braille.dictionnaire
{
    /// <summary>
    /// Dictionnaire SQLite de stockage des mots et de leur status dans les dictionnaires
    /// </summary>
    public class BaseSQlite
    {
        public static string dbFilename = "protection.db";

        private BaseSQlite() { }

        public static bool dbExists()
        {
            return File.Exists(Path.Combine(Globals.AppData.FullName, dbFilename));
        }

        public static void Redeploy()
        {
            if(dbExists()) {
                try {
                    File.Delete(Path.Combine(Globals.AppData.FullName, dbFilename));
                }
                catch (Exception e) {
                    //return;
                }
            }
            lock (dbFilename) {
                try {
                    byte[] rawDb = Properties.Resources.protection;
                    File.WriteAllBytes(Path.Combine(Globals.AppData.FullName, dbFilename), rawDb);
                    //onInfo?.Invoke("La base de donnée locale est installé");
                }
                catch (Exception e1) {
                    //onInfo?.Invoke($"Impossible d'installer la base de donnée locale: {e1.Message}");
                    //onInfo?.Invoke("Essai d'installation depuis le dépot distant");
                }
            }
        }

        public static void CheckForUpdates(
            OnInfoCallback onInfo = null
        )
        {
            onInfo?.Invoke("Recherche de mise à jour de la base de donnée de protection...");
            if (!dbExists()) {
                onInfo?.Invoke("Base non installé, installation de la base local...");
                lock (dbFilename) {
                    try {
                        byte[] rawDb = Properties.Resources.protection;
                        File.WriteAllBytes(Path.Combine(Globals.AppData.FullName, dbFilename), rawDb);
                        onInfo?.Invoke("La base de donnée locale est installé");
                    } catch (Exception e1) {
                        onInfo?.Invoke($"Impossible d'installer la base de donnée locale: {e1.Message}");
                        onInfo?.Invoke("Essai d'installation depuis le dépot distant");
                    }
                }
            }
            try {
                onInfo?.Invoke("Récupération des mises à jours ...");
                Depot.update(onInfo);
                FileInfo remoteDb = Depot.getProtectionDB();
                if (remoteDb != null && remoteDb.Exists) {
                    remoteDb.CopyTo(Path.Combine(Globals.AppData.FullName, dbFilename));
                }
                onInfo?.Invoke("Base de donnée mise à jour");
            } catch (Exception e2) {
                onInfo?.Invoke($"Impossible de déployer la base de protection distante : {e2.Message}");
                if (!dbExists()) {
                    onInfo?.Invoke("Les informations de protection ne seront pas accessibles lors du traitement.");
                }
                //onError?.Invoke(ex);
            }
            
        }

        // For hibernate
        public static ISessionFactory CreateSessionFactory(
            OnInfoCallback onInfo = null,
            OnErrorCallback onError = null
        )
        {
            if (!Directory.Exists(Globals.AppData.FullName))
            {
                Directory.CreateDirectory(Globals.AppData.FullName);
            }
            
            if (!dbExists())
            {
                CheckForUpdates(onInfo);
            }

            FluentConfiguration config = Fluently
                .Configure()
                .Database(
                    SQLiteConfiguration.Standard.UsingFile(
                        Path.Combine(Globals.AppData.FullName, dbFilename)
                    )
                )
                .Mappings(m =>
                {
                    m.FluentMappings.AddFromAssemblyOf<BaseSQlite>();
                });
            if (!dbExists())
            {
                config = config.ExposeConfiguration(
                    (c) =>
                    {
                        SchemaExport schema = new SchemaExport(c);
                        if (!dbExists())
                        {
                            schema.Create(false, true);
                        }
                        else { }
                        // this NHibernate tool takes a configuration (with mapping info in)
                        // and exports a database schema from it
                    }
                );
            }
            return config.BuildSessionFactory();
        }

        /// <summary>
        /// Pour une liste de mot données, extrait les données de la base de dictionnaires
        /// </summary>
        /// <param name="texte"></param>
        /// <returns></returns>
        public static List<Mot> getWordsData(List<string> texte)
        {
            using (var dbSession = CreateSessionFactory().OpenSession())
            {
                return dbSession
                    .Query<Mot>()
                    .Where(m => texte.Any(t => t.ToLower() == m.Texte))
                    .ToList();
            }
        }

        public static void SetOrUpdateWord(Mot mot)
        {
            using (var dbSession = CreateSessionFactory().OpenSession())
            {
                dbSession.SaveOrUpdate(mot);
            }
        }

        /// <summary>
        /// Not implemented yet :
        /// un essai de structure acceleratrice pour l'analyse de dictionnaires
        /// avant fusion avec la base de données
        /// </summary>
        /// <param name="memory"></param>
        public static void AddFrom(AnalyseurDeDictionnaires memory)
        {
            // merge the memory with the database
            using (var dbSession = CreateSessionFactory().OpenSession())
            {
                IQueryable<Mot> motsQuery = dbSession.Query<Mot>();
                using (var transaction = dbSession.BeginTransaction())
                {
                    // isolé les nouveaux mots
                    // pour chaque mot qui n'est pas dans la base de données, l'ajouter
                    List<string> textesInDB = motsQuery.Select(d => d.Texte).ToList();
                    foreach (var mot in memory.mots)
                    {
                        if (!textesInDB.Contains(mot.Key))
                        {
                            dbSession.Save(mot.Value);
                        }
                    }
                    transaction.Commit();
                }
            }
        }

        public static void updateFromCSV(
            string csvPath,
            OnInfoCallback info = null,
            CancellationTokenSource canceler = null
        )
        {
            info?.Invoke($"Analyse du fichier {csvPath}...");
            using (var reader = new StreamReader(csvPath))
            {
                var headers = reader.ReadLine().Split(';').ToList();
                int indexSelector = -1;
                for (int i = 0; i < headers.Count && indexSelector < 0; i++)
                {
                    if (
                        headers[i].ToLower().StartsWith("mot")
                        || headers[i].ToLower().StartsWith("texte")
                    )
                    {
                        indexSelector = i;
                    }
                }
                if (indexSelector < 0)
                {
                    throw new Exception(
                        "Aucune colonne ne permet d'identifier un mot dans la base de données"
                    );
                }
                Dictionary<string, Mot> mots = new Dictionary<string, Mot>();
                using (var dbSession = CreateSessionFactory().OpenSession())
                {
                    mots = dbSession.Query<Mot>().ToDictionary(m => m.Texte);
                }
                List<Mot> toUpdateOrInsert = new List<Mot>();
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(';').Select(s => s.Trim()).ToList();
                    List<string> columns = new List<string>();
                    Mot mot;
                    if (mots.ContainsKey(values[indexSelector].ToLower().Trim()))
                    {
                        mot = mots[values[indexSelector].ToLower().Trim()];
                    }
                    else
                    {
                        mot = new Mot()
                        {
                            Texte = values[indexSelector].ToLower().Trim(),
                            DateAjout = DateTime.Now.ToString(),
                            Abreviations = 0,
                            Protections = 0,
                            Documents = 0,
                            ToujoursDemander = 0,
                            Commentaires = ""
                        };
                    }
                    for (int i = 0; i < values.Count; i++)
                    {
                        if (i == indexSelector)
                        {
                            continue;
                        }
                        switch (headers[i].ToLower())
                        {
                            case "protections":
                                mot.Protections = long.Parse(values[i]);
                                break;
                            case "abreviations":
                                mot.Abreviations = long.Parse(values[i]);
                                break;
                            case "documents":
                                mot.Documents = long.Parse(values[i]);
                                break;
                            case "toujoursdemander":
                                if (
                                    values[i].ToLower() == "vrai"
                                    || values[i].ToLower() == "true"
                                    || values[i] == "1"
                                )
                                {
                                    mot.ToujoursDemander = 1;
                                }
                                else if (
                                    values[i].ToLower() == "faux" || values[i].ToLower() == "false"
                                )
                                {
                                    mot.ToujoursDemander = 0;
                                }
                                else
                                {
                                    mot.ToujoursDemander = int.Parse(values[i]);
                                }
                                mot.ToujoursDemander = int.Parse(values[i]);
                                break;
                            case "commentaires":
                                info?.Invoke(
                                    $"Nouveau commentaire pour {mot.Texte} : {values[i]}..."
                                );
                                mot.Commentaires = values[i];
                                break;
                            default:
                                break;
                        }
                    }
                    toUpdateOrInsert.Add(mot);
                }
                using (var dbSession = CreateSessionFactory().OpenSession())
                {
                    using (var transaction = dbSession.BeginTransaction())
                    {
                        foreach (var mot in toUpdateOrInsert)
                        {
                            dbSession.SaveOrUpdate(mot);
                        }
                        transaction.Commit();
                    }
                }
            }
        }

        public static void AnalyserDictionnaires(
            List<string> dictionnaires,
            Utils.OnInfoCallback info = null,
            CancellationTokenSource canceler = null
        )
        {
            using (var dbSession = CreateSessionFactory().OpenSession())
            {
                List<string> dictionnairesInDb = dbSession
                    .Query<Dictionnaire>()
                    .Select(d => d.Nom)
                    .ToList();
                List<Task<AncienDictionnaireDeTravail>> tasks = new List<Task<AncienDictionnaireDeTravail>>();
                for (int s = 0; s < dictionnaires.Count; s++)
                {
                    string d = dictionnaires[s];
                    if (canceler != null && canceler.Token.IsCancellationRequested)
                    {
                        return;
                    }
                    if (dictionnairesInDb.Contains(Path.GetFileName(d)))
                    {
                        info?.Invoke(
                            $"Le dictionnaire {Path.GetFileName(d)} est déjà dans la base de données"
                        );
                        continue;
                    }
                    info?.Invoke($"Ajout du dictionnaire {Path.GetFileName(d)}");
                    tasks.Add(AncienDictionnaireDeTravail.FromDictionnaryFile(d, canceler));
                }
                int p = 0;
                int max = tasks.Count;
                info?.Invoke(
                    $"{tasks.Count} Dictionnaires a analyser",
                    new Tuple<int, int>(p, max)
                );
                while (tasks.Count > 0)
                {
                    if (canceler != null && canceler.Token.IsCancellationRequested)
                    {
                        return;
                    }
                    int d = Task.WaitAny(tasks.ToArray());
                    AncienDictionnaireDeTravail result = tasks[d].Result;
                    info?.Invoke(
                        $"Ajout de {result.NomDictionnaire} : {result.Mots.Count} mots ..."
                    );
                    List<string> keys = result.Mots.Keys.ToList();
                    IQueryable<Mot> motsQuery = dbSession.Query<Mot>();
                    List<Mot> alreadyIn = motsQuery.Where(m => keys.Contains(m.Texte)).ToList();

                    List<Mot> notIn = result.Mots
                        .Where(kv => !alreadyIn.Select(m => m.Texte).Contains(kv.Key))
                        .Select(
                            kv =>
                                new Mot()
                                {
                                    Texte = kv.Key,
                                    DateAjout = DateTime.Now.ToString(),
                                    Abreviations = kv.Value == Statut.ABREGE ? 1 : 0,
                                    Protections = kv.Value == Statut.PROTEGE ? 1 : 0,
                                    Documents = 1,
                                    ToujoursDemander = 0,
                                    Commentaires = ""
                                }
                        )
                        .ToList();

                    using (var transaction = dbSession.BeginTransaction())
                    {
                        info?.Invoke($"{notIn.Count} nouveaux mots ...");
                        foreach (var mot in notIn)
                        {
                            dbSession.Save(mot);
                        }
                        transaction.Commit();
                    }
                    using (var transaction = dbSession.BeginTransaction())
                    {
                        info?.Invoke($"{alreadyIn.Count} mots mis à jour ...");
                        foreach (var mot in alreadyIn)
                        {
                            mot.Abreviations += result.Mots[mot.Texte] == Statut.ABREGE ? 1 : 0;
                            mot.Protections += (result.Mots[mot.Texte] == Statut.PROTEGE ? 1 : 0);
                            mot.Documents += 1;

                            dbSession.Update(mot);
                        }
                        transaction.Commit();
                    }
                    try
                    {
                        using (var transaction = dbSession.BeginTransaction())
                        {
                            info?.Invoke($"{result.NomDictionnaire} ajoutés a la base ...");
                            dbSession.Save(
                                new Dictionnaire()
                                {
                                    Nom = result.NomDictionnaire,
                                    DateAjout = DateTime.Now.ToString()
                                }
                            );
                            transaction.Commit();
                        }
                    }
                    catch (Exception e)
                    {
                        info?.Invoke(
                            $"Erreur lors de l'ajout de {result.NomDictionnaire} dans la base : {e.Message}"
                        );
                    }
                    p++;
                    info?.Invoke(
                        $"Analyse de {result.NomDictionnaire} terminer",
                        new Tuple<int, int>(p, max)
                    );
                    tasks.RemoveAt(d);
                }
            }
        }

        public static AnalyseurDeDictionnaires Load()
        {
            AnalyseurDeDictionnaires result = new AnalyseurDeDictionnaires();
            using (var dbSession = CreateSessionFactory().OpenSession())
            {
                dbSession.Query<Mot>().ToList().ForEach(d => result.mots.Add(d.Texte, d));
            }
            return result;
        }
    }
}
