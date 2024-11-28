using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using fr.avh.archivage;
using fr.avh.braille.dictionnaire.Entities;

namespace fr.avh.braille.dictionnaire
{
    
    /// <summary>
    /// Analyseur de dictionnaires récupérants de tous les mots et de leurs occurences 
    /// et status dans la liste des dictionnaires
    /// Version en mémoire de la base de donnée des dictionnaires pour simplifier l'analyse
    /// </summary>
    public class AnalyseurDeDictionnaires
    {

        /// <summary>
        /// Liste des mots trouvés dans les dictionnaires indexés par leur texte
        /// </summary>
        public Dictionary<string, Mot> mots;
        

        public AnalyseurDeDictionnaires()
        {
            mots = new Dictionary<string, Mot>();
        }


        public void AnalyserDictionnaires(
            List<string> dictionnaires,
            bool reanalyser = false,
            Utils.OnInfoCallback info = null,
            CancellationTokenSource canceler = null)
        {
            
            List<Task<AncienDictionnaireDeTravail>> tasks = new List<Task<AncienDictionnaireDeTravail>>();
            for(int s = 0; s < dictionnaires.Count; s++ ){

                string d = dictionnaires[s];
                tasks.Add(AncienDictionnaireDeTravail.FromDictionnaryFile(d, canceler));
                info?.Invoke($"Analyse du dictionnaire {Path.GetFileNameWithoutExtension(d)}");
                if (canceler != null && canceler.Token.IsCancellationRequested) {
                    return;
                }
            }
            int p = 0;
            int max = tasks.Count;
            info?.Invoke($"{tasks.Count} Dictionnaires a analyser", new Tuple<int, int>(p, max));
            while (tasks.Count > 0) {
                if (canceler != null && canceler.Token.IsCancellationRequested) {
                    return;
                }
                int d = Task.WaitAny(tasks.ToArray());
                AncienDictionnaireDeTravail result = tasks[d].Result;

                info?.Invoke($"Analyse de {result.NomDictionnaire} : {result.Mots.Count} mots ...");
                foreach(var _mot in result.Mots) {
                    if (canceler != null && canceler.Token.IsCancellationRequested) {
                        return;
                    }
                    Mot mot;
                    if (!mots.ContainsKey(_mot.Key)) {
                        mot = new Mot()
                        {
                            Texte = _mot.Key,
                            DateAjout = DateTime.Now.ToString(),
                            Abreviations = 0,
                            Protections = 0,
                            ToujoursDemander = 0,
                            Documents = 0,
                            Commentaires = ""
                        };
                    } else {
                        mot = mots[_mot.Key];
                    }
                    mot.Documents += 1;
                    switch (_mot.Value) {
                        case Statut.ABREGE:
                            mot.Abreviations += 1;
                            break;
                        case Statut.PROTEGE:
                            mot.Protections += 1;
                            break;
                    }
                    mots[_mot.Key] = mot;
                }
                p++;
                info?.Invoke($"Analyse de {result.NomDictionnaire} terminer", new Tuple<int, int>(p, max));
                tasks.RemoveAt(d);
            }
        }
    }
}
