using LibGit2Sharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using static fr.avh.archivage.Utils;

namespace fr.avh.braille.dictionnaire
{
    /// <summary>
    /// Classe de gestion du dépot distant hébergeant la base de donnée des informations de protections.
    /// Le dépot conserve également les fichiers dictionnaires ayant servis a la construction de la base,
    /// ainsi qu'un fichier CSV contenant les informations complémentaires par mot ne pouvant être
    /// récupéré à partir des dictionnaires (i.e. les commentaires et l'indicateur de demande systématique)
    /// 
    /// TODO: La gestion par git n'est pas fonctionnelle (probleme de déploiement de libgit2sharp et logique de synchro a voir)
    /// A voir si nous trouvons un fournisseur de base de donnée sqlite (Turso peut etre) pour gérer la base distante
    /// Ou si nous mettons en place notre propre serveur SQLite.
    /// 
    /// Pour le moment, les mises à jour de la base de donnée sont faites avec les mises à jour du complément.
    /// </summary>
    public static class Depot
    {
        static string repositoryUrl = "https://gitlab.com/studio-paon-avh/protection_db.git";
        static string repositoryClone = Path.Combine(Globals.AppData.FullName, "protection_db.git");
        static Signature appSignature = new Signature("protection_db", "protection_db@avh.asso.fr", DateTimeOffset.Now);

        /// <summary>
        /// Récupération des mises à jour du dépot distant
        /// </summary>
        /// <param name="onInfo"></param>
        public static void update(OnInfoCallback onInfo = null)
        {
            // Note NP 2024/05/03 : 
            // Probleme déploiement libgit2sharp... l'outil ne déploie pas les dlls de git a la publication
            
            // Note NP 2024/04/26 :
            // certains antivirus peuvent interférer avec la résolution des certificats ssl
            // (i.e. ESET a un SSL filter qui bypass ça)
            // Hack pour désactiver le controle SSL et eviter les SSL filter qui bloque la connexion a gitlab.com...
            // Addendum : le certificat est bien valide dans les structures privé sous-jacente
            // Le probleme semle venir du validateur de LibGit2Sharp qui le considère invalide quand meme
            FetchOptions fetchOptions = new FetchOptions();
            fetchOptions.CertificateCheck = (Certificate certificate, bool valid, string host) =>
            {
                // Essai de correction du probleme sans supprimer la chaine de certification
                //try {
                //    CertificateX509 test = (CertificateX509)certificate;
                //    X509Certificate2 x509Certificate2 = new X509Certificate2(((CertificateX509)certificate).Certificate);
                // Le probleme est la : Verify renvoi false alors que le meme certificat, si je test dans chrome, est considéré valide
                // Il semblerait que la chaine de verification n'arrive pas a établir la validation d'une autorité
                // Il faudra que je check moi meme la chaine de certification ou que je la customize ... oh god ...
                // https://stackoverflow.com/questions/10137208/x509certificate2-verify-method-always-return-false-for-the-valid-certificate
                //    bool testValid = x509Certificate2.Verify();
                //}
                //catch (Exception) {
                //    return valid;
                //}
                return true;
            };
            // clone or pull updates from the remote git repository
            if (!Directory.Exists(repositoryClone)) {
                onInfo?.Invoke($"Récupération du dépot {repositoryUrl} dans {repositoryClone}");
                CloneOptions options = new CloneOptions();
                options.FetchOptions.CertificateCheck = fetchOptions.CertificateCheck;
                string result = Repository.Clone(repositoryUrl, repositoryClone, options) ;
            }
            Repository repo = new Repository(repositoryClone);
            Branch main = repo.Head.TrackedBranch;
            if (main == null) {
                main = repo.Branches["main"];
            };
            onInfo?.Invoke($"Récupération des mises à jours de {repositoryUrl} ...");
            MergeResult res = Commands.Pull(repo, appSignature, new PullOptions()
            {
                MergeOptions = new MergeOptions()
                {
                    FastForwardStrategy = FastForwardStrategy.FastForwardOnly,
                    FileConflictStrategy = CheckoutFileConflictStrategy.Theirs
                },
                FetchOptions = fetchOptions
            });
            // TODO : gestion des conflits entre les dépots distants et local lors de la synchro
            switch(res.Status) {
                case MergeStatus.UpToDate:
                case MergeStatus.FastForward:
                case MergeStatus.NonFastForward:
                case MergeStatus.Conflicts:
                    break;
            }
        }

        /// <summary>
        /// Récupération du chemin du fichier de la base de donnée dans le dépot cloné
        /// </summary>
        /// <returns></returns>
        public static FileInfo getProtectionDB()
        {
            return new FileInfo(Path.Combine(repositoryClone, BaseSQlite.dbFilename));
        }

        /// <summary>
        /// Liste des fichiers dictionnaires stockées dans le dépot
        /// </summary>
        /// <returns></returns>
        public static List<string> getDictionnaries()
        {
            update();
            return Directory.GetFiles(Path.Combine(repositoryClone, "dictionnaires"))
                .Where(f => f.EndsWith(".dic") || f.EndsWith(".ddic"))
                .ToList();
        }

    }
}
