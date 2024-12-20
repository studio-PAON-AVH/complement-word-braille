using System;
using System.Collections.Generic;
using System.Deployment.Application;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;

namespace fr.avh.braille.addin
{
    public static class AddinUpdater
    {
        public static System.Version GetVersion()
        {
            if (ApplicationDeployment.IsNetworkDeployed) {
                return ApplicationDeployment.CurrentDeployment.CurrentVersion;
            }
            return new System.Version(0, 0, 0, 0); // debug mode
        }
        public static ulong computeVersionComparator(string version)
        {
            // version réduite avec seulement 0 a 255 par champs
            Regex versionSeek = new Regex(@"(\d+)(\.(\d+)(\.(\d+)(\.(\d+))?)?)?");
            Match found = versionSeek.Match(version);
            if (found.Success) {
                byte major = Math.Min(byte.Parse(found.Groups[1].Value),(byte)255);
                byte minor = found.Groups.Count >= 3 && found.Groups[3].Value != "" ? Math.Min(byte.Parse(found.Groups[3].Value), (byte)255) : (byte)0;
                byte build = found.Groups.Count >= 5 && found.Groups[5].Value != "" ? Math.Min(byte.Parse(found.Groups[5].Value), (byte)255) : (byte)0;
                byte patch = found.Groups.Count >= 7 && found.Groups[7].Value != "" ? Math.Min(byte.Parse(found.Groups[7].Value), (byte)255) : (byte)0;
                
                return (ulong)((major << 24) + (minor << 16) + (build << 8) + patch);
            }
            return 0;
        }
        public static void CheckForUpdate()
        {
            System.Version version = GetVersion();
            DateTime current = DateTime.Now;
            TimeSpan lastCheck = current - OptionsComplement.Instance.LastUpdateCheck;
            if (version != new System.Version(0, 0, 0, 0) && lastCheck > TimeSpan.FromDays(1)) {
                OptionsComplement.Instance.LastUpdateCheck = current;
                OptionsComplement.Instance.Sauvegarder();
                // Test de disponisibilité d'une mise à jour
                try {
                    string sourceURL = Properties.Settings.Default.UpdateUrl;
                    // from https://stackoverflow.com/questions/10822509/the-request-was-aborted-could-not-create-ssl-tls-secure-channel

                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
                        | SecurityProtocolType.Tls11
                        | SecurityProtocolType.Tls12
                        | SecurityProtocolType.Ssl3;
                    // allows for validation of SSL conversations
                    ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(sourceURL);
                    request.AllowAutoRedirect = true;
                    request.Proxy.Credentials = CredentialCache.DefaultCredentials;
                    HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                    string lastTag = response.ResponseUri.ToString().Split('/').Last();
                    response.Close();
                    if (computeVersionComparator(lastTag) > computeVersionComparator(version.ToString())) {
                        MessageBoxResult dr = System.Windows.MessageBox.Show(
                            "Une nouvelle version du complément braille est disponible au téléchargement.\r\nVoulez la télécharger ?\r\n" +
                                "(Vous allez être redirigez vers la page de téléchargement)",
                            "Nouvelle version du complément braille",
                            MessageBoxButton.OKCancel,
                            MessageBoxImage.Information
                        );
                        if (dr == MessageBoxResult.OK)
                            System.Diagnostics.Process.Start(Properties.Settings.Default.UpdateUrl);
                    } else {
                        //MessageBox.Show("Your already have the latest version of the plugin.", "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception e) {
                    // Pas de vérification de mise à jour
                    dictionnaire.Globals.logAsync("Impossible de récupérer les mises à jours du complément word");
                    dictionnaire.Globals.logAsync(e);
                }

            }
        }
    }
}
