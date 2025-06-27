using System;
using System.IO;
using System.Xml;

namespace fr.avh.braille.addin
{
    /// <summary>
    /// Singleton de gestion des options du compléments
    /// </summary>
    public sealed class OptionsComplement
    {
        private static readonly string CONFIG = Path.Combine(
            fr.avh.braille.dictionnaire.Globals.AppData.FullName,
            "config_addin.xml"
            );

        private static readonly Lazy<OptionsComplement> lazy = new Lazy<OptionsComplement>(() => new OptionsComplement());

        public static OptionsComplement Instance => lazy.Value;

        private OptionsComplement()
        {
            if (File.Exists(CONFIG)) {
                // Charger les options
                XmlDocument config = new XmlDocument();
                config.Load(CONFIG);
                XmlNode preprotec = config.GetElementsByTagName("preprotectionauto").Item(0);
                if (preprotec != null) {
                    ActiverPreProtectionAuto = bool.Parse(preprotec.InnerText);
                }
                try {
                    XmlNode lastUpdateCheck = config.GetElementsByTagName("lastupdatecheck").Item(0);
                    if (lastUpdateCheck != null) {
                        LastUpdateCheck = DateTime.Parse(preprotec.InnerText);
                    }
                } catch (Exception e) {
                    //ActiverPreProtectionAuto = false;
                }
                
            }
        }

        public void Sauvegarder()
        {
            XmlDocument config = new XmlDocument();
            XmlElement root = config.CreateElement("config");
            config.AppendChild(root);

            XmlElement preprotectionauto = config.CreateElement("preprotectionauto");
            preprotectionauto.InnerText = ActiverPreProtectionAuto.ToString();
            root.AppendChild(preprotectionauto);

            XmlElement lastUpdateCheck = config.CreateElement("lastupdatecheck");
            lastUpdateCheck.InnerText = lastUpdateCheck.ToString();
            root.AppendChild(lastUpdateCheck);

            config.Save(CONFIG);
        }

        /// <summary>
        /// Option d'activation de la pre protection automatique
        /// </summary>
        public bool ActiverPreProtectionAuto { get; set; } = false;
        public DateTime LastUpdateCheck { get; set; } = DateTime.MinValue;
    }


}
