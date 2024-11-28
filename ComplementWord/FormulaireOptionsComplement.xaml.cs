
using System;
using System.IO;
using System.Windows;
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

            config.Save(CONFIG);
        }

        /// <summary>
        /// Option d'activation de la pre protection automatique
        /// </summary>
        public bool ActiverPreProtectionAuto { get; set; } = false;
    }

    

    /// <summary>
    /// Logique d'interaction pour OptionsComplement.xaml
    /// </summary>
    public partial class FormulaireOptionsComplement : Window
    {
        public FormulaireOptionsComplement()
        {
            InitializeComponent();
            ActiverPreProtectionAuto.IsChecked = OptionsComplement.Instance.ActiverPreProtectionAuto;
        }

        private void FermerFenetre_Click(object sender, RoutedEventArgs e)
        {
            OptionsComplement.Instance.Sauvegarder();
            Close();
        }

        private void ActiverPreProtectionAuto_Checked(object sender, RoutedEventArgs e)
        {
            OptionsComplement.Instance.ActiverPreProtectionAuto = true;
            OptionsComplement.Instance.Sauvegarder();
        }

        private void ActiverPreProtectionAuto_Unchecked(object sender, RoutedEventArgs e)
        {
            OptionsComplement.Instance.ActiverPreProtectionAuto = false;
            OptionsComplement.Instance.Sauvegarder();
        }
    }
}
