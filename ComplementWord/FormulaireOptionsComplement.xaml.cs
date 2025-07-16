
using System;
using System.IO;
using System.Windows;
using System.Xml;

namespace fr.avh.braille.addin
{
    

    /// <summary>
    /// Logique d'interaction pour OptionsComplement.xaml
    /// </summary>
    public partial class FormulaireOptionsComplement : Window
    {
        public FormulaireOptionsComplement()
        {
            InitializeComponent();
            ActiverPreProtectionAuto.IsChecked = OptionsComplement.Instance.ActiverPreProtectionStatistique;
        }

        private void FermerFenetre_Click(object sender, RoutedEventArgs e)
        {
            OptionsComplement.Instance.Sauvegarder();
            Close();
        }

        private void ActiverPreProtectionAuto_Checked(object sender, RoutedEventArgs e)
        {
            OptionsComplement.Instance.ActiverPreProtectionStatistique = true;
            OptionsComplement.Instance.Sauvegarder();
        }

        private void ActiverPreProtectionAuto_Unchecked(object sender, RoutedEventArgs e)
        {
            OptionsComplement.Instance.ActiverPreProtectionStatistique = false;
            OptionsComplement.Instance.Sauvegarder();
        }
    }
}
