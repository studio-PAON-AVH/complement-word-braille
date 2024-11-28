using System.Windows;
using System.Windows.Forms;

namespace fr.avh.braille.addin
{
    /// <summary>
    /// Logique d'interaction pour ProgressDialog.xaml
    /// </summary>
    public partial class ProgressDialog : Window
    {

        public void SetProgress(int progress, int maxValue)
        {
            ProgressIndicator.Value = progress;
            ProgressIndicator.Maximum = maxValue;
        }

        public void AddMessage(string message)
        {
            ProgressMessages.AppendText(message + "\r\n");
            ProgressMessages.ScrollToEnd();
        }
        
        public ProgressDialog()
        {
            InitializeComponent();
        }
    }
}
