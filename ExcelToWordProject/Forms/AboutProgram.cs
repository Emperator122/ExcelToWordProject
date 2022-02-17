using System.Windows.Forms;

namespace ExcelToWordProject.Forms
{
    public partial class AboutProgramForm : Form
    {
        public AboutProgramForm()
        {
            InitializeComponent();
            versionLabel.Text = "Версия: " + Application.ProductVersion;
        }
    }
}
