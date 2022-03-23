using System;
using System.Windows.Forms;
using ExcelToWordProject.Syllabus;
using ExcelToWordProject.Syllabus.Tags;

namespace ExcelToWordProject.Forms
{
    public partial class TextBlocksForm : Form
    {
        private readonly SyllabusParameters _parameters;

        public TextBlocksForm(SyllabusParameters parameters)
        {
            _parameters = parameters;
            InitializeComponent();
            InittializeCustomComponents();
        }

        private void TextBlocksForm_Load(object sender, EventArgs e)
        {

        }

        private void OnGoToConditionsButtonClick(string tagName)
        {
            var form = new TextBlockConditionsForm(tagName, _parameters, this);
            form.Show();
            Hide();
        }

        private void AddNewTagButton_Click(object sender, EventArgs e)
        {
            var form = new TextBlockAddForm("", this);
            form.ShowDialog();
        }

        private void RemoveTag_Click(object sender)
        {
            if ((((sender as MenuItem)?.Parent as ContextMenu)?.SourceControl as PictureBox)?.Parent?.Tag is string tagKey)
            {
                TextBlockTag.RemoveTagFromDatabase(tagKey);
                Refresh();
            }
        }

        public override void Refresh()
        {
            base.Refresh();
            Controls.Clear();
            InitializeComponent();
            InittializeCustomComponents();
        }
    }
}
