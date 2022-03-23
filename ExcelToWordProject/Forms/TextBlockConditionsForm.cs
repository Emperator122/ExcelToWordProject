using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using ExcelToWordProject.Syllabus;
using ExcelToWordProject.Syllabus.Tags;

namespace ExcelToWordProject.Forms
{
    public partial class TextBlockConditionsForm : Form
    {
        private readonly string _tagName;
        private List<TextBlockTag> _tags;
        private readonly SyllabusParameters _parameters;
        private readonly TextBlocksForm _parentForm;

        public TextBlockConditionsForm(string tagName, SyllabusParameters parameters, TextBlocksForm parentForm = null)
        {
            _tagName = tagName;
            _parameters = parameters;
            _parentForm = parentForm;
            InitializeComponent();
            FetchTags();
            InitializeCustomComponents();
        }

        private void FetchTags()
        {
            _tags = _parameters.TextBlockTags
                .Where(tag => tag.Key == _tagName)
                .OrderBy(tag => tag.Conditions.Length)
                .ToList();

            if (_tags.First()?.Conditions.Length > 0)
            {
                // TODO: реализовать обработку отсутствия дефолтного значения
            }
        }

        private void TextBlockConditionsForm_Load(object sender, EventArgs e)
        {

        }

        private void TextBlockConditionsForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            _parentForm.Refresh();
            _parentForm?.Show();
            Dispose();
        }

        private void addNewConditionButton_Click(object sender, EventArgs e)
        {
            var form = new TextBlockAddForm(_tagName, this);
            form.ShowDialog();
        }

        public override void Refresh()
        {
            base.Refresh();
            textBlocksWrapper.Controls.Clear();
            FetchTags();
            InitializeCustomComponents();
        }
        private void TagRemove_Click(object sender)
        {
            if ((sender as PictureBox)?.Parent?.Tag is TextBlockTag tag)
            {
                tag.RemoveFromDatabase();
                if (_tags.Count == 1)
                {
                    Close();
                }
                else
                {
                    Refresh();
                }
            }
        }

        private void TagEdit_Click(object sender)
        {
            if ((sender as PictureBox)?.Parent?.Tag is TextBlockTag tag)
            {
                var form = new TextBlockAddForm(tag, this);
                form.ShowDialog();
            }
        }


    }
}
