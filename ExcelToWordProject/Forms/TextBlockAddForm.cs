using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelToWordProject.Syllabus.Tags;

namespace ExcelToWordProject.Forms
{
    public partial class TextBlockAddForm : Form
    {
        private readonly TextBlockTag _tag;
        private readonly Form _parent;
        public TextBlockAddForm(string tagName, Form parent = null)
        {
            _tag = new TextBlockTag(tagName, Array.Empty<TextBlockCondition>(), "");
            _parent = parent;
            Initialize();
        }

        public TextBlockAddForm(TextBlockTag tag, Form parent = null)
        {
            _tag = tag;
            _parent = parent;
            Initialize();
        }

        private void Initialize()
        {
            InitializeComponent();
            FillComponents();
        }

        private void FillComponents()
        {
            tagKeyTextBox.Text = _tag.Key;

            if (_tag.HasId)
                tagValueTextBox.Text = _tag.GetValue2();
            foreach (var condition in _tag.Conditions)
            {
                conditionsGridView.Rows.Add(condition.TagName, condition.Condition);
            }

        }

        private void SaveTag()
        {
            _tag.Key = tagKeyTextBox.Text;
            var conditions = new List<TextBlockCondition>();
            foreach (DataGridViewRow row in conditionsGridView.Rows)
            {
                var tagName = row.Cells["TagNameColumn"].Value?.ToString();
                var tagCondition = row.Cells["TagValueColumn"].Value?.ToString();
                if(string.IsNullOrEmpty(tagName))
                    continue;
                conditions.Add(new TextBlockCondition(tagName, tagCondition));
            }

            _tag.Conditions = conditions.ToArray();

            if (!_tag.HasDefaultValue)
            {
                TextBlockTag.SetDefaultValue(_tag.Key, "");
                if (_tag.IsDefault)
                {
                    return;
                }
                MessageBox.Show(@"Значение по умолчанию было создано автоматически", @"Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            if (_tag.CanStoreInDataBase)
            {
                _tag.SaveToDatabase(tagValueTextBox.Text); // TODO: unique error
            }
            else
            {
                MessageBox.Show(@"Данный набор условий уже существует", @"Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }




        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveTag();
            Dispose();
            _parent?.Refresh();
        }
    }

}
