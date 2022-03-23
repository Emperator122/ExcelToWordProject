using System;
using System.Collections.Generic;
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
            _tag = new TextBlockTag(tagName, Array.Empty<TextBlockCondition>(), String.Empty);
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
                conditionsGridView.Rows.Add(condition.TagName, condition.Condition, condition.EscapedDelimiter);
            }

            priorityTextBox.Text = _tag.Priority.ToString();
        }

        private void SaveTag()
        {
            // сборка параметров тега
            _tag.Key = tagKeyTextBox.Text;
            var conditions = new List<TextBlockCondition>();
            foreach (DataGridViewRow row in conditionsGridView.Rows)
            {
                var tagName = row.Cells["TagNameColumn"].Value?.ToString();
                var tagCondition = row.Cells["TagValueColumn"].Value?.ToString();
                var tagDelimiter = row.Cells["TagDelimiterColumn"].Value?.ToString();
                tagDelimiter = tagDelimiter == "" ? null : tagDelimiter;
                if (string.IsNullOrEmpty(tagName))
                    continue;
                conditions.Add(new TextBlockCondition(tagName, tagCondition, tagDelimiter));
            }
            _tag.Conditions = conditions.ToArray();

            var priority = 0;
            int.TryParse(priorityTextBox.Text, out priority);

            // запись в бд
            if (!_tag.CanStoreInDataBase) // контроль уникальности
            {
                MessageBox.Show(@"Данный набор условий уже существует", @"Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
                
            }

            _tag.SaveToDatabase(tagValueTextBox.Text, priority); // в любом случае сохраняем тег в бд

            // контроль дефолтности
            if (_tag.IsDefault && !_tag.CanBeDefault)
                _tag.SetIsDefaultState(false);

            if (_tag.HasDefaultValue) return;
            if (_tag.CanBeDefault)
            {
                _tag.SetIsDefaultState(true);
                return;
            }
            TextBlockTag.SetDefaultValue(_tag.Key, "");
                
            MessageBox.Show(@"Значение по умолчанию было создано автоматически", @"Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Information);





        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveTag();
            Dispose();
            _parent?.Refresh();
        }
    }

}
