using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using ExcelToWordProject.Syllabus.Tags;
using ExcelToWordProject.Utils;
// ReSharper disable LocalizableElement

namespace ExcelToWordProject.Forms
{
    public partial class TextBlockAddForm : Form
    {
        private bool IsFilePathSetting => valueModeTabs.SelectedTab.Name == "filePathTab";
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
            BuildToolTips();
        }

        private void FillComponents()
        {
            tagKeyTextBox.Text = _tag.Key;

            if (_tag.HasId)
            {
                if (_tag.IsFilePath)
                    templateFilePathTextBox.Text = _tag.GetValue2();
                else
                    tagValueTextBox.Text = _tag.GetValue2();
            }
            foreach (var condition in _tag.Conditions)
            {
                conditionsGridView.Rows.Add(condition.TagName, condition.Condition, condition.EscapedDelimiter);
            }

            priorityTextBox.Text = _tag.Priority.ToString();

            delimiterTextBox.Text = _tag.EscapedDelimiter;

            valueModeTabs.SelectTab(_tag.IsFilePath ? "filePathTab" : "textValueTab");
            OnTextBlockValueModeSwitch();
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

            int.TryParse(priorityTextBox.Text, out var priority);

            var isFilePath = IsFilePathSetting;

            var value = isFilePath ? templateFilePathTextBox.Text : tagValueTextBox.Text;

            _tag.EscapedDelimiter = isFilePath ? null : delimiterTextBox.Text;

            // запись в бд
            if (!_tag.CanStoreInDataBase) // контроль уникальности
            {
                MessageBox.Show(@"Данный набор условий уже существует", @"Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
                
            }

            _tag.SaveToDatabase(value, priority, isFilePath); // в любом случае сохраняем тег в бд

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
            if (IsFilePathSetting && !IsValidDocumentPath(templateFilePathTextBox.Text))
            {
                ShowPathError();
                return;
            }
            SaveTag();
            Dispose();
            _parent?.Refresh();
        }

        private void valueModeTabs_SelectedIndexChanged(object sender, EventArgs e)
        {
            OnTextBlockValueModeSwitch();
        }

        bool IsValidDocumentPath(string path)
        {
            var pathParts = path.Split('\\');
            return pathParts.Length > 0 && pathParts[0] == "TextBlocksDocuments";
        }

        private void templateFilePathButton_Click(object sender, EventArgs e)
        {
            using (var fileDialog = new OpenFileDialog())
            {
                fileDialog.Filter = @"Word документы|*.doc;*.docx";
                Directory.CreateDirectory($"{AppContext.BaseDirectory}\\TextBlocksDocuments");
                fileDialog.InitialDirectory = $"{AppContext.BaseDirectory}\\TextBlocksDocuments";
                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    var relativePath =
                        PathUtils.MakeRelativePath(AppContext.BaseDirectory, fileDialog.FileName);

                    if (!IsValidDocumentPath(relativePath))
                    {
                        ShowPathError();
                        return;
                    }
                    
                    templateFilePathTextBox.Text = relativePath;
                }
            }
        }
    }

}
