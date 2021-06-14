using ExcelToWordProject.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelToWordProject
{
    public partial class DefaultTagSettingsForm : Form
    {
        // Настройки для генерации таблички с параметрами тегов
        string[] names = new string[] { "rowIndexTextBox", "columnIndexTextBox", "tagTextBox", "listTextBox" };
        string[] titles = new string[] { "Индекс стобца", "Индекс строки", "Тег", "Рабочий лист", "Удалить" };
        int defaultTextBoxWidth = 180;
        int defaultMargin = 10;


        public SyllabusParameters syllabusParameters;
        public DefaultTagSettingsForm(SyllabusParameters syllabusParameters)
        {
            InitializeComponent();
            this.syllabusParameters = syllabusParameters;
            Control[] controls = GenerateSmartTagsSettingsElements(defaultTagsPanel);
            
            defaultTagsPanel.Controls.AddRange(controls);
            foreach (Control control in controls)
                control.BringToFront();
        }


        protected Control[] GenerateSmartTagsSettingsElements(Control parent)
        {
            List<Control> result = new List<Control>();

            // получим все смарт теги
            List<DefaultSyllabusTag> defaultSyllabusTags = new List<DefaultSyllabusTag>();
            syllabusParameters.Tags.ForEach(tag => {
                if (tag is DefaultSyllabusTag)
                    defaultSyllabusTags.Add(tag as DefaultSyllabusTag);
            });


            Panel headerPanel = new Panel();
            headerPanel.Height = 20;
            headerPanel.Name = "headerPanel";
            headerPanel.Top = defaultMargin;
            headerPanel.AutoSize = true;
            headerPanel.Dock = DockStyle.Top;
            for (int i = 0; i < titles.Length; i++)
            {

                Label label = new Label();
                label.Text = titles[i];
                label.Top = 0;
                label.Left = i * (defaultTextBoxWidth + defaultMargin) + defaultMargin;
                label.Font = new Font(FontFamily.GenericSansSerif, 12, FontStyle.Bold);
                label.AutoSize = true;
                headerPanel.Controls.Add(label);
            }
            result.Add(headerPanel);



            for (int i = 0; i < defaultSyllabusTags.Count(); i++)
            {
                // текущий тег
                DefaultSyllabusTag tag = defaultSyllabusTags[i];

                Panel panel = GenerateSmartTagRow(i, tag, parent);

                result.Add(panel);
            }
            return result.ToArray();
        }

        protected Panel GenerateSmartTagRow(int i, DefaultSyllabusTag tag, Control parent)
        {
            Panel panel = new Panel();
            panel.Height = 26;
            panel.AutoSize = true;
            panel.Left = 0;
            panel.Top = (i + 1) * (panel.Height + defaultMargin);
            panel.Dock = DockStyle.Top;
            panel.Margin = new Padding(10);
            

            // Добавим все тектовые поля
            string[] textBoxValues = new string[] { tag.RowIndex.ToString(), tag.ColumnIndex.ToString(),
                                                        tag.Key, tag.ListName};
            for (int j = 0; j < names.Length; j++)
            {
                TextBox textBox = new TextBox();
                textBox.Width = defaultTextBoxWidth;
                textBox.Top = 5;
                textBox.Left = j * (textBox.Width + defaultMargin) + defaultMargin;
                textBox.Name = names[j];
                textBox.Text = textBoxValues[j];

                panel.Controls.Add(textBox);
            }


            // И кнопку удаления
            PictureBox removeButton = new PictureBox();
            removeButton.Width = 26;
            removeButton.Height = 26;
            removeButton.Top = 0;
            removeButton.Name = "removeButton";
            removeButton.Left = (titles.Length - 1) * (defaultTextBoxWidth + defaultMargin) + defaultMargin + 20;
            removeButton.Image = Properties.Resources.remove;
            removeButton.Cursor = Cursors.Hand;
            removeButton.SizeMode = PictureBoxSizeMode.StretchImage;
            removeButton.Click += (Object sender, EventArgs e) => {
                syllabusParameters.Tags.Remove(tag);
                panel.Dispose();
                parent.Controls.Remove(panel);
            };


            panel.Controls.Add(removeButton);

            return panel;
        }

        private void DefaultTagSettingsForm_Resize(object sender, EventArgs e)
        {

        }

        private void Button1_Click(object sender, EventArgs e)
        {
            // получим все обычные теги
            List<DefaultSyllabusTag> defaultSyllabusTags = new List<DefaultSyllabusTag>();
            syllabusParameters.Tags.ForEach(tag => {
                if (tag is DefaultSyllabusTag)
                    defaultSyllabusTags.Add(tag as DefaultSyllabusTag);
            });
            defaultSyllabusTags.Reverse();

            // Загоним информацию из филдов в теги
            int i = 0;
            foreach (Control child in defaultTagsPanel.Controls)
            {
                if (child.Name != "headerPanel")
                {
                    DefaultSyllabusTag tag = defaultSyllabusTags[i];
                    tag.RowIndex = Convert.ToInt32((child.Controls["rowIndexTextBox"] as TextBox).Text);
                    tag.ColumnIndex = Convert.ToInt32((child.Controls["columnIndexTextBox"] as TextBox).Text);
                    tag.Key = (child.Controls["tagTextBox"] as TextBox).Text;
                    tag.ListName = (child.Controls["listTextBox"] as TextBox).Text;

                    i++;
                }
            }
            ConfigManager.SaveConfigData(syllabusParameters);

            Close();
            Dispose();
        }

        private void AddButton_Click(object sender, EventArgs e)
        {
            DefaultSyllabusTag tag = new DefaultSyllabusTag(0, 0, "", "");
            syllabusParameters.Tags.Add(tag);
            defaultTagsPanel.Controls.Add(GenerateSmartTagRow(defaultTagsPanel.Controls.Count, tag, defaultTagsPanel));
            defaultTagsPanel.Controls[defaultTagsPanel.Controls.Count - 1].BringToFront();
        }
    }
}
