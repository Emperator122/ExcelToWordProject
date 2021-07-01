using ExcelToWordProject.Forms;
using ExcelToWordProject.Syllabus;
using ExcelToWordProject.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Media;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelToWordProject
{
    public partial class DefaultTagSettingsForm : Form
    {
        // Настройки для генерации таблички с параметрами тегов
        string[] names = new string[] { "rowIndexTextBox", "columnIndexTextBox", "tagTextBox", "listTextBox"};
        string[] titles = new string[] { "Индекс строки", "Индекс столбца", "Тег", "Рабочий лист", "Рег. выражение",  "Описание", "Копировать", "Удалить" };
        int defaultTextBoxWidth = 180;
        int defaultMargin = 10;

        Bitmap clipboardIcon = Properties.Resources.clipboards;
        Bitmap infoIcon = Properties.Resources.information;
        Bitmap removeIcon = Properties.Resources.remove;
        Bitmap regexIcon = Properties.Resources.regExp;


        public SyllabusParameters syllabusParameters;
        public List<BaseSyllabusTag> Tags;
        public DefaultTagSettingsForm(SyllabusParameters syllabusParameters)
        {
            InitializeComponent();

            this.syllabusParameters = syllabusParameters;

            Tags = new List<BaseSyllabusTag>(syllabusParameters.Tags);

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
            Tags.ForEach(tag => {
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

            // Кнопку регулярок
            PictureBox regExpButton = new PictureBox()
            {
                Width = 26,
                Height = 26,
                Top = 0,
                Left = (titles.Length - 4) * (defaultTextBoxWidth + defaultMargin) + defaultMargin + 20,
                Image = regexIcon,
                Cursor = Cursors.Hand,
                SizeMode = PictureBoxSizeMode.StretchImage,
            };
            regExpButton.Click += (Object sender, EventArgs e) =>
            {
                RegularExpressionEditForm regularExpressionEditForm = new RegularExpressionEditForm(tag);
                regularExpressionEditForm.ShowDialog();
            };
            panel.Controls.Add(regExpButton);

            // Кнопку информации
            PictureBox infoButton = new PictureBox()
            {
                Width = 26,
                Height = 26,
                Top = 0,
                Left = (titles.Length - 3) * (defaultTextBoxWidth + defaultMargin) + defaultMargin + 20,
                Image = infoIcon,
                Cursor = Cursors.Hand,
                SizeMode = PictureBoxSizeMode.StretchImage,
            };
            infoButton.Click += (Object sender, EventArgs e) =>
            {
                DescriptionEditForm editForm = new DescriptionEditForm(tag);
                editForm.ShowDialog();
                

            };
            panel.Controls.Add(infoButton);


            // Кнопку копирования
            PictureBox copyButton = new PictureBox()
            {
                Width = 26,
                Height = 26,
                Top = 0,
                Left = (titles.Length - 2) * (defaultTextBoxWidth + defaultMargin) + defaultMargin + 20,
                Image = clipboardIcon,
                Cursor = Cursors.Hand,
                SizeMode = PictureBoxSizeMode.StretchImage,
            };
            copyButton.Click += (Object sender, EventArgs e) =>
            {
                Clipboard.SetData(DataFormats.Text, tag.Tag);
                SystemSounds.Beep.Play();
            };
            panel.Controls.Add(copyButton);


            // Кнопку удаления
            PictureBox removeButton = new PictureBox()
            {
                Width = 26,
                Height = 26,
                Top = 0,
                Name = "removeButton",
                Left = (titles.Length - 1) * (defaultTextBoxWidth + defaultMargin) + defaultMargin + 20,
                Image = removeIcon,
                Cursor = Cursors.Hand,
                SizeMode = PictureBoxSizeMode.StretchImage
            };

            removeButton.Click += (Object sender, EventArgs e) => {
                DialogResult dialogResult = MessageBox.Show("Вы уверены, что хотите удалить " +
                    "данный тег?", "Внимание", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dialogResult == DialogResult.Yes)
                {
                    Tags.Remove(tag);
                    panel.Dispose();
                    parent.Controls.Remove(panel);
                }
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
            Tags.ForEach(tag => {
                if (tag is DefaultSyllabusTag)
                    defaultSyllabusTags.Add(tag as DefaultSyllabusTag);
            });
            defaultSyllabusTags.Reverse(); // из-за Dock'a приходится инвертировать

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

            syllabusParameters.Tags = Tags;

            ConfigManager.SaveConfigData(syllabusParameters);

            Close();
            Dispose();
        }

        private void AddButton_Click(object sender, EventArgs e)
        {
            DefaultSyllabusTag tag = new DefaultSyllabusTag(0, 0, "", "");
            Tags.Add(tag);
            Panel row = GenerateSmartTagRow(defaultTagsPanel.Controls.Count, tag, defaultTagsPanel);
            row.Visible = false;
            defaultTagsPanel.Controls.Add(row);
            defaultTagsPanel.Controls[defaultTagsPanel.Controls.Count - 1].BringToFront();
            row.Visible = true;
        }

    }
}
