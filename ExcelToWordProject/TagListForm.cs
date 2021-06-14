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
    public partial class TagListForm : Form
    {
        // Настройки для генерации таблички с параметрами тегов
        string[] titles = new string[] { "Тег", "Описание", "Копировать"};
        int defaultTextBoxWidth = 110;
        int defaultMargin = 5;
        Bitmap infoIcon = Properties.Resources.information;
        Bitmap clipboardIcon = Properties.Resources.clipboards;

        List<BaseSyllabusTag> Tags;
        public TagListForm(List<BaseSyllabusTag> tags)
        {
            InitializeComponent();
            Tags = tags;
            tagsPanel.Controls.AddRange(GenerateSmartTagsSettingsElements(tagsPanel));
            topMostCheckBox.Checked = TopMost;
        }

        protected Control[] GenerateSmartTagsSettingsElements(Control parent)
        {
            List<Control> result = new List<Control>();

            Panel headerPanel = new Panel();
            headerPanel.Name = "headerPanel";
            headerPanel.Height = 25;
            headerPanel.Top = defaultMargin;
            headerPanel.AutoSize = true;
            for (int i = 0; i < titles.Length; i++)
            {

                Label label = new Label();
                label.Text = titles[i];
                label.Top = 0;
                label.Left = i * (defaultTextBoxWidth + defaultMargin) + defaultMargin;
                label.Font = new Font(FontFamily.GenericSansSerif, 12, FontStyle.Bold);
                label.Width = defaultTextBoxWidth;
                label.AutoEllipsis = true;
                headerPanel.Controls.Add(label);
            }
            result.Add(headerPanel);



            for (int i = 0; i < Tags.Count(); i++)
            {
                // текущий тег
                BaseSyllabusTag tag = Tags[i];

                Panel panel = GenerateSmartTagRow(i, tag, parent);

                result.Add(panel);
            }
            return result.ToArray();
        }

        protected Panel GenerateSmartTagRow(int i, BaseSyllabusTag tag, Control parent)
        {
            Panel panel = new Panel();
            panel.Height = 26;
            panel.AutoSize = true;
            panel.Left = 0;
            panel.Top = (i + 1) * (panel.Height + defaultMargin);
            panel.Margin = new Padding(10);

            // Инфу о типе тега
            Label label = new Label();
            label.Width = defaultTextBoxWidth;
            label.Top = 0;
            label.AutoEllipsis = true;
            label.Left = (titles.Length - 3) * (label.Width + defaultMargin) + defaultMargin;
            label.Text = tag.Key;
            label.Font = new Font("Arial", 12, FontStyle.Bold);
            panel.Tag = tag;
            panel.Controls.Add(label);

            // Добавляем кнопку информации
            PictureBox infoButton = new PictureBox()
            {
                Width = 26,
                Height = 26,
                Top = 0,
                Left = (titles.Length - 2) * (defaultTextBoxWidth + defaultMargin) + defaultMargin + 20,
                Image = infoIcon,
                Cursor = Cursors.Hand,
                SizeMode = PictureBoxSizeMode.StretchImage,
            };
            infoButton.Click += (Object sender, EventArgs e) =>
            {
                MessageBox.Show(tag.Description, "Описание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };
            panel.Controls.Add(infoButton);


            // Кнопку копирования
            PictureBox copyButton = new PictureBox()
            {
                Width = 26,
                Height = 26,
                Top = 0,
                Left = (titles.Length - 1) * (defaultTextBoxWidth + defaultMargin) + defaultMargin + 20,
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

            return panel;
        }

        private void ПоверхВсехОконToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void TopMostCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            TopMost = topMostCheckBox.Checked;
        }
    }
}
