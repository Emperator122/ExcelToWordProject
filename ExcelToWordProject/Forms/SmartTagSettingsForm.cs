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
    partial class SmartTagSettingsForm : Form
    {
        // Настройки для генерации таблички с параметрами тегов
        string[] names = new string[] { "indexTextBox", "tagTextBox" };
        string[] titles = new string[] { "Индекс стобца", "Тег", "Тип тега", "Описание", "Скопировать" };
        int defaultTextBoxWidth = 180;
        int defaultMargin = 10;

        Bitmap infoIcon = Properties.Resources.information;
        Bitmap clipboardIcon = Properties.Resources.clipboards;



        public SyllabusParameters syllabusParameters;
        public SmartTagSettingsForm(SyllabusParameters syllabusParameters)
        {
            InitializeComponent();
            // получим все смарт теги
            List<SmartSyllabusTag> smartSyllabusTags = new List<SmartSyllabusTag>();
            syllabusParameters.Tags.ForEach(tag => {
                if (tag is SmartSyllabusTag)
                    smartSyllabusTags.Add(tag as SmartSyllabusTag);
            });

            this.syllabusParameters = syllabusParameters;
            smartTagsPanel.Controls.AddRange(GenerateSmartTagsSettingsElements(smartTagsPanel, smartSyllabusTags));


            tagActivator.Nodes.AddRange(GenerateTreeNodes(smartSyllabusTags));
            foreach (TreeNode node in tagActivator.Nodes)
                node.ExpandAll();

        }


        public void SetStatePanelsWithListNamed(string listName, bool state)
        {
            foreach (Control item in smartTagsPanel.Controls)
            {
                if(item.Name != "headerPanel")
                {
                    SmartSyllabusTag tag = item.Tag as SmartSyllabusTag;
                    if(tag.ListName == listName)
                    {
                        tag.Active = state;
                        item.Enabled = state;
                    }
                }
            }
        }

        protected TreeNode[] GenerateTreeNodes(List<SmartSyllabusTag> smartSyllabusTags)
        {
            List<TreeNode> result = new List<TreeNode>();

            SmartSyllabusTag tempTag = smartSyllabusTags.Find(tag => tag.ListName == syllabusParameters.PlanListName);

            // Сгенерируем корневую ноду
            TreeNode rootNode = new TreeNode("Теги с листа " + syllabusParameters.PlanListName) {
                Checked = tempTag.Active,
                Name = "planCheckBox"
            };

            result.Add(rootNode);

            // И дочернюю
            tempTag = smartSyllabusTags.Find(tag => tag.ListName == syllabusParameters.ModulesContentListName);
            TreeNode node = new TreeNode("Теги с листа " + syllabusParameters.ModulesContentListName)
            {
                Checked = tempTag.Active,
                Name = "contentCheckBox"
            };
            rootNode.Nodes.Add(node);

            return result.ToArray();
        }

        protected Control[] GenerateSmartTagsSettingsElements(Control parent, List<SmartSyllabusTag> smartSyllabusTags)
        {
            List<Control> result = new List<Control>();

            Panel headerPanel = new Panel();
            headerPanel.Height = 20;
            headerPanel.Width = parent.Width;
            headerPanel.Name = "headerPanel";
            headerPanel.AutoSize = true; 
            headerPanel.Top = defaultMargin;
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


            // Поля для ввода (и не только) инфы о смарт тегах
            for (int i = 0; i < smartSyllabusTags.Count(); i++)
            {
                // текущий тег
                SmartSyllabusTag tag = smartSyllabusTags[i];


                Panel row = GenerateSmartTagRow(i, tag, parent);
                result.Add(row);
            }

            return result.ToArray();
        }

        protected Panel GenerateSmartTagRow(int i, SmartSyllabusTag tag, Control parent)
        {
            Panel panel = new Panel();
            panel.Height = 25;
            panel.Width = parent.Width;
            panel.AutoSize = true;
            panel.Left = 0;
            panel.Enabled = tag.Active;
            panel.Top = (i + 1) * (panel.Height + defaultMargin);

            // Добавим все тектовые поля
            string[] textBoxValues = new string[] { tag.ColumnIndex.ToString(), tag.Key };
            for (int j = 0; j < names.Length; j++)
            {
                TextBox textBox = new TextBox();
                textBox.Width = defaultTextBoxWidth;
                textBox.Top = 0;
                textBox.Left = j * (textBox.Width + defaultMargin) + defaultMargin;
                textBox.Name = names[j];
                textBox.Text = textBoxValues[j];
                textBox.Enabled = !(j == 0 && textBoxValues[j] == "-1"); 

                panel.Controls.Add(textBox);
            }


            // Инфу о типе тега
            Label label = new Label();
            label.Width = defaultTextBoxWidth;
            label.Top = 0;
            label.Left = (titles.Length - 3) * (label.Width + defaultMargin) + defaultMargin;
            label.Text = tag.Type.ToString();
            label.Font = new Font(FontFamily.GenericSansSerif, 10, FontStyle.Italic);
            label.AutoEllipsis = true;
            panel.Tag = tag;
            panel.Controls.Add(label);

            // Кнопку информации
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
            PictureBox copyButton = new PictureBox() {
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


        private void Button1_Click(object sender, EventArgs e)
        {
            // получим все смарт теги
            List<SmartSyllabusTag> smartSyllabusTags = new List<SmartSyllabusTag>();
            syllabusParameters.Tags.ForEach(tag => {
                if (tag is SmartSyllabusTag)
                    smartSyllabusTags.Add(tag as SmartSyllabusTag);
            });

            // Загоним информацию из филдов в теги
            int i = 0;
            foreach(Control child in smartTagsPanel.Controls)
            {
                if(child.Name != "headerPanel")
                {
                    SmartSyllabusTag tag = smartSyllabusTags[i];
                    tag.ColumnIndex = Convert.ToInt32((child.Controls["indexTextBox"] as TextBox).Text);
                    tag.Key = (child.Controls["tagTextBox"] as TextBox).Text;

                    i++;
                }
            }
            ConfigManager.SaveConfigData(syllabusParameters);

            Close();
            Dispose();
        }




        private void TagActivator_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if (!e.Node.Checked)
            {
                foreach (TreeNode node in e.Node.Nodes)
                    node.Checked = false;


            } else if(e.Node.Parent != null)
            {
                e.Node.Parent.Checked = true;
            }


            switch (e.Node.Name)
            {
                case "contentCheckBox":
                    SetStatePanelsWithListNamed(syllabusParameters.ModulesContentListName, e.Node.Checked);
                    break;

                case "planCheckBox":
                    SetStatePanelsWithListNamed(syllabusParameters.PlanListName, e.Node.Checked);
                    break;

            }

        }
    }
}
