using ExcelToWordProject.Syllabus;
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

namespace ExcelToWordProject.Forms
{
    public partial class GetDataParametersForm : Form
    {
        List<string> Errors = new List<string>();
        SyllabusParameters SyllabusParameters;

        int defaultTextBoxWidth = 250;
        int defaultMargin = 10;
        string[] names = new string[] { "headerTitle" };
        string[] titles = new string[] { "Тип заголовка", "Заголовок"};

        public GetDataParametersForm(SyllabusParameters syllabusParameters)
        {
            InitializeComponent();

            SyllabusParameters = syllabusParameters;

            yearsTextBox.Text = OtherUtils.ListToDelimiteredString(";", "", new List<int>(SyllabusParameters.ModulesYears));
            planListNameTextBox.Text = SyllabusParameters.PlanListName;
            modulesContentListNameTextBox.Text = SyllabusParameters.ModulesContentListName;

            planHeaderTitlesPanel.Controls.AddRange(GeneratePlanHeaderTitleSettings());

            planHeaderRowIndexTextBox.Text = syllabusParameters.PlanListHeaderRowIndex.ToString();
        }

        void ApplyYearsFilter()
        {
            if (yearsTextBox.Text.Trim() == "")
            {
                SyllabusParameters.ModulesYears = new int[0];
                return;
            }

            string[] tempYears = yearsTextBox.Text.Split(';');

            List<int> years = new List<int>();
            foreach (string yearString in tempYears)
            {
                int year;
                if (int.TryParse(yearString, out year))
                    years.Add(year);
            }

            if (years.Count > 0)
                SyllabusParameters.ModulesYears = years.ToArray();
            else
                Errors.Add("[*]\tВозможно неправильно заполнен фильтр курсов!");
            
        }

        void ApplyListNames()
        {
            SyllabusParameters.Tags.ForEach(tag =>
            {
                if (tag.ListName == SyllabusParameters.PlanListName)
                    tag.ListName = planListNameTextBox.Text;
                else if (tag.ListName == SyllabusParameters.ModulesContentListName)
                    tag.ListName = modulesContentListNameTextBox.Text;
            });
            SyllabusParameters.PlanListName = planListNameTextBox.Text;
            SyllabusParameters.ModulesContentListName = modulesContentListNameTextBox.Text;
        }

        void ApplyPlanHeaderParameters()
        {
            int rowIndex;
            if (!int.TryParse(planHeaderRowIndexTextBox.Text, out rowIndex))
                Errors.Add("[*]\tВозможно неправильно заполнен индекс строки заголовков!");
            else
                SyllabusParameters.PlanListHeaderRowIndex = rowIndex;

            foreach (Control child in planHeaderTitlesPanel.Controls)
            {
                if (child.Name != "headerPanel")
                {
                    TextBox textBox = child.Controls["headerTitle"] as TextBox;
                    SyllabusParameters.planListHeaderNames[textBox.Tag as string] = textBox.Text;
                }
            }
        }


        protected Control[] GeneratePlanHeaderTitleSettings()
        {
            List<Control> result = new List<Control>();

            Panel headerPanel = new Panel();
            headerPanel.Height = 20;
            headerPanel.Name = "headerPanel";
            headerPanel.Top = defaultMargin;
            headerPanel.AutoSize = true;
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

            int k = 0;
            foreach (var kv in SyllabusParameters.planListHeaderNames)
            {
                Panel panel = GeneratePlanHeaderTitleSettingsRow(k, kv);
                result.Add(panel);
                k++;
            }
            return result.ToArray();
        }

        protected Panel GeneratePlanHeaderTitleSettingsRow(int i, KeyValuePair<string, string> kv)
        {
            Panel panel = new Panel();
            panel.Height = 70;
            panel.AutoSize = true;
            panel.Left = 0;
            panel.Top = (i + 1) * (panel.Height + defaultMargin);
            panel.Margin = new Padding(10);


            // Добавим все тектовые поля
            string[] values = new string[] { kv.Key, kv.Value};

            Label label = new Label()
            {
                Width = defaultTextBoxWidth,
                AutoSize = false,
                AutoEllipsis = true,
                Text = values[0],
                Left = defaultMargin,
                Top = 5,
            };
            panel.Controls.Add(label);

            TextBox textBox = new TextBox();
            textBox.Width = defaultTextBoxWidth;
            textBox.Top = 5;
            textBox.Left = 1 * (textBox.Width + defaultMargin) + defaultMargin;
            textBox.Name = names[0];
            textBox.Text = values[1];
            textBox.Multiline = true;
            textBox.Height = panel.Height;
            textBox.ScrollBars = ScrollBars.Both;
            textBox.Tag = kv.Key;
            textBox.WordWrap = false;
            

            panel.Controls.Add(textBox);

            return panel;
        }



        private void SaveAndQuitButton_Click(object sender, EventArgs e)
        {
            Errors.Clear();
            ApplyYearsFilter();
            ApplyListNames();
            ApplyPlanHeaderParameters();

            if(Errors.Count > 0)
            {
                string errorMessage = "При сохранениии произошли следующие ошибки:\r\n\r\n";
                Errors.ForEach(error => errorMessage += error + "\r\n");
                errorMessage += "\r\nИзменения, связанные с этими ошибками, не будут сохранены. Попытайтесь исправить ошибки.";

                MessageBox.Show(errorMessage, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            ConfigManager.SaveConfigData(SyllabusParameters);

            Close();
            Dispose();
        }
    }
}
