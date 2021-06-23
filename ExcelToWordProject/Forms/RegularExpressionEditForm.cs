using ExcelToWordProject.Syllabus;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelToWordProject.Forms
{
    public partial class RegularExpressionEditForm : Form
    {
        BaseSyllabusTag tag;
        public RegularExpressionEditForm(BaseSyllabusTag tag)
        {
            InitializeComponent();
            this.tag = tag;
            // Получим все возмодные настройки регулярного выражения
            // и дропнем оттуда вариант "None"
            Array enumValues = Enum.GetValues(typeof(RegexOptions));
            RegexOptions[] enumValuesWithoutNone = new RegexOptions[enumValues.Length-1];
            Array.Copy(enumValues, 1, enumValuesWithoutNone, 0, enumValues.Length - 1);

            // Загрузим настройки в checkList
            // и выберем активные варианты
            regexSettings.DataSource = enumValuesWithoutNone;
            for (int i = 0; i < regexSettings.Items.Count; i++)
            {
                RegexOptions item = (RegexOptions)regexSettings.Items[i];
                if (tag.RegularEx.RegexOptions.HasFlag(item))
                    regexSettings.SetItemChecked(i, true);
            }

            // Выведем текст регулярки и индекс группы
            regExpTextBox.Text = tag.RegularEx.Expression;
            regexGroupIndexTextBox.Text = tag.RegularEx.GroupIndex.ToString();

        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            // Сохраним все данные в объект

            tag.RegularEx.Expression = regExpTextBox.Text;
            tag.RegularEx.GroupIndex = Convert.ToInt32(regexGroupIndexTextBox.Text);

            // Заполнение enum с настройками
            tag.RegularEx.RegexOptions = RegexOptions.None;
            foreach (var item in regexSettings.CheckedItems)
                tag.RegularEx.RegexOptions |= (RegexOptions)item;

            CancelButton1_Click(sender, e);
        }

        private void CancelButton1_Click(object sender, EventArgs e)
        {
            Close();
            Dispose();
        }

        /// <summary>
        /// Проверка регулярного выражения
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TestRegExpButton_Click(object sender, EventArgs e)
        {
            // enum с выбранными настройками
            RegexOptions options = RegexOptions.None;
            foreach (var item in regexSettings.CheckedItems)
                options |= (RegexOptions)item;

            try
            {
                // получим результат
                var match = Regex.Match(testRegExpTextBox.Text, regExpTextBox.Text, options);

                // вывод служебной информации
                groupsCountLabel.Text = "Количество групп: " + match.Groups.Count;

                //  если юзер ввел некорректный номер группы,
                // то выдадим ошибку
                int groupNumber = Convert.ToInt32(regexGroupIndexTextBox.Text);
                if (groupNumber >= match.Groups.Count)
                    MessageBox.Show("Найдено " + match.Groups.Count + " групп. Вы выбрали группу " +
                        "с индексом " + groupNumber + ". " +
                        "Индексация с нуля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else // иначе вывод результата
                    MessageBox.Show(match.Groups[groupNumber].Value, "Результат");
            }
            catch(Exception ex)
            {
                MessageBox.Show("Произошла ошибка:\r\n"+ex.Message ,"Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
           
        }
    }
}
