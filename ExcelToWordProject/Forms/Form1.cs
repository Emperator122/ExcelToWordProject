using ExcelToWordProject.Forms;
using ExcelToWordProject.Models;
using ExcelToWordProject.Syllabus;
using ExcelToWordProject.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xceed.Words.NET;

namespace ExcelToWordProject
{
    public partial class MainForm : Form
    {
        bool LockButtons = false;
        DefaultTagSettingsForm DefaultTagSettingsForm;
        SmartTagSettingsForm SmartTagSettingsForm;
        TagListForm TagListForm;
        AboutProgramForm aboutProgramForm;
        SyllabusParameters syllabusParameters;
        public MainForm()
        {
            InitializeComponent();

            syllabusParameters = ConfigManager.GetConfigData();

            SetToolTips();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Module module = new Module("0", "ваыва", "фваываыва;ываыва");
            Type moduleType = module.GetType();
            moduleType.GetField("Name");

        }

        private void SetToolTips()
        {
            ToolTip toolTip = new ToolTip();
            toolTip.SetToolTip(convertButton, "Конвертировать");
            toolTip.SetToolTip(filePathButton, "Выберите файл");
            toolTip.SetToolTip(templateFilePathButton, "Выберите файл");
            toolTip.SetToolTip(folderPathButton, "Выберите папку");
        }

        private async void ConvertButton_Click(object sender, EventArgs e)
        {
            if (LockButtons)
                return;
            LockButtons = true;
            string selectedFilePath = filePathTextBox.Text;
            string templateFilePath = templateFilePathTextBox.Text;
            string resultFolderPath = resultFolderPathTextBox.Text;
            SyllabusExcelReader syllabusExcelReader = null;
            SyllabusDocWriter syllabusDocWriter = null;
            /*
            syllabusExcelReader = new SyllabusExcelReader(selectedFilePath, syllabusParameters);
            syllabusDocWriter = new SyllabusDocWriter(syllabusExcelReader, syllabusParameters);
            status.Text = "Генерация файлов...";
            syllabusDocWriter.ConvertToDocx(resultFolderPath, templateFilePath, resultFilePrefixTextBox.Text,
                new Progress<int>(percent => progressBar1.Value = percent));
            */
            try
            {
                syllabusExcelReader = new SyllabusExcelReader(selectedFilePath, syllabusParameters);
                syllabusDocWriter = new SyllabusDocWriter(syllabusExcelReader, syllabusParameters);

                // Проверка на активные смарт теги при неправильном файле
                if (syllabusParameters.HasActiveSmartTags && !syllabusExcelReader.IsSyllabusFile)
                {
                    DialogResult dialogResult = MessageBox.Show("Возможно данный файл не является " +
                        "файлом учебного плана, но у вас активны \"умные\" теги. Это может стать причиной " +
                        "сбоя в работе программы.\r\nОтключить \"умные\" теги?", "Внимание!", MessageBoxButtons.YesNoCancel);
                    switch (dialogResult)
                    {
                        case DialogResult.Yes:
                            syllabusParameters.DisableSmartTags();
                            break;
                        case DialogResult.Cancel:
                            return;
                    }
                }

                
                status.Text = "Генерация файлов...";
                await Task.Run(()=> syllabusDocWriter.ConvertToDocx(resultFolderPath, templateFilePath, resultFilePrefixTextBox.Text,
                    new Progress<int>(percent => progressBar1.Value = percent)));
            }
            catch(Exception ex)
            {
                MessageBox.Show("Произошла ошибка:\r\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                syllabusExcelReader?.CloseStreams();
                status.Text = "Ожидание...";
                LockButtons = false;
            }
        }

        private void FilePathButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Excel документы|*.xls;*.xlsx";
            if(fileDialog.ShowDialog() == DialogResult.OK)
            {
                filePathTextBox.Text = fileDialog.FileName;
            }
        }

        private void FolderPathButton_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDialog = new FolderBrowserDialog();
            folderDialog.SelectedPath = Application.StartupPath;
            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                resultFolderPathTextBox.Text = folderDialog.SelectedPath;
            }
        }

        private void TemplateFilePathButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Word документы|*.doc;*.docx";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                templateFilePathTextBox.Text = fileDialog.FileName;
            }
        }

        private void SmartTagsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (SmartTagSettingsForm == null || SmartTagSettingsForm.IsDisposed)
            {
                SmartTagSettingsForm = new SmartTagSettingsForm(syllabusParameters);
                SmartTagSettingsForm.Show();
            }
            else
                SmartTagSettingsForm.Focus();
        }

        private void DefaultТегиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (DefaultTagSettingsForm == null || DefaultTagSettingsForm.IsDisposed)
            {
                DefaultTagSettingsForm = new DefaultTagSettingsForm(syllabusParameters);
                DefaultTagSettingsForm.Show();
            }
            else
                DefaultTagSettingsForm.Focus();
        }

        private void СписокТеговToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (TagListForm == null || TagListForm.IsDisposed)
            {
                TagListForm = new TagListForm(syllabusParameters.Tags);
                TagListForm.Show();
            }
            else
                TagListForm.Focus();
        }

        private void ОПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (TagListForm == null || TagListForm.IsDisposed)
            {
                aboutProgramForm = new AboutProgramForm();
                aboutProgramForm.ShowDialog();
            }
            else
                aboutProgramForm.Focus();
        }
    }
}
