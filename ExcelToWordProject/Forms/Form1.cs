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
        GetDataParametersForm getDataParametersForm;


        SyllabusParameters syllabusParameters;

        string[] selectedExcels = new string[0];
        public MainForm()
        {
            InitializeComponent();

            syllabusParameters = ConfigManager.GetConfigData();

            SetToolTips();

        }

        private void SwitchExcelLoadingMode(bool single)
        {
            if(single)
            {
                excelFilesLabel.Visible = false;
                excelFilesLabelClear.Visible = false;
                filePathTextBox.Visible = true;
                selectedExcels = new string[0];
            }
            else
            {
                excelFilesLabel.Visible = true;
                excelFilesLabelClear.Visible = true;
                filePathTextBox.Visible = false;
                filePathTextBox.Text = "";
                excelFilesLabel.Text = "Выбрано " + selectedExcels.Length + " файлов";
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {


        }

        private void SetToolTips()
        {
            ToolTip toolTip = new ToolTip();
            toolTip.SetToolTip(convertButton, "Конвертировать");
            toolTip.SetToolTip(filePathButton, "Выберите файл");
            toolTip.SetToolTip(templateFilePathButton, "Выберите файл");
            toolTip.SetToolTip(folderPathButton, "Выберите папку");
        }

        public async Task ConvertProcessing(string selectedFilePath, string templateFilePath, string resultFolderPath, string prefix)
        {
            SyllabusExcelReader syllabusExcelReader = null;
            SyllabusDocWriter syllabusDocWriter = null;

            try
            {
                syllabusExcelReader = new SyllabusExcelReader(selectedFilePath, syllabusParameters);
                syllabusDocWriter = new SyllabusDocWriter(syllabusExcelReader, syllabusParameters);
                // Проверка на активные смарт теги при неправильном файле
                if (syllabusParameters.HasActiveSmartTags && !syllabusExcelReader.IsSyllabusFile)
                {
                    DialogResult dialogResult = MessageBox.Show("Возможно данный файл " +
                        "("+selectedFilePath+") не является " +
                        "файлом учебного плана, но у вас активны \"умные\" теги. Это может стать причиной " +
                        "сбоя в работе программы.\r\nОтключить \"умные\" теги?", "Внимание!", MessageBoxButtons.YesNoCancel);
                    switch (dialogResult)
                    {
                        case DialogResult.Yes:
                            // Склонируем новые параметры
                            SyllabusParameters tempParameters = ConfigManager.GetConfigData();

                            // Закроем старые потоки чтения
                            syllabusExcelReader.Dispose();
                            syllabusDocWriter.Dispose();

                            // Отключим умные теги
                            tempParameters.DisableSmartTags();
                            syllabusExcelReader = new SyllabusExcelReader(selectedFilePath, tempParameters);
                            syllabusDocWriter = new SyllabusDocWriter(syllabusExcelReader, tempParameters);
                            break;
                        case DialogResult.Cancel:
                            return;
                    }
                }

                //syllabusDocWriter.ConvertToDocx(resultFolderPath, templateFilePath, prefix,
                //            new Progress<int>(percent => progressBar1.Value = percent));

                await Task.Run(() => syllabusDocWriter.ConvertToDocx(resultFolderPath, templateFilePath, prefix,
                            new Progress<int>(percent => progressBar1.Value = percent)));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка:\r\n" + ex.Message+"" +
                    "\r\nФайл:" + selectedFilePath, 
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                syllabusExcelReader?.CloseStreams();
                
            }
        }

        private async void ConvertButton_Click(object sender, EventArgs e)
        {
            if (LockButtons)
                return;
            LockButtons = true;
            string selectedFilePath = filePathTextBox.Text;
            string templateFilePath = templateFilePathTextBox.Text;
            string resultFolderPath = resultFolderPathTextBox.Text;
            
            
            if(selectedExcels.Length <= 1)
                 await ConvertProcessing(selectedFilePath, templateFilePath, resultFolderPath, resultFilePrefixTextBox.Text);
            else if(selectedExcels.Length > 1)
                for (int i = 0; i < selectedExcels.Length; i++)
                {
                    status.Text = "Файл " + (i + 1) + " из " + selectedExcels.Length + "...";

                    string fileName = Path.GetFileNameWithoutExtension(selectedExcels[i]);
                    string folderPath = Path.Combine(resultFolderPath, fileName);

                    try
                    {
                        Directory.CreateDirectory(folderPath);
                        await ConvertProcessing(
                            selectedExcels[i], 
                            templateFilePath,
                            folderPath, 
                            resultFilePrefixTextBox.Text);
                    }
                    catch
                    {
                        string prefix = "[" + Path.GetFileNameWithoutExtension(selectedExcels[i]) + "] " + resultFilePrefixTextBox.Text;
                        await ConvertProcessing(selectedExcels[i], templateFilePath, resultFolderPath, prefix);
                    }
                }

            LockButtons = false;
            status.Text = "Ожидание...";

        }

        private void FilePathButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Excel документы|*.xls;*.xlsx";
            fileDialog.Multiselect = true;
            if(fileDialog.ShowDialog() == DialogResult.OK)
            {
                filePathTextBox.Text = fileDialog.FileName;
                selectedExcels = fileDialog.FileNames;
                if(selectedExcels.Length > 1)
                    SwitchExcelLoadingMode(false);
                else
                    SwitchExcelLoadingMode(true);
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

        private void ExcelFilesLabelClear_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SwitchExcelLoadingMode(true);
        }

        private void ConstantsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (getDataParametersForm == null || getDataParametersForm.IsDisposed)
            {
                getDataParametersForm = new GetDataParametersForm(syllabusParameters);
                getDataParametersForm.Show();
            }
            else
                getDataParametersForm.Focus();
        }
    }
}
