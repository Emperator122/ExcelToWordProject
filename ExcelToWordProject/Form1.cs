using ExcelToWordProject.Models;
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
using Xceed.Words.NET;

namespace ExcelToWordProject
{
    public partial class MainForm : Form
    {
        SyllabusParameters syllabusParameters;
        public MainForm()
        {
            InitializeComponent();

            syllabusParameters = ConfigManager.GetConfigData();


        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private async void ConvertButton_Click(object sender, EventArgs e)
        {
            string selectedFilePath = filePathTextBox.Text;
            string templateFilePath = templateFilePathTextBox.Text;
            string resultFolderPath = resultFolderPathTextBox.Text;
            SyllabusReader syllabusReader = new SyllabusReader(selectedFilePath, syllabusParameters);
            status.Text = "Генерация файлов...";
            try
            {
                await Task.Run(()=> syllabusReader.ConvertToDocx(resultFolderPath, templateFilePath, resultFilePrefixTextBox.Text,
                    new Progress<int>(percent => progressBar1.Value = percent)));
            }
            finally
            {
                syllabusReader.CloseStreams();
                status.Text = "Ожидание...";
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
            SmartTagSettingsForm smartTagSettingsForm = new SmartTagSettingsForm(syllabusParameters);
            smartTagSettingsForm.Show();
        }

        private void DefaultТегиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DefaultTagSettingsForm defaultTagSettingsForm = new DefaultTagSettingsForm(syllabusParameters);
            defaultTagSettingsForm.Show();
        }
    }
}
