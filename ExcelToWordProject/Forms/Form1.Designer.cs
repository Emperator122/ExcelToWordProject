namespace ExcelToWordProject
{
    partial class MainForm
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.defaultТегиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.smartTagsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.constantsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.списокТеговToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.оПрограммеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.bodyPanel = new System.Windows.Forms.Panel();
            this.bodyGroupBox = new System.Windows.Forms.GroupBox();
            this.excelFilesLabelClear = new System.Windows.Forms.LinkLabel();
            this.excelFilesLabel = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.resultFilePrefixLabel = new System.Windows.Forms.Label();
            this.templateFilePathButton = new System.Windows.Forms.PictureBox();
            this.resultFilePrefixTextBox = new System.Windows.Forms.TextBox();
            this.templateFilePathLabel = new System.Windows.Forms.Label();
            this.templateFilePathTextBox = new System.Windows.Forms.TextBox();
            this.convertButton = new System.Windows.Forms.PictureBox();
            this.filePathButton = new System.Windows.Forms.PictureBox();
            this.folderPathButton = new System.Windows.Forms.PictureBox();
            this.resultFolderPathLabel = new System.Windows.Forms.Label();
            this.resultFolderPathTextBox = new System.Windows.Forms.TextBox();
            this.filePathLabel = new System.Windows.Forms.Label();
            this.filePathTextBox = new System.Windows.Forms.TextBox();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.status = new System.Windows.Forms.ToolStripStatusLabel();
            this.menuStrip1.SuspendLayout();
            this.bodyPanel.SuspendLayout();
            this.bodyGroupBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.templateFilePathButton)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.convertButton)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.filePathButton)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.folderPathButton)).BeginInit();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1,
            this.списокТеговToolStripMenuItem,
            this.оПрограммеToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(9, 2, 0, 2);
            this.menuStrip1.Size = new System.Drawing.Size(626, 24);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.defaultТегиToolStripMenuItem,
            this.smartTagsToolStripMenuItem,
            this.constantsToolStripMenuItem});
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(79, 20);
            this.toolStripMenuItem1.Text = "Настройки";
            // 
            // defaultТегиToolStripMenuItem
            // 
            this.defaultТегиToolStripMenuItem.Name = "defaultТегиToolStripMenuItem";
            this.defaultТегиToolStripMenuItem.Size = new System.Drawing.Size(245, 22);
            this.defaultТегиToolStripMenuItem.Text = "Обычные теги";
            this.defaultТегиToolStripMenuItem.Click += new System.EventHandler(this.DefaultТегиToolStripMenuItem_Click);
            // 
            // smartTagsToolStripMenuItem
            // 
            this.smartTagsToolStripMenuItem.Name = "smartTagsToolStripMenuItem";
            this.smartTagsToolStripMenuItem.Size = new System.Drawing.Size(245, 22);
            this.smartTagsToolStripMenuItem.Text = "\"Умные\" теги";
            this.smartTagsToolStripMenuItem.Click += new System.EventHandler(this.SmartTagsToolStripMenuItem_Click);
            // 
            // constantsToolStripMenuItem
            // 
            this.constantsToolStripMenuItem.Name = "constantsToolStripMenuItem";
            this.constantsToolStripMenuItem.Size = new System.Drawing.Size(245, 22);
            this.constantsToolStripMenuItem.Text = "Параметры получения данных";
            this.constantsToolStripMenuItem.Click += new System.EventHandler(this.ConstantsToolStripMenuItem_Click);
            // 
            // списокТеговToolStripMenuItem
            // 
            this.списокТеговToolStripMenuItem.Name = "списокТеговToolStripMenuItem";
            this.списокТеговToolStripMenuItem.Size = new System.Drawing.Size(92, 20);
            this.списокТеговToolStripMenuItem.Text = "Список тегов";
            this.списокТеговToolStripMenuItem.Click += new System.EventHandler(this.СписокТеговToolStripMenuItem_Click);
            // 
            // оПрограммеToolStripMenuItem
            // 
            this.оПрограммеToolStripMenuItem.Name = "оПрограммеToolStripMenuItem";
            this.оПрограммеToolStripMenuItem.Size = new System.Drawing.Size(94, 20);
            this.оПрограммеToolStripMenuItem.Text = "О программе";
            this.оПрограммеToolStripMenuItem.Click += new System.EventHandler(this.ОПрограммеToolStripMenuItem_Click);
            // 
            // bodyPanel
            // 
            this.bodyPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.bodyPanel.Controls.Add(this.bodyGroupBox);
            this.bodyPanel.Location = new System.Drawing.Point(0, 24);
            this.bodyPanel.Margin = new System.Windows.Forms.Padding(4);
            this.bodyPanel.Name = "bodyPanel";
            this.bodyPanel.Padding = new System.Windows.Forms.Padding(15, 7, 15, 14);
            this.bodyPanel.Size = new System.Drawing.Size(626, 352);
            this.bodyPanel.TabIndex = 2;
            // 
            // bodyGroupBox
            // 
            this.bodyGroupBox.Controls.Add(this.excelFilesLabelClear);
            this.bodyGroupBox.Controls.Add(this.excelFilesLabel);
            this.bodyGroupBox.Controls.Add(this.progressBar1);
            this.bodyGroupBox.Controls.Add(this.resultFilePrefixLabel);
            this.bodyGroupBox.Controls.Add(this.templateFilePathButton);
            this.bodyGroupBox.Controls.Add(this.resultFilePrefixTextBox);
            this.bodyGroupBox.Controls.Add(this.templateFilePathLabel);
            this.bodyGroupBox.Controls.Add(this.templateFilePathTextBox);
            this.bodyGroupBox.Controls.Add(this.convertButton);
            this.bodyGroupBox.Controls.Add(this.filePathButton);
            this.bodyGroupBox.Controls.Add(this.folderPathButton);
            this.bodyGroupBox.Controls.Add(this.resultFolderPathLabel);
            this.bodyGroupBox.Controls.Add(this.resultFolderPathTextBox);
            this.bodyGroupBox.Controls.Add(this.filePathLabel);
            this.bodyGroupBox.Controls.Add(this.filePathTextBox);
            this.bodyGroupBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.bodyGroupBox.Location = new System.Drawing.Point(15, 7);
            this.bodyGroupBox.Margin = new System.Windows.Forms.Padding(4);
            this.bodyGroupBox.Name = "bodyGroupBox";
            this.bodyGroupBox.Padding = new System.Windows.Forms.Padding(4);
            this.bodyGroupBox.Size = new System.Drawing.Size(596, 331);
            this.bodyGroupBox.TabIndex = 1;
            this.bodyGroupBox.TabStop = false;
            // 
            // excelFilesLabelClear
            // 
            this.excelFilesLabelClear.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.excelFilesLabelClear.Location = new System.Drawing.Point(479, 45);
            this.excelFilesLabelClear.Name = "excelFilesLabelClear";
            this.excelFilesLabelClear.Size = new System.Drawing.Size(77, 26);
            this.excelFilesLabelClear.TabIndex = 12;
            this.excelFilesLabelClear.TabStop = true;
            this.excelFilesLabelClear.Text = "Очистить";
            this.excelFilesLabelClear.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.excelFilesLabelClear.Visible = false;
            this.excelFilesLabelClear.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.ExcelFilesLabelClear_LinkClicked);
            // 
            // excelFilesLabel
            // 
            this.excelFilesLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.excelFilesLabel.Location = new System.Drawing.Point(10, 45);
            this.excelFilesLabel.Name = "excelFilesLabel";
            this.excelFilesLabel.Size = new System.Drawing.Size(471, 26);
            this.excelFilesLabel.TabIndex = 11;
            this.excelFilesLabel.Text = "Выбрано 0 файлов";
            this.excelFilesLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.excelFilesLabel.Visible = false;
            // 
            // progressBar1
            // 
            this.progressBar1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar1.Location = new System.Drawing.Point(13, 295);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(468, 27);
            this.progressBar1.TabIndex = 10;
            // 
            // resultFilePrefixLabel
            // 
            this.resultFilePrefixLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.resultFilePrefixLabel.AutoSize = true;
            this.resultFilePrefixLabel.Font = new System.Drawing.Font("Arial", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.resultFilePrefixLabel.Location = new System.Drawing.Point(10, 210);
            this.resultFilePrefixLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.resultFilePrefixLabel.Name = "resultFilePrefixLabel";
            this.resultFilePrefixLabel.Size = new System.Drawing.Size(173, 34);
            this.resultFilePrefixLabel.TabIndex = 5;
            this.resultFilePrefixLabel.Text = "Префикс к имени\r\nрезультирующего файла";
            // 
            // templateFilePathButton
            // 
            this.templateFilePathButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.templateFilePathButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.templateFilePathButton.Image = global::ExcelToWordProject.Properties.Resources.document;
            this.templateFilePathButton.Location = new System.Drawing.Point(562, 106);
            this.templateFilePathButton.Name = "templateFilePathButton";
            this.templateFilePathButton.Size = new System.Drawing.Size(26, 26);
            this.templateFilePathButton.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.templateFilePathButton.TabIndex = 9;
            this.templateFilePathButton.TabStop = false;
            this.templateFilePathButton.Click += new System.EventHandler(this.TemplateFilePathButton_Click);
            // 
            // resultFilePrefixTextBox
            // 
            this.resultFilePrefixTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.resultFilePrefixTextBox.Location = new System.Drawing.Point(191, 214);
            this.resultFilePrefixTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.resultFilePrefixTextBox.Name = "resultFilePrefixTextBox";
            this.resultFilePrefixTextBox.Size = new System.Drawing.Size(397, 26);
            this.resultFilePrefixTextBox.TabIndex = 4;
            // 
            // templateFilePathLabel
            // 
            this.templateFilePathLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.templateFilePathLabel.AutoSize = true;
            this.templateFilePathLabel.Location = new System.Drawing.Point(10, 84);
            this.templateFilePathLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.templateFilePathLabel.Name = "templateFilePathLabel";
            this.templateFilePathLabel.Size = new System.Drawing.Size(217, 18);
            this.templateFilePathLabel.TabIndex = 8;
            this.templateFilePathLabel.Text = "Пусть к Word файлу-шаблону";
            // 
            // templateFilePathTextBox
            // 
            this.templateFilePathTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.templateFilePathTextBox.Location = new System.Drawing.Point(13, 106);
            this.templateFilePathTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.templateFilePathTextBox.Name = "templateFilePathTextBox";
            this.templateFilePathTextBox.Size = new System.Drawing.Size(542, 26);
            this.templateFilePathTextBox.TabIndex = 7;
            // 
            // convertButton
            // 
            this.convertButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.convertButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.convertButton.Image = global::ExcelToWordProject.Properties.Resources.convert;
            this.convertButton.Location = new System.Drawing.Point(525, 258);
            this.convertButton.Name = "convertButton";
            this.convertButton.Size = new System.Drawing.Size(64, 64);
            this.convertButton.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.convertButton.TabIndex = 6;
            this.convertButton.TabStop = false;
            this.convertButton.Click += new System.EventHandler(this.ConvertButton_Click);
            // 
            // filePathButton
            // 
            this.filePathButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.filePathButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.filePathButton.Image = global::ExcelToWordProject.Properties.Resources.document;
            this.filePathButton.Location = new System.Drawing.Point(562, 45);
            this.filePathButton.Name = "filePathButton";
            this.filePathButton.Size = new System.Drawing.Size(26, 26);
            this.filePathButton.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.filePathButton.TabIndex = 5;
            this.filePathButton.TabStop = false;
            this.filePathButton.Click += new System.EventHandler(this.FilePathButton_Click);
            // 
            // folderPathButton
            // 
            this.folderPathButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.folderPathButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.folderPathButton.Image = global::ExcelToWordProject.Properties.Resources.folder;
            this.folderPathButton.Location = new System.Drawing.Point(562, 167);
            this.folderPathButton.Name = "folderPathButton";
            this.folderPathButton.Size = new System.Drawing.Size(26, 26);
            this.folderPathButton.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.folderPathButton.TabIndex = 4;
            this.folderPathButton.TabStop = false;
            this.folderPathButton.Click += new System.EventHandler(this.FolderPathButton_Click);
            // 
            // resultFolderPathLabel
            // 
            this.resultFolderPathLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.resultFolderPathLabel.AutoSize = true;
            this.resultFolderPathLabel.Location = new System.Drawing.Point(10, 145);
            this.resultFolderPathLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.resultFolderPathLabel.Name = "resultFolderPathLabel";
            this.resultFolderPathLabel.Size = new System.Drawing.Size(221, 18);
            this.resultFolderPathLabel.TabIndex = 3;
            this.resultFolderPathLabel.Text = "Путь к результирующей папке";
            // 
            // resultFolderPathTextBox
            // 
            this.resultFolderPathTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.resultFolderPathTextBox.Location = new System.Drawing.Point(13, 167);
            this.resultFolderPathTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.resultFolderPathTextBox.Name = "resultFolderPathTextBox";
            this.resultFolderPathTextBox.Size = new System.Drawing.Size(542, 26);
            this.resultFolderPathTextBox.TabIndex = 2;
            // 
            // filePathLabel
            // 
            this.filePathLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.filePathLabel.AutoSize = true;
            this.filePathLabel.Location = new System.Drawing.Point(10, 23);
            this.filePathLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.filePathLabel.Name = "filePathLabel";
            this.filePathLabel.Size = new System.Drawing.Size(224, 18);
            this.filePathLabel.TabIndex = 1;
            this.filePathLabel.Text = "Путь к исходному Excel файлу";
            // 
            // filePathTextBox
            // 
            this.filePathTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.filePathTextBox.Location = new System.Drawing.Point(13, 45);
            this.filePathTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.filePathTextBox.Name = "filePathTextBox";
            this.filePathTextBox.Size = new System.Drawing.Size(542, 26);
            this.filePathTextBox.TabIndex = 0;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.status});
            this.statusStrip1.Location = new System.Drawing.Point(0, 380);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(626, 22);
            this.statusStrip1.TabIndex = 3;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // status
            // 
            this.status.Name = "status";
            this.status.Size = new System.Drawing.Size(73, 17);
            this.status.Text = "Ожидание...";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(626, 402);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.bodyPanel);
            this.Controls.Add(this.menuStrip1);
            this.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MinimumSize = new System.Drawing.Size(642, 441);
            this.Name = "MainForm";
            this.Text = "Граббер данных из Excel";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.bodyPanel.ResumeLayout(false);
            this.bodyGroupBox.ResumeLayout(false);
            this.bodyGroupBox.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.templateFilePathButton)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.convertButton)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.filePathButton)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.folderPathButton)).EndInit();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem defaultТегиToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem smartTagsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem constantsToolStripMenuItem;
        private System.Windows.Forms.Panel bodyPanel;
        private System.Windows.Forms.GroupBox bodyGroupBox;
        private System.Windows.Forms.TextBox filePathTextBox;
        private System.Windows.Forms.Label filePathLabel;
        private System.Windows.Forms.Label resultFolderPathLabel;
        private System.Windows.Forms.TextBox resultFolderPathTextBox;
        private System.Windows.Forms.PictureBox folderPathButton;
        private System.Windows.Forms.PictureBox filePathButton;
        private System.Windows.Forms.PictureBox convertButton;
        private System.Windows.Forms.PictureBox templateFilePathButton;
        private System.Windows.Forms.Label templateFilePathLabel;
        private System.Windows.Forms.TextBox templateFilePathTextBox;
        private System.Windows.Forms.Label resultFilePrefixLabel;
        private System.Windows.Forms.TextBox resultFilePrefixTextBox;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel status;
        private System.Windows.Forms.ToolStripMenuItem списокТеговToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem оПрограммеToolStripMenuItem;
        private System.Windows.Forms.LinkLabel excelFilesLabelClear;
        private System.Windows.Forms.Label excelFilesLabel;
    }
}

