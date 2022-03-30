using System;
using System.IO;
using System.Windows.Forms;

namespace ExcelToWordProject.Forms
{
    partial class TextBlockAddForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void BuildToolTips()
        {
            toolTip1.ReshowDelay = 1;
            toolTip1.InitialDelay = 1;
            toolTip1.AutoPopDelay = 10000;
            toolTip1.SetToolTip(tagKeyTextBox, "Строка. Ключ, который будет использоваться в шаблоне (без треугольных скобок).");
            toolTip1.SetToolTip(delimiterTextBox, "Строка. Разделитель, по которому будет разбито значение блока текста, для переноса по абзацам в документе. " +
                                                  "Для многострочного текста рекомендуется использовать \\r\\n (либо \\n). " +
                                                  "Для работы тег должен быть единственной строкой в абзаце шаблона.");
            toolTip1.SetToolTip(priorityTextBox, "Целое число. Приоритет тега в рамках данного блока текста при выводе. Нужен если мы имеем два набора условий, " +
                                                 "которые одновременно выполняются. Значение по умолчанию всегда имеет самый низкий приоритет.");
            toolTip1.SetToolTip(groupBox3, "При выполнении данных условий, будет произведена замена тега в шаблоне на его содержимое.\r\n" +
                                           "Ключ тега - ключ \"умного\" или обычного тега, на значение которого налагется условие;\r\n" +
                                           "Значение  - значение \"умного\" или обычного тега, соответсвующего ключу, при котором будет произведена замена " +
                                           "на содержимое блока текста. Значение пишетя без разделителя;\r\n" +
                                           "Разделитель - строка (или символ), по которой должно быть разбито значение \"умного\" или обычного тега " +
                                           "при проверке условий. " +
                                           "Например, пусть \"умный\" или обычный тег может принять значение 'val1;val2;val3;val4' " +
                                           "(в любой из комбинаций). Тогда для того чтобы наложить условие на значение 'val1', мы указываем разделитель " +
                                           "';', а в поле 'Значение' пишем 'val1'.");

            toolTip1.SetToolTip(groupBox2, "Значение, которое принимает блок текста. Может быть как строкой, так и содержимым внешнего .docx документа. " +
                                           "На строку замена может быть произведена в любом месте документа. На разделенную строку или на содержимое " +
                                           "внешенго docx замена может быть произведена только в отдельном абзаце.");
        }

        private void OnTextBlockValueModeSwitch()
        {
            delimiterTextBox.Enabled = !IsFilePathSetting;

        }

        private void ShowPathError()
        {
            MessageBox.Show("Ошибка пути к документу. Документы должны храниться в папке \"TextBlocksDocuments\" " +
                            $"из директории программы ({Path.Combine(AppContext.BaseDirectory, "TextBlocksDocuments")}).");
        }


        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.tagKeyTextBox = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.valueModeTabs = new System.Windows.Forms.TabControl();
            this.textValueTab = new System.Windows.Forms.TabPage();
            this.tagValueTextBox = new System.Windows.Forms.TextBox();
            this.filePathTab = new System.Windows.Forms.TabPage();
            this.label2 = new System.Windows.Forms.Label();
            this.templateFilePathButton = new System.Windows.Forms.PictureBox();
            this.templateFilePathLabel = new System.Windows.Forms.Label();
            this.templateFilePathTextBox = new System.Windows.Forms.TextBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.conditionsGridView = new System.Windows.Forms.DataGridView();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.priorityTextBox = new System.Windows.Forms.TextBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.delimiterTextBox = new System.Windows.Forms.TextBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.TagNameColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TagValueColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TagDelimiterColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.valueModeTabs.SuspendLayout();
            this.textValueTab.SuspendLayout();
            this.filePathTab.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.templateFilePathButton)).BeginInit();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.conditionsGridView)).BeginInit();
            this.groupBox4.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.SuspendLayout();
            // 
            // tagKeyTextBox
            // 
            this.tagKeyTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tagKeyTextBox.Location = new System.Drawing.Point(9, 24);
            this.tagKeyTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.tagKeyTextBox.Name = "tagKeyTextBox";
            this.tagKeyTextBox.Size = new System.Drawing.Size(339, 26);
            this.tagKeyTextBox.TabIndex = 0;
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.tagKeyTextBox);
            this.groupBox1.Location = new System.Drawing.Point(13, 13);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox1.Size = new System.Drawing.Size(359, 69);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Ключ";
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.valueModeTabs);
            this.groupBox2.Location = new System.Drawing.Point(13, 398);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox2.Size = new System.Drawing.Size(595, 346);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Содержимое блока текста";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(299, 18);
            this.label1.TabIndex = 1;
            this.label1.Text = "Вы можете выбрать один из вариантов:";
            // 
            // valueModeTabs
            // 
            this.valueModeTabs.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.valueModeTabs.Controls.Add(this.textValueTab);
            this.valueModeTabs.Controls.Add(this.filePathTab);
            this.valueModeTabs.Location = new System.Drawing.Point(0, 55);
            this.valueModeTabs.Name = "valueModeTabs";
            this.valueModeTabs.SelectedIndex = 0;
            this.valueModeTabs.Size = new System.Drawing.Size(594, 291);
            this.valueModeTabs.TabIndex = 0;
            this.valueModeTabs.SelectedIndexChanged += new System.EventHandler(this.valueModeTabs_SelectedIndexChanged);
            // 
            // textValueTab
            // 
            this.textValueTab.Controls.Add(this.tagValueTextBox);
            this.textValueTab.Location = new System.Drawing.Point(4, 27);
            this.textValueTab.Name = "textValueTab";
            this.textValueTab.Padding = new System.Windows.Forms.Padding(3);
            this.textValueTab.Size = new System.Drawing.Size(586, 260);
            this.textValueTab.TabIndex = 0;
            this.textValueTab.Text = "Текстовое значение";
            this.textValueTab.UseVisualStyleBackColor = true;
            // 
            // tagValueTextBox
            // 
            this.tagValueTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tagValueTextBox.Location = new System.Drawing.Point(7, 4);
            this.tagValueTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.tagValueTextBox.MaxLength = 0;
            this.tagValueTextBox.Multiline = true;
            this.tagValueTextBox.Name = "tagValueTextBox";
            this.tagValueTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.tagValueTextBox.Size = new System.Drawing.Size(572, 249);
            this.tagValueTextBox.TabIndex = 0;
            this.tagValueTextBox.WordWrap = false;
            // 
            // filePathTab
            // 
            this.filePathTab.Controls.Add(this.label2);
            this.filePathTab.Controls.Add(this.templateFilePathButton);
            this.filePathTab.Controls.Add(this.templateFilePathLabel);
            this.filePathTab.Controls.Add(this.templateFilePathTextBox);
            this.filePathTab.Location = new System.Drawing.Point(4, 27);
            this.filePathTab.Name = "filePathTab";
            this.filePathTab.Padding = new System.Windows.Forms.Padding(3);
            this.filePathTab.Size = new System.Drawing.Size(586, 260);
            this.filePathTab.TabIndex = 1;
            this.filePathTab.Text = "Путь к файлу";
            this.filePathTab.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.SystemColors.ControlDark;
            this.label2.Location = new System.Drawing.Point(7, 73);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(430, 18);
            this.label2.TabIndex = 13;
            this.label2.Text = "*  Для замены тег должен находиться в отдельном абзаце";
            // 
            // templateFilePathButton
            // 
            this.templateFilePathButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.templateFilePathButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.templateFilePathButton.Image = global::ExcelToWordProject.Properties.Resources.document;
            this.templateFilePathButton.Location = new System.Drawing.Point(554, 43);
            this.templateFilePathButton.Name = "templateFilePathButton";
            this.templateFilePathButton.Size = new System.Drawing.Size(26, 26);
            this.templateFilePathButton.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.templateFilePathButton.TabIndex = 12;
            this.templateFilePathButton.TabStop = false;
            this.templateFilePathButton.Click += new System.EventHandler(this.templateFilePathButton_Click);
            // 
            // templateFilePathLabel
            // 
            this.templateFilePathLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.templateFilePathLabel.AutoEllipsis = true;
            this.templateFilePathLabel.AutoSize = true;
            this.templateFilePathLabel.Location = new System.Drawing.Point(7, 21);
            this.templateFilePathLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.templateFilePathLabel.Name = "templateFilePathLabel";
            this.templateFilePathLabel.Size = new System.Drawing.Size(465, 18);
            this.templateFilePathLabel.TabIndex = 11;
            this.templateFilePathLabel.Text = "Путь к Word файлу, на содержимое которого будет заменен тег";
            // 
            // templateFilePathTextBox
            // 
            this.templateFilePathTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.templateFilePathTextBox.Location = new System.Drawing.Point(7, 43);
            this.templateFilePathTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.templateFilePathTextBox.Name = "templateFilePathTextBox";
            this.templateFilePathTextBox.ReadOnly = true;
            this.templateFilePathTextBox.Size = new System.Drawing.Size(540, 26);
            this.templateFilePathTextBox.TabIndex = 10;
            // 
            // groupBox3
            // 
            this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox3.Controls.Add(this.conditionsGridView);
            this.groupBox3.Location = new System.Drawing.Point(13, 86);
            this.groupBox3.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox3.Size = new System.Drawing.Size(595, 307);
            this.groupBox3.TabIndex = 4;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Условия";
            // 
            // conditionsGridView
            // 
            this.conditionsGridView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.conditionsGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.conditionsGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.TagNameColumn,
            this.TagValueColumn,
            this.TagDelimiterColumn});
            this.conditionsGridView.Location = new System.Drawing.Point(9, 26);
            this.conditionsGridView.Name = "conditionsGridView";
            this.conditionsGridView.Size = new System.Drawing.Size(575, 274);
            this.conditionsGridView.TabIndex = 1;
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Location = new System.Drawing.Point(13, 751);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(594, 37);
            this.button1.TabIndex = 5;
            this.button1.Text = "Сохранить";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // groupBox4
            // 
            this.groupBox4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox4.Controls.Add(this.priorityTextBox);
            this.groupBox4.Location = new System.Drawing.Point(507, 13);
            this.groupBox4.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox4.Size = new System.Drawing.Size(101, 69);
            this.groupBox4.TabIndex = 3;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Приоритет";
            // 
            // priorityTextBox
            // 
            this.priorityTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.priorityTextBox.Location = new System.Drawing.Point(9, 24);
            this.priorityTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.priorityTextBox.Name = "priorityTextBox";
            this.priorityTextBox.Size = new System.Drawing.Size(81, 26);
            this.priorityTextBox.TabIndex = 0;
            // 
            // groupBox5
            // 
            this.groupBox5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox5.Controls.Add(this.delimiterTextBox);
            this.groupBox5.Location = new System.Drawing.Point(380, 13);
            this.groupBox5.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox5.Size = new System.Drawing.Size(119, 69);
            this.groupBox5.TabIndex = 4;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Разделитель";
            // 
            // delimiterTextBox
            // 
            this.delimiterTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.delimiterTextBox.Location = new System.Drawing.Point(9, 24);
            this.delimiterTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.delimiterTextBox.Name = "delimiterTextBox";
            this.delimiterTextBox.Size = new System.Drawing.Size(99, 26);
            this.delimiterTextBox.TabIndex = 0;
            // 
            // toolTip1
            // 
            this.toolTip1.ShowAlways = true;
            // 
            // TagNameColumn
            // 
            this.TagNameColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.TagNameColumn.FillWeight = 40F;
            this.TagNameColumn.HeaderText = "Ключ тега";
            this.TagNameColumn.Name = "TagNameColumn";
            // 
            // TagValueColumn
            // 
            this.TagValueColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.TagValueColumn.FillWeight = 40F;
            this.TagValueColumn.HeaderText = "Значение";
            this.TagValueColumn.Name = "TagValueColumn";
            // 
            // TagDelimiterColumn
            // 
            this.TagDelimiterColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.TagDelimiterColumn.FillWeight = 20F;
            this.TagDelimiterColumn.HeaderText = "Разделитель";
            this.TagDelimiterColumn.Name = "TagDelimiterColumn";
            // 
            // TextBlockAddForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(621, 803);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "TextBlockAddForm";
            this.ShowIcon = false;
            this.Text = "Добавление блока текста";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.valueModeTabs.ResumeLayout(false);
            this.textValueTab.ResumeLayout(false);
            this.textValueTab.PerformLayout();
            this.filePathTab.ResumeLayout(false);
            this.filePathTab.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.templateFilePathButton)).EndInit();
            this.groupBox3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.conditionsGridView)).EndInit();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TextBox tagKeyTextBox;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox tagValueTextBox;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.DataGridView conditionsGridView;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.TextBox priorityTextBox;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.TextBox delimiterTextBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TabControl valueModeTabs;
        private System.Windows.Forms.TabPage textValueTab;
        private System.Windows.Forms.TabPage filePathTab;
        private System.Windows.Forms.PictureBox templateFilePathButton;
        private System.Windows.Forms.Label templateFilePathLabel;
        private System.Windows.Forms.TextBox templateFilePathTextBox;
        private System.Windows.Forms.Label label2;
        private ToolTip toolTip1;
        private DataGridViewTextBoxColumn TagNameColumn;
        private DataGridViewTextBoxColumn TagValueColumn;
        private DataGridViewTextBoxColumn TagDelimiterColumn;
    }
}
