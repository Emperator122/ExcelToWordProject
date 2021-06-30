namespace ExcelToWordProject.Forms
{
    partial class RegularExpressionEditForm
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.regExpTextBox = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.regexSettings = new System.Windows.Forms.CheckedListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.regexGroupIndexTextBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupsCountLabel = new System.Windows.Forms.Label();
            this.testRegExpButton = new System.Windows.Forms.Button();
            this.testRegExpTextBox = new System.Windows.Forms.TextBox();
            this.CancelButton1 = new System.Windows.Forms.Button();
            this.SaveButton = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // regExpTextBox
            // 
            this.regExpTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.regExpTextBox.Location = new System.Drawing.Point(7, 26);
            this.regExpTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.regExpTextBox.Name = "regExpTextBox";
            this.regExpTextBox.Size = new System.Drawing.Size(532, 26);
            this.regExpTextBox.TabIndex = 0;
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.regexSettings);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.regexGroupIndexTextBox);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.regExpTextBox);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(546, 187);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Регулярное выражение";
            // 
            // regexSettings
            // 
            this.regexSettings.FormattingEnabled = true;
            this.regexSettings.Location = new System.Drawing.Point(303, 89);
            this.regexSettings.Name = "regexSettings";
            this.regexSettings.Size = new System.Drawing.Size(236, 88);
            this.regexSettings.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(300, 67);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(93, 18);
            this.label2.TabIndex = 3;
            this.label2.Text = "Параметры";
            // 
            // regexGroupIndexTextBox
            // 
            this.regexGroupIndexTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.regexGroupIndexTextBox.Location = new System.Drawing.Point(7, 89);
            this.regexGroupIndexTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.regexGroupIndexTextBox.Name = "regexGroupIndexTextBox";
            this.regexGroupIndexTextBox.Size = new System.Drawing.Size(188, 26);
            this.regexGroupIndexTextBox.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(4, 67);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(115, 18);
            this.label1.TabIndex = 1;
            this.label1.Text = "Индекс группы";
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.groupsCountLabel);
            this.groupBox2.Controls.Add(this.testRegExpButton);
            this.groupBox2.Controls.Add(this.testRegExpTextBox);
            this.groupBox2.Location = new System.Drawing.Point(12, 219);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(546, 232);
            this.groupBox2.TabIndex = 2;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Тестовый текст (для проверки)";
            // 
            // groupsCountLabel
            // 
            this.groupsCountLabel.AutoSize = true;
            this.groupsCountLabel.Location = new System.Drawing.Point(6, 163);
            this.groupsCountLabel.Name = "groupsCountLabel";
            this.groupsCountLabel.Size = new System.Drawing.Size(141, 18);
            this.groupsCountLabel.TabIndex = 6;
            this.groupsCountLabel.Text = "Количетсво групп: ";
            // 
            // testRegExpButton
            // 
            this.testRegExpButton.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.testRegExpButton.Location = new System.Drawing.Point(7, 196);
            this.testRegExpButton.Name = "testRegExpButton";
            this.testRegExpButton.Size = new System.Drawing.Size(532, 30);
            this.testRegExpButton.TabIndex = 1;
            this.testRegExpButton.Text = "Проверить";
            this.testRegExpButton.UseVisualStyleBackColor = true;
            this.testRegExpButton.Click += new System.EventHandler(this.TestRegExpButton_Click);
            // 
            // testRegExpTextBox
            // 
            this.testRegExpTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.testRegExpTextBox.Location = new System.Drawing.Point(7, 26);
            this.testRegExpTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.testRegExpTextBox.Multiline = true;
            this.testRegExpTextBox.Name = "testRegExpTextBox";
            this.testRegExpTextBox.Size = new System.Drawing.Size(532, 119);
            this.testRegExpTextBox.TabIndex = 0;
            // 
            // CancelButton1
            // 
            this.CancelButton1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CancelButton1.Location = new System.Drawing.Point(341, 458);
            this.CancelButton1.Margin = new System.Windows.Forms.Padding(4);
            this.CancelButton1.Name = "CancelButton1";
            this.CancelButton1.Size = new System.Drawing.Size(217, 31);
            this.CancelButton1.TabIndex = 4;
            this.CancelButton1.Text = "Отмена";
            this.CancelButton1.UseVisualStyleBackColor = true;
            this.CancelButton1.Click += new System.EventHandler(this.CancelButton1_Click);
            // 
            // SaveButton
            // 
            this.SaveButton.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.SaveButton.Location = new System.Drawing.Point(12, 458);
            this.SaveButton.Margin = new System.Windows.Forms.Padding(4);
            this.SaveButton.Name = "SaveButton";
            this.SaveButton.Size = new System.Drawing.Size(217, 31);
            this.SaveButton.TabIndex = 3;
            this.SaveButton.Text = "Сохранить и закрыть";
            this.SaveButton.UseVisualStyleBackColor = true;
            this.SaveButton.Click += new System.EventHandler(this.SaveButton_Click);
            // 
            // RegularExpressionEditForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(570, 499);
            this.Controls.Add(this.CancelButton1);
            this.Controls.Add(this.SaveButton);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "RegularExpressionEditForm";
            this.ShowIcon = false;
            this.Text = "Редактирование регулярного выражения";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TextBox regExpTextBox;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button testRegExpButton;
        private System.Windows.Forms.TextBox testRegExpTextBox;
        private System.Windows.Forms.Button CancelButton1;
        private System.Windows.Forms.Button SaveButton;
        private System.Windows.Forms.TextBox regexGroupIndexTextBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckedListBox regexSettings;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label groupsCountLabel;
    }
}