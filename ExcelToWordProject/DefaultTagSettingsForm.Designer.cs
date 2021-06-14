namespace ExcelToWordProject
{
    partial class DefaultTagSettingsForm
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
            this.defaultTagsGroupBox = new System.Windows.Forms.GroupBox();
            this.addButton = new System.Windows.Forms.PictureBox();
            this.defaultTagsPanel = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.defaultTagsGroupBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.addButton)).BeginInit();
            this.SuspendLayout();
            // 
            // defaultTagsGroupBox
            // 
            this.defaultTagsGroupBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.defaultTagsGroupBox.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.defaultTagsGroupBox.Controls.Add(this.addButton);
            this.defaultTagsGroupBox.Controls.Add(this.defaultTagsPanel);
            this.defaultTagsGroupBox.Location = new System.Drawing.Point(21, 10);
            this.defaultTagsGroupBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.defaultTagsGroupBox.Name = "defaultTagsGroupBox";
            this.defaultTagsGroupBox.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.defaultTagsGroupBox.Size = new System.Drawing.Size(1285, 482);
            this.defaultTagsGroupBox.TabIndex = 1;
            this.defaultTagsGroupBox.TabStop = false;
            this.defaultTagsGroupBox.Text = "Список тегов";
            // 
            // addButton
            // 
            this.addButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.addButton.Image = global::ExcelToWordProject.Properties.Resources.plus;
            this.addButton.Location = new System.Drawing.Point(1213, 410);
            this.addButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.addButton.Name = "addButton";
            this.addButton.Size = new System.Drawing.Size(64, 64);
            this.addButton.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.addButton.TabIndex = 1;
            this.addButton.TabStop = false;
            this.addButton.Click += new System.EventHandler(this.AddButton_Click);
            // 
            // defaultTagsPanel
            // 
            this.defaultTagsPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.defaultTagsPanel.AutoScroll = true;
            this.defaultTagsPanel.Location = new System.Drawing.Point(8, 22);
            this.defaultTagsPanel.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.defaultTagsPanel.Name = "defaultTagsPanel";
            this.defaultTagsPanel.Size = new System.Drawing.Size(1268, 380);
            this.defaultTagsPanel.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Location = new System.Drawing.Point(1036, 500);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(270, 42);
            this.button1.TabIndex = 2;
            this.button1.Text = "Сохранить и закрыть";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // DefaultTagSettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1311, 555);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.defaultTagsGroupBox);
            this.DoubleBuffered = true;
            this.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "DefaultTagSettingsForm";
            this.Text = "Настройка обычных тегов";
            this.Resize += new System.EventHandler(this.DefaultTagSettingsForm_Resize);
            this.defaultTagsGroupBox.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.addButton)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox defaultTagsGroupBox;
        private System.Windows.Forms.Panel defaultTagsPanel;
        private System.Windows.Forms.PictureBox addButton;
        private System.Windows.Forms.Button button1;
    }
}