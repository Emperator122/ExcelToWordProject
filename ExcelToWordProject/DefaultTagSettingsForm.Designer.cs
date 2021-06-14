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
            this.defaultTagsGroupBox.Location = new System.Drawing.Point(14, 7);
            this.defaultTagsGroupBox.Name = "defaultTagsGroupBox";
            this.defaultTagsGroupBox.Size = new System.Drawing.Size(883, 423);
            this.defaultTagsGroupBox.TabIndex = 1;
            this.defaultTagsGroupBox.TabStop = false;
            this.defaultTagsGroupBox.Text = "Список тегов";
            // 
            // addButton
            // 
            this.addButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.addButton.Image = global::ExcelToWordProject.Properties.Resources.plus;
            this.addButton.Location = new System.Drawing.Point(813, 353);
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
            this.defaultTagsPanel.Location = new System.Drawing.Point(5, 16);
            this.defaultTagsPanel.Name = "defaultTagsPanel";
            this.defaultTagsPanel.Size = new System.Drawing.Size(872, 331);
            this.defaultTagsPanel.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Location = new System.Drawing.Point(710, 436);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(187, 30);
            this.button1.TabIndex = 2;
            this.button1.Text = "Сохранить и закрыть";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // DefaultTagSettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(909, 480);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.defaultTagsGroupBox);
            this.DoubleBuffered = true;
            this.Name = "DefaultTagSettingsForm";
            this.Text = "DefaultTagSettingsForm";
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