namespace ExcelToWordProject
{
    partial class SmartTagSettingsForm
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
            this.smartTagsGroupBox = new System.Windows.Forms.GroupBox();
            this.smartTagsPanel = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.tagActivator = new ExcelToWordProject.Components.MyTreeView();
            this.smartTagsGroupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // smartTagsGroupBox
            // 
            this.smartTagsGroupBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.smartTagsGroupBox.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.smartTagsGroupBox.Controls.Add(this.smartTagsPanel);
            this.smartTagsGroupBox.Location = new System.Drawing.Point(12, 10);
            this.smartTagsGroupBox.Name = "smartTagsGroupBox";
            this.smartTagsGroupBox.Size = new System.Drawing.Size(559, 339);
            this.smartTagsGroupBox.TabIndex = 0;
            this.smartTagsGroupBox.TabStop = false;
            this.smartTagsGroupBox.Text = "\"Умные\" теги";
            // 
            // smartTagsPanel
            // 
            this.smartTagsPanel.AutoScroll = true;
            this.smartTagsPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.smartTagsPanel.Location = new System.Drawing.Point(3, 16);
            this.smartTagsPanel.Name = "smartTagsPanel";
            this.smartTagsPanel.Size = new System.Drawing.Size(553, 320);
            this.smartTagsPanel.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Location = new System.Drawing.Point(384, 365);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(187, 30);
            this.button1.TabIndex = 1;
            this.button1.Text = "Сохранить и закрыть";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // tagActivator
            // 
            this.tagActivator.CheckBoxes = true;
            this.tagActivator.Location = new System.Drawing.Point(12, 352);
            this.tagActivator.Name = "tagActivator";
            this.tagActivator.Size = new System.Drawing.Size(306, 52);
            this.tagActivator.TabIndex = 2;
            this.tagActivator.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.TagActivator_AfterCheck);
            // 
            // SmartTagSettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(583, 407);
            this.Controls.Add(this.tagActivator);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.smartTagsGroupBox);
            this.Name = "SmartTagSettingsForm";
            this.Text = "SmartTagSettingsForm";
            this.smartTagsGroupBox.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox smartTagsGroupBox;
        private System.Windows.Forms.Panel smartTagsPanel;
        private System.Windows.Forms.Button button1;
        private Components.MyTreeView tagActivator;
    }
}