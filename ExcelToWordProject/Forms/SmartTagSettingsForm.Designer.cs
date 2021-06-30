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
            this.button1 = new System.Windows.Forms.Button();
            this.tagActivator = new ExcelToWordProject.Components.MyTreeView();
            this.smartTagsPanel = new System.Windows.Forms.Panel();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Location = new System.Drawing.Point(800, 475);
            this.button1.Margin = new System.Windows.Forms.Padding(4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(280, 41);
            this.button1.TabIndex = 1;
            this.button1.Text = "Сохранить и закрыть";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // tagActivator
            // 
            this.tagActivator.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.tagActivator.BackColor = System.Drawing.SystemColors.Control;
            this.tagActivator.CheckBoxes = true;
            this.tagActivator.Location = new System.Drawing.Point(18, 457);
            this.tagActivator.Margin = new System.Windows.Forms.Padding(4);
            this.tagActivator.Name = "tagActivator";
            this.tagActivator.Size = new System.Drawing.Size(457, 71);
            this.tagActivator.TabIndex = 2;
            this.tagActivator.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.TagActivator_AfterCheck);
            // 
            // smartTagsPanel
            // 
            this.smartTagsPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.smartTagsPanel.AutoScroll = true;
            this.smartTagsPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.smartTagsPanel.Location = new System.Drawing.Point(18, 13);
            this.smartTagsPanel.Margin = new System.Windows.Forms.Padding(4);
            this.smartTagsPanel.Name = "smartTagsPanel";
            this.smartTagsPanel.Size = new System.Drawing.Size(1062, 436);
            this.smartTagsPanel.TabIndex = 3;
            // 
            // SmartTagSettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1099, 534);
            this.Controls.Add(this.smartTagsPanel);
            this.Controls.Add(this.tagActivator);
            this.Controls.Add(this.button1);
            this.DoubleBuffered = true;
            this.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "SmartTagSettingsForm";
            this.ShowIcon = false;
            this.Text = "Настройка \"умных\" тегов";
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button button1;
        private Components.MyTreeView tagActivator;
        private System.Windows.Forms.Panel smartTagsPanel;
    }
}