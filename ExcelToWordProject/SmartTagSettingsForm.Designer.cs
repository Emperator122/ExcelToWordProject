﻿namespace ExcelToWordProject
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
            this.smartTagsGroupBox.Location = new System.Drawing.Point(18, 13);
            this.smartTagsGroupBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.smartTagsGroupBox.Name = "smartTagsGroupBox";
            this.smartTagsGroupBox.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.smartTagsGroupBox.Size = new System.Drawing.Size(1063, 439);
            this.smartTagsGroupBox.TabIndex = 0;
            this.smartTagsGroupBox.TabStop = false;
            this.smartTagsGroupBox.Text = "\"Умные\" теги";
            // 
            // smartTagsPanel
            // 
            this.smartTagsPanel.AutoScroll = true;
            this.smartTagsPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.smartTagsPanel.Location = new System.Drawing.Point(4, 23);
            this.smartTagsPanel.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.smartTagsPanel.Name = "smartTagsPanel";
            this.smartTagsPanel.Size = new System.Drawing.Size(1055, 412);
            this.smartTagsPanel.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Location = new System.Drawing.Point(800, 475);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
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
            this.tagActivator.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.tagActivator.Name = "tagActivator";
            this.tagActivator.Size = new System.Drawing.Size(457, 71);
            this.tagActivator.TabIndex = 2;
            this.tagActivator.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.TagActivator_AfterCheck);
            // 
            // SmartTagSettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1099, 534);
            this.Controls.Add(this.tagActivator);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.smartTagsGroupBox);
            this.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "SmartTagSettingsForm";
            this.Text = "Настройка \"умных\" тегов";
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