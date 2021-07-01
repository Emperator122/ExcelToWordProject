namespace ExcelToWordProject
{
    partial class TagListForm
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
            this.tagsPanel = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.topMostCheckBox = new System.Windows.Forms.CheckBox();
            this.hideInactiveCheckBox = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // tagsPanel
            // 
            this.tagsPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tagsPanel.AutoScroll = true;
            this.tagsPanel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.tagsPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tagsPanel.Location = new System.Drawing.Point(13, 45);
            this.tagsPanel.Margin = new System.Windows.Forms.Padding(4);
            this.tagsPanel.Name = "tagsPanel";
            this.tagsPanel.Size = new System.Drawing.Size(541, 448);
            this.tagsPanel.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.Font = new System.Drawing.Font("Arial", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(13, 9);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(541, 32);
            this.label1.TabIndex = 2;
            this.label1.Text = "Общий список тегов";
            this.label1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // topMostCheckBox
            // 
            this.topMostCheckBox.AutoSize = true;
            this.topMostCheckBox.Location = new System.Drawing.Point(13, 500);
            this.topMostCheckBox.Name = "topMostCheckBox";
            this.topMostCheckBox.Size = new System.Drawing.Size(155, 22);
            this.topMostCheckBox.TabIndex = 3;
            this.topMostCheckBox.Text = "Поверх всех окон";
            this.topMostCheckBox.UseVisualStyleBackColor = true;
            this.topMostCheckBox.CheckedChanged += new System.EventHandler(this.TopMostCheckBox_CheckedChanged);
            // 
            // hideInactiveCheckBox
            // 
            this.hideInactiveCheckBox.AutoSize = true;
            this.hideInactiveCheckBox.Location = new System.Drawing.Point(367, 500);
            this.hideInactiveCheckBox.Name = "hideInactiveCheckBox";
            this.hideInactiveCheckBox.Size = new System.Drawing.Size(187, 22);
            this.hideInactiveCheckBox.TabIndex = 4;
            this.hideInactiveCheckBox.Text = "Спрятать неактивные";
            this.hideInactiveCheckBox.UseVisualStyleBackColor = true;
            this.hideInactiveCheckBox.CheckedChanged += new System.EventHandler(this.HideInactiveCheckBox_CheckedChanged);
            // 
            // TagListForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(567, 527);
            this.Controls.Add(this.hideInactiveCheckBox);
            this.Controls.Add(this.topMostCheckBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tagsPanel);
            this.DoubleBuffered = true;
            this.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.Name = "TagListForm";
            this.ShowIcon = false;
            this.Text = "Список тегов";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel tagsPanel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox topMostCheckBox;
        private System.Windows.Forms.CheckBox hideInactiveCheckBox;
    }
}