using System;
using System.Drawing;
using System.Windows.Forms;
using ExcelToWordProject.Syllabus.Tags;

namespace ExcelToWordProject.Forms
{
    partial class TextBlockConditionsForm
    {
        private const int TextBlockPanelHeight = 64;
        private const int FormWidth = 400;
        private const int IconButtonsSize = 32;
        private readonly Bitmap _arrowIcon = Properties.Resources.right_arrow;
        private readonly Bitmap _removeIcon = Properties.Resources.remove;
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

        private void InitializeCustomComponents()
        {
            Text = _tagName;
            BuildTextBlockTagsLayout();
        }

        private void BuildTextBlockTagsLayout()
        {
            TableLayoutPanel tableLayoutPanel1 = new TableLayoutPanel()
            {
                RowCount = 0,
                ColumnCount = 0,
                Top = 0,
                Left = 0,
                Height = textBlocksWrapper.Height,
                Width = textBlocksWrapper.Width,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowOnly,
                Margin = new Padding(5, 5, 5, 5),
            };

            foreach (var textBlockTag in _tags)
            {
                tableLayoutPanel1.Controls.Add(GetTextBlockPanel(textBlockTag));
            }

            textBlocksWrapper.Controls.Add(tableLayoutPanel1);
        }

        private Panel GetTextBlockPanel(TextBlockTag tag)
        {
            // панель
            var panel = new TableLayoutPanel
            {
                Tag = tag,
                ColumnCount = 3,
                RowCount = 1,
                Top = 0,
                Left = 0,
                Width = textBlocksWrapper.Width - 25,
                MinimumSize = new Size(0,TextBlockPanelHeight),
                Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
                BorderStyle = BorderStyle.FixedSingle,
                ColumnStyles =
                {
                    new ColumnStyle(SizeType.Percent, 15),
                    new ColumnStyle(SizeType.Percent, 70),
                    new ColumnStyle(SizeType.Percent, 15),
                }
            };

            // иконка удаления
            var removePicture = new PictureBox
            {
                SizeMode = PictureBoxSizeMode.StretchImage,
                Width = IconButtonsSize,
                Height = IconButtonsSize,
                Image = _removeIcon,
                Anchor = AnchorStyles.Left,
                Cursor = Cursors.Hand
            };
            removePicture.Click += (sender, args) => TagRemove_Click(sender);
            panel.Controls.Add(removePicture);

            // условия
            string conditionsString = tag.ConditionsToGuiString();
            var conditionsLabel = new Label
            {
                Text = conditionsString != "" ? conditionsString : "Значение по умолчанию",
                AutoSize = true,
            };
            var titlePanel = new TableLayoutPanel
            {
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                MaximumSize = new Size(0, 96),
                Controls =
                {
                    new Label
                    {
                        Font = new System.Drawing.Font(Font, FontStyle.Bold),
                        Text = @"Условия:",
                    },
                    conditionsLabel,
                    new Panel
                    {
                        Height = 10,
                    },
                    new Label
                    {
                        AutoSize = true,
                        ForeColor = Color.DarkGray,
                        Text = $"Приоритет: {tag.Priority}",
                    },
                }
            };
            panel.Controls.Add(titlePanel);
            conditionsLabel.MaximumSize = new Size(conditionsLabel.Parent.Width, 0);
            titlePanel.Resize += (sender, args) =>
                conditionsLabel.MaximumSize = new Size(conditionsLabel.Parent.Width, 0);
            titlePanel.AutoScroll = true;



            var arrowPicture = new PictureBox
            {
                SizeMode = PictureBoxSizeMode.StretchImage,
                Width = IconButtonsSize,
                Height = IconButtonsSize,
                Image = _arrowIcon,
                Anchor = AnchorStyles.Right,
                Cursor = Cursors.Hand,
            };
            arrowPicture.Click += (object sender, EventArgs args) => TagEdit_Click(sender);

            panel.Controls.Add(arrowPicture);
            return panel;
        }


        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.textBlocksWrapper = new System.Windows.Forms.Panel();
            this.addNewConditionButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBlocksWrapper
            // 
            this.textBlocksWrapper.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBlocksWrapper.AutoScroll = true;
            this.textBlocksWrapper.Location = new System.Drawing.Point(0, 0);
            this.textBlocksWrapper.Name = "textBlocksWrapper";
            this.textBlocksWrapper.Size = new System.Drawing.Size(384, 563);
            this.textBlocksWrapper.TabIndex = 2;
            // 
            // addNewConditionButton
            // 
            this.addNewConditionButton.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.addNewConditionButton.Location = new System.Drawing.Point(12, 571);
            this.addNewConditionButton.Name = "addNewConditionButton";
            this.addNewConditionButton.Size = new System.Drawing.Size(360, 40);
            this.addNewConditionButton.TabIndex = 1;
            this.addNewConditionButton.Text = "Добавить новое условие";
            this.addNewConditionButton.UseVisualStyleBackColor = true;
            this.addNewConditionButton.Click += new System.EventHandler(this.addNewConditionButton_Click);
            // 
            // TextBlockConditionsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(384, 620);
            this.Controls.Add(this.addNewConditionButton);
            this.Controls.Add(this.textBlocksWrapper);
            this.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "TextBlockConditionsForm";
            this.ShowIcon = false;
            this.Text = "TextBlockConditionsForm";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.TextBlockConditionsForm_FormClosed);
            this.Load += new System.EventHandler(this.TextBlockConditionsForm_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel textBlocksWrapper;
        private System.Windows.Forms.Button addNewConditionButton;
    }
}