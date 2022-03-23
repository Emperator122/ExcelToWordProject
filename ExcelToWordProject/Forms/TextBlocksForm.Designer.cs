using System;
using System.Drawing;
using System.Windows.Forms;


namespace ExcelToWordProject.Forms
{
    partial class TextBlocksForm
    {
        private const int TextBlockPanelHeight = 64;
        private const int FormWidth = 400;
        private const int IconButtonsSize = 32;
        private readonly Bitmap _arrowIcon = Properties.Resources.right_arrow;
        private readonly Bitmap _moreIcon = Properties.Resources.more;

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

        private void InittializeCustomComponents()
        {
            Width = FormWidth;
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
            var contextMenu = GetMoreActionsContextMenu();
            foreach (var textBlockTag in _parameters.TextBlockTags.GroupedByKey())
            {
                tableLayoutPanel1.Controls.Add(GetTextBlockPanel(textBlockTag.Key, contextMenu));
            }

            textBlocksWrapper.Controls.Add(tableLayoutPanel1);
        }

        private Panel GetTextBlockPanel(string tagName, ContextMenu contextMenu)
        {
            // панель
            var panel = new TableLayoutPanel
            {
                Tag = tagName,
                ColumnCount = 3,
                RowCount = 1,
                Top = 0,
                Left = 0,
                Height = TextBlockPanelHeight,
                Width = textBlocksWrapper.Width - 25,
                Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
                BorderStyle = BorderStyle.FixedSingle,
                ColumnStyles =
                {
                    new ColumnStyle(SizeType.Percent, 15),
                    new ColumnStyle(SizeType.Percent, 70),
                    new ColumnStyle(SizeType.Percent, 15),
                }
            };

            // иконка доп. действий
            var morePicture = GetMorePictureBox(contextMenu);
            panel.Controls.Add(morePicture);

            // описание
            var titlePanel = GetTitlePanel(tagName);
            panel.Controls.Add(titlePanel);


            // кнопка "далее"
            var arrowPicture = GetArrowPictureBox(tagName);
            panel.Controls.Add(arrowPicture);
            return panel;
        }


        private PictureBox GetMorePictureBox(ContextMenu menu)
        {
            var pictureBox = new PictureBox
            {
                SizeMode = PictureBoxSizeMode.StretchImage,
                Width = IconButtonsSize,
                Height = IconButtonsSize,
                Image = _moreIcon,
                Anchor = AnchorStyles.Left,
                Cursor = Cursors.Hand
            };

            pictureBox.MouseClick +=
                (sender, args) =>
                {
                    if(args.Button == MouseButtons.Left)
                        menu.Show(pictureBox, new Point(args.X, args.Y));
                };

            return pictureBox;
        }

        private TableLayoutPanel GetTitlePanel(string tagName)
        {
            return new TableLayoutPanel
            {
                Height = 48,
                Anchor = AnchorStyles.Right | AnchorStyles.Left,
                Controls =
                {
                    new Label
                    {
                        Anchor = AnchorStyles.Top | AnchorStyles.Bottom,
                        Font = new System.Drawing.Font(Font, FontStyle.Bold),
                        Text = @"Имя тега:",
                    },
                    new Label
                    {
                        Anchor = AnchorStyles.Top | AnchorStyles.Bottom,
                        Text = tagName,
                        AutoSize = true,
                        AutoEllipsis = true,
                    }
                }
            };
        }

        private PictureBox GetArrowPictureBox(string tagName)
        {
            var arrowPicture = new PictureBox
            {
                SizeMode = PictureBoxSizeMode.StretchImage,
                Width = IconButtonsSize,
                Height = IconButtonsSize,
                Image = _arrowIcon,
                Anchor = AnchorStyles.Right,
                Cursor = Cursors.Hand,
            };
            arrowPicture.Click += (object sender, EventArgs args) => OnGoToConditionsButtonClick(tagName);
            return arrowPicture;
        }

        private ContextMenu GetMoreActionsContextMenu()
        {
            var menu = new ContextMenu();
            menu.MenuItems.Add("Удалить", (sender, args) => RemoveTag_Click(sender));

            return menu;
        }
        

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.textBlocksWrapper = new System.Windows.Forms.Panel();
            this.addNewTagButton = new System.Windows.Forms.Button();
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
            this.textBlocksWrapper.Size = new System.Drawing.Size(384, 562);
            this.textBlocksWrapper.TabIndex = 1;
            // 
            // addNewTagButton
            // 
            this.addNewTagButton.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.addNewTagButton.Location = new System.Drawing.Point(12, 571);
            this.addNewTagButton.Name = "addNewTagButton";
            this.addNewTagButton.Size = new System.Drawing.Size(360, 40);
            this.addNewTagButton.TabIndex = 0;
            this.addNewTagButton.Text = "Добавить новый тег";
            this.addNewTagButton.UseVisualStyleBackColor = true;
            this.addNewTagButton.Click += new System.EventHandler(this.AddNewTagButton_Click);
            // 
            // TextBlocksForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(384, 620);
            this.Controls.Add(this.addNewTagButton);
            this.Controls.Add(this.textBlocksWrapper);
            this.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "TextBlocksForm";
            this.ShowIcon = false;
            this.Text = "Группы блоков текста";
            this.Load += new System.EventHandler(this.TextBlocksForm_Load);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Panel textBlocksWrapper;
        private Button addNewTagButton;
    }
}