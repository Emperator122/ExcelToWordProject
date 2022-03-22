﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Media;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelToWordProject.Syllabus;
using ExcelToWordProject.Syllabus.Tags;
using BorderStyle = System.Windows.Forms.BorderStyle;

namespace ExcelToWordProject.Forms
{
    public partial class TextBlocksForm : Form
    {
        private readonly SyllabusParameters _parameters;

        public TextBlocksForm(SyllabusParameters parameters)
        {
            _parameters = parameters;
            InitializeComponent();
            InittializeCustomComponents();
        }

        private void TextBlocksForm_Load(object sender, EventArgs e)
        {

        }

        private void OnGoToConditionsButtonClick(string tagName)
        {
            var form = new TextBlockConditionsForm(tagName, _parameters, this);
            form.Show();
            Hide();
        }

        private void addNewTagButton_Click(object sender, EventArgs e)
        {
            var form = new TextBlockAddForm("", this);
            form.ShowDialog();
        }

        public override void Refresh()
        {
            base.Refresh();
            Controls.Clear();
            InitializeComponent();
            InittializeCustomComponents();
        }
    }
}