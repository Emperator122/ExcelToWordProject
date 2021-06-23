using ExcelToWordProject.Syllabus;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelToWordProject
{
    public partial class DescriptionEditForm : Form
    {
        BaseSyllabusTag tag;
        public DescriptionEditForm(BaseSyllabusTag tag)
        {
            InitializeComponent();
            this.tag = tag;

            descriptionTextBox.Lines = tag.Description.Split('\n');
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            tag.Description = descriptionTextBox.Text;
            CancelButton1_Click(sender, e);
        }

        private void CancelButton1_Click(object sender, EventArgs e)
        {
            Close();
            Dispose();
        }
    }
}
