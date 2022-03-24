using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelToWordProject.Forms
{
    public partial class InputDataDialogForm : Form
    {
        private string Message { get; }
        private string Title { get; }
        private string LeftActionLabel { get; }
        private string RightActionLabel { get; }
        private Action<string> OnLeftAction { get; }
        private Action<string> OnRightAction { get; }


        public InputDataDialogForm(string message, string title, string leftActionLabel = "Ок", string rightActionLabel = "Отмена",
            Action<string> onLeftAction = null, Action<string> onRightAction = null)
        {
            Message = message;
            Title = title;
            LeftActionLabel = leftActionLabel;
            RightActionLabel = rightActionLabel;
            OnLeftAction = onLeftAction;
            OnRightAction = onRightAction;
            InitializeComponent();
        }

        private void InputDataDialogForm_Load(object sender, EventArgs e)
        {
            Size = tableLayoutPanel2.Size;

            Text = Title;
            label1.Text = Message;
            button1.Text = LeftActionLabel;
            button2.Text = RightActionLabel;
        }

        #pragma warning disable IDE1006 // Стили именования
        private void button1_Click(object sender, EventArgs e)
        {
            Close();
            OnLeftAction?.Invoke(textBox1.Text);
            Dispose();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            Close();
            OnRightAction?.Invoke(textBox1.Text);
            Dispose();
        }
        #pragma warning restore IDE1006 // Стили именования
    }
}
