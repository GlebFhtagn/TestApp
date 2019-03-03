using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TestApp
{
    public partial class AddTopForm : Form
    {
        String name;
        public AddTopForm(String text)
        {
            InitializeComponent();
            name = text;
            this.nameTopCompBox.Text = text;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (nameTopCompBox.Text != null && nameTopCompBox.Text != name)
            {
                name = nameTopCompBox.Text;
                DialogResult = DialogResult.OK;
            }
            else
            {
                nameTopCompBox.ForeColor = Color.Red;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Hide();
        }

        public void containAllert(String componentName)
        {
            nameTopCompBox.ForeColor = Color.Red;
            label2.Text = "Компонент с таким названием уже присутствует";
            this.name = componentName;
        }

        public String GetComponent()
        {
            return name;
        }

        internal void recursionFound(string componentName)
        {
            nameTopCompBox.ForeColor = Color.Red;
            label2.Text = "Рекурсия найдена";
            this.name = componentName;
        }

        internal void changeStatus(string message)
        {
            this.label2.Text = message;
        }
    }
}
