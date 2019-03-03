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
    public partial class AddForm : Form
    {
        String name = "";
        int quantity = -1;

        public AddForm()
        {
            InitializeComponent();
            name = "";
            quantity = -1;
        }

        public AddForm(String text, int quantity)
        {
            InitializeComponent();
            name = text;
            nameCompBox.Text = text;
            this.quantity = quantity;
            quantityCompBox.Text = quantity.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (nameCompBox.Text != null && quantityCompBox.Text != null &&
                   (name != nameCompBox.Text || quantity != int.Parse(quantityCompBox.Text)) && int.Parse(quantityCompBox.Text) != 0)
                {
                    nameCompBox.ForeColor = Color.Black;
                    quantityCompBox.ForeColor = Color.Red;
                    name = nameCompBox.Text;
                    quantity = int.Parse(quantityCompBox.Text);
                    this.DialogResult = DialogResult.OK;
                    this.Hide();
                }
                else
                {
                    quantityCompBox.ForeColor = Color.Red;
                    nameCompBox.ForeColor = Color.Red;
                }
            }
            catch(System.FormatException exc)
            {

                label3.Text = "Количество должно быть численным";
                quantityCompBox.ForeColor = Color.Red;
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Hide();
        }

        public String GetComponent()
        {
            return name;
        }

        public int GetQuantity()
        {
            return quantity;
        }

        public void changeStatus(String name)
        {
            this.label3.Text = name;
        }

        internal void containAllert(string componentName, int componentQuantity)
        {
            nameCompBox.ForeColor = Color.Red;
            label3.Text = "Недопустимое имя компонента";
            this.name = componentName;
        }

        internal void recursionFound(string componentName, int componentQuantity)
        {
            nameCompBox.ForeColor = Color.Red;
            label3.Text = "Найдена рекурсия";
            this.name = componentName;
        }
    }
}
