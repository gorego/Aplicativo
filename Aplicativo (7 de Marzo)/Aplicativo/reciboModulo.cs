using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Aplicativo
{
    public partial class reciboModulo : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();

        public reciboModulo(string recibo, string numRecibo, string modulo)
        {
            InitializeComponent();
            this.Text = numRecibo;
            label1.Text = "Recibo: " + numRecibo;
            label2.Text = "Modulo Actual: " + modulo;
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}
