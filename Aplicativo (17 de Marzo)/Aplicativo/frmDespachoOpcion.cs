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
    public partial class frmDespachoOpcion : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();

        public frmDespachoOpcion()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                frmCrearDespacho newFrm = new frmCrearDespacho(0,-1);
                this.Hide();
                newFrm.ShowDialog();
                this.Close();
            }
            else
            {
                frmCrearDespacho newFrm = new frmCrearDespacho(1,-1);
                this.Hide();
                newFrm.ShowDialog();
                this.Close();
            }
        }
    }
}
