using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Aplicativo
{
    public partial class frmDespacho : Form
    {
        public frmDespacho()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            frmDespachoOpcion newFrm = new frmDespachoOpcion();
            newFrm.Show();
        }
    }
}
