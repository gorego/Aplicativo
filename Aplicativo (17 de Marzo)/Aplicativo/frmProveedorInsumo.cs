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
    public partial class frmProveedorInsumo : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();

        public frmProveedorInsumo(string id)
        {
            InitializeComponent();
            cargarProveedores(id);
            totalCosto();
        }

        public void cargarProveedores(string id)
        {
            string query = "SELECT p.ID, (i.Codigo + ' ' + i.Marca + ' ' + i.Modelo) As Insumo, p.Cantidad, p.Fecha, p.Costo FROM Insumos AS i INNER JOIN proveedorInsumo AS p ON i.ID = p.Insumo WHERE p.Proveedor = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable supervisores = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(supervisores);
            dataGridView2.DataSource = supervisores;
            dataGridView2.Columns[0].Visible = false;
            dataGridView2.Columns[4].DefaultCellStyle.Format = "c";
        }

        public void totalCosto()
        {
            int total = 0;
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                total += Int32.Parse(dataGridView2.Rows[i].Cells[4].Value.ToString());
            }
            label1.Text = "Total: " + String.Format("{0:c}", total);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
