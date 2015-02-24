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
    public partial class frmPaqueteOP : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();

        public frmPaqueteOP(int op, int producto)
        {
            InitializeComponent();
            Variables.cargar(dataGridView14, "SELECT paq.Id, paq.numPaquete, prod.Codigo, paq.Bodega, paq.numPiezas, paq.dia, (Avg(info.ancho1) + Avg(info.ancho2))/2, (Avg(info.alto1) + Avg(info.alto2))/2, paq.Fecha, paq.Hora FROM (Paquete AS paq INNER JOIN Productos AS prod ON paq.Producto = prod.ID) INNER JOIN infoPaquete AS info ON paq.Id = info.Paquete WHERE (((paq.OP)="+op+") AND ((paq.Producto)="+producto+")) GROUP BY paq.Id, paq.numPaquete, prod.Codigo, paq.Bodega, paq.numPiezas, paq.dia, paq.porcentaje, paq.Fecha, paq.Hora");
            dataGridView14.Columns[1].FillWeight = 200;
            dataGridView14.Columns[2].FillWeight = 200;
            dataGridView14.Columns[1].HeaderText = "# de Paquete";
            dataGridView14.Columns[4].HeaderText = "# de Piezas";
            dataGridView14.Columns[5].HeaderText = "Día";
            dataGridView14.Columns[5].Visible = false;
            dataGridView14.Columns[6].HeaderText = "Ancho Promedio";
            dataGridView14.Columns[7].HeaderText = "Alto Promedio";
            parametrosRequerido(producto);
        }

        public void parametrosRequerido(int producto)
        {
            string query = "SELECT anchoProd, altoProd FROM Productos WHERE ID = " + producto;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                if (myReader.Read())
                {
                    label2.Text = myReader.GetValue(0).ToString();
                    label3.Text = myReader.GetValue(1).ToString();
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
        }

        private void dataGridView14_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
        }
    }
}
