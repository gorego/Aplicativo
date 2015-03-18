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
        int tipo = 0;

        public frmPaqueteOP(int op, int producto)
        {
            InitializeComponent();
            Variables.cargar(dataGridView14, "SELECT paq.ID, paq.Id, paq.numPaquete, prod.Codigo, paq.Bodega, paq.numPiezas, paq.dia, (Avg(info.ancho1) + Avg(info.ancho2))/2, (Avg(info.alto1) + Avg(info.alto2))/2, paq.Fecha, paq.Hora FROM (Paquete AS paq INNER JOIN Productos AS prod ON paq.Producto = prod.ID) INNER JOIN infoPaquete AS info ON paq.Id = info.Paquete WHERE (((paq.OP)="+op+") AND ((paq.Producto)="+producto+")) GROUP BY paq.Id, paq.numPaquete, prod.Codigo, paq.Bodega, paq.numPiezas, paq.dia, paq.porcentaje, paq.Fecha, paq.Hora");
            dataGridView14.Columns[2].FillWeight = 200;
            dataGridView14.Columns[3].FillWeight = 200;
            dataGridView14.Columns[2].HeaderText = "# de Paquete";
            dataGridView14.Columns[5].HeaderText = "# de Piezas";
            dataGridView14.Columns[6].HeaderText = "Día";
            dataGridView14.Columns[6].Visible = false;
            dataGridView14.Columns[7].HeaderText = "Ancho Promedio";
            dataGridView14.Columns[8].HeaderText = "Alto Promedio";
            parametrosRequerido(producto);
            numProductos(dataGridView14);
            dataGridView14.Columns[1].HeaderText = "#";
        }

        public frmPaqueteOP(string producto)
        {
            InitializeComponent();
            Variables.cargar(dataGridView14, "SELECT paq.ID, paq.Id, paq.numPaquete, prod.Codigo, paq.Bodega, paq.numPiezas, paq.dia, (Avg(info.ancho1)+Avg(info.ancho2))/2 AS Expr1, (Avg(info.alto1)+Avg(info.alto2))/2 AS Expr2, paq.Fecha, paq.Hora, h.OP, paq.OP FROM ((Paquete AS paq INNER JOIN Productos AS prod ON paq.Producto = prod.ID) INNER JOIN infoPaquete AS info ON paq.Id = info.Paquete) INNER JOIN historicoProduccion AS h ON paq.OP = h.ID WHERE (((paq.Producto)= " + producto + ")) AND (paq.volumenActual > 0) AND paq.Parcial = 0 GROUP BY paq.Id, paq.numPaquete, prod.Codigo, paq.Bodega, paq.numPiezas, paq.dia, paq.Fecha, paq.Hora, paq.porcentaje, h.OP, paq.OP;");
            dataGridView14.Columns[2].FillWeight = 200;
            dataGridView14.Columns[3].FillWeight = 200;
            dataGridView14.Columns[2].HeaderText = "# de Paquete";
            dataGridView14.Columns[5].HeaderText = "# de Piezas";
            dataGridView14.Columns[6].HeaderText = "Día";
            dataGridView14.Columns[6].Visible = false;
            dataGridView14.Columns[7].HeaderText = "Ancho Promedio";
            dataGridView14.Columns[8].HeaderText = "Alto Promedio";
            parametrosRequerido(Int32.Parse(producto));
            numProductos(dataGridView14);
            dataGridView14.Columns[0].Visible = false;
            dataGridView14.Columns[1].HeaderText = "#";
            dataGridView14.Columns[11].DefaultCellStyle.Font = new Font(dataGridView14.DefaultCellStyle.Font, FontStyle.Underline);
            dataGridView14.Columns[12].Visible = false;
            dataGridView14.Columns[11].HeaderText = "OP";
            tipo = 1;
        }

        public void numProductos(DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                data.Rows[i].Cells[1].Value = i + 1;
            }
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

        private void dataGridView14_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (tipo == 1)
            {
                if (dataGridView14.CurrentCell.ColumnIndex == 11)
                {
                    frmCrearProduccion newFrm = new frmCrearProduccion(Int32.Parse(dataGridView14.Rows[dataGridView14.CurrentCell.RowIndex].Cells[12].Value.ToString()), 1);
                    newFrm.Show();
                }
            }
        }
    }
}
