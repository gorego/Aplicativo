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
    public partial class frmAsistenciaOrdenes : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();

        public frmAsistenciaOrdenes()
        {
            InitializeComponent();
            cargarEmpleados();
        }

        public void cargarEmpleados()
        {
            while (dataGridView3.Rows.Count != 0)
            {
                dataGridView3.Rows.RemoveAt(0);
            }
            string query = "SELECT t.ID,(t.Nombres + ' ' + t.Apellidos), t.Cedula, c.Cargo FROM ((CargoLaboral AS c INNER JOIN Trabajadores AS t ON c.ID = t.Cargo) INNER JOIN produccionEmpleados AS e ON t.ID = e.Trabajador) INNER JOIN historicoProduccion AS h ON e.Orden = h.ID WHERE h.Estado <> 'Cerrada' GROUP BY t.ID, (t.Nombres + ' ' + t.Apellidos), t.Cedula, c.Cargo";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int i = 0;
            try
            {
                while (myReader.Read())
                {
                    dataGridView3.Rows.Add();
                    dataGridView3.Rows[i].Cells[0].Value = i + 1;
                    dataGridView3.Rows[i].Cells[1].Value = myReader.GetInt32(0);
                    dataGridView3.Rows[i].Cells[2].Value = myReader.GetString(1);
                    dataGridView3.Rows[i].Cells[3].Value = myReader.GetInt32(2);
                    dataGridView3.Rows[i].Cells[4].Value = myReader.GetString(3);
                    dataGridView3.Rows[i].Cells[12].Value = 0;
                    i++;
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

    }
}
