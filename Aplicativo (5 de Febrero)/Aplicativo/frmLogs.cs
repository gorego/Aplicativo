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
    public partial class frmLogs : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();

        public void cargarCargos()
        {
            while (dataGridView1.Rows.Count != 0)
            {
                dataGridView1.Rows.RemoveAt(0);
            }
            string query = "SELECT * FROM historicoIngresos ORDER BY ID desc";
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
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[i].Cells[0].Value = myReader.GetInt32(0);
                    dataGridView1.Rows[i].Cells[1].Value = myReader.GetString(1);
                    dataGridView1.Rows[i].Cells[2].Value = myReader.GetString(2);
                    dataGridView1.Rows[i].Cells[3].Value = myReader.GetString(3);
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

        public frmLogs()
        {
            InitializeComponent();
            cargarCargos();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Variables.imprimir(dataGridView1);
        }
    }
}
