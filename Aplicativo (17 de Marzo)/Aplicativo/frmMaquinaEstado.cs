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
    public partial class frmMaquinaEstado : Form
    {
        string connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        string tipoEstado = "";
        public frmMaquinaEstado(string Tipo)
        {
            InitializeComponent();
            tipoEstado = Tipo;
            if (tipoEstado.Equals("Maleza"))
                this.Text = "Estado de Maleza";
            if (tipoEstado.Equals("Cercas"))
                this.Text = "Estado de Cercas";
            cargarEstados();
        }

        public void cargarEstados()
        {
            string query = "SELECT * From Estados WHERE Tipo = '"+ tipoEstado + "'";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinarias = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(maquinarias);
            dataGridView2.DataSource = maquinarias;
            dataGridView2.Columns[0].Visible = false;
            dataGridView2.Columns[1].Visible = false;
        }

        public void agregarEstado() {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Estados (Tipo,Estado,Descripcion) VALUES (@Tipo,@Estado,@Descripcion)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Tipo", OleDbType.VarChar).Value = tipoEstado;
                cmd.Parameters.Add("@Estado", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = textBox2.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Estado agregado.");
                    conn.Close();
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show(ex.Source);
                    conn.Close();
                }
            }
            else
            {
                MessageBox.Show("Connection Failed");
            }
        }

        public void eliminarEstado() {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el estado " + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[2].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM Estados WHERE ID = " + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString());
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Estado eliminado.");
                        conn.Close();
                    }
                    catch (OleDbException ex)
                    {
                        MessageBox.Show(ex.Source);
                        conn.Close();
                    }
                }
                else
                {
                    MessageBox.Show("Connection Failed");
                }
                cargarEstados();
            }
        }

        private void btnEliminarAceite_Click(object sender, EventArgs e)
        {
            eliminarEstado();
        }

        private void btnAceite_Click(object sender, EventArgs e)
        {
            agregarEstado();
            cargarEstados();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[2].Value.ToString();
            textBox2.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[3].Value.ToString();
        }

        private void btnEliminarAceite_Click_1(object sender, EventArgs e)
        {
            eliminarEstado();
        }

        private void btnAceite_Click_1(object sender, EventArgs e)
        {
            agregarEstado();
            cargarEstados();
        }
    }
}
