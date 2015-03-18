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
    public partial class frmUnidadDeManejo : Form
    {
        String connectionString = Variables.connectionString;
                OleDbConnection conn = new OleDbConnection();

        public void cargarUnidades()
                {

                    string query = "SELECT u.ID, u.Unidad, SUM(l.areaPlantacion) As AreaPlanteada, SUM(l.areaEfectiva) As AreaEfectiva FROM UnidadDeManejo u LEFT JOIN Lotes l ON (u.ID = l.Unidad) GROUP BY u.id, u.Unidad";
                    //Ejecutar el query y llenar el GridView.
                    conn.ConnectionString = connectionString;
                    OleDbCommand cmd = new OleDbCommand(query, conn);
                    DataTable banco = new DataTable();
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(banco);
                    dataGridView1.DataSource = banco;
                    dataGridView1.Columns[0].Visible = false;
                    dataGridView1.Columns[1].HeaderText = "Unidad de Manejo";
                    dataGridView1.Columns[2].HeaderText = "Area plantacion";
                    dataGridView1.Columns[3].HeaderText = "Area efectiva";
                }

        public void modificarUnidad()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE UnidadDeManejo SET Unidad=@Unidad WHERE ID = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Unidad", OleDbType.VarChar).Value = textBox1.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Unidad de manejo modficiada.");
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

        public void agregarUnidad()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO UnidadDeManejo (Unidad) VALUES (@Unidad)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Unidad", OleDbType.VarChar).Value = textBox1.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Unidad de manejo agregada.");
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

        public frmUnidadDeManejo()
        {
            InitializeComponent();
            cargarUnidades();
        }

        private void btnAgregar_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Equals(""))
            {
                MessageBox.Show("Favor ingresar el nombre de la Unidad de Manejo", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                agregarUnidad();
                cargarUnidades();
            }
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnModificar_Click(object sender, EventArgs e)
        {
            modificarUnidad();
            cargarUnidades();
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar la unidad " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {

                string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM UnidadDeManejo WHERE id = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Unidad de manejo eliminada.");
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
                cargarUnidades();
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Variables.imprimir(dataGridView1);
        }
    }
}
