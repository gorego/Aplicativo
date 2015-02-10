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
    public partial class frmMaquinaPropietarios : Form
    {
        string connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        public frmMaquinaPropietarios()
        {
            InitializeComponent();
            cargarPropietarios();
        }

        public void cargarPropietarios()
        {
            listBox1.Items.Clear();
            string query = "SELECT * From propietariosMaquina WHERE Propietario <> 'N/A'";
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
                    listBox1.Items.Add(myReader.GetString(1));
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

        public void agregarPropietarios()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO propietariosMaquina (Propietario) VALUES (@Propietario)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Propietario", OleDbType.VarChar).Value = textBox1.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Propietario agregado.");
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

        public void eliminarPropietario()
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el propietario " + listBox1.SelectedItem + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM propietariosMaquina WHERE Propietario = '" + listBox1.SelectedItem.ToString() + "'");
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Propietario eliminado.");
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
                cargarPropietarios();
            }
        }

        private void btnAgregar_Click(object sender, EventArgs e)
        {
            agregarPropietarios();
            cargarPropietarios();
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            eliminarPropietario();
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
