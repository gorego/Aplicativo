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
    public partial class frmMunicipios : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        public frmMunicipios()
        {
            InitializeComponent();
            cargarMunicipios();
        }

        public void cargarMunicipios() {
            listBox1.Items.Clear();
            string query = "SELECT * From Municipio";
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

        public void agregarMuncipio() {
            if (!textBox1.Text.Equals(""))
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("INSERT INTO Municipio (Municipio) VALUES (@Municipio)");
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@Muncipio", OleDbType.VarChar).Value = textBox1.Text;
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Municipio agregado.");
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
            else
                MessageBox.Show("Favor ingresar nombre del municipio", "Error");
        }

        public void eliminarMuncipio() {
            if (listBox1.SelectedIndex != -1)
            {
                DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el municipio " + listBox1.SelectedItem + "?", "Confirmar", MessageBoxButtons.YesNo);

                if (dialogResult == DialogResult.Yes)
                {
                    conn.ConnectionString = connectionString;
                    OleDbCommand cmd = new OleDbCommand("DELETE FROM Municipio WHERE Municipio = '" + listBox1.SelectedItem.ToString() + "'");
                    cmd.Connection = conn;
                    conn.Open();

                    if (conn.State == ConnectionState.Open)
                    {
                        try
                        {
                            cmd.ExecuteNonQuery();
                            MessageBox.Show("Municipio eliminado.");
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
                    cargarMunicipios();
                }
            }
            else
                MessageBox.Show("Favor seleccionar municipio a eliminar","Error");
        }

        private void btnAgregar_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Equals(""))
            {
                MessageBox.Show("Favor ingresar el nombre del municipio.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                agregarMuncipio();
                cargarMunicipios();
            }            
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            eliminarMuncipio();
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
