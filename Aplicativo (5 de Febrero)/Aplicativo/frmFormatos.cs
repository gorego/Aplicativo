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
using System.IO;

namespace Aplicativo
{
    public partial class frmFormatos : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        public frmFormatos()
        {
            InitializeComponent();
            cargarFormatos();
            dataGridView1.Columns[1].DefaultCellStyle.Font = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Underline);
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public void cargarFormatos()
        {
            string query = "SELECT * FROM Formatos;";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable contratos = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(contratos);
            dataGridView1.DataSource = contratos;
            dataGridView1.Columns[0].Visible = false;
        }

        public void agregarFormato()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Formatos(Formato,Nombre) VALUES (@Formato,@Nombre)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                string contrato = textBox3.Text;
                if (!contrato.Equals(""))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos");
                    string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos", textBox1.Text+"*");
                    if (prueba.Length > 0)
                    {
                        if (File.Exists(prueba[0]))
                        {

                            File.Delete(prueba[0]);
                        }
                    }
                    using (FileStream fs = File.Open(contrato, FileMode.Open))
                    {
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos");
                        string ext = Path.GetExtension(contrato);
                        fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos\\" + textBox1.Text + ext));
                    }
                }
                cmd.Parameters.Add("@Formato", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@Nombre", OleDbType.VarChar).Value = textBox2.Text;

                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Formato agregado.");
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

        private void btnAgregar_Click(object sender, EventArgs e)
        {
            agregarFormato();
            cargarFormatos();
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
        }

        private void btnModificar_Click(object sender, EventArgs e)
        {
            string contrato = textBox3.Text;
            if (!contrato.Equals(""))
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos");
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos", textBox1.Text + "*");
                if (prueba.Length > 0)
                {
                    if (File.Exists(prueba[0]))
                    {

                        File.Delete(prueba[0]);
                    }
                }
                using (FileStream fs = File.Open(contrato, FileMode.Open))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos");
                    string ext = Path.GetExtension(contrato);
                    fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos\\" + textBox1.Text + ext));
                    MessageBox.Show("Contrato modificado.");
                }
            }
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el contrato " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {

                string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM Formato WHERE id = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Contrato eliminado.");
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos");
                        string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos", textBox1.Text+"*");
                        if (prueba.Length > 0)
                        {
                            if (File.Exists(prueba[0]))
                            {

                                File.Delete(prueba[0]);
                            }
                        }
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
                cargarFormatos();
            }
        }

        private void btnExaminar_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox3.Text = openFileDialog1.FileName;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentCell.ColumnIndex == 1)
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos");
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos\\",  dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString()+"*");
                if (prueba.Length > 0)
                {
                    if (File.Exists(prueba[0]))
                    {
                        System.Diagnostics.Process.Start(prueba[0]);
                    }
                }
                else
                {
                    MessageBox.Show("No se encuentra el archivo.", "Error");
                }
            }
            textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();
            textBox2.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Variables.imprimir(dataGridView1);
        }
    }
}
