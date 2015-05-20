using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;

namespace Aplicativo
{
    public partial class frmFSC : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();

        public frmFSC()
        {
            InitializeComponent();
            Variables.cargar(dataGridView1, "SELECT * FROM certificadoFSC");
        }

        public void certificadoFSC(int tipo)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand();
            if (tipo == 0)
                cmd = new OleDbCommand("INSERT INTO certificadoFSC (Certificado) VALUES (@Certificado)");
            else
                cmd = new OleDbCommand("UPDATE certificadoFSC SET Certificado = @Certificado WHERE ID =" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Certificado", OleDbType.VarChar).Value = textBox1.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    if(tipo == 0)
                        MessageBox.Show("Certificado FSC agregado.");
                    else
                        MessageBox.Show("Certificado FSC modificado.");
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

        public void subirArchivo()
        {
            string archivo = textBox4.Text;
            if (!archivo.Equals(""))
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Certificado");
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Certificado", textBox1.Text + "*");
                if (prueba.Length > 0)
                {
                    if (File.Exists(prueba[0]))
                    {

                        File.Delete(prueba[0]);
                    }
                }
                using (FileStream fs = File.Open(archivo, FileMode.Open))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Certificado");
                    string ext = Path.GetExtension(archivo);
                    fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Certificado\\" + textBox1.Text + ext));
                }
            }
        }
        
        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox4.Text = openFileDialog1.FileName;
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Certificado");
            if (!textBox1.Text.Equals("N/A"))
            {
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Certificado", textBox1.Text + "*");
                if (prueba.Length > 0)
                {
                    if (File.Exists(prueba[0]))
                    {
                        System.Diagnostics.Process.Start(prueba[0]);
                    }
                    else
                    {
                        MessageBox.Show("No se encuentra el archivo.", "Error");
                    }
                }
                else
                {
                    MessageBox.Show("No se encuentra el archivo.", "Error");
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!Variables.existe("SELECT * FROM certificadoFSC WHERE Certificado = '" + textBox1.Text + "'"))
            {
                certificadoFSC(0);
                subirArchivo();
                Variables.cargar(dataGridView1, "SELECT * FROM certificadoFSC");
            }
            else
            {
                MessageBox.Show("Certificado ya existe, favor ingresar uno nuevo.");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString().Equals(textBox1.Text))
            {
                certificadoFSC(1);
                subirArchivo();
                Variables.cargar(dataGridView1, "SELECT * FROM certificadoFSC");
            }
            else
            {
                if (!Variables.existe("SELECT * FROM certificadoFSC WHERE Certificado = '" + textBox1.Text + "'"))
                {
                    certificadoFSC(1);
                    subirArchivo();
                    Variables.cargar(dataGridView1, "SELECT * FROM certificadoFSC");
                }
                else
                {
                    MessageBox.Show("Certificado ya existe, favor ingresar uno nuevo.");
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el certificado # " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {

                string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM certificadoFSC WHERE id = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Certificado eliminado.");
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Certificado");
                        string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Certificado", textBox1.Text + "*");
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
                Variables.cargar(dataGridView1, "SELECT * FROM certificadoFSC");
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();
        }
    }
}
