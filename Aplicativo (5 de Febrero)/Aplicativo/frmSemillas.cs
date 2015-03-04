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
    public partial class frmSemillas : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();

        public frmSemillas()
        {
            InitializeComponent();
            Variables.cargar(dataGridView1, "SELECT * FROM Semillas");
            dataGridView1.Columns[1].DefaultCellStyle.Font = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Underline);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox4.Text = openFileDialog1.FileName;
        }

        public void modificarSemilla(int tipo)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand();
            if(tipo == 0)
                cmd = new OleDbCommand("INSERT INTO Semillas (Nombre) VALUES (@Nombre)");
            else
                cmd = new OleDbCommand("UPDATE Semillas SET Nombre = @Nombre WHERE ID =" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Nombre", OleDbType.VarChar).Value = textBox1.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Semilla agregada.");
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
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Semillas");
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Semillas", textBox1.Text + "*");
                if (prueba.Length > 0)
                {
                    if (File.Exists(prueba[0]))
                    {

                        File.Delete(prueba[0]);
                    }
                }
                using (FileStream fs = File.Open(archivo, FileMode.Open))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Semillas");
                    string ext = Path.GetExtension(archivo);
                    fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Semillas\\" + textBox1.Text + ext));
                }
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            modificarSemilla(0);
            subirArchivo();
            Variables.cargar(dataGridView1, "SELECT * FROM Semillas");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            modificarSemilla(1);
            subirArchivo();
            Variables.cargar(dataGridView1, "SELECT * FROM Semillas");
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Semillas");
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Semillas", textBox1.Text+ "*");
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
}
