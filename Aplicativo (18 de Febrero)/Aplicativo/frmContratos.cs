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
    public partial class frmContratos : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();

        public void cargarOperador()
        {
            string query = "SELECT * FROM Operador";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            textBox2.DataSource = ds.Tables[0];
            textBox2.DisplayMember = "Operador";
            textBox2.ValueMember = "ID";
            textBox2.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            textBox2.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void cargarContratos()
        {
            string query = "SELECT c.ID, c.Contrato, o.Operador, c.Ubicacion FROM Operador AS o INNER JOIN Contratos AS c ON o.ID = c.Operador WHERE c.Contrato <> 'N/A';";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable contratos = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(contratos);
            dataGridView1.DataSource = contratos;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[3].HeaderText = "Ubicación Fisica";
        }

        public void agregarContrato()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Contratos(Contrato,Operador,Ubicacion) VALUES (@Contrato,@Operador,@Ubicacion)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                string contrato = textBox3.Text;
                if (!contrato.Equals(""))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Contratos\\" + textBox1.Text);
                    string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Contratos\\" + textBox1.Text, "Contrato*");
                    if (prueba.Length > 0)
                    {
                        if (File.Exists(prueba[0]))
                        {

                            File.Delete(prueba[0]);
                        }
                    }
                    using (FileStream fs = File.Open(contrato, FileMode.Open))
                    {
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Contratos\\" + textBox1.Text);
                        string ext = Path.GetExtension(contrato);
                        fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Contratos\\" + textBox1.Text + "\\Contrato" + ext));
                    }
                }
                cmd.Parameters.Add("@Contrato", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@Operador", OleDbType.VarChar).Value = textBox2.SelectedValue;
                cmd.Parameters.Add("@Ubicacion", OleDbType.VarChar).Value = textBox4.Text;

                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Contrato agregado.");
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

        public void modificarContrato()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Contratos SET Contrato=@Contrato,Operador=@Operador,Ubicacion=@Ubicacion WHERE ID = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                string contrato = textBox3.Text;
                if (!contrato.Equals(""))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Contratos\\" + textBox1.Text);
                    string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Contratos\\" + textBox1.Text, "Contrato*");
                    if (prueba.Length > 0)
                    {
                        if (File.Exists(prueba[0]))
                        {

                            File.Delete(prueba[0]);
                        }
                    }
                    using (FileStream fs = File.Open(contrato, FileMode.Open))
                    {
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Contratos\\" + textBox1.Text);
                        string ext = Path.GetExtension(contrato);
                        fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Contratos\\" + textBox1.Text + "\\Contrato" + ext));
                    }
                }
                cmd.Parameters.Add("@Contrato", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@Operador", OleDbType.VarChar).Value = textBox2.SelectedValue;
                cmd.Parameters.Add("@Ubicacion", OleDbType.VarChar).Value = textBox4.Text;

                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Contrato modificado.");
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

        public frmContratos()
        {
            InitializeComponent();
            cargarContratos();
            cargarOperador();
            dataGridView1.Columns[1].DefaultCellStyle.Font = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Underline);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //string contrato = textBox3.Text;
            //if (!contrato.Equals(""))
            //{
            //    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Contratos\\" + textBox1.Text);
            //    string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Contratos\\" + textBox1.Text, "Contrato*");
            //    if (prueba.Length > 0)
            //    {
            //        if (File.Exists(prueba[0]))
            //        {

            //            File.Delete(prueba[0]);
            //        }
            //    }
            //    using (FileStream fs = File.Open(contrato, FileMode.Open))
            //    {
            //        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Contratos\\" + textBox1.Text);
            //        string ext = Path.GetExtension(contrato);
            //        fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Contratos\\" + textBox1.Text + "\\Contrato" + ext));
            //        MessageBox.Show("Contrato modificado.");
            //    }
            //}
            modificarContrato();
            cargarContratos();
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnAgregar_Click(object sender, EventArgs e)
        {
            if (!textBox1.Text.Equals(""))
            {
                if (!textBox1.Text.Equals(""))
                {
                    agregarContrato();
                    cargarContratos();
                    textBox1.Text = "";
                    textBox2.Text = "";
                    textBox3.Text = "";
                    textBox4.Text = "";
                }
                else
                {
                    MessageBox.Show("Favor ingresar el numero de contrato.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            else
            {
                MessageBox.Show("Favor seleccionar el operador.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el contrato " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {

                string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM Contratos WHERE id = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Contrato eliminado.");
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Contratos\\"+ textBox1.Text);
                        string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Contratos\\" + textBox1.Text, "Contrato*");
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
                cargarContratos();
            }
        }

        private void btnPredial_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox3.Text = openFileDialog1.FileName;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentCell.ColumnIndex == 2)
            {
                frmOrdenes newFrm = new frmOrdenes("Contrato", dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
                newFrm.Show();
            }
            if (dataGridView1.CurrentCell.ColumnIndex == 1)
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Contratos\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString());
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Contratos\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString(), "Contrato*");
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
            textBox4.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Variables.imprimir(dataGridView1);
        }
    }
}
