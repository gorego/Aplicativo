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
    public partial class frmClientes : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        public frmClientes()
        {
            InitializeComponent();
            cargarClientes();
            dataGridView1.Columns[1].DefaultCellStyle.Font = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Underline);
            linkLabel1.Visible = false;
            linkLabel2.Visible = false;
            linkLabel3.Visible = false;
        }

        public void cargarClientes()
        {
            string query = "SELECT * FROM Clientes WHERE Cliente <> 'N/A'";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable supervisores = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(supervisores);
            dataGridView1.DataSource = supervisores;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[3].HeaderText = "Camara de Comercio";
            dataGridView1.Columns[5].HeaderText = "Extranjero/Nacional";
            dataGridView1.Columns[6].HeaderText = "Nombre de Contacto";
            dataGridView1.Columns[7].HeaderText = "Telefono de Contacto";
            dataGridView1.Columns[8].HeaderText = "Direccion de Contacto";
            dataGridView1.Columns[9].HeaderText = "Email de Contacto";
            dataGridView1.Columns[10].HeaderText = "Ciudad de Contacto";
            dataGridView1.Columns[11].HeaderText = "Pais de Contacto";
            dataGridView1.Columns[12].HeaderText = "Codigo Postal";
            dataGridView1.Columns[13].HeaderText = "Codigo del Banco";
            dataGridView1.Columns[14].HeaderText = "# de Cuenta";
            dataGridView1.Columns[15].HeaderText = "Nombre del Banco";
            dataGridView1.Columns[16].HeaderText = "IBAN";
            dataGridView1.Columns[17].HeaderText = "Swift";
            dataGridView1.Columns[18].HeaderText = "ABA";
        }

        public void agregarCliente()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Clientes (Cliente,Rut,Camara,Nit,Nacional,Nombre,Telefono,Direccion,Email,Ciudad,Pais,Codigo,CodigoBanco,NumCuenta,NomBanco,IBAN,Swift,ABA) VALUES (@Cliente,@Rut,@Camara,@Nit,@Nacional,@Nombre,@Telefono,@Direccion,@Email,@Ciudad,@Pais,@Codigo,@CodigoBanco,@NumCuenta,@NomBanco,@IBAN,@Swift,@ABA)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                string camara = textBox12.Text;
                if (!camara.Equals(""))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text);
                    string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text, "CamaraDeComercio*");
                    if (prueba.Length > 0)
                    {
                        if (File.Exists(prueba[0]))
                        {

                            File.Delete(prueba[0]);
                        }
                    }
                    using (FileStream fs = File.Open(camara, FileMode.Open))
                    {
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text);
                        string ext = Path.GetExtension(camara);
                        fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text + "\\CamaraDeComercio" + ext));
                    }
                }
                string RUT = textBox13.Text;
                if (!RUT.Equals(""))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text);
                    string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text, "RUT*");
                    if (prueba.Length > 0)
                    {
                        if (File.Exists(prueba[0]))
                        {

                            File.Delete(prueba[0]);
                        }
                    }
                    using (FileStream fs = File.Open(RUT, FileMode.Open))
                    {
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text);
                        string ext = Path.GetExtension(RUT);
                        fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text + "\\RUT" + ext));
                    }
                }
                string Bancaria = textBox20.Text;
                if (!Bancaria.Equals(""))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text);
                    string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text, "CertificacionBancaria*");
                    if (prueba.Length > 0)
                    {
                        if (File.Exists(prueba[0]))
                        {

                            File.Delete(prueba[0]);
                        }
                    }
                    using (FileStream fs = File.Open(Bancaria, FileMode.Open))
                    {
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text);
                        string ext = Path.GetExtension(Bancaria);
                        fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text + "\\CertificacionBancaria" + ext));
                    }
                }
                cmd.Parameters.Add("@Proveedor", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@Rut", OleDbType.VarChar).Value = textBox2.Text;
                cmd.Parameters.Add("@Camara", OleDbType.VarChar).Value = textBox3.Text;
                cmd.Parameters.Add("@Nit", OleDbType.VarChar).Value = textBox4.Text;
                cmd.Parameters.Add("@Nacional", OleDbType.VarChar).Value = comboBox2.Text;
                cmd.Parameters.Add("@Nombre", OleDbType.VarChar).Value = textBox6.Text;
                cmd.Parameters.Add("@Telefono", OleDbType.VarChar).Value = textBox7.Text;
                cmd.Parameters.Add("@Direccion", OleDbType.VarChar).Value = textBox8.Text;
                cmd.Parameters.Add("@Email", OleDbType.VarChar).Value = textBox9.Text;
                cmd.Parameters.Add("@Ciudad", OleDbType.VarChar).Value = textBox10.Text;
                cmd.Parameters.Add("@Pais", OleDbType.VarChar).Value = textBox11.Text;
                cmd.Parameters.Add("@Codigo", OleDbType.VarChar).Value = textBox5.Text;
                cmd.Parameters.Add("@CodigoBanco", OleDbType.VarChar).Value = textBox17.Text;
                cmd.Parameters.Add("@NumCuenta", OleDbType.VarChar).Value = textBox16.Text;
                cmd.Parameters.Add("@NomBanco", OleDbType.VarChar).Value = textBox14.Text;
                cmd.Parameters.Add("@IBAN", OleDbType.VarChar).Value = textBox18.Text;
                cmd.Parameters.Add("@Swift", OleDbType.VarChar).Value = textBox19.Text;
                cmd.Parameters.Add("@ABA", OleDbType.VarChar).Value = textBox15.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Cliente agregado.");
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

        public void modificarCliente()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Clientes SET Cliente=@Cliente,Rut=@Rut,Camara=@Camara,Nit=@Nit,Nacional=@Nacional,Nombre=@Nombre,Telefono=@Telefono,Direccion=@Direccion,Email=@Email,Ciudad=@Ciudad,Pais=@Pais,Codigo=@Codigo,CodigoBanco=@CodigoBanco,NumCuenta=@NumCuenta,NomBanco=@NomBanco,IBAN=@IBAN,Swift=@Swift,ABA=@ABA WHERE ID = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                string camara = textBox12.Text;
                if (!camara.Equals(""))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text);
                    string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text, "CamaraDeComercio*");
                    if (prueba.Length > 0)
                    {
                        if (File.Exists(prueba[0]))
                        {

                            File.Delete(prueba[0]);
                        }
                    }
                    using (FileStream fs = File.Open(camara, FileMode.Open))
                    {
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text);
                        string ext = Path.GetExtension(camara);
                        fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text + "\\CamaraDeComercio" + ext));
                    }
                }
                string RUT = textBox13.Text;
                if (!RUT.Equals(""))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text);
                    string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text, "RUT*");
                    if (prueba.Length > 0)
                    {
                        if (File.Exists(prueba[0]))
                        {

                            File.Delete(prueba[0]);
                        }
                    }
                    using (FileStream fs = File.Open(RUT, FileMode.Open))
                    {
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text);
                        string ext = Path.GetExtension(RUT);
                        fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text + "\\RUT" + ext));
                    }
                }
                string Bancaria = textBox20.Text;
                if (!Bancaria.Equals(""))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text);
                    string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text, "CertificacionBancaria*");
                    if (prueba.Length > 0)
                    {
                        if (File.Exists(prueba[0]))
                        {

                            File.Delete(prueba[0]);
                        }
                    }
                    using (FileStream fs = File.Open(Bancaria, FileMode.Open))
                    {
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text);
                        string ext = Path.GetExtension(Bancaria);
                        fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text + "\\CertificacionBancaria" + ext));
                    }
                }
                cmd.Parameters.Add("@Proveedor", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@Rut", OleDbType.VarChar).Value = textBox2.Text;
                cmd.Parameters.Add("@Camara", OleDbType.VarChar).Value = textBox3.Text;
                cmd.Parameters.Add("@Nit", OleDbType.VarChar).Value = textBox4.Text;
                cmd.Parameters.Add("@Nacional", OleDbType.VarChar).Value = comboBox2.Text;
                cmd.Parameters.Add("@Nombre", OleDbType.VarChar).Value = textBox6.Text;
                cmd.Parameters.Add("@Telefono", OleDbType.VarChar).Value = textBox7.Text;
                cmd.Parameters.Add("@Direccion", OleDbType.VarChar).Value = textBox8.Text;
                cmd.Parameters.Add("@Email", OleDbType.VarChar).Value = textBox9.Text;
                cmd.Parameters.Add("@Ciudad", OleDbType.VarChar).Value = textBox10.Text;
                cmd.Parameters.Add("@Pais", OleDbType.VarChar).Value = textBox11.Text;
                cmd.Parameters.Add("@Codigo", OleDbType.VarChar).Value = textBox5.Text;
                cmd.Parameters.Add("@CodigoBanco", OleDbType.VarChar).Value = textBox17.Text;
                cmd.Parameters.Add("@NumCuenta", OleDbType.VarChar).Value = textBox16.Text;
                cmd.Parameters.Add("@NomBanco", OleDbType.VarChar).Value = textBox14.Text;
                cmd.Parameters.Add("@IBAN", OleDbType.VarChar).Value = textBox18.Text;
                cmd.Parameters.Add("@Swift", OleDbType.VarChar).Value = textBox19.Text;
                cmd.Parameters.Add("@ABA", OleDbType.VarChar).Value = textBox15.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Cliente modificado.");
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

        public void buscarClientes()
        {
            string query = "SELECT * FROM Clientes ";
            int i = 0;
            if (!textBox1.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Cliente LIKE '%" + textBox1.Text + "%'";
            }
            if (!textBox2.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Rut LIKE '%" + textBox2.Text + "%'";
            }
            if (!textBox3.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Camara LIKE '%" + textBox3.Text + "%'";
            }
            if (!textBox4.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Nit LIKE '%" + textBox4.Text + "%'";
            }
            if (!comboBox2.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Nacional LIKE '%" + comboBox2.Text + "%'";
            }
            if (!textBox6.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Nombre LIKE '%" + textBox6.Text + "%'";
            }
            if (!textBox7.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Telefono LIKE '%" + textBox7.Text + "%'";
            }
            if (!textBox8.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Direccion LIKE '%" + textBox8.Text + "%'";
            }
            if (!textBox9.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Email LIKE '%" + textBox9.Text + "%'";
            }
            if (!textBox10.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Ciudad LIKE '%" + textBox10.Text + "%'";
            }
            if (!textBox11.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Pais LIKE '%" + textBox11.Text + "%'";
            }
            if (!textBox5.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Codigo LIKE '%" + textBox5.Text + "%'";
            }
            if (!textBox17.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "CodigoBanco LIKE '%" + textBox17.Text + "%'";
            }
            if (!textBox16.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "NumCuenta LIKE '%" + textBox16.Text + "%'";
            }
            if (!textBox14.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "NomBanco LIKE '%" + textBox14.Text + "%'";
            }
            if (!textBox18.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "IBAN LIKE '%" + textBox18.Text + "%'";
            }
            if (!textBox19.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Swift LIKE '%" + textBox19.Text + "%'";
            }
            if (!textBox15.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "ABA LIKE '%" + textBox15.Text + "%'";
            }
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable supervisores = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(supervisores);
            dataGridView1.DataSource = supervisores;
            dataGridView1.Columns[0].Visible = false;
        }

        public void reiniciarTablero()
        {
            linkLabel1.Visible = false;
            linkLabel2.Visible = false;
            linkLabel3.Visible = false;
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            comboBox2.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            textBox5.Text = "";
            textBox17.Text = "";
            textBox16.Text = "";
            textBox14.Text = "";
            textBox18.Text = "";
            textBox19.Text = "";
            textBox15.Text = "";
            textBox12.Text = "";
            textBox13.Text = "";
            textBox20.Text = "";
        }


        private void button6_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnPredial_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox12.Text = openFileDialog1.FileName;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox13.Text = openFileDialog1.FileName;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox20.Text = openFileDialog1.FileName;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            reiniciarTablero();
        }

        private void dataGridView1_CellClick_1(object sender, DataGridViewCellEventArgs e)
        {
            linkLabel1.Visible = true;
            linkLabel2.Visible = true;
            linkLabel3.Visible = true;
            textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();
            textBox2.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString();
            textBox3.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString();
            textBox4.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[4].Value.ToString();
            comboBox2.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Value.ToString();
            textBox6.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[6].Value.ToString();
            textBox7.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[7].Value.ToString();
            textBox8.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[8].Value.ToString();
            textBox9.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[9].Value.ToString();
            textBox10.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[10].Value.ToString();
            textBox11.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[11].Value.ToString();
            textBox5.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[12].Value.ToString();
            textBox17.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[13].Value.ToString();
            textBox16.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[14].Value.ToString();
            textBox14.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[15].Value.ToString();
            textBox18.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[16].Value.ToString();
            textBox19.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[17].Value.ToString();
            textBox15.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[18].Value.ToString();
        }

        private void btnPredial_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox12.Text = openFileDialog1.FileName;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            agregarCliente();
            cargarClientes();
            reiniciarTablero();
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox13.Text = openFileDialog1.FileName;
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox20.Text = openFileDialog1.FileName;
        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            modificarCliente();
            cargarClientes();
            reiniciarTablero();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el cliente " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {

                string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM Clientes WHERE id = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Cliente eliminado.");
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text);
                        string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + textBox1.Text, "CamaraDeComercio*");
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
                cargarClientes();
            }
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            buscarClientes();
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            reiniciarTablero();
        }

        private void linkLabel1_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString());
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString(), "CamaraDeComercio*");
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

        private void linkLabel2_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString());
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString(), "RUT*");
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

        private void linkLabel3_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString());
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Clientes\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString(), "CertificacionBancaria*");
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

        private void textBox12_TextChanged(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Variables.imprimir(dataGridView1);
        }
    }
}
