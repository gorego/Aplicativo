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
    public partial class frmProveedor : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();

        public frmProveedor()
        {
            InitializeComponent();
            cargarProveedores();
            dataGridView1.Columns[1].DefaultCellStyle.Font = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Underline);
            dataGridView1.Columns[10].DefaultCellStyle.Font = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Underline);
            linkLabel1.Visible = false;
            linkLabel2.Visible = false;
            linkLabel3.Visible = false;
        }

        public void cargarProveedores()
        {
            string query = "SELECT * FROM Proveedores WHERE Proveedor <> 'N/A'";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable supervisores = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(supervisores);
            dataGridView1.DataSource = supervisores;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[3].HeaderText = "Camara de Comercio";
            dataGridView1.Columns[5].HeaderText = "Tipo de Proveedor";
            dataGridView1.Columns[6].HeaderText = "Extranjero/Nacional";
            dataGridView1.Columns[7].HeaderText = "Nombre de Contacto";
            dataGridView1.Columns[8].HeaderText = "Telefono de Contacto";
            dataGridView1.Columns[9].HeaderText = "Direccion de Contacto";
            dataGridView1.Columns[10].HeaderText = "Email de Contacto";
            dataGridView1.Columns[11].HeaderText = "Ciudad de Contacto";
            dataGridView1.Columns[12].HeaderText = "Pais de Contacto";
            dataGridView1.Columns[13].HeaderText = "Codigo Postal";
            dataGridView1.Columns[14].HeaderText = "Codigo del Banco";
            dataGridView1.Columns[15].HeaderText = "# de Cuenta";
            dataGridView1.Columns[16].HeaderText = "Nombre del Banco";
            dataGridView1.Columns[17].HeaderText = "IBAN";
            dataGridView1.Columns[18].HeaderText = "Swift";
            dataGridView1.Columns[19].HeaderText = "ABA";
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

        public void agregarProveedor()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Proveedores (Proveedor,Rut,Camara,Nit,Tipo,Nacional,Nombre,Telefono,Direccion,Email,Ciudad,Pais,Codigo,CodigoBanco,NumCuenta,NomBanco,IBAN,Swift,ABA) VALUES (@Proveedor,@Rut,@Camara,@Nit,@Tipo,@Nacional,@Nombre,@Telefono,@Direccion,@Email,@Ciudad,@Pais,@Codigo,@CodigoBanco,@NumCuenta,@NomBanco,@IBAN,@Swift,@ABA)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                string camara = textBox12.Text;
                if (!camara.Equals(""))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text);
                    string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text, "CamaraDeComercio*");
                    if (prueba.Length > 0)
                    {
                        if (File.Exists(prueba[0]))
                        {

                            File.Delete(prueba[0]);
                        }
                    }
                    using (FileStream fs = File.Open(camara, FileMode.Open))
                    {
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text);
                        string ext = Path.GetExtension(camara);
                        fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text + "\\CamaraDeComercio" + ext));
                    }
                }
                string RUT = textBox13.Text;
                if (!RUT.Equals(""))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text);
                    string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text, "RUT*");
                    if (prueba.Length > 0)
                    {
                        if (File.Exists(prueba[0]))
                        {

                            File.Delete(prueba[0]);
                        }
                    }
                    using (FileStream fs = File.Open(RUT, FileMode.Open))
                    {
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text);
                        string ext = Path.GetExtension(RUT);
                        fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text + "\\RUT" + ext));
                    }
                }
                string Bancaria = textBox20.Text;
                if (!Bancaria.Equals(""))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text);
                    string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text, "CertificacionBancaria*");
                    if (prueba.Length > 0)
                    {
                        if (File.Exists(prueba[0]))
                        {

                            File.Delete(prueba[0]);
                        }
                    }
                    using (FileStream fs = File.Open(Bancaria, FileMode.Open))
                    {
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text);
                        string ext = Path.GetExtension(Bancaria);
                        fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text + "\\CertificacionBancaria" + ext));
                    }
                }
                cmd.Parameters.Add("@Proveedor", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@Rut", OleDbType.VarChar).Value = textBox2.Text;
                cmd.Parameters.Add("@Camara", OleDbType.VarChar).Value = textBox3.Text;
                cmd.Parameters.Add("@Nit", OleDbType.VarChar).Value = textBox4.Text;
                cmd.Parameters.Add("@Tipo", OleDbType.VarChar).Value = comboBox1.Text;
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
                    MessageBox.Show("Proveedor agregado.");
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

        public void modificarProveedor()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Proveedores SET Proveedor=@Proveedor,Rut=@Rut,Camara=@Camara,Nit=@Nit,Tipo=@Tipo,Nacional=@Nacional,Nombre=@Nombre,Telefono=@Telefono,Direccion=@Direccion,Email=@Email,Ciudad=@Ciudad,Pais=@Pais,Codigo=@Codigo,CodigoBanco=@CodigoBanco,NumCuenta=@NumCuenta,NomBanco=@NomBanco,IBAN=@IBAN,Swift=@Swift,ABA=@ABA WHERE ID = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                string camara = textBox12.Text;
                if (!camara.Equals(""))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text);
                    string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text, "CamaraDeComercio*");
                    if (prueba.Length > 0)
                    {
                        if (File.Exists(prueba[0]))
                        {

                            File.Delete(prueba[0]);
                        }
                    }
                    using (FileStream fs = File.Open(camara, FileMode.Open))
                    {
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text);
                        string ext = Path.GetExtension(camara);
                        fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text + "\\CamaraDeComercio" + ext));
                    }
                }
                string RUT = textBox13.Text;
                if (!RUT.Equals(""))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text);
                    string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text, "RUT*");
                    if (prueba.Length > 0)
                    {
                        if (File.Exists(prueba[0]))
                        {

                            File.Delete(prueba[0]);
                        }
                    }
                    using (FileStream fs = File.Open(RUT, FileMode.Open))
                    {
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text);
                        string ext = Path.GetExtension(RUT);
                        fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text + "\\RUT" + ext));
                    }
                }
                string Bancaria = textBox20.Text;
                if (!Bancaria.Equals(""))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text);
                    string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text, "CertificacionBancaria*");
                    if (prueba.Length > 0)
                    {
                        if (File.Exists(prueba[0]))
                        {

                            File.Delete(prueba[0]);
                        }
                    }
                    using (FileStream fs = File.Open(Bancaria, FileMode.Open))
                    {
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text);
                        string ext = Path.GetExtension(Bancaria);
                        fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text + "\\CertificacionBancaria" + ext));
                    }
                }
                cmd.Parameters.Add("@Proveedor", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@Rut", OleDbType.VarChar).Value = textBox2.Text;
                cmd.Parameters.Add("@Camara", OleDbType.VarChar).Value = textBox3.Text;
                cmd.Parameters.Add("@Nit", OleDbType.VarChar).Value = textBox4.Text;
                cmd.Parameters.Add("@Tipo", OleDbType.VarChar).Value = comboBox1.Text;
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
                    MessageBox.Show("Proveedor modificado.");
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

        public void buscarProveedor()
        {
            string query = "SELECT * FROM Proveedores ";
            int i = 0;
            if (!textBox1.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Proveedor LIKE '%" + textBox1.Text + "%'";
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
            if (!comboBox1.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Tipo LIKE '%" + comboBox1.Text + "%'";
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

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentCell.ColumnIndex == 1)
            {
                frmProveedorInsumo newFrm = new frmProveedorInsumo(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
                newFrm.Show();
            }
            linkLabel1.Visible = true;
            linkLabel2.Visible = true;
            linkLabel3.Visible = true;
            linkLabel4.Visible = true;
            textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();
            textBox2.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString();
            textBox3.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString();
            textBox4.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[4].Value.ToString();
            comboBox1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Value.ToString();
            comboBox2.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[6].Value.ToString();
            textBox6.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[7].Value.ToString();
            textBox7.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[8].Value.ToString();
            textBox8.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[9].Value.ToString();
            textBox9.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[10].Value.ToString();
            textBox10.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[11].Value.ToString();
            textBox11.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[12].Value.ToString();
            textBox5.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[13].Value.ToString();
            textBox17.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[14].Value.ToString();
            textBox16.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[15].Value.ToString();
            textBox14.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[16].Value.ToString();
            textBox18.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[17].Value.ToString();
            textBox19.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[18].Value.ToString();
            textBox15.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[19].Value.ToString();
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
            comboBox1.Text = "";
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
        }

        private void button5_Click(object sender, EventArgs e)
        {
            reiniciarTablero();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el proveedor " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {

                string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM Proveedores WHERE id = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Proveedor eliminado.");
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text);
                        string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + textBox1.Text, "Contrato*");
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
                cargarProveedores();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            modificarProveedor();
            cargarProveedores();
            reiniciarTablero();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            buscarProveedor();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString());
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString(), "CamaraDeComercio*");
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

        public void agregarTransportador(int id)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Transportadores (Nombres,Apellidos,Cedula,Licencia,Placa,Proveedor) VALUES (@Nombres,@Apellidos,@Cedula,@Licencia,@Placa,@Proveedor)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Nombres", OleDbType.VarChar).Value = "N/A";
                cmd.Parameters.Add("@Apellidos", OleDbType.VarChar).Value = "N/A";
                cmd.Parameters.Add("@Cedula", OleDbType.VarChar).Value = "N/A";
                cmd.Parameters.Add("@Licencia", OleDbType.VarChar).Value = 0;
                cmd.Parameters.Add("@Placa", OleDbType.VarChar).Value = "N/A";
                cmd.Parameters.Add("@Proveedor", OleDbType.VarChar).Value = id;
                try
                {
                    cmd.ExecuteNonQuery();
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

        public int getMaxID()
        {
            string query = "SELECT MAX(ID) FROM Proveedores";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int id = 0;
            try
            {
                if (myReader.Read())
                {
                    id = myReader.GetInt32(0);
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return id;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            agregarProveedor();
            cargarProveedores();
            int id = getMaxID();
            agregarTransportador(id);
            reiniciarTablero();
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

        private void linkLabel1_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString());
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString(), "RUT*");
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

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString());
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Proveedores\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString(), "CertificacionBancaria*");
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

        private void button6_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmTransportador newFrm = new frmTransportador(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            newFrm.Show();
        }

        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Variables.imprimir(dataGridView1);
        }        
    }
}
