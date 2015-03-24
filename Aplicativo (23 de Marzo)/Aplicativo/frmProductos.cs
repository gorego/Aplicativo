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
    public partial class frmProductos : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        public frmProductos()
        {
            InitializeComponent();
            cargarProductos();
        }

        public void agregarProducto()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Productos(Codigo,Caracteristica,Variabilidad,Condiciones,Precio,anchoProd,altoProd,largoProd,anchoFact,altoFact,largoFact,anchoEmp,altoEmp,largoEmp,distSeparador,xSeparador,ySeparador,numAnchoEmp,numAltoEmp,Especie) VALUES (@Codigo,@Caracteristica,@Variabilidad,@Condiciones,@Precio,@anchoProd,@altoProd,@largoProd,@anchoFact,@altoFact,@largoFact,@anchoEmp,@altoEmp,@largoEmp,@distSeparador,@xSeparador,@ySeparador,@numAnchoEmp,@numAltoEmp,Especie)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                string descripcion = textBox21.Text;
                if (!descripcion.Equals(""))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Anexos\\Productos\\" + txtCodigo.Text);
                    string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Productos\\" + txtCodigo.Text, "Descripcion*");
                    if (prueba.Length > 0)
                    {
                        if (File.Exists(prueba[0]))
                        {

                            File.Delete(prueba[0]);
                        }
                    }
                    using (FileStream fs = File.Open(descripcion, FileMode.Open))
                    {
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Productos\\" + txtCodigo.Text);
                        string ext = Path.GetExtension(descripcion);
                        fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Productos\\" + txtCodigo.Text + "\\Descripcion" + ext));
                    }
                }
                cmd.Parameters.Add("@Codigo", OleDbType.VarChar).Value = txtCodigo.Text;
                cmd.Parameters.Add("@Caracteristica", OleDbType.VarChar).Value = comboBox1.Text;
                cmd.Parameters.Add("@Variabilidad", OleDbType.VarChar).Value = textBox11.Text;
                cmd.Parameters.Add("@Condiciones", OleDbType.VarChar).Value = textBox12.Text;
                cmd.Parameters.Add("@Precio", OleDbType.VarChar).Value = textBox14.Text;
                cmd.Parameters.Add("@anchoProd", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@altoProd", OleDbType.VarChar).Value = textBox2.Text;
                cmd.Parameters.Add("@largoProd", OleDbType.VarChar).Value = textBox3.Text;
                cmd.Parameters.Add("@anchoFact", OleDbType.VarChar).Value = textBox10.Text;
                cmd.Parameters.Add("@altoFact", OleDbType.VarChar).Value = textBox9.Text;
                cmd.Parameters.Add("@largoFact", OleDbType.VarChar).Value = textBox8.Text;
                cmd.Parameters.Add("@anchoEmp", OleDbType.VarChar).Value = textBox17.Text;
                cmd.Parameters.Add("@altoEmp", OleDbType.VarChar).Value = textBox16.Text;
                cmd.Parameters.Add("@largoEmp", OleDbType.VarChar).Value = textBox15.Text;                
                cmd.Parameters.Add("@distSeparador", OleDbType.VarChar).Value = textBox20.Text;
                cmd.Parameters.Add("@xSeparador", OleDbType.VarChar).Value = textBox18.Text;
                cmd.Parameters.Add("@ySeparador", OleDbType.VarChar).Value = textBox19.Text;
                cmd.Parameters.Add("@numAnchoEmp", OleDbType.VarChar).Value = textBox23.Text;
                cmd.Parameters.Add("@numAltoEmp", OleDbType.VarChar).Value = textBox22.Text;
                cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = txtEspecie.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Producto agregado.");
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

        public void modificarProducto()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Productos SET Codigo=@Codigo,Caracteristica=@Caracteristica,Variabilidad=@Variabilidad,Condiciones=@Condiciones,Precio=@Precio,anchoProd=@anchoProd,altoProd=@altoProd,largoProd=@largoProd,anchoFact=@anchoFact,altoFact=@altoFact,largoFact=@largoFact,anchoEmp=@anchoEmp,altoEmp=@altoEmp,largoEmp=@largoEmp,distSeparador=@distSeparador,xSeparador=@xSeparador,ySeparador=@ySeparador,numAnchoEmp=@numAnchoEmp,numAltoEmp=@numAltoEmp,Especie=@Especie WHERE ID = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                string descripcion = textBox21.Text;
                if (!descripcion.Equals(""))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Productos\\" + txtCodigo.Text);
                    string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Productos\\" + txtCodigo.Text, "Descripcion*");
                    if (prueba.Length > 0)
                    {
                        if (File.Exists(prueba[0]))
                        {

                            File.Delete(prueba[0]);
                        }
                    }
                    using (FileStream fs = File.Open(descripcion, FileMode.Open))
                    {
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Productos\\" + txtCodigo.Text);
                        string ext = Path.GetExtension(descripcion);
                        fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Productos\\" + txtCodigo.Text + "\\Descripcion" + ext));
                    }
                }
                cmd.Parameters.Add("@Codigo", OleDbType.VarChar).Value = txtCodigo.Text;
                cmd.Parameters.Add("@Caracteristica", OleDbType.VarChar).Value = comboBox1.Text;
                cmd.Parameters.Add("@Variabilidad", OleDbType.VarChar).Value = textBox11.Text;
                cmd.Parameters.Add("@Condiciones", OleDbType.VarChar).Value = textBox12.Text;
                cmd.Parameters.Add("@Precio", OleDbType.VarChar).Value = textBox14.Text;
                cmd.Parameters.Add("@anchoProd", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@altoProd", OleDbType.VarChar).Value = textBox2.Text;
                cmd.Parameters.Add("@largoProd", OleDbType.VarChar).Value = textBox3.Text;
                cmd.Parameters.Add("@anchoFact", OleDbType.VarChar).Value = textBox10.Text;
                cmd.Parameters.Add("@altoFact", OleDbType.VarChar).Value = textBox9.Text;
                cmd.Parameters.Add("@largoFact", OleDbType.VarChar).Value = textBox8.Text;
                cmd.Parameters.Add("@anchoEmp", OleDbType.VarChar).Value = textBox17.Text;
                cmd.Parameters.Add("@altoEmp", OleDbType.VarChar).Value = textBox16.Text;
                cmd.Parameters.Add("@largoEmp", OleDbType.VarChar).Value = textBox15.Text;                
                cmd.Parameters.Add("@distSeparador", OleDbType.VarChar).Value = textBox20.Text;
                cmd.Parameters.Add("@xSeparador", OleDbType.VarChar).Value = textBox18.Text;
                cmd.Parameters.Add("@ySeparador", OleDbType.VarChar).Value = textBox19.Text;
                cmd.Parameters.Add("@numAnchoEmp", OleDbType.VarChar).Value = textBox23.Text;
                cmd.Parameters.Add("@numAltoEmp", OleDbType.VarChar).Value = textBox22.Text;
                cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = txtEspecie.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Producto modificado.");
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

        public void cargarProductos()
        {
            string query = "SELECT * FROM Productos";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable supervisores = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(supervisores);
            dataGridView1.DataSource = supervisores;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[5].DefaultCellStyle.Format = "c";
            dataGridView1.Columns[6].HeaderText = "Ancho Producción";
            dataGridView1.Columns[7].HeaderText = "Alto Producción";
            dataGridView1.Columns[8].HeaderText = "Largo Producción";
            dataGridView1.Columns[9].HeaderText = "Ancho Facturación";
            dataGridView1.Columns[10].HeaderText = "Alto Facturación";
            dataGridView1.Columns[11].HeaderText = "Largo Facturación";
            dataGridView1.Columns[12].HeaderText = "Ancho Empaque";
            dataGridView1.Columns[13].HeaderText = "Alto Empaque";
            dataGridView1.Columns[14].HeaderText = "Largo Empaque";            
            dataGridView1.Columns[15].HeaderText = "Distancia Separador";
            dataGridView1.Columns[16].HeaderText = "Alto Separador";
            dataGridView1.Columns[17].HeaderText = "Ancho Separador";
            dataGridView1.Columns[18].HeaderText = "Cantidad Ancho Empaque";
            dataGridView1.Columns[19].HeaderText = "Cantidad Alto Empaque";
        }

        public void buscarProductos()
        {
            string query = "SELECT * FROM Productos ";
            int i = 0;
            if (!txtCodigo.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Codigo LIKE '%" + txtCodigo.Text + "%'";
            }
            if (!textBox1.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "anchoProd LIKE '%" + textBox1.Text + "%'";
            }
            if (!textBox2.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "altoProd LIKE '%" + textBox2.Text + "%'";
            }
            if (!textBox3.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "largoProd LIKE '%" + textBox3.Text + "%'";
            }
            if (!comboBox1.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Caracteristica LIKE '%" + comboBox1.Text + "%'";
            }
            if (!textBox11.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Variabilidad LIKE '%" + textBox11.Text + "%'";
            }
            if (!textBox12.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Condiciones LIKE '%" + textBox12.Text + "%'";
            }
            if (!textBox14.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Precio LIKE '%" + textBox14.Text + "%'";
            }
            if (!textBox10.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "anchoFact LIKE '%" + textBox10.Text + "%'";
            }
            if (!textBox9.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "altoFact LIKE '%" + textBox9.Text + "%'";
            }
            if (!textBox8.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "largoFact LIKE '%" + textBox8.Text + "%'";
            }
            if (!textBox17.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "anchoEmp LIKE '%" + textBox17.Text + "%'";
            }
            if (!textBox16.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "altoEmp LIKE '%" + textBox16.Text + "%'";
            }
            if (!textBox15.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "largoEmp LIKE '%" + textBox15.Text + "%'";
            }
            if (!textBox20.Text.Equals("") && !textBox20.Text.Equals("0"))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "distSeparador LIKE '%" + textBox20.Text + "%'";
            }
            if (!textBox18.Text.Equals("") && !textBox18.Text.Equals("0"))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "xSeparador LIKE '%" + textBox18.Text + "%'";
            }
            if (!textBox19.Text.Equals("") && !textBox19.Text.Equals("0"))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "ySeparador LIKE '%" + textBox19.Text + "%'";
            }
            if (!textBox23.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "numAnchoEmp LIKE '%" + textBox23.Text + "%'";
            }
            if (!textBox22.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "numAltoEmp LIKE '%" + textBox22.Text + "%'";
            }
            if (!txtEspecie.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Especie LIKE '%" + txtEspecie.Text + "%'";
            }
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable supervisores = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(supervisores);
            dataGridView1.DataSource = supervisores;
            dataGridView1.Columns[0].Visible = false;
        }

        public void eliminarProducto()
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el producto " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {

                string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM Productos WHERE id = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Producto eliminado.");
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Productos\\" + txtCodigo.Text);
                        string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Productos\\" + txtCodigo.Text, "Descripcion*");
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
            }
        }

        public void reiniciarTablero()
        {
            linkLabel1.Visible = false;
            txtCodigo.Text = "";
            comboBox1.Text = "";
            textBox1.Text = "0";
            textBox2.Text = "0";
            textBox3.Text = "0";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "0";
            textBox9.Text = "0";
            textBox10.Text = "0";
            textBox11.Text = "";
            textBox12.Text = "0";
            textBox13.Text = "0";
            textBox14.Text = "0";
            textBox15.Text = "0";
            textBox16.Text = "0";
            textBox17.Text = "0";                        
            textBox18.Text = "0";
            textBox23.Text = "0";
            textBox22.Text = "0";
            textBox19.Text = "0";
            textBox20.Text = "0";
            textBox21.Text = "";
            txtEspecie.Text = "";
        }

        private void btnExaminar_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox21.Text = openFileDialog1.FileName;
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (!textBox1.Text.Equals("") && !textBox2.Text.Equals("") && !textBox3.Text.Equals("") && !txtEspecie.Text.Equals(""))
            {
                textBox6.Text = ((Double.Parse(textBox1.Text) * Double.Parse(textBox2.Text) * Double.Parse(textBox3.Text)) / 1000000000).ToString();
                textBox5.Text = ((Double.Parse(textBox6.Text) * 424)).ToString();
                string cod = "";
                if (comboBox1.SelectedIndex == 0)
                    cod = "DIM ";
                else if (comboBox1.SelectedIndex == 1)
                    cod = "SD ";
                else if (comboBox1.SelectedIndex == 2)
                    cod = "SC ";
                else
                    cod = "SLATS ";
                string especie = "";
                if (txtEspecie.Text.Equals("Teca"))
                    especie = "TG ";
                else if (txtEspecie.Text.Equals("Melina"))
                    especie = "GA ";
                else if (txtEspecie.Text.Equals("Ceiba Roja"))
                    especie = "BQ ";
                else if (txtEspecie.Text.Equals("Eucalipto"))
                    especie = "ET ";
                else
                    especie = "RF ";
                txtCodigo.Text = cod + especie + "" + textBox1.Text + "-" + textBox2.Text + "-" + textBox3.Text;
            }
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void textBox2_Leave(object sender, EventArgs e)
        {
            if (!textBox1.Text.Equals("") && !textBox2.Text.Equals("") && !textBox3.Text.Equals("") && !txtEspecie.Text.Equals(""))
            {
                textBox6.Text = ((Double.Parse(textBox1.Text) * Double.Parse(textBox2.Text) * Double.Parse(textBox3.Text)) / 1000000000).ToString();
                textBox5.Text = ((Double.Parse(textBox6.Text) * 424)).ToString();
                string cod = "";
                if (comboBox1.SelectedIndex == 0)
                    cod = "DIM ";
                else if (comboBox1.SelectedIndex == 1)
                    cod = "SD ";
                else if (comboBox1.SelectedIndex == 2)
                    cod = "SC ";
                else
                    cod = "SLATS ";
                string especie = "";
                if (txtEspecie.Text.Equals("Teca"))
                    especie = "TG ";
                else if (txtEspecie.Text.Equals("Melina"))
                    especie = "GA ";
                else if (txtEspecie.Text.Equals("Ceiba Roja"))
                    especie = "BQ ";
                else if (txtEspecie.Text.Equals("Eucalipto"))
                    especie = "ET ";
                else
                    especie = "RF ";
                txtCodigo.Text = cod + especie + "" + textBox1.Text + "-" + textBox2.Text + "-" + textBox3.Text;
            }
        }

        private void textBox3_Leave(object sender, EventArgs e)
        {
            if (!textBox1.Text.Equals("") && !textBox2.Text.Equals("") && !textBox3.Text.Equals("") && !txtEspecie.Text.Equals(""))
            {                                      
                textBox6.Text = ((Double.Parse(textBox1.Text) * Double.Parse(textBox2.Text) * Double.Parse(textBox3.Text)) / 1000000000).ToString();
                textBox5.Text = ((Double.Parse(textBox6.Text) * 424)).ToString();
                string cod = "";
                if (comboBox1.SelectedIndex == 0)
                    cod = "DIM ";
                else if (comboBox1.SelectedIndex == 1)
                    cod = "SD ";
                else if (comboBox1.SelectedIndex == 2)
                    cod = "SC ";
                else
                    cod = "SLATS ";
                string especie = "";
                if (txtEspecie.Text.Equals("Teca"))
                    especie = "TG ";
                else if (txtEspecie.Text.Equals("Melina"))
                    especie = "GA ";
                else if (txtEspecie.Text.Equals("Ceiba Roja"))
                    especie = "BQ ";
                else if (txtEspecie.Text.Equals("Eucalipto"))
                    especie = "ET ";
                else
                    especie = "RF ";
                txtCodigo.Text = cod + especie + "" + textBox1.Text + "-" + textBox2.Text + "-" + textBox3.Text;
            }
        }

        private void textBox10_Leave(object sender, EventArgs e)
        {
            if (!textBox10.Text.Equals("") && !textBox9.Text.Equals("") && !textBox8.Text.Equals(""))
            {
                textBox7.Text = ((Double.Parse(textBox10.Text) * Double.Parse(textBox9.Text) * Double.Parse(textBox8.Text)) / 1000000000).ToString();
                textBox4.Text = ((Double.Parse(textBox7.Text) * 424)).ToString();
            }
        }

        private void textBox9_Leave(object sender, EventArgs e)
        {
            if (!textBox10.Text.Equals("") && !textBox9.Text.Equals("") && !textBox8.Text.Equals(""))
            {
                textBox7.Text = ((Double.Parse(textBox10.Text) * Double.Parse(textBox9.Text) * Double.Parse(textBox8.Text)) / 1000000000).ToString();
                textBox4.Text = ((Double.Parse(textBox7.Text) * 424)).ToString();
            }
        }

        private void textBox8_Leave(object sender, EventArgs e)
        {
            if (!textBox10.Text.Equals("") && !textBox9.Text.Equals("") && !textBox8.Text.Equals(""))
            {
                textBox7.Text = ((Double.Parse(textBox10.Text) * Double.Parse(textBox9.Text) * Double.Parse(textBox8.Text)) / 1000000000).ToString();
                textBox4.Text = ((Double.Parse(textBox7.Text) * 424)).ToString();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {            
            if (comboBox1.SelectedIndex == 1)
            {
                label20.Visible = true;
                label21.Visible = true;
                label22.Visible = true;
                textBox18.Visible = true;
                textBox19.Visible = true;
                textBox20.Visible = true;
            }
            else
            {
                label20.Visible = false;
                label21.Visible = false;
                label22.Visible = false;
                textBox18.Visible = false;
                textBox19.Visible = false;
                textBox20.Visible = false;
            }
            if (!textBox1.Text.Equals("") && !textBox2.Text.Equals("") && !textBox3.Text.Equals("") && !txtEspecie.Text.Equals(""))
            {
                string cod = "";
                if (comboBox1.SelectedIndex == 0)
                    cod = "DIM ";
                else if (comboBox1.SelectedIndex == 1)
                    cod = "SD ";
                else if (comboBox1.SelectedIndex == 2)
                    cod = "SC ";
                else
                    cod = "SLATS ";
                string especie = "";
                if (txtEspecie.Text.Equals("Teca"))
                    especie = "TG ";
                else if (txtEspecie.Text.Equals("Melina"))
                    especie = "GA ";
                else if (txtEspecie.Text.Equals("Ceiba Roja"))
                    especie = "BQ ";
                else if (txtEspecie.Text.Equals("Eucalipto"))
                    especie = "ET ";
                else
                    especie = "RF ";
                txtCodigo.Text = cod + especie + "" + textBox1.Text + "-" + textBox2.Text + "-" + textBox3.Text;
            }
        }

        private void btnAgregar_Click(object sender, EventArgs e)
        {
            agregarProducto();
            cargarProductos();
            reiniciarTablero();
        }

        private void btnReiniciar_Click(object sender, EventArgs e)
        {
            reiniciarTablero();
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            textBox1.Text = "";
        }

        private void textBox2_Enter(object sender, EventArgs e)
        {
            textBox2.Text = "";
        }

        private void textBox3_Enter(object sender, EventArgs e)
        {
            textBox3.Text = "";
        }

        private void textBox10_Enter(object sender, EventArgs e)
        {
            textBox10.Text = "";
        }

        private void textBox9_Enter(object sender, EventArgs e)
        {
            textBox9.Text = "";
        }

        private void textBox8_Enter(object sender, EventArgs e)
        {
            textBox8.Text = "";
        }

        private void textBox17_Enter(object sender, EventArgs e)
        {
            textBox17.Text = "";
        }

        private void textBox16_Enter(object sender, EventArgs e)
        {
            textBox16.Text = "";
        }

        private void textBox15_Enter(object sender, EventArgs e)
        {
            textBox15.Text = "";
        }

        private void textBox13_Enter(object sender, EventArgs e)
        {
            textBox13.Text = "";
        }

        private void textBox12_Enter(object sender, EventArgs e)
        {
            textBox12.Text = "";
        }

        private void textBox14_Enter(object sender, EventArgs e)
        {
            textBox14.Text = "";
        }

        private void textBox18_Enter(object sender, EventArgs e)
        {
            textBox18.Text = "";
        }

        private void textBox19_Enter(object sender, EventArgs e)
        {
            textBox19.Text = "";
        }

        private void textBox20_Enter(object sender, EventArgs e)
        {
            textBox20.Text = "";
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            linkLabel1.Visible = true;
            txtCodigo.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();
            comboBox1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString();
            textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[6].Value.ToString();
            textBox2.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[7].Value.ToString();
            textBox3.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[8].Value.ToString();
            textBox8.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[11].Value.ToString();
            textBox9.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[10].Value.ToString();
            textBox10.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[9].Value.ToString();
            textBox11.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString();
            textBox12.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[4].Value.ToString();
            textBox14.Text = String.Format("{0:c}",dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Value);
            textBox15.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[14].Value.ToString();
            textBox16.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[13].Value.ToString();
            textBox17.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[12].Value.ToString();
            textBox18.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[16].Value.ToString();
            textBox19.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[17].Value.ToString();
            textBox20.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[15].Value.ToString();
            textBox23.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[18].Value.ToString();
            textBox22.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[19].Value.ToString();
            textBox21.Text = "";
            textBox6.Text = ((Double.Parse(textBox1.Text) * Double.Parse(textBox2.Text) * Double.Parse(textBox3.Text)) / 1000000000).ToString();
            textBox5.Text = ((Double.Parse(textBox6.Text) * 424)).ToString();
            textBox7.Text = ((Double.Parse(textBox10.Text) * Double.Parse(textBox9.Text) * Double.Parse(textBox8.Text)) / 1000000000).ToString();
            textBox4.Text = ((Double.Parse(textBox7.Text) * 424)).ToString();
            textBox13.Text = ((Double.Parse(textBox22.Text) * Double.Parse(textBox23.Text))).ToString();
            txtEspecie.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[20].Value.ToString();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Productos\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString());
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Productos\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString(), "Descripcion*");
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

        private void button1_Click(object sender, EventArgs e)
        {
            modificarProducto();
            cargarProductos();
            reiniciarTablero();
        }
        
        private void btnEliminar_Click(object sender, EventArgs e)
        {
            eliminarProducto();
            cargarProductos();
            reiniciarTablero();
        }

        private void btnBusqueda_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Equals("0") && textBox2.Text.Equals("0") && textBox3.Text.Equals("0") && textBox10.Text.Equals("0") && textBox9.Text.Equals("0") && textBox8.Text.Equals("0") && textBox17.Text.Equals("0") && textBox16.Text.Equals("0") && textBox15.Text.Equals("0") && textBox13.Text.Equals("0") && textBox12.Text.Equals("0") && textBox14.Text.Equals("0"))
            {
                cargarProductos();
            }
            else
            {
                buscarProductos();
            }            
        }

        private void textBox23_Enter(object sender, EventArgs e)
        {
            textBox23.Text = "";
        }

        private void textBox23_Leave(object sender, EventArgs e)
        {
            if (!textBox23.Text.Equals("") && !textBox22.Text.Equals(""))
            {
                textBox13.Text = (Double.Parse(textBox23.Text) * Double.Parse(textBox22.Text)).ToString();
            }
        }

        private void textBox22_Leave(object sender, EventArgs e)
        {
            if (!textBox23.Text.Equals("") && !textBox22.Text.Equals(""))
            {
                textBox13.Text = (Double.Parse(textBox23.Text) * Double.Parse(textBox22.Text)).ToString();
            }
        }

        private void textBox22_Enter(object sender, EventArgs e)
        {
            textBox22.Text = "";
        }

        private void txtEspecie_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!textBox1.Text.Equals("") && !textBox2.Text.Equals("") && !textBox3.Text.Equals("") && !txtEspecie.Text.Equals(""))
            {
                textBox6.Text = ((Double.Parse(textBox1.Text) * Double.Parse(textBox2.Text) * Double.Parse(textBox3.Text)) / 1000000000).ToString();
                textBox5.Text = ((Double.Parse(textBox6.Text) * 424)).ToString();
                string cod = "";
                if (comboBox1.SelectedIndex == 0)
                    cod = "DIM ";
                else if (comboBox1.SelectedIndex == 1)
                    cod = "SD ";
                else if (comboBox1.SelectedIndex == 2)
                    cod = "SC ";
                else
                    cod = "SLATS ";
                string especie = "";
                if (txtEspecie.Text.Equals("Teca"))
                    especie = "TG ";
                else if (txtEspecie.Text.Equals("Melina"))
                    especie = "GA ";
                else if (txtEspecie.Text.Equals("Ceiba Roja"))
                    especie = "BQ ";
                else if (txtEspecie.Text.Equals("Eucalipto"))
                    especie = "ET ";
                else
                    especie = "RF ";
                txtCodigo.Text = cod + especie + "" + textBox1.Text + "-" + textBox2.Text + "-" + textBox3.Text;
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Variables.imprimir(dataGridView1);
        }

        private void textBox12_Leave(object sender, EventArgs e)
        {
            if (!textBox12.Text.Contains("%"))
            {
                textBox12.Text = textBox12.Text + "%";
            }
        }
    }
}
