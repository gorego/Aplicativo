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
    public partial class frmRegistrarPaquete : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        int end = 0;
        int numEntradas = 0;
        int orden;
        int dia = 0;
        bool sw = false;
        int paquete = 0;
        int entradas = 0;

        public frmRegistrarPaquete(int op, int diaOrden)
        {
            InitializeComponent();
            orden = op;
            dia = diaOrden;
            this.Text = "Registrar Paquete OP # " + getNombreOP(op);
            cargar(comboBox1, "SELECT op.Id, op.Producto, p.Codigo, op.Cantidad, op.Volumen FROM Productos AS p INNER JOIN produccionProducto AS op ON p.ID = op.Producto WHERE Orden = " + op, "Codigo", "Producto");
            comboBox1.SelectedItem = null;
            end = 1;
        }

        public void cargar(ComboBox combo, string query, string display, string value)
        {
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            combo.DataSource = ds.Tables[0];
            combo.DisplayMember = display;
            combo.ValueMember = value;
            combo.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            combo.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public string getNombreOP(int op)
        {
            string query = "SELECT OP FROM historicoProduccion WHERE id = " + op;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            string value = "";
            try
            {
                if (myReader.Read())
                {
                    value = myReader.GetValue(0).ToString();
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return value;
        }

        public void getProducto(string op)
        {
            string query = "SELECT numAnchoEmp,numAltoEmp FROM Productos WHERE Codigo = '" + op + "'";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                if (myReader.Read())
                {
                    textBox23.Text = myReader.GetValue(0).ToString();
                    textBox22.Text = myReader.GetValue(1).ToString();
                    textBox13.Text = (myReader.GetDouble(0) * myReader.GetDouble(1)).ToString();
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

        public void getPaquete(string producto)
        {
            string query = "SELECT * FROM Paquete WHERE Producto = " + producto + " AND OP = " + orden + " AND porcentaje <> 1";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                if (myReader.Read())
                {
                    paquete = myReader.GetInt32(0);
                    radioButton2.Checked = true;
                    comboBox2.Text = myReader.GetValue(4).ToString();
                    textBox1.Text = (double.Parse(textBox13.Text) - double.Parse(myReader.GetValue(5).ToString())).ToString();
                    comboBox2.Enabled = true;
                    sw = true;
                }
                else
                    sw = false;
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
        }

        public void getAnchoAlto(int producto,DataGridView data)
        {
            string query = "SELECT ancho1,ancho2,alto1,alto2 FROM infoPaquete WHERE Paquete = " + producto;
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
                    data.Rows.Add();
                    data.Rows[i].Cells[0].Value = myReader.GetDouble(0);
                    data.Rows[i].Cells[1].Value = myReader.GetDouble(1);
                    data.Rows[i].Cells[2].Value = myReader.GetDouble(2);
                    data.Rows[i].Cells[3].Value = myReader.GetDouble(3);
                    numEntradas++;
                    entradas++;
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
                
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                textBox1.ReadOnly = false;
                comboBox2.Enabled = false;
            }
            else
            {
                textBox1.ReadOnly = true;
                textBox1.Text = "0";
                comboBox2.Enabled = true;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (end == 1)
            {
                getProducto(comboBox1.Text);
                getPaquete(comboBox1.SelectedValue.ToString());
                if (sw)
                    getAnchoAlto(paquete,dataGridView1);
            }
        }

        public void reiniciar()
        {
            textBox2.Text = "0";
            textBox3.Text = "0";
            textBox5.Text = "0";
            textBox4.Text = "0";
        }

        public void registrarPaquete()
        {
            if (!textBox2.Text.Equals("0") && !textBox2.Text.Equals(""))
            {
                if (!textBox3.Text.Equals("0") && !textBox3.Text.Equals(""))
                {
                    if (!textBox5.Text.Equals("0") && !textBox5.Text.Equals(""))
                    {
                        if (!textBox4.Text.Equals("0") && !textBox4.Text.Equals(""))
                        {
                            dataGridView1.Rows.Add();
                            dataGridView1.Rows[numEntradas].Cells[0].Value = textBox2.Text;
                            dataGridView1.Rows[numEntradas].Cells[1].Value = textBox3.Text;
                            dataGridView1.Rows[numEntradas].Cells[2].Value = textBox5.Text;
                            dataGridView1.Rows[numEntradas].Cells[3].Value = textBox4.Text;
                            numEntradas++;
                            reiniciar();
                            textBox2.Focus();
                        }
                        else
                            MessageBox.Show("Favor insertar un numero superior a 0 en el Alto 2.");
                    }
                    else
                        MessageBox.Show("Favor insertar un numero superior a 0 en el Alto 1.");
                }
                else
                    MessageBox.Show("Favor insertar un numero superior a 0 en el Ancho 2.");
            }
            else
                MessageBox.Show("Favor insertar un numero superior a 0 en el Ancho 1.");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            registrarPaquete();
        }

        public int getMaxID()
        {
            string query = "SELECT MAX(num) FROM Paquete WHERE OP = " + orden;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int id = 0;
            string value = "";
            try
            {
                if (myReader.Read())
                {
                    value = myReader.GetValue(0).ToString();
                    if (value.Equals(""))
                        id = 0;
                    else
                        id = Int32.Parse(value);
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

        public int getMaxPaquete()
        {
            string query = "SELECT MAX(ID) FROM Paquete";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int id = 0;
            string value = "";
            try
            {
                if (myReader.Read())
                {
                    value = myReader.GetValue(0).ToString();
                    if (value.Equals(""))
                        id = 0;
                    else
                        id = Int32.Parse(value);
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

        public void agregarRegistro(int id, string nombre)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Paquete (numPaquete,Producto,OP,Bodega,numPiezas,num,dia,porcentaje,Fecha,Hora) VALUES (@numPaquete,@Producto,@OP,@Bodega,@numPiezas,@num,@dia,@porcentaje,@Fecha,@Hora)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@numPaquete", OleDbType.VarChar).Value = "P - " + nombre + " - " + id.ToString().PadLeft(4, '0');
                cmd.Parameters.Add("@Producto", OleDbType.VarChar).Value = comboBox1.SelectedValue;
                cmd.Parameters.Add("@OP", OleDbType.VarChar).Value = orden;
                if(!comboBox2.Text.Equals(""))
                    cmd.Parameters.Add("@Bodega", OleDbType.VarChar).Value = comboBox2.Text;
                else
                    cmd.Parameters.Add("@Bodega", OleDbType.VarChar).Value = 0;
                if(!radioButton1.Checked)
                    cmd.Parameters.Add("@numPiezas", OleDbType.VarChar).Value = textBox1.Text;
                else
                    cmd.Parameters.Add("@numPiezas", OleDbType.VarChar).Value = textBox13.Text;
                cmd.Parameters.Add("@num", OleDbType.VarChar).Value = id;
                cmd.Parameters.Add("@dia", OleDbType.VarChar).Value = dia;
                if (textBox1.ReadOnly == false)
                    cmd.Parameters.Add("@porcentaje", OleDbType.VarChar).Value = Math.Round(double.Parse(textBox1.Text) / double.Parse(textBox13.Text), 1, MidpointRounding.AwayFromZero);
                else
                    cmd.Parameters.Add("@porcentaje", OleDbType.VarChar).Value = 1;
                cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = DateTime.Now.ToString("dd") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.Year;
                cmd.Parameters.Add("@Hora", OleDbType.VarChar).Value = DateTime.Now.ToString("HH") + ":" + DateTime.Now.ToString("mm") + ":" + DateTime.Now.ToString("ss"); 
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

        public void modificarRegistro(int paquete)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Paquete SET Bodega=@Bodega,numPiezas=numPiezas+@piezas,dia=@dia,porcentaje=porcentaje+@porcentaje2,Fecha=@Fecha,Hora=@Hora WHERE ID = " + paquete);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                if (!comboBox2.Text.Equals(""))
                    cmd.Parameters.Add("@Bodega", OleDbType.VarChar).Value = comboBox2.Text;
                else
                    cmd.Parameters.Add("@Bodega", OleDbType.VarChar).Value = 0;
                cmd.Parameters.Add("@piezas", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@dia", OleDbType.VarChar).Value = dia;
                if (textBox1.ReadOnly == false)
                    cmd.Parameters.Add("@porcentaje2", OleDbType.VarChar).Value = Math.Round(double.Parse(textBox1.Text) / double.Parse(textBox13.Text), 1, MidpointRounding.AwayFromZero);
                else
                    cmd.Parameters.Add("@porcentaje2", OleDbType.VarChar).Value = 1;
                cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = DateTime.Now.ToString("dd") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.Year;
                cmd.Parameters.Add("@Hora", OleDbType.VarChar).Value = DateTime.Now.ToString("HH") + ":" + DateTime.Now.ToString("mm") + ":" + DateTime.Now.ToString("ss");
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

        public void agregarAnchoAlto(int id, DataGridView data, int entradas)
        {
            for (int i = entradas; i < data.Rows.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("INSERT INTO infoPaquete (Paquete,ancho1,ancho2,alto1,alto2) VALUES (@Paquete,@ancho1,@ancho2,@alto1,@alto2)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Paquete", OleDbType.VarChar).Value = id;
                cmd.Parameters.Add("@ancho1", OleDbType.VarChar).Value = data.Rows[i].Cells[0].Value;
                cmd.Parameters.Add("@ancho2", OleDbType.VarChar).Value = data.Rows[i].Cells[1].Value;
                cmd.Parameters.Add("@alto1", OleDbType.VarChar).Value = data.Rows[i].Cells[2].Value;
                cmd.Parameters.Add("@alto2", OleDbType.VarChar).Value = data.Rows[i].Cells[3].Value;
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
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (numEntradas < 5)
            {
                if (radioButton2.Checked)
                {
                    if (!sw)
                    {
                        int id = getMaxID();
                        string nombre = getNombreOP(orden);
                        agregarRegistro(id + 1, nombre);
                        int id2 = getMaxPaquete();
                        agregarAnchoAlto(id2, dataGridView1,0);
                        MessageBox.Show("Paquete Registrado.");
                        this.Close();
                    }
                    else
                        MessageBox.Show("Favor llenar 5 registros.");
                }
                else
                    MessageBox.Show("Favor llenar 5 registros.");
            }
            else
            {
                if (!sw)
                {
                    int id = getMaxID();
                    string nombre = getNombreOP(orden);
                    agregarRegistro(id + 1, nombre);
                    int id2 = getMaxPaquete();
                    agregarAnchoAlto(id2, dataGridView1,0);
                    MessageBox.Show("Paquete Registrado.");
                    this.Close();
                }
                else
                {
                    modificarRegistro(paquete);
                    agregarAnchoAlto(paquete, dataGridView1,entradas);
                    MessageBox.Show("Paquete terminado.");
                    this.Close();
                }
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                registrarPaquete();
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                registrarPaquete();
            }
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                registrarPaquete();
            }
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                registrarPaquete();
            }
        }

    }
}
