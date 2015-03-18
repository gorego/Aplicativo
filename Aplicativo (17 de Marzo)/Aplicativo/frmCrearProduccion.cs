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
using System.Globalization;
using System.IO;

namespace Aplicativo
{
    public partial class frmCrearProduccion : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        List<string> empleados = new List<string>();
        int OP = 0;
        int tipousuario = 0;
        int semanaExpedicion;
        string fechaExpedicion, nomOrden;

        public frmCrearProduccion(int tipo)
        {
            InitializeComponent();
            Variables.cargar(txtOutputs, "SELECT * FROM Productos", "Codigo");
            txtOutputs.SelectedItem = null;
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd/MM/yyyy";
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "HH:mm"; // Only use hours and minutes
            dateTimePicker2.ShowUpDown = true;
            dateTimePicker3.Format = DateTimePickerFormat.Custom;
            dateTimePicker3.CustomFormat = "dd/MM/yyyy";
            label1.Text += "  " + DateTime.Now.Day + "/" + DateTime.Now.Month + "/" + DateTime.Now.Year;
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            DateTime date1 = DateTime.Now;
            Calendar cal = dfi.Calendar;
            label8.Text = "Semana de Expidición: " + cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek).ToString();
            Variables.cargar(txtCliente, "SELECT * FROM Clientes", "Cliente");
            Variables.cargar(txtEmpleado, "SELECT ID, (Nombres + '  ' + Apellidos) As nombre FROM Trabajadores", "nombre");
            Variables.cargar(comboBox5, "SELECT ID,Nombre FROM Cuadrilla", "Nombre");
            txtCliente.SelectedItem = null;
            txtEmpleado.SelectedItem = null;
            comboBox5.SelectedItem = null;
        }

        public frmCrearProduccion(int orden, int tipo)
        {
            InitializeComponent();
            Variables.cargar(txtOutputs, "SELECT * FROM Productos", "Codigo");
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd/MM/yyyy";
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "HH:mm"; // Only use hours and minutes
            dateTimePicker2.ShowUpDown = true;
            dateTimePicker3.Format = DateTimePickerFormat.Custom;
            dateTimePicker3.CustomFormat = "dd/MM/yyyy";
            Variables.cargar(txtCliente, "SELECT * FROM Clientes", "Cliente");
            Variables.cargar(txtEmpleado, "SELECT ID, (Nombres + '  ' + Apellidos) As nombre FROM Trabajadores", "nombre");
            Variables.cargar(comboBox5, "SELECT ID,Nombre FROM Cuadrilla", "Nombre");
            txtCliente.SelectedItem = null;
            txtEmpleado.SelectedItem = null;
            txtOutputs.SelectedItem = null;
            comboBox5.SelectedItem = null;
            OP = orden;
            tipousuario = tipo;
            getOrden(orden);
            button5.Text = "Modificar Orden";
            if (tipo == 1)
            {
                dateTimePicker1.Enabled = false;
                dateTimePicker2.Enabled = false;
                dateTimePicker3.Enabled = false;
                button6.Enabled = false;
                txtTipo.Enabled = false;
                txtOutputs.Enabled = false;
                txtCliente.Enabled = false;
                comboBox1.Enabled = false;
                txtDestino.Enabled = false;
                textBox1.Enabled = false;
                textBox2.Enabled = false;
                //textBox3.Enabled = false;
                //textBox8.Enabled = false;
                //textBox7.Enabled = false;
                textBox6.Enabled = false;
                //textBox5.Enabled = false;
                button2.Enabled = false;
                button3.Enabled = false;
                button5.Text = "Asignar Empleados a la Orden de Producción";
            }
            getEmpleados(orden);
            getProductos(orden);
        }

        public void getProductos(int id)
        {
            while (dataGridView1.Rows.Count != 0)
            {
                dataGridView1.Rows.RemoveAt(0);
            }
            string query = "SELECT pp.Id, pp.Orden, pp.Producto, p.Codigo, pp.Cantidad, pp.Volumen FROM Productos AS p INNER JOIN produccionProducto AS pp ON p.ID = pp.Producto WHERE Orden = " + id;
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
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[i].Cells[0].Value = myReader.GetInt32(0);
                    dataGridView1.Rows[i].Cells[1].Value = myReader.GetInt32(2);
                    dataGridView1.Rows[i].Cells[2].Value = myReader.GetString(3);
                    dataGridView1.Rows[i].Cells[3].Value = myReader.GetDouble(4);
                    dataGridView1.Rows[i].Cells[4].Value = myReader.GetDouble(5);
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


        public void getEmpleados(int id)
        {
            listBox2.Items.Clear();
            empleados.Clear();
            string query = "SELECT i.ID, i.Orden, (m.Nombres + ' ' + m.Apellidos) As nombre, m.ID FROM Trabajadores AS m INNER JOIN produccionEmpleados AS i ON m.ID = i.Trabajador WHERE i.Orden = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                while (myReader.Read())
                {
                    listBox2.Items.Add(myReader.GetString(2));
                    empleados.Add(myReader.GetInt32(3).ToString());
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

        public void getOrden(int orden)
        {
            string query = "SELECT * FROM historicoProduccion WHERE ID = " + orden;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int cliente = 0;
            string especie = "";
            try
            {
                if (myReader.Read())
                {
                    fechaExpedicion = myReader.GetString(1);                    
                    label1.Text = "Fecha de Expedición: " + myReader.GetString(1);
                    DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
                    DateTime date1 = DateTime.ParseExact(myReader.GetString(1), "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                    Calendar cal = dfi.Calendar;
                    semanaExpedicion = cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);
                    label8.Text = "Semana de Expidición: " + cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek).ToString();
                    dateTimePicker1.Value = DateTime.ParseExact(myReader.GetString(2), "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                    dateTimePicker2.Value = DateTime.ParseExact(myReader.GetString(3), "HH:mm", System.Globalization.CultureInfo.InvariantCulture);
                    dateTimePicker3.Value = DateTime.ParseExact(myReader.GetString(4), "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                    txtTipo.Text = myReader.GetString(5);
                    especie = myReader.GetString(6);
                    cliente = myReader.GetInt32(7);
                    txtDestino.Text = myReader.GetString(8);
                    textBox6.Text = myReader.GetString(9);
                    nomOrden = myReader.GetString(11);
                    label2.Text = "Orden de Producción #: " + nomOrden;
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
                txtCliente.SelectedValue = cliente;
                comboBox1.Text = especie;
            }
        }
 
        private void button2_Click(object sender, EventArgs e)
        {
            if (!txtOutputs.Text.Equals(""))
            {
                if (!textBox1.Equals("") && !textBox2.Equals(""))
                {
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[dataGridView1.Rows.Count-1].Cells[1].Value = txtOutputs.SelectedValue;
                    dataGridView1.Rows[dataGridView1.Rows.Count-1].Cells[2].Value = txtOutputs.Text;
                    dataGridView1.Rows[dataGridView1.Rows.Count-1].Cells[3].Value = textBox1.Text;
                    dataGridView1.Rows[dataGridView1.Rows.Count-1].Cells[4].Value = textBox2.Text;
                }
            }
            else
            {
                MessageBox.Show("Favor seleccionar un producto.", "Error");
            }
        }

        private void txtOutputs_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtOutputs.SelectedItem != null && !txtOutputs.SelectedValue.ToString().Equals("System.Data.DataRowView"))
            {
                textBox1.Text = "1";
                textBox2.Text = getVolumen(Int32.Parse(txtOutputs.SelectedValue.ToString())).ToString();
            }
        }

        public double getVolumen(int id)
        {
            double volumen = 0;
            string query = "SELECT anchoProd,altoProd,largoProd,numAnchoEmp,numAltoEmp FROM Productos WHERE ID = " + id;
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
                    volumen = (((Double.Parse(myReader.GetValue(0).ToString())) * (Double.Parse(myReader.GetValue(1).ToString())) * (Double.Parse(myReader.GetValue(2).ToString()))) / 1000000000) * (Double.Parse(myReader.GetValue(3).ToString()) * Double.Parse(myReader.GetValue(4).ToString()));
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return volumen;
        }

        public string getDestino(int id)
        {
            string destino = "";
            string query = "SELECT Nacional FROM Clientes WHERE ID = " + id;
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
                    destino = myReader.GetString(0);
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return destino;
        }

        public double getEmpaque(int id)
        {
            double cantidad = 0;
            string query = "SELECT numAnchoEmp,numAltoEmp FROM Productos WHERE ID = " + id;
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
                    cantidad = (Double.Parse(myReader.GetValue(0).ToString()) * Double.Parse(myReader.GetValue(1).ToString()));
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return cantidad;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (!textBox1.Text.Equals("") && !textBox2.Text.Equals(""))
            {
                double volumen = 0;
                bool isNum = double.TryParse(textBox1.Text.Trim(), out volumen);
                bool isNum2 = double.TryParse(textBox2.Text.Trim(), out volumen);
                if (isNum && isNum2)
                {                    
                    textBox2.Text = (double.Parse(textBox1.Text) * getVolumen(Int32.Parse(txtOutputs.SelectedValue.ToString()))).ToString();
                }
                else
                {
                    MessageBox.Show("Favor digitar un numero valido.", "Error");
                }
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
        }

        private void textBox2_Leave(object sender, EventArgs e)
        {
            //if (!textBox1.Text.Equals("") && !textBox2.Text.Equals(""))
            //{
            //    double volumen = 0;
            //    bool isNum = double.TryParse(textBox1.Text.Trim(), out volumen);
            //    bool isNum2 = double.TryParse(textBox2.Text.Trim(), out volumen);
            //    if (isNum && isNum2)
            //    {
            //        textBox1.Text = (double.Parse(textBox2.Text) / getEmpaque(Int32.Parse(txtOutputs.SelectedValue.ToString()))).ToString();
            //    }
            //    else
            //    {
            //        MessageBox.Show("Favor digitar un numero valido.", "Error");
            //    }
            //}
        }

        private void button3_Click(object sender, EventArgs e)
        {            
                dataGridView1.Rows.RemoveAt(dataGridView1.CurrentCell.RowIndex);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int dias = (DateTime.Now - dateTimePicker1.Value).Days;
            frmProcesamientoFormatos newFrm = new frmProcesamientoFormatos(OP, dias, tipousuario);
            if (!newFrm.IsDisposed)
            {
                this.Hide();
                newFrm.ShowDialog();
                this.Close();
            }  
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            DateTime date1 = dateTimePicker1.Value;
            Calendar cal = dfi.Calendar;
            label9.Text = "Semana de Inicio: " + cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek).ToString();
        }

        private void txtCliente_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtCliente.SelectedItem != null && !txtCliente.SelectedValue.ToString().Equals("System.Data.DataRowView"))
            {
                txtDestino.Text = getDestino(Int32.Parse(txtCliente.SelectedValue.ToString()));
            }
        }

        private void radioButton13_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton13.Checked)
            {
                textBox6.Visible = true;
                linkLabel1.Visible = true;
                textBox4.Visible = true;
                button6.Visible = true;
            }
            else
            {
                textBox6.Visible = false;
                linkLabel1.Visible = false;
                textBox4.Visible = false;
                button6.Visible = false;
            }
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            //if (!textBox7.Text.Equals(""))
            //{
            //    double volumen = 0;
            //    bool isNum = double.TryParse(textBox7.Text.Trim(), out volumen);
            //    if (isNum)
            //    {
            //        if (volumen != 0)
            //        {
            //            textBox8.Visible = true;
            //            label16.Visible = true;
            //        }
            //        else
            //        {
            //            textBox8.Visible = false;
            //            label16.Visible = false;
            //        }
            //    }
            //    else
            //    {
            //        MessageBox.Show("Favor digitar un numero valido.", "Error");
            //    }
            //}
            //else
            //{
            //    textBox8.Visible = false;
            //    label16.Visible = false;
            //}
        }

        private void radioButton13_TextChanged(object sender, EventArgs e)
        {

        }

        private void frmCrearProduccion_Load(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }

        private void label18_Click(object sender, EventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        public void cargarProductos(int producto, DataGridView dataGridView1)
        {
            string query = "SELECT * FROM Productos WHERE ID = " + producto;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable supervisores = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(supervisores);
            dataGridView1.DataSource = supervisores;
            dataGridView1.Columns[0].Visible = false;
            if (tipousuario == 1)
            {
                dataGridView1.Columns[3].Visible = false;
                dataGridView1.Columns[4].Visible = false;
                dataGridView1.Columns[5].Visible = false;
                dataGridView1.Columns[9].Visible = false;
                dataGridView1.Columns[10].Visible = false;
                dataGridView1.Columns[11].Visible = false;
            }
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

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            cargarProductos(Int32.Parse(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString()),dataGridView2);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (txtEmpleado.Text.Equals(""))
            {
                MessageBox.Show("Favor seleccionar un empleado.");
            }
            else
            {
                listBox2.Items.Add(txtEmpleado.Text);
                empleados.Add(txtEmpleado.SelectedValue.ToString());
            }      
        }

        public void getEmpleadosCuadrilla(string id)
        {
            string query = "SELECT i.ID, i.Cuadrilla, (m.Nombres + ' ' + m.Apellidos) As nombre, m.ID FROM Trabajadores AS m INNER JOIN empleadoCuadrilla AS i ON m.ID = i.Trabajador WHERE i.Cuadrilla = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                while (myReader.Read())
                {
                    listBox2.Items.Add(myReader.GetString(2));
                    empleados.Add(myReader.GetInt32(3).ToString());
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

        private void button9_Click(object sender, EventArgs e)
        {
            if (comboBox5.Text.Equals(""))
                MessageBox.Show("Favor seleccionar la cuadrilla deseada.");
            else
                getEmpleadosCuadrilla(comboBox5.SelectedValue.ToString());
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (listBox2.Text.Equals(""))
            {
                MessageBox.Show("Favor seleccionar un empleado.");
            }
            else
            {
                empleados.RemoveAt(listBox2.SelectedIndex);
                listBox2.Items.Remove(listBox2.SelectedItem);
            }
        }

        public void agregarOrden(int id)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO historicoProduccion (fechaExpedicion,fechaInicio,horaInicio,fechaFinal,Tipo,Especie,Cliente,Nacional,FSC,Estado,OP) VALUES (@fechaExpedicion,@fechaInicio,@horaInicio,@fechaFinal,@Tipo,@Especie,@Cliente,@Nacional,@FSC,@Estado,@OP)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@fechaExpedicion", OleDbType.VarChar).Value = DateTime.Now.ToString("dd") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.Year;
                cmd.Parameters.Add("@fechaInicio", OleDbType.VarChar).Value = dateTimePicker1.Value.ToString("dd") + "/" + dateTimePicker1.Value.ToString("MM") + "/" + dateTimePicker1.Value.Year;
                cmd.Parameters.Add("@horaInicio", OleDbType.VarChar).Value = dateTimePicker2.Value.ToString("HH") + ":" + dateTimePicker2.Value.ToString("mm");
                cmd.Parameters.Add("@fechaFinal", OleDbType.VarChar).Value = dateTimePicker3.Value.ToString("dd") + "/" + dateTimePicker3.Value.ToString("MM") + "/" + dateTimePicker3.Value.Year;
                cmd.Parameters.Add("@Tipo", OleDbType.VarChar).Value = txtTipo.Text;
                cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = comboBox1.Text;
                cmd.Parameters.Add("@Cliente", OleDbType.VarChar).Value = txtCliente.SelectedValue;
                cmd.Parameters.Add("@Nacional", OleDbType.VarChar).Value = txtDestino.Text;
                if(radioButton13.Checked)
                    cmd.Parameters.Add("@FSC", OleDbType.VarChar).Value = textBox6.Text;
                else
                    cmd.Parameters.Add("@FSC", OleDbType.VarChar).Value = 0;
                cmd.Parameters.Add("@Estado", OleDbType.VarChar).Value = "Activa";
                cmd.Parameters.Add("@OP", OleDbType.VarChar).Value = "OP-" + id.ToString().PadLeft(4, '0') + "-" + DateTime.Now.Year;
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

        public void modificarOrden(int id)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE historicoProduccion SET fechaInicio=@fechaInicio,horaInicio=@horaInicio,fechaFinal=@fechaFinal,Tipo=@Tipo,Especie=@Especie,Cliente=@Cliente,Nacional=@Nacional,FSC=@FSC WHERE ID = " + id);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@fechaInicio", OleDbType.VarChar).Value = dateTimePicker1.Value.ToString("dd") + "/" + dateTimePicker1.Value.ToString("MM") + "/" + dateTimePicker1.Value.Year;
                cmd.Parameters.Add("@horaInicio", OleDbType.VarChar).Value = dateTimePicker2.Value.ToString("HH") + ":" + dateTimePicker2.Value.ToString("mm");
                cmd.Parameters.Add("@fechaFinal", OleDbType.VarChar).Value = dateTimePicker3.Value.ToString("dd") + "/" + dateTimePicker3.Value.ToString("MM") + "/" + dateTimePicker3.Value.Year;
                cmd.Parameters.Add("@Tipo", OleDbType.VarChar).Value = txtTipo.Text;
                cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = comboBox1.Text;
                cmd.Parameters.Add("@Cliente", OleDbType.VarChar).Value = txtCliente.SelectedValue;
                cmd.Parameters.Add("@Nacional", OleDbType.VarChar).Value = txtDestino.Text;
                if (radioButton13.Checked)
                    cmd.Parameters.Add("@FSC", OleDbType.VarChar).Value = textBox6.Text;
                else
                    cmd.Parameters.Add("@FSC", OleDbType.VarChar).Value = 0;
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


        private void button11_Click(object sender, EventArgs e)
        {
            if (tipousuario == 0)
            {
                frmHistoricoProduccion newFrm = new frmHistoricoProduccion(0);
                this.Hide();
                newFrm.ShowDialog();
                this.Close();
            }
            else
            {
                this.Close();
            }
        }

        public int getMaxID()
        {
            string query = "SELECT MAX(ID) FROM historicoProduccion";
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

        public void agregarOrdenProductos(int id)
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("INSERT INTO produccionProducto(Orden,Producto,Cantidad,Volumen) VALUES (@Orden,@Producto,@Cantidad,@Volumen)");
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@Orden", OleDbType.VarChar).Value = id;
                    cmd.Parameters.Add("@Producto", OleDbType.VarChar).Value = dataGridView1.Rows[i].Cells[1].Value;
                    cmd.Parameters.Add("@Cantidad", OleDbType.VarChar).Value = dataGridView1.Rows[i].Cells[3].Value;
                    cmd.Parameters.Add("@Volumen", OleDbType.VarChar).Value = dataGridView1.Rows[i].Cells[4].Value;
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
            }
        }

        public void eliminarOrdenProductos(int id)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("DELETE FROM produccionProducto WHERE Orden = " + id);
            cmd.Connection = conn;
            conn.Open();

            if (conn.State == ConnectionState.Open)
            {
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


        public void agregarOrdenEmpleados(int id)
        {
            for (int i = 0; i < empleados.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("INSERT INTO produccionEmpleados(Orden,Trabajador) VALUES (@Orden,@Trabajador)");
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@Orden", OleDbType.VarChar).Value = id;
                    cmd.Parameters.Add("@Trabajador", OleDbType.VarChar).Value = empleados[i];
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
            }
        }

        public void eliminarOrdenEmpleados(int id)
        {
            for (int i = 0; i < empleados.Count + 1; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM produccionEmpleados WHERE Orden = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
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

        public bool existeAñoActual()
        {
            string query = "SELECT * FROM historicoProduccion WHERE OP like '%" + DateTime.Now.Year + "%'";
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
                    return true;
                }
                else
                {
                    return false;
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

        private void button5_Click_1(object sender, EventArgs e)
        {
            if (!txtTipo.Text.Equals(""))
            {
                if (!comboBox1.Text.Equals("")) 
                {
                    if (!txtCliente.Text.Equals(""))
                    {
                        if (button5.Text.Equals("Generar Orden"))
                        {
                            int id = getMaxID();
                            int id2 = id;
                            if (!existeAñoActual())
                                id2 = 0;          
                            agregarOrden(id2+1);
                            id++;
                            agregarOrdenEmpleados(id);
                            agregarOrdenProductos(id);
                        }
                        else if (button5.Text.Equals("Modificar Orden"))
                        {
                            modificarOrden(OP);
                            eliminarOrdenEmpleados(OP);
                            agregarOrdenEmpleados(OP);
                            eliminarOrdenProductos(OP);
                            agregarOrdenProductos(OP);
                        }
                        if (tipousuario == 0)
                        {
                            frmHistoricoProduccion newFrm = new frmHistoricoProduccion(0);
                            this.Hide();
                            newFrm.ShowDialog();
                            this.Close();
                        }
                        else
                        {
                            this.Close();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Favor seleccionar el cliente.", "Error");
                    }
                }
                else
                {
                    MessageBox.Show("Favor seleccionar especie.", "Error");
                }
            }
            else
            {
                MessageBox.Show("Favor seleccionar un tipo de OP.", "Error");
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text.Equals(""))
                Variables.cargar(txtOutputs,"SELECT * FROM Productos WHERE Especie = '" + comboBox1.Text + "'","Codigo");
            else
                Variables.cargar(txtOutputs, "SELECT * FROM Productos WHERE Especie = '" + comboBox1.Text + "' AND Caracteristica = '" + comboBox2.Text + "'", "Codigo");
            txtOutputs.SelectedItem = null;
            textBox1.Text = "";
            textBox2.Text = "";
        }

        private void comboBox2_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (comboBox1.Text.Equals(""))
                Variables.cargar(txtOutputs, "SELECT * FROM Productos WHERE Caracteristica = '" + comboBox2.Text + "'", "Codigo");
            else
                Variables.cargar(txtOutputs, "SELECT * FROM Productos WHERE Especie = '" + comboBox1.Text + "' AND Caracteristica = '" + comboBox2.Text + "'", "Codigo");
            txtOutputs.SelectedItem = null;
            textBox1.Text = "";
            textBox2.Text = "";
        }

        private void button10_Click(object sender, EventArgs e)
        {
            txtTipo.Text = "";
            comboBox1.Text = "";
            txtOutputs.Text = "";
            textBox1.Text = "";
            textBox2.Text = "";
            txtCliente.Text = "";
            textBox6.Text = "0";
            textBox4.Text = "";
            comboBox2.Text = "";
            comboBox5.Text = "";
            txtEmpleado.Text = "";
            txtDestino.Text = "";
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        public void imprimirOP()
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos");
            Microsoft.Office.Interop.Excel.Application XcelApp = new Microsoft.Office.Interop.Excel.Application();
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos\\", "OP*");
            XcelApp.Application.Workbooks.Add(prueba[0]);
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            DateTime date1 = dateTimePicker1.Value;
            Calendar cal = dfi.Calendar;
            XcelApp.Cells[5, "C"] = fechaExpedicion;
            XcelApp.Cells[5, "I"] = OP;
            XcelApp.Cells[8, "C"] = dateTimePicker1.Text + " - " + dateTimePicker2.Text;
            XcelApp.Cells[10, "C"] = dateTimePicker3.Text;
            XcelApp.Cells[8, "I"] = cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek).ToString();
            XcelApp.Cells[10, "I"] = semanaExpedicion;
            XcelApp.Cells[14, "C"] = txtTipo.Text;
            XcelApp.Cells[14, "I"] = comboBox1.Text;
            XcelApp.Cells[17, "C"] = txtCliente.Text;
            XcelApp.Cells[17, "I"] = txtDestino.Text;
            XcelApp.Cells[22, "B"] = getFormatoSeparado(listBox2);
            int row = 23;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                XcelApp.Cells[row, "G"] = dataGridView1.Rows[i].Cells[2].Value;
                XcelApp.Cells[row, "H"] = dataGridView1.Rows[i].Cells[3].Value;
                XcelApp.Cells[row, "I"] = dataGridView1.Rows[i].Cells[4].Value;
                row++;
            }

            XcelApp.Visible = true;
        }

        public string getFormatoSeparado(ListBox lb)
        {
            string texto = "";
            for (int i = 0; i < lb.Items.Count; i++)
            {
                if (i == 0)
                    texto += lb.Items[i].ToString();
                else
                    texto += "\n" + lb.Items[i].ToString();
            }
            return texto;
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            imprimirOP();
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            int dias = (DateTime.Now - dateTimePicker1.Value).Days;
            frmResumenProcesamiento newFrm = new frmResumenProcesamiento(OP, dias, tipousuario);
            newFrm.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }
    }
}
