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
using System.Globalization;

namespace Aplicativo
{
    public partial class frmCrearDespacho : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        List<string> productos = new List<string>();
        List<string> codigo = new List<string>();
        string[,] ADF006 = new string[,] { { "Tractor", "Combustible", "Aceite Hyd - Caja", "Aceite Motor" }, { "Horas", "Gal", "Litros", "Litros" } };
        string[] Reporte = new string[] { "Mangueras", "Filtro Combustible", "Llantas", "Otro" };
        int index = 0;
        int desp = 0;
        int tipo2 = 0;
        int count = 0;
        bool sw = true;
        bool sw2 = false;

        public frmCrearDespacho(int tipo, int despacho)
        {
            InitializeComponent();
            tipo2 = tipo;
            desp = despacho;
            Variables.cargar(comboBox1, "SELECT ID, (Nombres + Apellidos) As Nombre FROM Trabajadores", "Nombre");
            //cargarEmpleados((getMaxDespacho() + 1));
            cargarADF006(dataGridView9, dataGridView7, dataGridView8, "Tractor");
            cargarADF006(dataGridView4, dataGridView6, dataGridView5, "Mini Cargador");
            Variables.cargar(comboBox3, "SELECT * FROM Clientes", "Cliente");
            Variables.cargar(comboBox4, "SELECT * FROM Proveedores", "Proveedor");
            Variables.cargar(comboBox2, "SELECT ID, (Nombres + ' ' + Apellidos) As Nombre FROM Transportadores", "Nombre");
            comboBox3.SelectedItem = null;
            comboBox4.SelectedItem = null;
            if (getInfoMaquina("Tractor") != 0)
            {
                dataGridView9.Rows[0].Cells[3].Value = getInfoMaquina("Tractor");
                dataGridView8.Rows[0].Cells[3].Value = getInfoMaquina("Tractor");
            }
            if (getInfoMaquina("Mini Cargador") != 0)
            {
                dataGridView4.Rows[0].Cells[3].Value = getInfoMaquina("Mini Cargador");
                dataGridView5.Rows[0].Cells[3].Value = getInfoMaquina("Mini Cargador");
            }
            if (tipo == 1)
            {
                Variables.cargar(dataGridView1, "SELECT Productos.ID, Productos.Codigo, ((((Productos.anchoProd * Productos.altoProd * Productos.largoProd)/1000000000) * (Productos.numAnchoEmp* Productos.numAltoEmp)) * COUNT(Paquete.ID)) FROM Productos INNER JOIN Paquete ON Productos.ID = Paquete.Producto WHERE Paquete.volumenActual > 0 AND Paquete.Porcentaje = 1 GROUP BY Productos.ID, Productos.Codigo, (((Productos.anchoProd * Productos.altoProd * Productos.largoProd)/1000000000) * (Productos.numAnchoEmp* Productos.numAltoEmp))");
                //Variables.cargar(dataGridView1, "SELECT historicoProduccion.ID, historicoProduccion.OP, Sum(Paquete.volumenActual) FROM historicoProduccion INNER JOIN Paquete ON historicoProduccion.Id = Paquete.OP GROUP BY historicoProduccion.ID, historicoProduccion.OP HAVING SUM(Paquete.volumenActual) > 0 ORDER BY historicoProduccion.ID DESC");
                dataGridView1.Columns[2].HeaderText = "Vol. Producido";
                dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                getDespachoOrden(desp);
                cargarFormatoInput(dataGridView13);
                eliminarDespachoOrden(dataGridView1);
                //agregarDespachoOrden(dataGridView1);
                dataGridView13.Columns[4].Visible = false;
            }
            else if (tipo == 0)
            {
                Variables.cargar(dataGridView1, "SELECT Pila,SUM(volumenActual) FROM reciboCliente GROUP BY Pila HAVING SUM(volumenActual) > 0 ORDER BY Pila Desc");
                dataGridView1.Columns[0].Visible = true;
                dataGridView1.Columns[1].HeaderText = "Vol. Pila";                
                dataGridView1.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                cargarFormatoInput(dataGridView13);
            }
            if (despacho != -1)
            {
                button3.Text = "Modificar Despacho";
                cargarDespacho(despacho);
                cargarSemanal(despacho);
                comboBox4.SelectedValue = getProveedor(comboBox2.SelectedValue.ToString());
                cargarADF(desp.ToString(), "ADF006", dataGridView4, "Mini Cargador");
                cargarADF006(desp.ToString(), "ADF006-2-Mini Cargador", dataGridView5);
                cargarDaños(desp.ToString(), dataGridView6, "ADF006-Mini Cargador");
                cargarADF(desp.ToString(), "ADF006", dataGridView9, "Tractor");
                cargarADF006(desp.ToString(), "ADF006-2-Tractor", dataGridView8);
                cargarDaños(desp.ToString(), dataGridView7, "ADF006-Tractor");
            }
            else
            {
                tabControl1.TabPages.RemoveAt(1);
            }

            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd/MM/yyyy";
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "HH:mm"; // Only use hours and minutes
            dateTimePicker2.ShowUpDown = true;
        }

        public void formatoIndex(DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                data.Rows[i].Cells[0].Value = i + 1;
            }
        }

        public void cargarDespacho(int despacho)
        {
            string query = "SELECT * FROM Despacho WHERE ID = " + despacho;
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
                    this.Text = "Despacho " + myReader.GetString(1);
                    comboBox3.SelectedValue = myReader.GetInt32(2);
                    comboBox2.SelectedValue = myReader.GetInt32(3);
                    textBox1.Text = myReader.GetString(4);
                    textBox2.Text = myReader.GetString(5);
                    if (myReader.GetInt32(6) != 0)
                    {
                        radioButton13.Checked = true;
                        textBox6.Text = myReader.GetInt32(6).ToString();
                    }

                    dateTimePicker1.Value = DateTime.ParseExact(myReader.GetString(7), "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                    dateTimePicker2.Value = DateTime.ParseExact(myReader.GetString(8), "HH:mm", System.Globalization.CultureInfo.InvariantCulture);
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

        public void getDespachoOrden(int despacho)
        {
            string query = "SELECT Productos.ID, Productos.Codigo FROM (despachoPaqueteOrdenes INNER JOIN Paquete ON despachoPaqueteOrdenes.Paquete = Paquete.Id) INNER JOIN Productos ON despachoPaqueteOrdenes.Producto = Productos.ID WHERE despachoPaqueteOrdenes.Despacho = " + despacho + " GROUP BY Productos.ID, Productos.Codigo";
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
                    productos.Add(myReader.GetInt32(0).ToString());
                    codigo.Add(myReader.GetString(1));
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

        public void eliminarDespachoOrden(DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                for (int j = 0; j < productos.Count; j++)
                {
                    if (Int32.Parse(data.Rows[i].Cells[0].Value.ToString()) == Int32.Parse(productos[j]))
                    {
                        productos.RemoveAt(j);
                        codigo.RemoveAt(j);
                    }
                }              
            }
        }

        public void agregarDespachoOrden(DataGridView data)
        {
            for (int i = 0; i < productos.Count; i++)
            {

                data.Rows.Add();
                data.Rows[i].Cells[0].Value = Int32.Parse(productos[i]);
                data.Rows[i].Cells[1].Value = codigo[i];
                data.Rows[i].Cells[2].Value = 0;
            }
        }

        public int getMaxDespacho()
        {
            string query = "SELECT MAX(id) FROM Despacho";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int id = 0;
            try
            {
                while (myReader.Read())
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

        public void clearRecibos(DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                if (double.Parse(data.Rows[i].Cells[5].Value.ToString()) == 0)
                {
                    data.Rows.RemoveAt(i);
                    i--;
                }
            }
        }

        public void cargarADF006(DataGridView data1, DataGridView data2, DataGridView data3, string tipo)
        {
            crearADF006(data1, "SELECT * FROM Insumos WHERE Descripcion LIKE '%", "Modelo", ADF006, tipo);
            formato(data2, Reporte);
            formato(data3, "Recorridos");
        }

        public void cargarADF(string despacho, string adf, DataGridView data, string tipo)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                if (adf.Contains("ADF006") && !adf.Equals("ADF006-2"))
                {
                    if (i == 0)
                        adf = "ADF006-1-" + tipo;
                    else
                        adf = "ADF006-" + tipo;
                }
                string query = "SELECT * FROM ControlDespacho WHERE Despacho = " + despacho + " AND Detalle = '" + data.Rows[i].Cells[2].Value.ToString() + "' AND ADF = '" + adf + "'";
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
                        if (!myReader.IsDBNull(3))
                            if (myReader.GetInt32(3) != 0)
                                data.Rows[i].Cells[3].Value = myReader.GetInt32(3);
                        data.Rows[i].Cells[1].Value = myReader.GetInt32(0);
                        for (int j = 5, k = 5; j < 11; j++, k = k + 2)
                        {
                            data.Rows[i].Cells[k].Value = myReader.GetDouble(j).ToString();
                        }
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
        }

        public void cargarADF006(string despacho, string adf, DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                string query = "SELECT * FROM ControlDespacho WHERE Despacho = " + despacho + " AND ADF = '" + adf + "'";
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
                        if (!myReader.IsDBNull(3))
                            if (myReader.GetInt32(3) != 0)
                                data.Rows[i].Cells[3].Value = myReader.GetInt32(3);
                        data.Rows[i].Cells[1].Value = myReader.GetInt32(0);
                        if (!myReader.IsDBNull(4))
                            data.Rows[i].Cells[2].Value = myReader.GetString(4);
                        for (int j = 5, k = 5; j < 11; j++, k = k + 2)
                        {
                            data.Rows[i].Cells[k].Value = myReader.GetDouble(j).ToString();
                        }
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
        }

        public void cargarDaños(string despacho, DataGridView data, string ADF)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                string query = "SELECT * FROM DañosDespacho WHERE Despacho = " + despacho + " AND Detalle = '" + data.Rows[i].Cells[1].Value.ToString() + "' AND ADF = '" + ADF + "'";
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
                        for (int j = 1; j < 7; j++)
                        {
                            if (myReader.GetInt32(j + 2) == 1)
                                data.Rows[i].Cells[j+1].Value = true;
                        }
                        data.Rows[i].Cells[8].Value = myReader.GetString(10);
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
        }


        public int getInfoMaquina(string tipo)
        {
            string query = "SELECT ID FROM Maquinarias WHERE Tipo = '" + tipo + "'";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int maquina = 0;
            try
            {
                while (myReader.Read())
                {
                    maquina = myReader.GetInt32(0);
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return maquina;
        }

        public void formato(DataGridView data, string[] formato)
        {
            for (int i = 0; i < formato.Length; i++)
            {
                data.Rows.Add();
                data.Rows[i].Cells[0].Value = i + 1;
                data.Rows[i].Cells[1].Value = formato[i];
            }
        }

        public void formato(DataGridView data, string formato)
        {
            data.Rows.Add();
            for (int j = 4; j < 16; j = j + 2)
            {
                data.Rows[0].Cells[j].Value = formato;
                data.Rows[0].Cells[j + 1].Value = "0";
            }
        }

        public void formato2(DataGridView data, string[,] formato)
        {
            for (int i = 0; i < (formato.Length) / 2; i++)
            {
                data.Rows.Add();
                data.Rows[i].Cells[0].Value = i + 1;
                data.Rows[i].Cells[2].Value = formato[0, i];
                for (int j = 4; j < 16; j = j + 2)
                {
                    data.Rows[i].Cells[j].Value = formato[1, i];
                }
                for (int j = 5; j < 17; j = j + 2)
                {
                    data.Rows[i].Cells[j].Value = "0";
                }
            }
        }

        public void crearADF006(DataGridView data, string query, string display, string[,] formatos, string tipo)
        {
            DataGridViewComboBoxColumn combo = new DataGridViewComboBoxColumn();
            combo.HeaderText = "Modelo";
            data.Columns.Add("Column1", "#");
            data.Columns.Add("Column2", "ID");
            data.Columns[1].Visible = false;
            data.Columns.Add("Column3", "Detalle");
            data.Columns.Add(combo);
            data.Columns[0].ReadOnly = true;
            data.Columns[2].ReadOnly = true;
            for (int i = 0; i < 6; i++)
            {
                data.Columns.Add("Column" + i + 4, "Unidad");
                data.Columns[4 + (i * 2)].ReadOnly = true;
                data.Columns.Add("Column" + i + 5, "Cantidad");
            }
            data.Columns.Add("Column20", "Total");
            data.Columns[0].FillWeight = 40;
            data.Columns[2].FillWeight = 200;
            data.Columns[3].FillWeight = 300;
            formato2(data, formatos);
            cargarDetalleADF006(query, display, data, formatos, tipo);
        }

        public void cargarDetalleADF006(string query, string display, DataGridView data, string[,] formato, string tipo)
        {
            DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)(data.Rows[0].Cells[3]);
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            string query2 = "SELECT ID, (Tipo + ' / ' + Marca + ' / ' + Placa) As Maquina FROM Maquinarias WHERE Tipo = '" + tipo + "'";
            OleDbCommand cmd = new OleDbCommand(query2, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            combo.DataSource = ds.Tables[0];
            combo.DisplayMember = "Maquina";
            combo.ValueMember = "ID";
            for (int i = 1; i < (formato.Length) / 2; i++)
            {
                combo = (DataGridViewComboBoxCell)(data.Rows[i].Cells[3]);
                //Ejecutar el query y llenar el ComboBox.
                conn.ConnectionString = connectionString;
                cmd = new OleDbCommand(query + formato[0, i] + "%'", conn);
                maquinaria = new DataTable();
                da = new OleDbDataAdapter(cmd);
                ds = new DataSet();

                da.Fill(ds);
                combo.DataSource = ds.Tables[0];
                combo.DisplayMember = display;
                combo.ValueMember = "ID";
            }
        }

        public void agregarEmpleado(string id)
        {
            string query = "SELECT t.ID, (t.Nombres + ' ' + t.Apellidos), t.Cedula, c.Cargo FROM CargoLaboral AS c INNER JOIN Trabajadores AS t ON c.ID = t.Cargo WHERE t.ID = " + id;
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
                    dataGridView3.Rows.Add();
                    dataGridView3.Rows[index].Cells[0].Value = index + 1;
                    dataGridView3.Rows[index].Cells[1].Value = myReader.GetInt32(0);
                    dataGridView3.Rows[index].Cells[2].Value = myReader.GetString(1);
                    dataGridView3.Rows[index].Cells[3].Value = myReader.GetInt32(2);
                    dataGridView3.Rows[index].Cells[4].Value = myReader.GetString(3);
                    dataGridView3.Rows[index].Cells[12].Value = 0;
                    index++;
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

        private void button1_Click(object sender, EventArgs e)
        {
            agregarEmpleado(comboBox1.SelectedValue.ToString());
        }

        private void dataGridView3_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            selectAll();
            Contador();
        }

        public void selectAll()
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                DataGridViewCheckBoxCell ch2 = new DataGridViewCheckBoxCell();
                ch2 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[12];
                if ((bool)ch2.FormattedValue == true && sw == false)
                {
                    for (int j = 5; j < 11; j++)
                    {
                        dataGridView3.Rows[i].Cells[j].Value = true;
                    }
                    dataGridView3.Rows[i].Cells[13].Value = 6;
                    dataGridView3.Rows[i].Cells[12].Value = false;
                    sw = true;
                }
                else
                {
                    sw = false;
                }

            }
        }

        public void Contador()
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                int total = 0;
                for (int j = 5; j < 12; j++)
                {
                    DataGridViewCheckBoxCell ch1 = new DataGridViewCheckBoxCell();
                    ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[j];
                    if ((bool)ch1.FormattedValue == true)
                        total++;
                }
                dataGridView3.Rows[i].Cells[13].Value = total;
            }
        }

        private void dataGridView3_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            Variables.dirtyCell(dataGridView3);
        }

        public void cargarEmpleados(int despacho)
        {
            while (dataGridView3.Rows.Count != 0)
            {
                dataGridView3.Rows.RemoveAt(0);
            }
            string query = "SELECT t.ID, (t.Nombres + ' ' + t.Apellidos), t.Cedula, c.Cargo FROM adf002Despacho INNER JOIN (CargoLaboral AS c INNER JOIN Trabajadores AS t ON c.ID = t.Cargo) ON adf002Despacho.Trabajador = t.ID WHERE Despacho = " + despacho;
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
                    dataGridView3.Rows.Add();
                    dataGridView3.Rows[index].Cells[0].Value = index + 1;
                    dataGridView3.Rows[index].Cells[1].Value = myReader.GetInt32(0);
                    dataGridView3.Rows[index].Cells[2].Value = myReader.GetString(1);
                    dataGridView3.Rows[index].Cells[3].Value = myReader.GetInt32(2);
                    dataGridView3.Rows[index].Cells[4].Value = myReader.GetString(3);
                    dataGridView3.Rows[index].Cells[12].Value = 0;
                    index++;
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

        public void cargarRecibos(string pila, DataGridView data)
        {
            data.Rows.Clear();
            string query = "SELECT r.Id,r.volumenActual, r.Motivo, r.Diametro, r.Largo, r.Cantidad, r.Especie, r.numRecibo, l.Lote FROM reciboCliente AS r INNER JOIN Lotes AS l ON r.Lote = l.Codigo WHERE Pila = " + pila;
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
                    data.Rows[i].Cells[0].Value = i + 1;
                    data.Rows[i].Cells[1].Value = myReader.GetInt32(0);
                    data.Rows[i].Cells[3].Value = myReader.GetString(7);
                    data.Rows[i].Cells[4].Value = myReader.GetString(8);
                    data.Rows[i].Cells[5].Value = myReader.GetDouble(1);
                    data.Rows[i].Cells[6].Value = "100%";
                    data.Rows[i].Cells[7].Value = myReader.GetDouble(1);
                    data.Rows[i].Cells[8].Value = myReader.GetString(6);
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

        public void cargarPaquetes(string op, DataGridView data)
        {
            data.Rows.Clear();
            //string query = "SELECT Paquete.Id, Paquete.volumenActual, Productos.Especie, Paquete.numPaquete FROM Paquete INNER JOIN Productos ON Paquete.Producto = Productos.ID WHERE Paquete.OP = " + op;
            string query = "SELECT Paquete.Id, Paquete.volumenActual, Productos.Especie, Paquete.numPaquete FROM Paquete INNER JOIN Productos ON Paquete.Producto = Productos.ID WHERE Paquete.Producto = " + op;
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
                    data.Rows[i].Cells[0].Value = i + 1;
                    data.Rows[i].Cells[1].Value = myReader.GetInt32(0);
                    data.Rows[i].Cells[3].Value = myReader.GetString(3);
                    data.Rows[i].Cells[4].Value = "";
                    data.Rows[i].Cells[5].Value = myReader.GetDouble(1);
                    data.Rows[i].Cells[6].Value = "100%";
                    data.Rows[i].Cells[7].Value = myReader.GetDouble(1);
                    data.Rows[i].Cells[8].Value = myReader.GetString(2);
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

        public void cargarPaquetesAsignadas(int op, DataGridView data)
        {
            string query = "SELECT ID, Paquete,Porcentaje,Volumen,volumenOriginal FROM despachoPaqueteOrdenes WHERE Despacho = " + op;
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
                    for (int j = 0; j < data.Rows.Count; j++)
                    {
                        if (data.Rows[j].Cells[1].Value.ToString().Equals(myReader.GetInt32(1).ToString()))
                        {
                            data.Rows[j].Cells[2].Value = true;
                            data.Rows[j].Cells[5].Value = myReader.GetDouble(4);
                            data.Rows[j].Cells[6].Value = myReader.GetString(2);
                            data.Rows[j].Cells[7].Value = myReader.GetDouble(3);
                        }
                    }
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

        public void cargarPilaAsignadas(int op, DataGridView data)
        {
            string query = "SELECT ID, Pila,Porcentaje,Volumen,volumenOriginal FROM despachoPilaOrdenes WHERE Despacho = " + op;
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
                    for (int j = 0; j < data.Rows.Count; j++)
                    {
                        if (data.Rows[j].Cells[1].Value.ToString().Equals(myReader.GetInt32(1).ToString()))
                        {
                            data.Rows[j].Cells[2].Value = true;
                            data.Rows[j].Cells[5].Value = myReader.GetDouble(4);
                            data.Rows[j].Cells[6].Value = myReader.GetString(2);
                            data.Rows[j].Cells[7].Value = myReader.GetDouble(3);
                        }
                    }
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


        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (tipo2 == 0)
            {
                cargarRecibos(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString(), dataGridView13);
                cargarPilaAsignadas(desp, dataGridView13);
            }
            else
            {
                cargarPaquetes(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString(), dataGridView13);
                cargarPaquetesAsignadas(desp, dataGridView13);
            }
            clearRecibos(dataGridView13);
            formatoIndex(dataGridView13);
        }

        public void cargarFormatoInput(DataGridView data)
        {
            DataGridViewCheckBoxColumn check = new DataGridViewCheckBoxColumn();
            check.HeaderText = "";
            DataGridViewComboBoxColumn combo = new DataGridViewComboBoxColumn();
            //combo.HeaderText = "Cuchilla";
            data.Columns.Add("Column1", "#");
            data.Columns.Add("Column2", "ID");
            data.Columns[1].Visible = false;
            data.Columns.Add(check);
            data.Columns.Add("Column3", "Recibo");
            data.Columns.Add("Column6", "Lote");
            data.Columns.Add("Column4", "Volumen (m3)");
            //cargarDetalle(combo, "SELECT * FROM Cuchillas", "Codigo", data);
            //data.Columns.Add(combo);
            combo = new DataGridViewComboBoxColumn();
            combo.HeaderText = "%";
            combo.Items.Add("100%");
            combo.Items.Add("75%");
            combo.Items.Add("50%");
            combo.Items.Add("25%");
            combo.Items.Add("0%");
            data.Columns.Add(combo);
            data.Columns.Add("Column5", "Volumen Seleccionado (m3)");
            data.Columns.Add("Column6", "Especie");
            data.Columns[0].ReadOnly = true;
            data.Columns[3].ReadOnly = true;
            data.Columns[4].ReadOnly = true;
            data.Columns[5].ReadOnly = true;
            data.Columns[7].ReadOnly = true;
            //combo = new DataGridViewComboBoxColumn();
            //data.Columns.Add(combo);
            data.Columns[0].FillWeight = 40;
            data.Columns[2].FillWeight = 40;
            data.Columns[3].FillWeight = 150;
        }

        private void dataGridView13_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView13.CurrentCell != null)
            {
                if (dataGridView13.CurrentCell.ColumnIndex == 6)
                {
                    if (dataGridView13.Rows[dataGridView13.CurrentCell.RowIndex].Cells[dataGridView13.CurrentCell.ColumnIndex].Value.ToString().Equals("100%"))
                        dataGridView13.Rows[dataGridView13.CurrentCell.RowIndex].Cells[7].Value = double.Parse(dataGridView13.Rows[dataGridView13.CurrentCell.RowIndex].Cells[5].Value.ToString()) * 1;
                    else if (dataGridView13.Rows[dataGridView13.CurrentCell.RowIndex].Cells[dataGridView13.CurrentCell.ColumnIndex].Value.ToString().Equals("75%"))
                        dataGridView13.Rows[dataGridView13.CurrentCell.RowIndex].Cells[7].Value = double.Parse(dataGridView13.Rows[dataGridView13.CurrentCell.RowIndex].Cells[5].Value.ToString()) * 0.75;
                    else if (dataGridView13.Rows[dataGridView13.CurrentCell.RowIndex].Cells[dataGridView13.CurrentCell.ColumnIndex].Value.ToString().Equals("50%"))
                        dataGridView13.Rows[dataGridView13.CurrentCell.RowIndex].Cells[7].Value = double.Parse(dataGridView13.Rows[dataGridView13.CurrentCell.RowIndex].Cells[5].Value.ToString()) * 0.5;
                    else if (dataGridView13.Rows[dataGridView13.CurrentCell.RowIndex].Cells[dataGridView13.CurrentCell.ColumnIndex].Value.ToString().Equals("25%"))
                        dataGridView13.Rows[dataGridView13.CurrentCell.RowIndex].Cells[7].Value = double.Parse(dataGridView13.Rows[dataGridView13.CurrentCell.RowIndex].Cells[5].Value.ToString()) * 0.25;
                    else
                        dataGridView13.Rows[dataGridView13.CurrentCell.RowIndex].Cells[7].Value = double.Parse(dataGridView13.Rows[dataGridView13.CurrentCell.RowIndex].Cells[5].Value.ToString()) * 0;
                }
            }
        }

        private void dataGridView13_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            Variables.dirtyCell(dataGridView13);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox4.Text = openFileDialog1.FileName;
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox4.SelectedItem != null && !comboBox4.SelectedValue.ToString().Equals("System.Data.DataRowView"))
            {
                comboBox2.SelectedItem = null;
                Variables.cargar(comboBox2, "SELECT ID, (Nombres + ' ' + Apellidos) As Nombre FROM Transportadores WHERE Proveedor = " + comboBox4.SelectedValue, "Nombre");
                if (comboBox2.Items.Count > 0)
                    comboBox2.Enabled = true;
                else
                    comboBox2.Enabled = false;
            }
        }

        public bool semanaExiste(int despacho)
        {
            string query = "SELECT * FROM ADF002Despacho WHERE Despacho = " + despacho;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                if (myReader.HasRows)
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

        public bool empleadoSemanaExiste(int despacho, string empleado)
        {
            string query = "SELECT * FROM ADF002Despacho WHERE Despacho = " + despacho + " AND Trabajador = " + empleado;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                if (myReader.HasRows)
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

        public void crearSemanal(int despacho)
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("INSERT INTO adf002Despacho (Trabajador,Despacho,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,Domingo,Estado,Editable) VALUES (@Trabajador,@Despacho,@Lunes,@Martes,@Miercoles,@Jueves,@Viernes,@Sabado,@Domingo,@Estado,@Editable)");
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@Trabajador", OleDbType.VarChar).Value = dataGridView3.Rows[i].Cells[1].Value.ToString();
                    cmd.Parameters.Add("@Despacho", OleDbType.VarChar).Value = despacho;
                    DataGridViewCheckBoxCell ch1 = new DataGridViewCheckBoxCell();
                    ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[5];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[6];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[7];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[8];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[9];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[10];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[11];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Domingo", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Domingo", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@Estado", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@Editable", OleDbType.VarChar).Value = 0;
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

        public void empleadoCrearSemanal(int despacho, int i)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO adf002Despacho (Trabajador,Despacho,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,Domingo,Estado,Editable) VALUES (@Trabajador,@Despacho,@Lunes,@Martes,@Miercoles,@Jueves,@Viernes,@Sabado,@Domingo,@Estado,@Editable)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Trabajador", OleDbType.VarChar).Value = dataGridView3.Rows[i].Cells[1].Value.ToString();
                cmd.Parameters.Add("@Despacho", OleDbType.VarChar).Value = despacho;
                DataGridViewCheckBoxCell ch1 = new DataGridViewCheckBoxCell();
                ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[5];
                if ((bool)ch1.FormattedValue == true)
                    cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = 1;
                else
                    cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = 0;
                ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[6];
                if ((bool)ch1.FormattedValue == true)
                    cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = 1;
                else
                    cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = 0;
                ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[7];
                if ((bool)ch1.FormattedValue == true)
                    cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = 1;
                else
                    cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = 0;
                ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[8];
                if ((bool)ch1.FormattedValue == true)
                    cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = 1;
                else
                    cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = 0;
                ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[9];
                if ((bool)ch1.FormattedValue == true)
                    cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = 1;
                else
                    cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = 0;
                ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[10];
                if ((bool)ch1.FormattedValue == true)
                    cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = 1;
                else
                    cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = 0;
                ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[11];
                if ((bool)ch1.FormattedValue == true)
                    cmd.Parameters.Add("@Domingo", OleDbType.VarChar).Value = 1;
                else
                    cmd.Parameters.Add("@Domingo", OleDbType.VarChar).Value = 0;
                cmd.Parameters.Add("@Estado", OleDbType.VarChar).Value = 0;
                cmd.Parameters.Add("@Editable", OleDbType.VarChar).Value = 0;
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

        public void empleadoModificarSemanal(int despacho, int i)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE adf002Despacho SET Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado,Domingo=@Domingo WHERE Despacho= " + despacho + " AND Trabajador = " + dataGridView3.Rows[i].Cells[1].Value);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                DataGridViewCheckBoxCell ch1 = new DataGridViewCheckBoxCell();
                ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[5];
                if ((bool)ch1.FormattedValue == true)
                    cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = 1;
                else
                    cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = 0;
                ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[6];
                if ((bool)ch1.FormattedValue == true)
                    cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = 1;
                else
                    cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = 0;
                ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[7];
                if ((bool)ch1.FormattedValue == true)
                    cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = 1;
                else
                    cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = 0;
                ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[8];
                if ((bool)ch1.FormattedValue == true)
                    cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = 1;
                else
                    cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = 0;
                ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[9];
                if ((bool)ch1.FormattedValue == true)
                    cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = 1;
                else
                    cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = 0;
                ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[10];
                if ((bool)ch1.FormattedValue == true)
                    cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = 1;
                else
                    cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = 0;
                ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[11];
                if ((bool)ch1.FormattedValue == true)
                    cmd.Parameters.Add("@Domingo", OleDbType.VarChar).Value = 1;
                else
                    cmd.Parameters.Add("@Domingo", OleDbType.VarChar).Value = 0;
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

        public void modificarSemanal(int despacho)
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (!empleadoSemanaExiste(despacho,dataGridView3.Rows[i].Cells[1].Value.ToString()))
                {
                    empleadoCrearSemanal(despacho, i);
                }
                else
                {
                    empleadoModificarSemanal(despacho, i);
                }
            }
        }

        public void cargarSemanal(int despacho)
        {
            string query = "SELECT t.ID, (t.Nombres + ' ' + t.Apellidos), t.Cedula, c.Cargo, a.Lunes, a.Martes, a.Miercoles, a.Jueves, a.Viernes, a.Sabado, a.Domingo FROM adf002Despacho AS a INNER JOIN (CargoLaboral AS c INNER JOIN Trabajadores AS t ON c.ID = t.Cargo) ON a.Trabajador = t.ID WHERE Despacho = " + despacho;
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
                    dataGridView3.Rows.Add();
                    dataGridView3.Rows[index].Cells[0].Value = index + 1;
                    dataGridView3.Rows[index].Cells[1].Value = myReader.GetInt32(0);
                    dataGridView3.Rows[index].Cells[2].Value = myReader.GetString(1);
                    dataGridView3.Rows[index].Cells[3].Value = myReader.GetInt32(2);
                    dataGridView3.Rows[index].Cells[4].Value = myReader.GetString(3);
                    dataGridView3.Rows[index].Cells[12].Value = 0;
                    for (int j = 5; j < 12; j++)
                    {
                        if (myReader.GetInt32(j - 1) == 1)
                            dataGridView3.Rows[index].Cells[j].Value = true;
                    }
                    index++;
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

        private void radioButton13_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton13.Checked)
            {
                textBox6.Visible = true;
                label7.Visible = true;
            }
            else
            {
                textBox6.Visible = false;
                label7.Visible = false;

            }
        }

        public void modificarDias(int despacho, string tipo)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Trabajadores AS t INNER JOIN adf002Despacho AS f ON t.ID = f.Trabajador SET t.diasLaborados = (t.diasLaborados " + tipo + " f.Lunes " + tipo + " f.Martes " + tipo + " f.Miercoles " + tipo + " f.Jueves " + tipo + " f.Viernes " + tipo + " f.Sabado " + tipo + " f.Domingo),f.Estado=1, f.Editable=0 WHERE f.Despacho = @Despacho");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Despacho", OleDbType.VarChar).Value = despacho;
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

        private void button3_Click(object sender, EventArgs e)
        {
            if (button3.Text.Equals("Crear Despacho"))
            {
                int id = getMaxID();
                int id2 = id;
                if (!existeAñoActual())
                    id2 = 0;
                crearDespacho(id2 + 1);
                crearSemanal(id2 + 1);
                modificarDias(id2 + 1, "+");
                if (ADFExiste(desp.ToString(), "ADF006-Mini Cargador") == false)
                {
                    crearADF((id2+1).ToString(), "ADF006", dataGridView4, "Mini Cargador");
                    crearADF((id2 + 1).ToString(), "ADF006-2", dataGridView5, "Mini Cargador");
                    crearDaños((id2 + 1).ToString(), "ADF006-Mini Cargador", dataGridView6);
                }
                if (ADFExiste((id2 + 1).ToString(), "ADF006-Tractor") == false)
                {
                    crearADF((id2 + 1).ToString(), "ADF006", dataGridView9, "Tractor");
                    crearADF((id2 + 1).ToString(), "ADF006-2", dataGridView8, "Tractor");
                    crearDaños((id2 + 1).ToString(), "ADF006-Tractor", dataGridView7);
                }
                MessageBox.Show("Despacho generado.\nFavor llenar los materiales para despachar.");
                tabControl1.TabPages.Add(tabPage3);
                desp = id2 + 1;
                button3.Text.Equals("Modificar Despacho");
            }
            else
            {
                modificarDespacho(desp);
                modificarDias(desp,"-");
                modificarSemanal(desp);
                modificarDias(desp,"+");
                modificarADF(desp.ToString(), "ADF006", dataGridView4, "Mini Cargador");
                modificarADF2(desp.ToString(), "ADF006-2-Mini Cargador", dataGridView5);
                modificarDaños(desp.ToString(), "ADF006-Mini Cargador", dataGridView6);
                modificarADF(desp.ToString(), "ADF006", dataGridView9, "Tractor");
                modificarADF2(desp.ToString(), "ADF006-2-Tractor", dataGridView8);
                modificarDaños(desp.ToString(), "ADF006-Tractor", dataGridView7);
                MessageBox.Show("Despacho modificado.");
                frmDespacho newFrm = new frmDespacho();
                this.Hide();
                newFrm.Show();
                this.Close();
            }
        }

        public void modificarADF(string despacho, string adf, DataGridView data, string tipo)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd;
                int j = 0;
                if (adf.Contains("ADF006") && !adf.Equals("ADF006-2"))
                {
                    if (i == 0)
                        adf = "ADF006-1-" + tipo;
                    else
                        adf = "ADF006-" + tipo;
                }
                if (adf.Equals("ADF006-2"))
                {
                    adf = adf + "-" + tipo;
                }
                if (data.Rows[i].Cells[3].Value != null)
                {
                    cmd = new OleDbCommand("UPDATE ControlDespacho SET Modelo=@Modelo,Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado WHERE Despacho = " + despacho + " AND Detalle = '" + data.Rows[i].Cells[2].Value.ToString() + "' AND ADF = '" + adf + "'");
                    j = 0;
                }
                else
                {
                    cmd = new OleDbCommand("UPDATE ControlDespacho SET Modelo=@Modelo,Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado WHERE Despacho = " + despacho + " AND Detalle = '" + data.Rows[i].Cells[2].Value.ToString() + "' AND ADF = '" + adf + "'");
                    j = 1;
                }
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    if (j == 0)
                        cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = data.Rows[i].Cells[3].Value.ToString();
                    else
                        cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = data.Rows[i].Cells[5].Value.ToString();
                    cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = data.Rows[i].Cells[7].Value.ToString();
                    cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = data.Rows[i].Cells[9].Value.ToString();
                    cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = data.Rows[i].Cells[11].Value.ToString();
                    cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = data.Rows[i].Cells[13].Value.ToString();
                    cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = data.Rows[i].Cells[15].Value.ToString();
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

        public int getMaxID()
        {
            string query = "SELECT MAX(ID) FROM Despacho";
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

        public int getProveedor(string transpotador)
        {
            string query = "SELECT Proveedor FROM Transportadores WHERE ID = " + transpotador;
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
        
        public bool existeAñoActual()
        {
            string query = "SELECT * FROM Despacho WHERE Despacho like '%" + DateTime.Now.Year + "%'";
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

        public void crearDespacho(int id)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Despacho (Despacho,Cliente,Transportador,numICA,numICAWeb,FSC,Fecha,Hora,Tipo) VALUES (@Despacho,@Cliente,@Transportador,@numICA,@numICAWeb,@FSC,@Fecha,@Hora,@Tipo)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Despacho", OleDbType.VarChar).Value = "D-" + id.ToString().PadLeft(4, '0') + "-" + DateTime.Now.Year;
                cmd.Parameters.Add("@Cliente", OleDbType.VarChar).Value = comboBox3.SelectedValue;
                cmd.Parameters.Add("@Transportador", OleDbType.VarChar).Value = comboBox2.SelectedValue;
                cmd.Parameters.Add("@numICA", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@numICAWeb", OleDbType.VarChar).Value = textBox2.Text;
                if (radioButton13.Checked)
                    cmd.Parameters.Add("@FSC", OleDbType.VarChar).Value = textBox6.Text;
                else
                    cmd.Parameters.Add("@FSC", OleDbType.VarChar).Value = 0;
                cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = dateTimePicker1.Value.ToString("dd") + "/" + dateTimePicker1.Value.ToString("MM") + "/" + dateTimePicker1.Value.Year;
                cmd.Parameters.Add("@Hora", OleDbType.VarChar).Value = dateTimePicker2.Value.ToString("HH") + ":" + dateTimePicker2.Value.ToString("mm");
                if(tipo2==0)
                    cmd.Parameters.Add("@Tipo", OleDbType.VarChar).Value = "Madera Rolliza";
                else
                    cmd.Parameters.Add("@Tipo", OleDbType.VarChar).Value = "Otro";
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

        public void modificarDespacho(int id)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Despacho SET Cliente=@Cliente,Transportador=@Transportador,numICA=@numICA,numICAWeb=@numICAWeb,FSC=@FSC,Fecha=@Fecha,Hora=@Hora WHERE ID = " + id);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Cliente", OleDbType.VarChar).Value = comboBox3.SelectedValue;
                cmd.Parameters.Add("@Transportador", OleDbType.VarChar).Value = comboBox2.SelectedValue;
                cmd.Parameters.Add("@numICA", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@numICAWeb", OleDbType.VarChar).Value = textBox2.Text;
                if (radioButton13.Checked)
                    cmd.Parameters.Add("@FSC", OleDbType.VarChar).Value = textBox6.Text;
                else
                    cmd.Parameters.Add("@FSC", OleDbType.VarChar).Value = 0;
                cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = dateTimePicker1.Value.ToString("dd") + "/" + dateTimePicker1.Value.ToString("MM") + "/" + dateTimePicker1.Value.Year;
                cmd.Parameters.Add("@Hora", OleDbType.VarChar).Value = dateTimePicker2.Value.ToString("HH") + ":" + dateTimePicker2.Value.ToString("mm");
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

        public bool ADFExiste(string despacho, string adf)
        {
            string query = "SELECT * FROM ControlDespacho WHERE Despacho = " + despacho + " AND adf = '" + adf + "'";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                if (myReader.HasRows)
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

        public void crearADF(string orden, string adf, DataGridView data, string tipo)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                conn.ConnectionString = connectionString;
                int j = 0;
                OleDbCommand cmd;
                if (data.Rows[i].Cells[3].Value != null)
                {
                    cmd = new OleDbCommand("INSERT INTO ControlDespacho(Unidad,Despacho,Detalle,Modelo,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,Estado,Editable,ADF) VALUES (@Unidad,@Despacho,@Detalle,@Modelo,@Lunes,@Martes,@Miercoles,@Jueves,@Viernes,@Sabado,@Estado,@Editable,@ADF)");
                    j = 0;
                }
                else
                {
                    cmd = new OleDbCommand("INSERT INTO ControlDespacho(Unidad,Despacho,Detalle,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,Estado,Editable,ADF) VALUES (@Unidad,@Despacho,@Detalle,@Lunes,@Martes,@Miercoles,@Jueves,@Viernes,@Sabado,@Estado,@Editable,@ADF)");
                    j = 1;
                }
                if (adf.Contains("ADF006") && !adf.Equals("ADF006-2"))
                {
                    if (i == 0)
                        adf = "ADF006-1-" + tipo;
                    else
                        adf = "ADF006-" + tipo;
                }
                if (adf.Equals("ADF006-2"))
                {
                    adf = adf + "-" + tipo;
                }
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@Unidad", OleDbType.VarChar).Value = data.Rows[i].Cells[4].Value.ToString();
                    cmd.Parameters.Add("@Despacho", OleDbType.VarChar).Value = orden;
                    if (data.Rows[i].Cells[2].Value != null)
                        cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = data.Rows[i].Cells[2].Value.ToString();
                    else
                        cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = "";
                    if (j == 0)
                        cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = data.Rows[i].Cells[3].Value.ToString();
                    cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = data.Rows[i].Cells[5].Value.ToString();
                    cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = data.Rows[i].Cells[7].Value.ToString();
                    cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = data.Rows[i].Cells[9].Value.ToString();
                    cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = data.Rows[i].Cells[11].Value.ToString();
                    cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = data.Rows[i].Cells[13].Value.ToString();
                    cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = data.Rows[i].Cells[15].Value.ToString();
                    cmd.Parameters.Add("@Estado", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@Editable", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@ADF", OleDbType.VarChar).Value = adf;
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

        public void crearDaños(string orden, string ADF, DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("INSERT INTO DañosDespacho(Despacho,Detalle,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,ADF,Descripcion) VALUES (@Despacho,@Detalle,@Lunes,@Martes,@Miercoles,@Jueves,@Viernes,@Sabado,@ADF,@Descripcion)");
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@Despacho", OleDbType.VarChar).Value = orden;
                    cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = data.Rows[i].Cells[1].Value.ToString();
                    DataGridViewCheckBoxCell ch1 = new DataGridViewCheckBoxCell();
                    ch1 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[2];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[3];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[4];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[5];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[6];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[7];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@ADF", OleDbType.VarChar).Value = ADF;
                    if (data.Rows[i].Cells[8].Value != null)
                        cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = data.Rows[i].Cells[8].Value.ToString();
                    else
                        cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = "";
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

        public void modificarADF2(string despacho, string adf, DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd;
                int j = 0;
                if (data.Rows[i].Cells[3].Value != null)
                {
                    if (data.Rows[i].Cells[1].Value != null)
                    {
                        cmd = new OleDbCommand("UPDATE ControlDespacho SET Detalle=@Detalle,Modelo=@Modelo,Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado WHERE ID = " + data.Rows[i].Cells[1].Value);
                    }
                    else
                    {
                        cmd = new OleDbCommand("UPDATE ControlDespacho SET Detalle=@Detalle,Modelo=@Modelo,Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado WHERE Despacho = " + despacho + " AND ADF = '" + adf + "'");
                    }
                    j = 0;
                }
                else
                {
                    if (data.Rows[i].Cells[1].Value != null)
                    {
                        cmd = new OleDbCommand("UPDATE ControlDespacho SET Detalle=@Detalle,Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado  WHERE ID = " + data.Rows[i].Cells[1].Value);
                    }
                    else
                    {
                        cmd = new OleDbCommand("UPDATE ControlDespacho SET Detalle=@Detalle,Modelo=@Modelo,Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado WHERE Despacho = " + despacho + " AND ADF = '" + adf + "'");
                    }
                    j = 1;
                }
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    if (data.Rows[i].Cells[2].Value != null)
                        cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = data.Rows[i].Cells[2].Value.ToString();
                    else
                        cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = "";
                    if (j == 0)
                        cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = data.Rows[i].Cells[3].Value.ToString();
                    else
                        cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = data.Rows[i].Cells[5].Value.ToString();
                    cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = data.Rows[i].Cells[7].Value.ToString();
                    cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = data.Rows[i].Cells[9].Value.ToString();
                    cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = data.Rows[i].Cells[11].Value.ToString();
                    cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = data.Rows[i].Cells[13].Value.ToString();
                    cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = data.Rows[i].Cells[15].Value.ToString();
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

        public void modificarDaños(string despacho, string ADF, DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("UPDATE DañosDespacho SET Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado,Descripcion=@Descripcion WHERE ADF = '" + ADF + "' AND Detalle = '" + data.Rows[i].Cells[1].Value.ToString() + "' AND Despacho = " + despacho);
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    DataGridViewCheckBoxCell ch1 = new DataGridViewCheckBoxCell();
                    ch1 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[2];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[3];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[4];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[5];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[6];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[7];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[8].Value != null)
                        cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = data.Rows[i].Cells[8].Value.ToString();
                    else
                        cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = "";
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

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (comboBox2.SelectedItem != null && !comboBox2.SelectedValue.ToString().Equals("System.Data.DataRowView"))
            //{
            //    comboBox4.SelectedValue = getProveedor(comboBox2.SelectedValue.ToString());
            //}
        }

        public void modificarPaquete(DataGridView data, int i, string tipo)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand();
            cmd = new OleDbCommand("UPDATE Paquete SET volumenActual = volumenActual " + tipo + " @volumen WHERE ID =  " + data.Rows[i].Cells[1].Value);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                if (tipo.Equals("+"))
                    cmd.Parameters.Add("@volumen", OleDbType.VarChar).Value = data.Rows[i].Cells[5].Value;
                else
                    cmd.Parameters.Add("@volumen", OleDbType.VarChar).Value = data.Rows[i].Cells[7].Value;
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

        public void eliminarAsignacionPaquete(string producto)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("DELETE FROM despachoPaqueteOrdenes WHERE Despacho = " + desp + " AND Producto = " + producto);
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

        public void eliminarAsignacionPaquete()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("DELETE FROM despachoPaqueteOrdenes WHERE Despacho = " + desp);
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


        public void agregarAsignacionPaquete(int i, DataGridView data, string producto)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO despachoPaqueteOrdenes (Paquete,Despacho,Porcentaje,Producto,volumen,volumenOriginal) VALUES (@Paquete,@Despacho,@Porcentaje,@Producto,@volumen,@volumenOriginal)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Paquete", OleDbType.VarChar).Value = data.Rows[i].Cells[1].Value;
                cmd.Parameters.Add("@Despacho", OleDbType.VarChar).Value = desp;
                cmd.Parameters.Add("@Porcentaje", OleDbType.VarChar).Value = data.Rows[i].Cells[6].Value;
                cmd.Parameters.Add("@Producto", OleDbType.VarChar).Value = producto;
                cmd.Parameters.Add("@volumen", OleDbType.VarChar).Value = data.Rows[i].Cells[7].Value;
                cmd.Parameters.Add("@volumenOriginal", OleDbType.VarChar).Value = data.Rows[i].Cells[5].Value;
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

        public void modificarPila(DataGridView data, int i, string tipo)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand();
            cmd = new OleDbCommand("UPDATE reciboCliente SET volumenActual = volumenActual " + tipo + " @volumen WHERE ID =  " + data.Rows[i].Cells[1].Value);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                if (tipo.Equals("+"))
                    cmd.Parameters.Add("@volumen", OleDbType.VarChar).Value = data.Rows[i].Cells[5].Value;
                else
                    cmd.Parameters.Add("@volumen", OleDbType.VarChar).Value = data.Rows[i].Cells[7].Value;
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

        public void eliminarAsignacionPila(string pila)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("DELETE FROM despachoPilaOrdenes WHERE Despacho = " + desp + " AND numPila = " + pila);
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

        public void eliminarAsignacionPila()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("DELETE FROM despachoPilaOrdenes WHERE Despacho = " + desp);
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

        public void agregarAsignacionPila(int i, DataGridView data, string pila)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO despachoPilaOrdenes (Pila,Despacho,Porcentaje,numPila,volumen,volumenOriginal) VALUES (@Pila,@Despacho,@Porcentaje,@numPila,@volumen,@volumenOriginal)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Pila", OleDbType.VarChar).Value = data.Rows[i].Cells[1].Value;
                cmd.Parameters.Add("@Despacho", OleDbType.VarChar).Value = desp;
                cmd.Parameters.Add("@Porcentaje", OleDbType.VarChar).Value = data.Rows[i].Cells[6].Value;
                cmd.Parameters.Add("@numPila", OleDbType.VarChar).Value = pila;
                cmd.Parameters.Add("@volumen", OleDbType.VarChar).Value = data.Rows[i].Cells[7].Value;
                cmd.Parameters.Add("@volumenOriginal", OleDbType.VarChar).Value = data.Rows[i].Cells[5].Value;
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



        private void button2_Click(object sender, EventArgs e)
        {
            if (desp == -1)
            {
                desp = getMaxDespacho();
                sw2 = true;
            }

            if (tipo2 == 1)
            {
                if (Variables.existe("SELECT * FROM despachoPaqueteOrdenes WHERE Despacho = " + desp))
                {
                    for (int i = 0; i < dataGridView13.Rows.Count; i++)
                    {
                        if (Variables.existe("SELECT * FROM despachoPaqueteOrdenes WHERE Despacho = " + desp + " AND Paquete = " + Int32.Parse(dataGridView13.Rows[i].Cells[1].Value.ToString())))
                            modificarPaquete(dataGridView13, i, "+");
                    }
                    eliminarAsignacionPaquete(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
                }
                for (int i = 0; i < dataGridView13.Rows.Count; i++)
                {
                    if (Variables.existe("SELECT * FROM despachoPaqueteOrdenes WHERE Despacho = " + desp + " AND Paquete = " + Int32.Parse(dataGridView13.Rows[i].Cells[1].Value.ToString())))
                        modificarPaquete(dataGridView13, i, "+");
                    if (Convert.ToBoolean(dataGridView13.Rows[i].Cells[2].Value) == true)
                    {
                        agregarAsignacionPaquete(i, dataGridView13, dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
                        modificarPaquete(dataGridView13, i, "-");
                    }
                }
                MessageBox.Show("Pilas asignadas.");
            }
            else
            {
                if (Variables.existe("SELECT * FROM despachoPilaOrdenes WHERE Despacho = " + desp))
                {
                    for (int i = 0; i < dataGridView13.Rows.Count; i++)
                    {
                        if (Variables.existe("SELECT * FROM despachoPilaOrdenes WHERE Despacho = " + desp + " AND Pila = " + Int32.Parse(dataGridView13.Rows[i].Cells[1].Value.ToString())))
                            modificarPila(dataGridView13, i, "+");
                    }
                    eliminarAsignacionPila(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
                }
                for (int i = 0; i < dataGridView13.Rows.Count; i++)
                {
                    if (Variables.existe("SELECT * FROM despachoPilaOrdenes WHERE Despacho = " + desp + " AND Pila = " + Int32.Parse(dataGridView13.Rows[i].Cells[1].Value.ToString())))
                        modificarPila(dataGridView13, i, "+");
                    if (Convert.ToBoolean(dataGridView13.Rows[i].Cells[2].Value) == true)
                    {
                        agregarAsignacionPila(i, dataGridView13, dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
                        modificarPila(dataGridView13, i, "-");
                    }
                }
                MessageBox.Show("Paquetes asignados.");
            }
        }

        private void frmCrearDespacho_FormClosing(object sender, FormClosingEventArgs e)
        {
            //if (sw2)
            //{
            //    if (tipo2 == 1)
            //    {
            //        eliminarAsignacionPaquete();
            //    }
            //    else
            //    {
            //        eliminarAsignacionPila();
            //    }
            //}
        }

    }
}
