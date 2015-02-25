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

namespace Aplicativo
{
    public partial class frmProcesamientoFormatos : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();

        string[,] ADF006 = new string[,] { { "Tractor", "Combustible", "Aceite Hyd - Caja", "Aceite Motor" }, { "Horas", "Gal", "Litros", "Litros" } };
        string[,] ADF020 = new string[,] { { "TVS", "SVS", "HR 1000 (3)", "HR 1000 (4)", "SHR", "EDGER"}, { "Horas", "Horas", "Horas", "Horas", "Horas", "Horas"} };
        string[,] ADF0201 = new string[,] { { "Falta de Electricidad" }, { "Horas"} };
        string[] Reporte = new string[] { "Mangueras", "Filtro Combustible", "Llantas", "Otro" };
        int orden = 0, diaOrden = 0, tipo2 = 0, diaPaquete = 0, diaRecibos = 0, diaPaquete2 = 0, diaRecibos2 = 0, semana = 0, ano = DateTime.Now.Year;        

        public frmProcesamientoFormatos(int op, int dia, int tipo)
        {
            InitializeComponent();
            this.Text = "Formatos de la Orden #" + getNombreOP(op) + " Día #: " + (dia + 1) + " Fecha Actual: " + DateTime.Now.ToString("dd") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.Year;
            label58.Text = "Día #: " + (dia + 1) + "                     Fecha Actual: " + DateTime.Now.ToString("dd") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.Year;
            //cargarEmpleados(op);
            diaOrden = dia;
            orden = op;
            tipo2 = tipo;
            formatoADF020(dataGridView28, ADF020);
            formatoADF020(dataGridView1, ADF020);
            formatoADF020(dataGridView2, ADF0201);
            cargarADF006(dataGridView9, dataGridView7, dataGridView8,"Monta Carga");
            cargarADF006(dataGridView4, dataGridView6, dataGridView5,"Mini Cargador");
            cargarADF006(dataGridView12, dataGridView10, dataGridView11,"Chipeadora");
            if(getInfoMaquina("Monta Carga") != 0)
                dataGridView9.Rows[0].Cells[3].Value = getInfoMaquina("Monta Carga");
            if (getInfoMaquina("Mini Cargador") != 0)
                dataGridView4.Rows[0].Cells[3].Value = getInfoMaquina("Mini Cargador");
            if (getInfoMaquina("Chipeadora") != 0)
                dataGridView12.Rows[0].Cells[3].Value = getInfoMaquina("Chipeadora");
            cargarFormatoInput(dataGridView13);
            cargarRecibos(getEspecie(op), dataGridView13);
            cargarRecibosAsignadas(op,dataGridView13,diaOrden);
            clearRecibos(dataGridView13);
            cargarFormatoOutput(dataGridView14);
            cargarProductos(op, dataGridView14);
            cargarPaquetesTotales(dataGridView14);
            cargarPaquetesTotalesDiarios(dataGridView14,dia);
            //DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            //DateTime date1 = DateTime.ParseExact(DateTime.Now.ToString("dd") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.Year, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
            //Calendar cal = dfi.Calendar;
            //semana = cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);
            //cargarADF(semana, orden.ToString(), "ADF006", dataGridView4, ano, "Mini Cargador");
            //cargarADF006(semana, orden.ToString(), "ADF006-2-Mini Cargador", dataGridView5, ano);
            //cargarDaños(semana, orden.ToString(), dataGridView6, "ADF006-Mini Cargador", ano);
            //cargarADF(semana, orden.ToString(), "ADF006", dataGridView9, ano, "Monta Carga");
            //cargarADF006(semana, orden.ToString(), "ADF006-2-Monta Carga", dataGridView8, ano);
            //cargarDaños(semana, orden.ToString(), dataGridView7, "ADF006-Monta Carga", ano);
            //cargarADF(semana, orden.ToString(), "ADF006", dataGridView12, ano, "Chipeadora");
            //cargarADF006(semana, orden.ToString(), "ADF006-2-Chipeadora", dataGridView11, ano);
            //cargarDaños(semana, orden.ToString(), dataGridView10, "ADF006-Chipeadora", ano);
            dataGridView14.Columns[2].DefaultCellStyle.Font = new Font(dataGridView14.DefaultCellStyle.Font, FontStyle.Underline);

        }
        
        public void cargarADF(int semana, string orden, string adf, DataGridView data, int ano, string tipo)
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
                string query = "SELECT * FROM ControlOP WHERE Semana = " + semana + " AND Orden = " + orden + " AND Detalle = '" + data.Rows[i].Cells[2].Value.ToString() + "' AND ADF = '" + adf + "' AND Ano = " + ano;
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
                        if (!myReader.IsDBNull(4))
                            if (myReader.GetInt32(4) != 0)
                                data.Rows[i].Cells[3].Value = myReader.GetInt32(4);
                        data.Rows[i].Cells[1].Value = myReader.GetInt32(0);
                        for (int j = 6, k = 5; j < 12; j++, k = k + 2)
                        {
                            data.Rows[i].Cells[k].Value = myReader.GetInt32(j).ToString();
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
            data.Columns[0].ReadOnly = true;
            data.Columns[3].ReadOnly = true;
            data.Columns[4].ReadOnly = true;
            data.Columns[5].ReadOnly = true;
            data.Columns[7].ReadOnly = true;
            //combo = new DataGridViewComboBoxColumn();
            //data.Columns.Add(combo);
            data.Columns[0].FillWeight = 40;
            data.Columns[2].FillWeight = 40;            
        }

        public void cargarFormatoOutput(DataGridView data)
        {
            //combo.HeaderText = "Cuchilla";
            data.Columns.Add("Column1", "#");
            data.Columns.Add("Column2", "ID");
            data.Columns[1].Visible = false;
            data.Columns.Add("Column3", "Producto");
            data.Columns.Add("Column4", "# Paquetes Requeridos");
            data.Columns.Add("Column5", "# Paquetes Producidos");
            data.Columns.Add("Column6", "# Paquetes Producidos Hoy");
            data.Columns.Add("Column6", "Vol (m3) por Paquete");
            data.Columns.Add("Column6", "Vol (m3) Producidos Hoy");
            //combo = new DataGridViewComboBoxColumn();
            //data.Columns.Add(combo);
            data.Columns[0].FillWeight = 40;
        }

        public void cargarRecibos(string especie, DataGridView data)
        {
            string query = "SELECT r.Id,r.volumenActual, r.Motivo, r.Diametro, r.Largo, r.Cantidad, r.Modulo, r.numRecibo, l.Lote FROM Recibo AS r INNER JOIN Lotes AS l ON r.Lote = l.Codigo WHERE r.Especie = '" + especie + "' AND ((Month(r.Fecha)) = " + (DateTime.Now.Month - 3) + " OR Month(r.Fecha) = " + DateTime.Now.Month + ")";
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

        public void cargarRecibosAsignadas(int op, DataGridView data, int dia)
        {
            string query = "SELECT ID, Recibo,Porcentaje,Volumen,volumenOriginal FROM recibosOrdenes WHERE OP = " + op + " AND Dia = " + dia;
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

        public void cargarPaquetesTotales(DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                string query = "SELECT SUM(porcentaje) FROM Paquete WHERE OP = " + orden + " AND Producto = " + data.Rows[i].Cells[1].Value;
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
                        data.Rows[i].Cells[4].Value = myReader.GetValue(0);
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

        public void cargarPaquetesTotalesDiarios(DataGridView data, int dia)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                string query = "SELECT SUM(porcentaje) FROM Paquete WHERE OP = " + orden + " AND Producto = " + data.Rows[i].Cells[1].Value + " AND dia = " + dia;
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
                        if (!myReader.GetValue(0).ToString().Equals(""))
                        {
                            data.Rows[i].Cells[5].Value = myReader.GetValue(0);
                            data.Rows[i].Cells[7].Value = double.Parse(data.Rows[i].Cells[6].Value.ToString()) * double.Parse(myReader.GetValue(0).ToString());
                        }
                        else
                        {
                            data.Rows[i].Cells[5].Value = "0";
                            data.Rows[i].Cells[7].Value = "0";
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

        public void cargarProductos(int op, DataGridView data)
        {
            string query = "SELECT op.Id, op.Producto, p.Codigo, op.Cantidad, op.Volumen FROM Productos AS p INNER JOIN produccionProducto AS op ON p.ID = op.Producto WHERE Orden = " + op;
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
                    data.Rows[i].Cells[1].Value = myReader.GetInt32(1);
                    data.Rows[i].Cells[2].Value = myReader.GetString(2);
                    data.Rows[i].Cells[3].Value = myReader.GetDouble(3);
                    data.Rows[i].Cells[4].Value = 0;
                    data.Rows[i].Cells[5].Value = 0;
                    data.Rows[i].Cells[6].Value = (myReader.GetDouble(4)/myReader.GetDouble(3));
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

        public void formato(DataGridView data, string formato)
        {
            data.Rows.Add();
            for (int j = 4; j < 16; j = j + 2)
            {
                data.Rows[0].Cells[j].Value = formato;
                data.Rows[0].Cells[j + 1].Value = "0";
            }
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

        public void cargarADF006(DataGridView data1, DataGridView data2, DataGridView data3, string tipo)
        {
            crearADF006(data1, "SELECT * FROM Insumos WHERE Descripcion LIKE '%", "Modelo", ADF006, tipo);
            formato(data2, Reporte);
            formato(data3, "Recorridos");
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

        public void formatoADF020(DataGridView data, string[,] formatos)
        {
            data.Columns.Add("Column1", "#");
            data.Columns.Add("Column2", "ID");
            data.Columns[1].Visible = false;    
            data.Columns.Add("Column3", "Detalle");
            data.Columns[0].ReadOnly = true;
            data.Columns[2].ReadOnly = true;
            for (int i = 0; i < 6; i++)
            {
                data.Columns.Add("Column" + i + 3, "Unidad");
                data.Columns[3 + (i * 2)].ReadOnly = true;
                data.Columns.Add("Column" + i + 4, "Cantidad");
            }
            data.Columns.Add("Column20", "Total");
            data.Columns[0].FillWeight = 40;
            data.Columns[2].FillWeight = 200;
            formato(data, formatos);
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

        public void formato(DataGridView data, string[,] formato)
        {
            for (int i = 0; i < (formato.Length) / 2; i++)
            {
                data.Rows.Add();
                data.Rows[i].Cells[0].Value = i + 1;
                data.Rows[i].Cells[2].Value = formato[0, i];
                for (int j = 3; j < 15; j = j + 2)
                {
                    data.Rows[i].Cells[j].Value = formato[1, i];
                }
                for (int j = 4; j < 16; j = j + 2)
                {
                    data.Rows[i].Cells[j].Value = "0";
                }
            }
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

        public string getEspecie(int op)
        {
            string query = "SELECT Especie FROM historicoProduccion WHERE id = " + op;
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

        //public void cargarEmpleados(int orden)
        //{
        //    while (dataGridView3.Rows.Count != 0)
        //    {
        //        dataGridView3.Rows.RemoveAt(0);
        //    }
        //    string query = "SELECT t.ID, (t.Nombres + ' ' + t.Apellidos), t.Cedula, c.Cargo FROM CargoLaboral AS c INNER JOIN (Trabajadores AS t INNER JOIN produccionEmpleados AS s ON t.ID = s.Trabajador) ON c.ID = t.Cargo WHERE s.Orden = " + orden;
        //    //Ejecutar el query y llenar el GridView.
        //    conn.ConnectionString = connectionString;
        //    OleDbCommand cmd = new OleDbCommand(query, conn);
        //    cmd.Connection = conn;
        //    conn.Open();
        //    OleDbDataReader myReader = cmd.ExecuteReader();
        //    int i = 0;
        //    try
        //    {
        //        while (myReader.Read())
        //        {
        //            dataGridView3.Rows.Add();
        //            dataGridView3.Rows[i].Cells[0].Value = i + 1;
        //            dataGridView3.Rows[i].Cells[1].Value = myReader.GetInt32(0);
        //            dataGridView3.Rows[i].Cells[2].Value = myReader.GetString(1);
        //            dataGridView3.Rows[i].Cells[3].Value = myReader.GetInt32(2);
        //            dataGridView3.Rows[i].Cells[4].Value = myReader.GetString(3);
        //            dataGridView3.Rows[i].Cells[12].Value = 0;
        //            i++;
        //        }
        //    }
        //    finally
        //    {
        //        // always call Close when done reading.
        //        myReader.Close();
        //        // always call Close when done reading.
        //        conn.Close();
        //    }
        //}

        private void dataGridView28_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Contador(dataGridView28, 4, 16, ADF020);
        }

        public void Contador(DataGridView data, int inicio, int final, string[,] adf)
        {
            if (data.Rows.Count == (adf.Length) / 2)
            {
                for (int i = 0; i < data.Rows.Count; i++)
                {
                    int total = 0;
                    for (int j = inicio; j < final; j = j + 2)
                    {
                        int num;
                        if (!(data.Rows[i].Cells[j].Value == null))
                        {
                            bool isNum = Int32.TryParse(data.Rows[i].Cells[j].Value.ToString(), out num);
                            if (isNum)
                            {
                                total = total + Int32.Parse(data.Rows[i].Cells[j].Value.ToString());
                            }
                        }
                    }
                    data.Rows[i].Cells[data.Columns.Count - 1].Value = total;
                }
            }
        }

        private void dataGridView28_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            Variables.dirtyCell(dataGridView28);
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Contador(dataGridView1, 4, 16, ADF020);
        }

        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            Variables.dirtyCell(dataGridView1);
        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Contador(dataGridView2, 4, 16, ADF0201);
        }

        private void dataGridView2_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            Variables.dirtyCell(dataGridView2);
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

        private void button8_Click(object sender, EventArgs e)
        {
            frmRegistrarPaquete newFrm = new frmRegistrarPaquete(orden,diaPaquete);
            newFrm.Show();
        }

        private void dataGridView4_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Contador(dataGridView4, 5, 17, ADF006);
        }

        private void dataGridView4_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            Variables.dirtyCell(dataGridView4);
        }

        private void dataGridView5_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Contador(dataGridView5, 5, 17);
        }

        public void Contador(DataGridView data, int inicio, int final)
        {
            if (data.Rows.Count != 0)
            {
                for (int i = 0; i < data.Rows.Count; i++)
                {
                    int total = 0;
                    for (int j = inicio; j < final; j = j + 2)
                    {
                        int num;

                        if (!(data.Rows[i].Cells[j].Value == null))
                        {
                            bool isNum = Int32.TryParse(data.Rows[i].Cells[j].Value.ToString(), out num);
                            if (isNum)
                            {
                                total = total + Int32.Parse(data.Rows[i].Cells[j].Value.ToString());
                            }
                        }
                    }
                    data.Rows[i].Cells[data.Columns.Count - 1].Value = total;
                }
            }
        }

        private void dataGridView5_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            Variables.dirtyCell(dataGridView5);
        }

        private void dataGridView6_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void dataGridView9_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Contador(dataGridView9, 5, 7, ADF006);
        }

        private void dataGridView9_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            Variables.dirtyCell(dataGridView9);
        }

        private void dataGridView8_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Contador(dataGridView8, 5, 7);
        }

        private void dataGridView8_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            Variables.dirtyCell(dataGridView8);
        }

        private void dataGridView12_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Contador(dataGridView12, 5, 7, ADF006);
        }

        private void dataGridView12_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            Variables.dirtyCell(dataGridView12);
        }

        private void dataGridView11_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Contador(dataGridView11, 5, 7);
        }

        private void dataGridView11_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            Variables.dirtyCell(dataGridView11);
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmProcesamientoFormatos newFrm = new frmProcesamientoFormatos(orden,diaOrden-1,tipo2);
            this.Hide();
            newFrm.ShowDialog();
            this.Close();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmProcesamientoFormatos newFrm = new frmProcesamientoFormatos(orden, diaOrden + 1, tipo2);
            this.Hide();
            newFrm.ShowDialog();
            this.Close();
        }

        private void dataGridView14_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView14.CurrentCell.ColumnIndex == 2)
            {
                frmPaqueteOP newFrm = new frmPaqueteOP(orden,Int32.Parse(dataGridView14.Rows[dataGridView14.CurrentCell.RowIndex].Cells[1].Value.ToString()));
                newFrm.Show();
            }
        }

        public bool ADFExiste(int semana, string orden, string adf, int ano)
        {
            string query = "SELECT * FROM ControlOP WHERE Semana = " + semana + " AND Orden = " + orden + " AND adf = '" + adf + "' AND Ano = " + ano;
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

        public void modificarADF(int semana, string orden, string adf, DataGridView data, int ano, string tipo)
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
                    cmd = new OleDbCommand("UPDATE ControlOP SET Modelo=@Modelo,Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado WHERE Semana = " + semana + " AND Orden = " + orden + " AND Detalle = '" + data.Rows[i].Cells[2].Value.ToString() + "' AND ADF = '" + adf + "' AND Ano = " + ano);
                    j = 0;
                }
                else
                {
                    cmd = new OleDbCommand("UPDATE ControlOP SET Modelo=@Modelo,Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado WHERE Semana = " + semana + " AND Orden = " + orden + " AND Detalle = '" + data.Rows[i].Cells[2].Value.ToString() + "' AND ADF = '" + adf + "' AND Ano = " + ano);
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

        public void crearADF(int semana, string orden, string adf, DataGridView data, int ano, string tipo)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                conn.ConnectionString = connectionString;
                int j = 0;
                OleDbCommand cmd;
                if (data.Rows[i].Cells[3].Value != null)
                {
                    cmd = new OleDbCommand("INSERT INTO ControlOP(Semana,Unidad,Orden,Detalle,Modelo,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,Estado,Editable,ADF,Ano) VALUES (@Semana,@Unidad,@Orden,@Detalle,@Modelo,@Lunes,@Martes,@Miercoles,@Jueves,@Viernes,@Sabado,@Estado,@Editable,@ADF,@Ano)");
                    j = 0;
                }
                else
                {
                    cmd = new OleDbCommand("INSERT INTO ControlOP(Semana,Unidad,Orden,Detalle,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,Estado,Editable,ADF,Ano) VALUES (@Semana,@Unidad,@Orden,@Detalle,@Lunes,@Martes,@Miercoles,@Jueves,@Viernes,@Sabado,@Estado,@Editable,@ADF,@Ano)");
                    j = 1;
                }
                if (adf.Contains("ADF006") && !adf.Equals("ADF006-2"))
                {
                    if (i == 0)
                        adf = "ADF006-1-" + tipo;
                    else
                        adf = "ADF006-" + tipo ;
                }
                if (adf.Equals("ADF006-2"))
                {
                    adf = adf + "-" + tipo;
                }
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@Semana", OleDbType.VarChar).Value = semana;
                    cmd.Parameters.Add("@Unidad", OleDbType.VarChar).Value = data.Rows[i].Cells[4].Value.ToString();
                    cmd.Parameters.Add("@Orden", OleDbType.VarChar).Value = orden;
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
                    cmd.Parameters.Add("@Ano", OleDbType.VarChar).Value = ano;
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

        public void crearDaños(int semana, string orden, string ADF, DataGridView data, int ano)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("INSERT INTO DañosOP(Semana,Orden,Detalle,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,ADF,Descripcion,Ano) VALUES (@Semana,@Orden,@Detalle,@Lunes,@Martes,@Miercoles,@Jueves,@Viernes,@Sabado,@ADF,@Descripcion,@Ano)");
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@Semana", OleDbType.VarChar).Value = semana;
                    cmd.Parameters.Add("@Orden", OleDbType.VarChar).Value = orden;
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
                    cmd.Parameters.Add("@Ano", OleDbType.VarChar).Value = ano;
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

        public void modificarADF2(int semana, string orden, string adf, DataGridView data, int ano)
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
                        cmd = new OleDbCommand("UPDATE ControlOP SET Detalle=@Detalle,Modelo=@Modelo,Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado WHERE ID = " + data.Rows[i].Cells[1].Value);
                    }
                    else
                    {
                        cmd = new OleDbCommand("UPDATE ControlOP SET Detalle=@Detalle,Modelo=@Modelo,Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado WHERE Orden = " + orden + " AND ADF = '" + adf + "' AND Semana = " + semana + " AND Ano = " + ano);
                    }
                    j = 0;
                }
                else
                {
                    cmd = new OleDbCommand("UPDATE ControlOP SET Detalle=@Detalle,Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado  WHERE ID = " + data.Rows[i].Cells[1].Value);
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

        public void cargarDaños(int semana, string orden, DataGridView data, string ADF, int ano)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                string query = "SELECT * FROM DañosOP WHERE Semana = " + semana + " AND Orden = " + orden + " AND Detalle = '" + data.Rows[i].Cells[1].Value.ToString() + "' AND ADF = '" + ADF + "' AND Ano = " + ano;
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
                        for (int j = 2; j < 8; j++)
                        {
                            if (myReader.GetInt32(j + 2) == 1)
                                data.Rows[i].Cells[j].Value = true;
                        }
                        data.Rows[i].Cells[8].Value = myReader.GetString(11);
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

        public void cargarADF006(int semana, string orden, string adf, DataGridView data, int ano)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                string query = "SELECT * FROM ControlOP WHERE Semana = " + semana + " AND Orden = " + orden + " AND ADF = '" + adf + "' AND Ano = " + ano;
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
                        if (!myReader.IsDBNull(4))
                            if (myReader.GetInt32(4) != 0)
                                data.Rows[i].Cells[3].Value = myReader.GetInt32(4);
                        data.Rows[i].Cells[1].Value = myReader.GetInt32(0);
                        if (!myReader.IsDBNull(5))
                            data.Rows[i].Cells[2].Value = myReader.GetString(5);
                        for (int j = 6, k = 5; j < 12; j++, k = k + 2)
                        {
                            data.Rows[i].Cells[k].Value = myReader.GetInt32(j).ToString();
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

        public void modificarDaños(int semana, string orden, string ADF, DataGridView data, int ano)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("UPDATE DañosOP SET Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado,Descripcion=@Descripcion WHERE Semana = " + semana + " AND ADF = '" + ADF + "' AND Detalle = '" + data.Rows[i].Cells[1].Value.ToString() + "' AND Orden = " + orden + " AND Ano = " + ano);
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

        private void button6_Click(object sender, EventArgs e)
        {
            if (ADFExiste(semana, orden.ToString(), "ADF006-Mini Cargador", ano) == false)
            {
                crearADF(semana, orden.ToString(), "ADF006", dataGridView4, ano,"Mini Cargador");
                crearADF(semana, orden.ToString(), "ADF006-2", dataGridView5, ano, "Mini Cargador");
                crearDaños(semana, orden.ToString(), "ADF006-Mini Cargador", dataGridView6, ano);
            }
            else
            {
                modificarADF(semana, orden.ToString(), "ADF006", dataGridView4, ano, "Mini Cargador");
                modificarADF2(semana, orden.ToString(), "ADF006-2-Mini Cargador", dataGridView5, ano);
                modificarDaños(semana, orden.ToString(), "ADF006-Mini Cargador", dataGridView6, ano);
            }
            MessageBox.Show("Control de Equipo Mini Cargador registrado.");
        }

        public bool existeAsignacion(int dia)
        {
            string query = "SELECT * FROM recibosOrdenes WHERE OP = " + orden + " AND Dia = " + dia;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                if (myReader.Read())
                    return true;
                else
                    return false;
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
        }

        public bool existeReciboAsignado(int i, int dia)
        {
            string query = "SELECT * FROM recibosOrdenes WHERE OP = " + orden + " AND Recibo = " + i + " AND Dia = " + dia;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                if (myReader.Read())
                    return true;
                else
                    return false;
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
        }

        public void eliminarAsignacion(int dia)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("DELETE FROM recibosOrdenes WHERE OP = " + orden + " AND Dia = " + dia);
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

        public void agregarAsignacion(int i, DataGridView data, int dia)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO recibosOrdenes (Recibo,OP,Porcentaje,Dia,volumen,volumenOriginal) VALUES (@Recibo,@OP,@Porcentaje,@Dia,@volumen,@volumenOriginal)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Recibo", OleDbType.VarChar).Value = data.Rows[i].Cells[1].Value;
                cmd.Parameters.Add("@OP", OleDbType.VarChar).Value = orden;
                cmd.Parameters.Add("@Porcentaje", OleDbType.VarChar).Value = data.Rows[i].Cells[6].Value;
                cmd.Parameters.Add("@Dia", OleDbType.VarChar).Value = dia;
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
        public void modificarRecibo(DataGridView data, int i, string tipo)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand();
            cmd = new OleDbCommand("UPDATE Recibo SET volumenActual = volumenActual " + tipo + " @volumen WHERE ID =  " + data.Rows[i].Cells[1].Value);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                if(tipo.Equals("+"))
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

        //public void modificarRecibo(DataGridView data, int i, string tipo)
        //{
        //    conn.ConnectionString = connectionString;
        //    OleDbCommand cmd = new OleDbCommand();
        //    cmd = new OleDbCommand("UPDATE recibosOrdenes INNER JOIN Recibo ON recibosOrdenes.Recibo = Recibo.Id SET Recibo.volumenActual = Recibo.volumenActual " + tipo + " [recibosOrdenes].[volumen] WHERE Recibo.ID =  " + data.Rows[i].Cells[1].Value);
        //    cmd.Connection = conn;
        //    conn.Open();
        //    if (conn.State == ConnectionState.Open)
        //    {
        //        cmd.Parameters.Add("@volumen", OleDbType.VarChar).Value = data.Rows[i].Cells[7].Value;
        //        try
        //        {
        //            cmd.ExecuteNonQuery();
        //            conn.Close();
        //        }
        //        catch (OleDbException ex)
        //        {
        //            MessageBox.Show(ex.Source);
        //            conn.Close();
        //        }
        //    }
        //    else
        //    {
        //        MessageBox.Show("Connection Failed");
        //    }
        //}

        private void button9_Click(object sender, EventArgs e)
        {
            if (existeAsignacion(diaRecibos))
            {
                for (int i = 0; i < dataGridView13.Rows.Count; i++)
                {
                    if (existeReciboAsignado(Int32.Parse(dataGridView13.Rows[i].Cells[1].Value.ToString()),diaRecibos))
                        modificarRecibo(dataGridView13, i, "+");
                }
                eliminarAsignacion(diaRecibos);
            }                
            for (int i = 0; i < dataGridView13.Rows.Count; i++)
            {
                if(existeReciboAsignado(Int32.Parse(dataGridView13.Rows[i].Cells[1].Value.ToString()),diaRecibos))
                    modificarRecibo(dataGridView13, i, "+");
                if (Convert.ToBoolean(dataGridView13.Rows[i].Cells[2].Value) == true)
                {
                    agregarAsignacion(i, dataGridView13,diaRecibos);
                    modificarRecibo(dataGridView13, i, "-");
                }
            }
            MessageBox.Show("Inputs registrados.");
        }



        private void tabPage7_Enter(object sender, EventArgs e)
        {
            diaPaquete = diaOrden;
            diaPaquete2 = 0;
            label59.Text = "Día #: " + (diaPaquete + 1) + "                     Fecha Actual: " + DateTime.Now.ToString("dd") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.Year;
            cargarPaquetesTotales(dataGridView14);
            cargarPaquetesTotalesDiarios(dataGridView14, diaPaquete);
        }

        private void linkLabel6_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            diaPaquete--;
            diaPaquete2--;
            DateTime today = DateTime.Now;
            DateTime date = today.AddDays(diaPaquete2);
            label59.Text = "Día #: " + (diaPaquete + 1) + "                     Fecha: " + date.ToString("dd") + "/" + date.ToString("MM") + "/" + date.Year; cargarPaquetesTotales(dataGridView14); cargarPaquetesTotales(dataGridView14);
            cargarPaquetesTotalesDiarios(dataGridView14, diaPaquete);
        }

        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            diaPaquete++;
            diaPaquete2++;
            DateTime today = DateTime.Now;
            DateTime date = today.AddDays(diaPaquete2);
            label59.Text = "Día #: " + (diaPaquete + 1) + "                     Fecha: " + date.ToString("dd") + "/" + date.ToString("MM") + "/" + date.Year; cargarPaquetesTotales(dataGridView14); cargarPaquetesTotales(dataGridView14);
            cargarPaquetesTotalesDiarios(dataGridView14, diaPaquete);
        }

        public void clearDataGrid(DataGridView data)
        {
            data.Rows.Clear();
            data.Columns.Clear();
            data.Refresh();            
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

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            diaRecibos++;
            diaRecibos2++;
            DateTime today = DateTime.Now;
            DateTime date = today.AddDays(diaRecibos2);
            label58.Text = "Día #: " + (diaRecibos + 1) + "                     Fecha: " + date.ToString("dd") + "/" + date.ToString("MM") + "/" + date.Year;
            clearDataGrid(dataGridView13);
            cargarFormatoInput(dataGridView13);
            cargarRecibos(getEspecie(orden), dataGridView13);
            cargarRecibosAsignadas(orden, dataGridView13, diaRecibos);
            clearRecibos(dataGridView13);
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            diaRecibos--;
            diaRecibos2--;
            DateTime today = DateTime.Now;
            DateTime date = today.AddDays(diaRecibos2);
            label58.Text = "Día #: " + (diaRecibos + 1) + "                     Fecha: " + date.ToString("dd") + "/" + date.ToString("MM") + "/" + date.Year; 
            clearDataGrid(dataGridView13);
            cargarFormatoInput(dataGridView13);
            cargarRecibos(getEspecie(orden), dataGridView13);
            cargarRecibosAsignadas(orden, dataGridView13, diaRecibos);
            clearRecibos(dataGridView13);
        }

        private void tabPage3_Enter(object sender, EventArgs e)
        {
            diaRecibos = diaOrden;
            diaRecibos2 = 0;
            DateTime date = DateTime.Now;
            date.AddDays(diaRecibos2);
            label58.Text = "Día #: " + (diaRecibos + 1) + "                     Fecha Actual: " + date.ToString("dd") + "/" + date.ToString("MM") + "/" + date.Year; 
            clearDataGrid(dataGridView13);
            cargarFormatoInput(dataGridView13);
            cargarRecibos(getEspecie(orden), dataGridView13);
            cargarRecibosAsignadas(orden, dataGridView13, diaRecibos);
            clearRecibos(dataGridView13);
        }
        public string getSemana(DateTime date2)
        {
            System.Globalization.CultureInfo ci =
            System.Threading.Thread.CurrentThread.CurrentCulture;
            DayOfWeek fdow = ci.DateTimeFormat.FirstDayOfWeek;
            DayOfWeek today = date2.DayOfWeek;
            DateTime date = DateTime.Now.AddDays(-(today - fdow)).Date;
            DateTime date3 = date.AddDays(6);
            string semana = "Semana: " + date.ToString("dd") + "/" + date.ToString("MM") + "/" + date.Year + "    -    " + date3.ToString("dd") + "/" + date3.ToString("MM") + "/" + date3.Year;
            return semana;
        }

        public static DateTime FirstDateOfWeek(int year, int weekOfYear)
        {
            DateTime jan1 = new DateTime(year, 1, 1);
            int daysOffset = DayOfWeek.Thursday - jan1.DayOfWeek;

            DateTime firstThursday = jan1.AddDays(daysOffset);
            var cal = System.Globalization.CultureInfo.CurrentCulture.Calendar;
            int firstWeek = cal.GetWeekOfYear(firstThursday, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

            var weekNum = weekOfYear;
            if (firstWeek <= 1)
            {
                weekNum -= 1;
            }
            var result = firstThursday.AddDays(weekNum * 7);
            return result.AddDays(-3);
        }

        private void tabPage5_Enter(object sender, EventArgs e)
        {
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            DateTime date1 = DateTime.ParseExact(DateTime.Now.ToString("dd") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.Year, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
            Calendar cal = dfi.Calendar;
            semana = cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);
            DateTime date = DateTime.Now;
            DateTime first = (FirstDateOfWeek(DateTime.Now.Year, semana));
            DateTime last = (FirstDateOfWeek(DateTime.Now.Year, semana + 1)).AddDays(-1);
            label60.Text = "Semana: " + first.ToString("dd") + "/" + first.ToString("MM") + "/" + first.Year + "    -    " + last.ToString("dd") + "/" + last.ToString("MM") + "/" + last.Year;
            cargarADF(semana, orden.ToString(), "ADF006", dataGridView4, ano, "Mini Cargador");
            cargarADF006(semana, orden.ToString(), "ADF006-2-Mini Cargador", dataGridView5, ano);
            cargarDaños(semana, orden.ToString(), dataGridView6, "ADF006-Mini Cargador", ano);
        }

        private void tabPage6_Enter(object sender, EventArgs e)
        {
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            DateTime date1 = DateTime.ParseExact(DateTime.Now.ToString("dd") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.Year, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
            Calendar cal = dfi.Calendar;
            semana = cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);
            DateTime date = DateTime.Now;
            DateTime first = (FirstDateOfWeek(DateTime.Now.Year, semana));
            DateTime last = (FirstDateOfWeek(DateTime.Now.Year, semana + 1)).AddDays(-1);
            label61.Text = "Semana: " + first.ToString("dd") + "/" + first.ToString("MM") + "/" + first.Year + "    -    " + last.ToString("dd") + "/" + last.ToString("MM") + "/" + last.Year;
            cargarADF(semana, orden.ToString(), "ADF006", dataGridView9, ano, "Monta Carga");
            cargarADF006(semana, orden.ToString(), "ADF006-2-Monta Carga", dataGridView8, ano);
            cargarDaños(semana, orden.ToString(), dataGridView7, "ADF006-Monta Carga", ano);
        }

        private void tabPage1_Enter(object sender, EventArgs e)
        {
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            DateTime date1 = DateTime.ParseExact(DateTime.Now.ToString("dd") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.Year, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
            Calendar cal = dfi.Calendar;
            semana = cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);
            DateTime date = DateTime.Now;
            DateTime first = (FirstDateOfWeek(DateTime.Now.Year, semana));
            DateTime last = (FirstDateOfWeek(DateTime.Now.Year, semana + 1)).AddDays(-1);            
            label62.Text = "Semana: " + first.ToString("dd") + "/" + first.ToString("MM") + "/" + first.Year + "    -    " + last.ToString("dd") + "/" + last.ToString("MM") + "/" + last.Year;
            cargarADF(semana, orden.ToString(), "ADF006", dataGridView12, ano, "Chipeadora");
            cargarADF006(semana, orden.ToString(), "ADF006-2-Chipeadora", dataGridView11, ano);
            cargarDaños(semana, orden.ToString(), dataGridView10, "ADF006-Chipeadora", ano);
        }

        public int GetWeeksInYear(int year)
        {
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            DateTime date1 = new DateTime(year, 12, 31);
            Calendar cal = dfi.Calendar;
            return cal.GetWeekOfYear(date1, dfi.CalendarWeekRule,
                                                dfi.FirstDayOfWeek);
        }
        
        private void linkLabel8_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            semana--;
            DateTime date = DateTime.Now;
            DateTime first = (FirstDateOfWeek(DateTime.Now.Year, semana));
            DateTime last = (FirstDateOfWeek(DateTime.Now.Year, semana + 1)).AddDays(-1);
            ano = first.Year;
            label60.Text = "Semana: " + first.ToString("dd") + "/" + first.ToString("MM") + "/" + first.Year + "    -    " + last.ToString("dd") + "/" + last.ToString("MM") + "/" + last.Year;
            clearDataGrid(dataGridView4);
            dataGridView6.Rows.Clear();
            dataGridView5.Rows.Clear();
            cargarADF006(dataGridView4, dataGridView6, dataGridView5, "Mini Cargador");
            if (getInfoMaquina("Mini Cargador") != 0)
                dataGridView4.Rows[0].Cells[3].Value = getInfoMaquina("Mini Cargador");
            cargarADF(semana, orden.ToString(), "ADF006", dataGridView4, ano, "Mini Cargador");
            cargarADF006(semana, orden.ToString(), "ADF006-2-Mini Cargador", dataGridView5, ano);
            cargarDaños(semana, orden.ToString(), dataGridView6, "ADF006-Mini Cargador", ano);
        }

        private void linkLabel7_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            semana++;
            DateTime date = DateTime.Now;
            DateTime first = (FirstDateOfWeek(DateTime.Now.Year, semana));
            DateTime last = (FirstDateOfWeek(DateTime.Now.Year, semana + 1)).AddDays(-1);
            ano = first.Year;            
            label60.Text = "Semana: " + first.ToString("dd") + "/" + first.ToString("MM") + "/" + first.Year + "    -    " + last.ToString("dd") + "/" + last.ToString("MM") + "/" + last.Year;
            clearDataGrid(dataGridView4);
            dataGridView6.Rows.Clear();
            dataGridView5.Rows.Clear();
            cargarADF006(dataGridView4, dataGridView6, dataGridView5, "Mini Cargador");
            if (getInfoMaquina("Mini Cargador") != 0)
                dataGridView4.Rows[0].Cells[3].Value = getInfoMaquina("Mini Cargador");
            cargarADF(semana, orden.ToString(), "ADF006", dataGridView4, ano, "Mini Cargador");
            cargarADF006(semana, orden.ToString(), "ADF006-2-Mini Cargador", dataGridView5, ano);
            cargarDaños(semana, orden.ToString(), dataGridView6, "ADF006-Mini Cargador", ano);
        }

        private void linkLabel1_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            semana--;
            DateTime date = DateTime.Now;
            DateTime first = (FirstDateOfWeek(DateTime.Now.Year, semana));
            DateTime last = (FirstDateOfWeek(DateTime.Now.Year, semana + 1)).AddDays(-1);
            label63.Text = "Semana: " + first.ToString("dd") + "/" + first.ToString("MM") + "/" + first.Year + "    -    " + last.ToString("dd") + "/" + last.ToString("MM") + "/" + last.Year;
        }

        private void tabPage4_Enter(object sender, EventArgs e)
        {
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            DateTime date1 = DateTime.ParseExact(DateTime.Now.ToString("dd") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.Year, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
            Calendar cal = dfi.Calendar;
            semana = cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);
            DateTime date = DateTime.Now;
            DateTime first = (FirstDateOfWeek(DateTime.Now.Year, semana));
            DateTime last = (FirstDateOfWeek(DateTime.Now.Year, semana + 1)).AddDays(-1);
            label63.Text = "Semana: " + first.ToString("dd") + "/" + first.ToString("MM") + "/" + first.Year + "    -    " + last.ToString("dd") + "/" + last.ToString("MM") + "/" + last.Year;
        }

        private void linkLabel2_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            semana++;
            DateTime date = DateTime.Now;
            DateTime first = (FirstDateOfWeek(DateTime.Now.Year, semana));
            DateTime last = (FirstDateOfWeek(DateTime.Now.Year, semana + 1)).AddDays(-1);
            label63.Text = "Semana: " + first.ToString("dd") + "/" + first.ToString("MM") + "/" + first.Year + "    -    " + last.ToString("dd") + "/" + last.ToString("MM") + "/" + last.Year;
        }

        private void linkLabel10_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            semana--;
            DateTime date = DateTime.Now;
            DateTime first = (FirstDateOfWeek(DateTime.Now.Year, semana));
            DateTime last = (FirstDateOfWeek(DateTime.Now.Year, semana + 1)).AddDays(-1);
            label61.Text = "Semana: " + first.ToString("dd") + "/" + first.ToString("MM") + "/" + first.Year + "    -    " + last.ToString("dd") + "/" + last.ToString("MM") + "/" + last.Year;
            clearDataGrid(dataGridView9);
            dataGridView8.Rows.Clear();
            dataGridView7.Rows.Clear();
            cargarADF006(dataGridView9, dataGridView7, dataGridView8, "Monta Carga");
            if (getInfoMaquina("Monta Carga") != 0)
                dataGridView9.Rows[0].Cells[3].Value = getInfoMaquina("Monta Carga");
            cargarADF(semana, orden.ToString(), "ADF006", dataGridView9, ano, "Monta Carga");
            cargarADF006(semana, orden.ToString(), "ADF006-2-Monta Carga", dataGridView8, ano);
            cargarDaños(semana, orden.ToString(), dataGridView7, "ADF006-Monta Carga", ano);
        }

        private void linkLabel9_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            semana++;
            DateTime date = DateTime.Now;
            DateTime first = (FirstDateOfWeek(DateTime.Now.Year, semana));
            DateTime last = (FirstDateOfWeek(DateTime.Now.Year, semana + 1)).AddDays(-1);
            label61.Text = "Semana: " + first.ToString("dd") + "/" + first.ToString("MM") + "/" + first.Year + "    -    " + last.ToString("dd") + "/" + last.ToString("MM") + "/" + last.Year;
            clearDataGrid(dataGridView9);
            dataGridView8.Rows.Clear();
            dataGridView7.Rows.Clear();
            cargarADF006(dataGridView9, dataGridView7, dataGridView8, "Monta Carga");
            if (getInfoMaquina("Monta Carga") != 0)
                dataGridView9.Rows[0].Cells[3].Value = getInfoMaquina("Monta Carga");
            cargarADF(semana, orden.ToString(), "ADF006", dataGridView9, ano, "Monta Carga");
            cargarADF006(semana, orden.ToString(), "ADF006-2-Monta Carga", dataGridView8, ano);
            cargarDaños(semana, orden.ToString(), dataGridView7, "ADF006-Monta Carga", ano);
        }

        private void linkLabel12_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            semana--;
            DateTime date = DateTime.Now;
            DateTime first = (FirstDateOfWeek(DateTime.Now.Year, semana));
            DateTime last = (FirstDateOfWeek(DateTime.Now.Year, semana + 1)).AddDays(-1);
            label62.Text = "Semana: " + first.ToString("dd") + "/" + first.ToString("MM") + "/" + first.Year + "    -    " + last.ToString("dd") + "/" + last.ToString("MM") + "/" + last.Year;
            clearDataGrid(dataGridView12);
            dataGridView11.Rows.Clear();
            dataGridView10.Rows.Clear();
            cargarADF006(dataGridView12, dataGridView10, dataGridView11, "Chipeadora");
            if (getInfoMaquina("Chipeadora") != 0)
                dataGridView12.Rows[0].Cells[3].Value = getInfoMaquina("Chipeadora");
            cargarADF(semana, orden.ToString(), "ADF006", dataGridView12, ano, "Chipeadora");
            cargarADF006(semana, orden.ToString(), "ADF006-2-Chipeadora", dataGridView11, ano);
            cargarDaños(semana, orden.ToString(), dataGridView10, "ADF006-Chipeadora", ano);
        }

        private void linkLabel11_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            semana++;
            DateTime date = DateTime.Now;
            DateTime first = (FirstDateOfWeek(DateTime.Now.Year, semana));
            DateTime last = (FirstDateOfWeek(DateTime.Now.Year, semana + 1)).AddDays(-1);
            label62.Text = "Semana: " + first.ToString("dd") + "/" + first.ToString("MM") + "/" + first.Year + "    -    " + last.ToString("dd") + "/" + last.ToString("MM") + "/" + last.Year;
            clearDataGrid(dataGridView12);
            dataGridView11.Rows.Clear();
            dataGridView10.Rows.Clear();
            cargarADF006(dataGridView12, dataGridView10, dataGridView11, "Chipeadora");
            if (getInfoMaquina("Chipeadora") != 0)
                dataGridView12.Rows[0].Cells[3].Value = getInfoMaquina("Chipeadora");
            cargarADF(semana, orden.ToString(), "ADF006", dataGridView12, ano, "Chipeadora");
            cargarADF006(semana, orden.ToString(), "ADF006-2-Chipeadora", dataGridView11, ano);
            cargarDaños(semana, orden.ToString(), dataGridView10, "ADF006-Chipeadora", ano);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (ADFExiste(semana, orden.ToString(), "ADF006-Monta Carga", ano) == false)
            {
                crearADF(semana, orden.ToString(), "ADF006", dataGridView9, ano, "Monta Carga");
                crearADF(semana, orden.ToString(), "ADF006-2", dataGridView8, ano, "Monta Carga");
                crearDaños(semana, orden.ToString(), "ADF006-Monta Carga", dataGridView7, ano);
            }
            else
            {
                modificarADF(semana, orden.ToString(), "ADF006", dataGridView9, ano, "Monta Carga");
                modificarADF2(semana, orden.ToString(), "ADF006-2-Monta Carga", dataGridView8, ano);
                modificarDaños(semana, orden.ToString(), "ADF006-Monta Carga", dataGridView7, ano);
            }
            MessageBox.Show("Control de Equipo Monta Carga registrado.");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (ADFExiste(semana, orden.ToString(), "ADF006-Chipeadora", ano) == false)
            {
                crearADF(semana, orden.ToString(), "ADF006", dataGridView12, ano, "Chipeadora");
                crearADF(semana, orden.ToString(), "ADF006-2", dataGridView11, ano, "Chipeadora");
                crearDaños(semana, orden.ToString(), "ADF006-Mini Cargador", dataGridView10, ano);
            }
            else
            {
                modificarADF(semana, orden.ToString(), "ADF006", dataGridView12, ano, "Chipeadora");
                modificarADF2(semana, orden.ToString(), "ADF006-2-Chipeadora", dataGridView11, ano);
                modificarDaños(semana, orden.ToString(), "ADF006-Chipeadora", dataGridView10, ano);
            }
            MessageBox.Show("Control de Equipo Mini Cargador registrado.");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            frmCrearProduccion newFrm = new frmCrearProduccion(orden,tipo2);
            if (!newFrm.IsDisposed)
            {
                this.Hide();
                newFrm.ShowDialog();
                this.Close();
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            frmResumenProcesamiento newFrm = new frmResumenProcesamiento(orden);
            newFrm.Show();
        }

    }
}
