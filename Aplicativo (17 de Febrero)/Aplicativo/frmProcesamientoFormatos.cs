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
    public partial class frmProcesamientoFormatos : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();

        string[,] ADF006 = new string[,] { { "Tractor", "Combustible", "Aceite Hyd - Caja", "Aceite Motor" }, { "Horas", "Gal", "Litros", "Litros" } };
        string[,] ADF020 = new string[,] { { "TVS", "SVS", "HR 1000 (3)", "HR 1000 (4)", "SHR", "EDGER", "Bandit 990XP" }, { "Horas", "Horas", "Horas", "Horas", "Horas", "Horas", "Horas" } };
        string[,] ADF0201 = new string[,] { { "Falta de Electricidad" }, { "Horas"} };
        string[] Reporte = new string[] { "Mangueras", "Filtro Combustible", "Llantas", "Otro" };
        int orden = 0;
        int diaOrden = 0;
        int tipo2 = 0;

        public frmProcesamientoFormatos(int op, int dia, int tipo)
        {
            InitializeComponent();
            this.Text = "Formatos de la Orden #" + getNombreOP(op);
            label104.Text = "Día #: " + (dia + 1) + "                     Fecha Actual: " + DateTime.Now.ToString("dd") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.Year;
            cargarEmpleados(op);
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
            cargarFormatoOutput(dataGridView14);
            cargarProductos(op, dataGridView14);
            cargarPaquetesTotales(dataGridView14);
            cargarPaquetesTotalesDiarios(dataGridView14);
            dataGridView14.Columns[2].DefaultCellStyle.Font = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Underline);

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
            //combo = new DataGridViewComboBoxColumn();
            //data.Columns.Add(combo);
            data.Columns[0].FillWeight = 40;
        }

        public void cargarRecibos(string especie, DataGridView data)
        {
            string query = "SELECT r.Id,r.volumenActual, r.Motivo, r.Diametro, r.Largo, r.Cantidad, r.Modulo, r.numRecibo, l.Lote FROM Recibo AS r INNER JOIN Lotes AS l ON r.Lote = l.Codigo WHERE r.Especie = '" + especie + "'";
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

        public void cargarPaquetesTotalesDiarios(DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                string query = "SELECT SUM(porcentaje) FROM Paquete WHERE OP = " + orden + " AND Producto = " + data.Rows[i].Cells[1].Value + " AND dia = " + diaOrden;
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
                        if(!myReader.GetValue(0).ToString().Equals(""))
                            data.Rows[i].Cells[5].Value = myReader.GetValue(0);
                        else
                            data.Rows[i].Cells[5].Value = "0";
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

        public void cargarEmpleados(int orden)
        {
            while (dataGridView3.Rows.Count != 0)
            {
                dataGridView3.Rows.RemoveAt(0);
            }
            string query = "SELECT t.ID, (t.Nombres + ' ' + t.Apellidos), t.Cedula, c.Cargo FROM CargoLaboral AS c INNER JOIN (Trabajadores AS t INNER JOIN produccionEmpleados AS s ON t.ID = s.Trabajador) ON c.ID = t.Cargo WHERE s.Orden = " + orden;
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
                    dataGridView3.Rows.Add();
                    dataGridView3.Rows[i].Cells[0].Value = i + 1;
                    dataGridView3.Rows[i].Cells[1].Value = myReader.GetInt32(0);
                    dataGridView3.Rows[i].Cells[2].Value = myReader.GetString(1);
                    dataGridView3.Rows[i].Cells[3].Value = myReader.GetInt32(2);
                    dataGridView3.Rows[i].Cells[4].Value = myReader.GetString(3);
                    dataGridView3.Rows[i].Cells[12].Value = 0;
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
            frmRegistrarPaquete newFrm = new frmRegistrarPaquete(orden,diaOrden);
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

    }
}
