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
    public partial class frmCrearDespacho : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        string[,] ADF006 = new string[,] { { "Tractor", "Combustible", "Aceite Hyd - Caja", "Aceite Motor" }, { "Horas", "Gal", "Litros", "Litros" } };
        string[] Reporte = new string[] { "Mangueras", "Filtro Combustible", "Llantas", "Otro" };
        int index = 0;
        int tipo2 = 0;
        bool sw = true;

        public frmCrearDespacho(int tipo)
        {
            InitializeComponent();
            tipo2 = tipo;
            Variables.cargar(comboBox1, "SELECT ID, (Nombres + Apellidos) As Nombre FROM Trabajadores", "Nombre");
            //cargarEmpleados((getMaxDespacho() + 1));
            cargarADF006(dataGridView9, dataGridView7, dataGridView8, "Tractor");
            cargarADF006(dataGridView4, dataGridView6, dataGridView5, "Mini Cargador");
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
                Variables.cargar(dataGridView1, "SELECT historicoProduccion.ID, historicoProduccion.OP, Sum(Paquete.volumenActual) FROM historicoProduccion INNER JOIN Paquete ON historicoProduccion.Id = Paquete.OP GROUP BY historicoProduccion.ID, historicoProduccion.OP HAVING SUM(Paquete.volumenActual) > 0 ORDER BY historicoProduccion.ID DESC");
                dataGridView1.Columns[2].HeaderText = "Vol. Producción";
                dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                cargarFormatoInput(dataGridView13);
            }
            else
            {
                Variables.cargar(dataGridView1, "SELECT Pila,SUM(volumenActual) FROM reciboCliente GROUP BY Pila HAVING SUM(volumenActual) > 0 ORDER BY Pila Desc");
                dataGridView1.Columns[0].Visible = true;
                dataGridView1.Columns[1].HeaderText = "Vol. Pila";                
                dataGridView1.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                cargarFormatoInput(dataGridView13);
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

        public void cargarADF006(DataGridView data1, DataGridView data2, DataGridView data3, string tipo)
        {
            crearADF006(data1, "SELECT * FROM Insumos WHERE Descripcion LIKE '%", "Modelo", ADF006, tipo);
            formato(data2, Reporte);
            formato(data3, "Recorridos");
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


        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if(tipo2 ==0)
                cargarRecibos(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString(), dataGridView13);
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

    }
}
