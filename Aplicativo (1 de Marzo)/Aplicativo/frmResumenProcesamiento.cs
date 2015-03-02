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
    public partial class frmResumenProcesamiento : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        int op;

        public frmResumenProcesamiento(int orden)
        {
            InitializeComponent();
            this.Text = "Resumen de Orden #: " + getNombreOP(orden);
            cargarInputs(dataGridView13, orden);
            Variables.cargar(dataGridView16, "SELECT c.Cuchilla, Cuchillas.Codigo FROM (cuchillaAsignadas AS c INNER JOIN Maquinarias AS m ON c.Maquina = m.ID) INNER JOIN Cuchillas ON c.Cuchilla = Cuchillas.Id WHERE OP = " + orden);
            op = orden;
            cargarADF006(dataGridView4, dataGridView5, dataGridView6, "Mini Cargador", label15);
            cargarADF006(dataGridView8, dataGridView3, dataGridView7, "Monta Carga", label1);
            cargarADF006(dataGridView11, dataGridView9, dataGridView10, "Chipeadora", label5);
            cargarADF(dataGridView2,orden,"ADF020");
            cargarADFCantidad(dataGridView2, orden, "ADF020");
            cargarADF(dataGridView17, orden, "ADF020-3");
            cargarADFCantidad(dataGridView17, orden, "ADF020-3");
            label7.Text = "Total: " + getTotalADF020(dataGridView2,4) + " horas.";
            label21.Text = "Total: " + getTotalADF020(dataGridView17,4) + " horas.";
            cargarADFBajos(dataGridView1, op, "ADF020-2");
            label22.Text = "Total: " + getTotalADF020(dataGridView1, 4) + " horas.";
        }

        public void cargarADF006(DataGridView data1, DataGridView data2, DataGridView data3, string tipo, Label label)
        {
            cargarADF(data1, op, "ADF006", tipo);
            if (data1.Rows.Count > 0)
                cargarADFCantidad(data1, op, "ADF006", tipo);
            cargarADF2(data3, op, "ADF006-2",tipo);
            cargarADF2Cantidad(data3, op, "ADF006-2",tipo);
            cargarADFDaños(data2, op, "ADF006",tipo);
            data1.Columns[data1.Columns.Count - 1].DefaultCellStyle.Format = "c";
            label.Text = "Total: " + getTotalADF006(data1);
        }

        public string getTotalADF006(DataGridView data)
        {
            string total = "";
            int valor = 0;
            for (int i = 0; i < data.Rows.Count; i++)
            {
                valor += Int32.Parse(data.Rows[i].Cells[7].Value.ToString());
            }
            total = String.Format("{0:c}", valor);
            return total;
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

        public void cargarInputs(DataGridView data, int orden)
        {
            string query = "SELECT recibosOrdenes.Dia, Recibo.numRecibo, Lotes.Lote, recibosOrdenes.Porcentaje, recibosOrdenes.volumen FROM (recibosOrdenes INNER JOIN Recibo ON recibosOrdenes.Recibo = Recibo.Id) INNER JOIN Lotes ON Recibo.Lote = Lotes.Codigo WHERE recibosOrdenes.OP = " + orden + " ORDER BY Dia;";
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
                    data.Rows[i].Cells[2].Value = myReader.GetString(1);
                    data.Rows[i].Cells[3].Value = myReader.GetString(2);
                    data.Rows[i].Cells[4].Value = myReader.GetString(3);
                    data.Rows[i].Cells[5].Value = myReader.GetDouble(4);
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

        private void dataGridView16_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            Variables.cargar(dataGridView12, "SELECT ID,Evento,Fecha,Hora FROM cuchillaEvento WHERE Cuchilla = " + dataGridView16.Rows[dataGridView16.CurrentCell.RowIndex].Cells[0].Value.ToString() + " AND OP = " + op);
            dataGridView12.Columns[1].HeaderText = "Afilado/Triscado";
            Variables.cargar(dataGridView15, "SELECT ID,Hora,Fecha FROM cuchillaHoras WHERE Cuchilla = " + dataGridView16.Rows[dataGridView16.CurrentCell.RowIndex].Cells[0].Value.ToString() + " AND OP = " + op);
            label13.Text = "Total : " + getTotal(dataGridView15) + " horas registradas.";
            getTotalAfiTris(dataGridView12);
        }

        public double getTotal(DataGridView data)
        {
            double valor = 0;
            for (int i = 0; i < data.Rows.Count; i++)
            {
                valor += double.Parse(data.Rows[i].Cells[1].Value.ToString());
            }
            return valor;
        }

        public double getTotalADF020(DataGridView data, int indice)
        {
            double valor = 0;
            for (int i = 0; i < data.Rows.Count; i++)
            {
                valor += double.Parse(data.Rows[i].Cells[indice].Value.ToString());
            }
            return valor;
        }

        public void getTotalAfiTris(DataGridView data)
        {
            int total = 0, total2 = 0;
            for (int i = 0; i < data.Rows.Count; i++)
            {
                if (data.Rows[i].Cells[1].Value.ToString().Equals("Afilada"))
                    total++;
                else
                    total2++;
            }
            label14.Text = "# de Afiladas : " + total + "     # de Triscadas : " + total2;
        }

        public void cargarADF(DataGridView data, int orden, string adf, string tipo)
        {
            while (data.Rows.Count != 0)
            {
                data.Rows.RemoveAt(0);
            }
            int i = 0;
            if (adf.Equals("ADF006"))
            {
                cargarEquipoADF(data, orden, "ADF006-1", tipo);
                i++;
            }
            string query = "SELECT c.Modelo, c.Detalle, (i.Clase + ' ' + i.Marca + ' ' + i.Modelo), c.Unidad, i.Costo_Unitario FROM ControlOP As c INNER JOIN Insumos As i ON c.Modelo = i.ID WHERE c.Orden = " + orden + " AND ADF = '" + (adf + "-" + tipo) + "' GROUP BY c.Modelo, c.Detalle, (i.Clase + ' ' + i.Marca + ' ' + i.Modelo), c.Unidad, i.Costo_Unitario";
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
                    data.Rows.Add();
                    data.Rows[i].Cells[0].Value = i + 1;
                    data.Rows[i].Cells[1].Value = myReader.GetInt32(0);
                    data.Rows[i].Cells[2].Value = myReader.GetString(1);
                    if (myReader.GetInt32(0) != 0)
                        data.Rows[i].Cells[3].Value = myReader.GetString(2);
                    else
                        data.Rows[i].Cells[3].Value = "";
                    data.Rows[i].Cells[4].Value = myReader.GetString(3);
                    //data.Rows[i].Cells[5].Value = myReader.GetValue(4).ToString();
                    data.Rows[i].Cells[6].Value = String.Format("{0:c}", myReader.GetDouble(4));
                    //data.Rows[i].Cells[7].Value = String.Format("{0:c}", (Int32.Parse(myReader.GetValue(4).ToString()) * myReader.GetInt32(5)).ToString());
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

        public void cargarEquipoADF(DataGridView data, int orden, string adf, string tipo)
        {
            string query = "SELECT c.Modelo, c.Detalle, (i.Tipo + ' ' + i.Marca + ' ' + i.Modelo), c.Unidad FROM ControlOP As c INNER JOIN Maquinarias As i ON c.Modelo = i.ID WHERE c.Orden = " + orden + " AND ADF = '" + (adf + "-" + tipo) + "' GROUP BY c.Modelo, c.Detalle, (i.Tipo + ' ' + i.Marca + ' ' + i.Modelo), c.Unidad";
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
                    data.Rows[i].Cells[2].Value = myReader.GetString(1);
                    if (myReader.GetInt32(0) != 0)
                        data.Rows[i].Cells[3].Value = myReader.GetString(2);
                    else
                        data.Rows[i].Cells[3].Value = "";
                    data.Rows[i].Cells[4].Value = myReader.GetString(3);
                    data.Rows[i].Cells[6].Value = 0;
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

        public void cargarADFCantidad(DataGridView data, int orden, string adf, string tipo)
        {
            int index = 0;
            if (adf.Equals("ADF006"))
            {
                    cargarCantidadEquipoADF(data, orden, "ADF006-1", tipo);
                    index++;
            }
            for (int i = index; i < data.Rows.Count; i++)
            {
                string query = "SELECT SUM(Lunes + Martes + Miercoles + Jueves + Viernes + Sabado) FROM ControlOP WHERE Orden = " + orden + " AND ADF = '" + (adf + "-" + tipo) + "' AND Modelo = " + data.Rows[i].Cells[1].Value.ToString() + " AND Detalle = '" + data.Rows[i].Cells[2].Value.ToString() + "'";
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
                        if (!myReader.GetValue(0).ToString().Equals(""))
                        {
                            data.Rows[i].Cells[5].Value = myReader.GetValue(0).ToString();
                            var cultureInfo = new System.Globalization.CultureInfo("en-US");
                            double valor = double.Parse(data.Rows[i].Cells[6].Value.ToString(), System.Globalization.NumberStyles.Currency);
                            data.Rows[i].Cells[7].Value = valor * Int32.Parse(myReader.GetValue(0).ToString());
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

        public void cargarCantidadEquipoADF(DataGridView data, int orden, string adf, string tipo)
        {
            int i = 0;
            string query = "SELECT SUM(Lunes + Martes + Miercoles + Jueves + Viernes + Sabado) FROM ControlOP WHERE Orden = " + orden + " AND ADF = '" + (adf + "-" + tipo) + "' AND Modelo = " + data.Rows[i].Cells[1].Value.ToString() + " AND Detalle = '" + data.Rows[i].Cells[2].Value.ToString() + "'";
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
                    data.Rows[i].Cells[5].Value = myReader.GetValue(0).ToString();
                    var cultureInfo = new System.Globalization.CultureInfo("en-US");
                    int valor = Int32.Parse(data.Rows[i].Cells[6].Value.ToString(), cultureInfo);
                    data.Rows[i].Cells[7].Value = valor * Int32.Parse(myReader.GetValue(0).ToString());
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

        public void cargarADF2(DataGridView data, int orden, string adf, string tipo)
        {
            while (data.Rows.Count != 0)
            {
                data.Rows.RemoveAt(0);
            }
            string query = "SELECT c.Modelo, c.Detalle, (i.Tipo + ' ' + i.Marca + ' ' + i.Modelo), c.Unidad FROM ControlOP As c INNER JOIN Maquinarias As i ON c.Modelo = i.ID WHERE c.Orden = " + orden + " AND ADF = '" + (adf + "-" + tipo) + "' GROUP BY c.Modelo, c.Detalle, (i.Tipo + ' ' + i.Marca + ' ' + i.Modelo), c.Unidad";
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
                    data.Rows[i].Cells[2].Value = myReader.GetString(1);
                    if (myReader.GetInt32(0) != 0)
                        data.Rows[i].Cells[3].Value = myReader.GetString(2);
                    else
                        data.Rows[i].Cells[3].Value = "";
                    data.Rows[i].Cells[4].Value = myReader.GetString(3);
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

        public void cargarADF2Cantidad(DataGridView data, int orden, string adf, string tipo)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                string query = "SELECT SUM(Lunes + Martes + Miercoles + Jueves + Viernes + Sabado) FROM ControlOP WHERE Orden = " + orden + " AND ADF = '" + (adf + "-" + tipo) + "' AND Modelo = " + data.Rows[i].Cells[1].Value.ToString() + " AND Detalle = '" + data.Rows[i].Cells[2].Value.ToString() + "'";
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
                        data.Rows[i].Cells[5].Value = myReader.GetValue(0).ToString();
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

        public void cargarADFDaños(DataGridView data, int orden, string adf, string tipo)
        {
            while (data.Rows.Count != 0)
            {
                data.Rows.RemoveAt(0);
            }
            string query = "SELECT Detalle, Descripcion, SUM(Lunes+Martes+Miercoles+Jueves+Viernes+Sabado) FROM DañosOP WHERE ADF = '" + (adf + "-" + tipo) + "' AND Orden = " + orden + " GROUP BY Detalle, Descripcion";
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
                    data.Rows[i].Cells[2].Value = myReader.GetString(0);
                    data.Rows[i].Cells[3].Value = myReader.GetString(1);
                    data.Rows[i].Cells[4].Value = myReader.GetValue(2).ToString();
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

        public void cargarADFBajos(DataGridView data, int orden, string adf)
        {
            while (data.Rows.Count != 0)
            {
                data.Rows.RemoveAt(0);
            }
            string query = "SELECT c.Detalle, SUM(Lunes+Martes+Miercoles+Jueves+Viernes+Sabado), (m.Modelo + ' / ' + m.Marca + ' / ' + m.Placa) As Maquina FROM ControlOP AS c INNER JOIN Maquinarias AS m ON c.Modelo = m.ID WHERE ADF = 'ADF020' AND Orden = 4 GROUP BY Detalle, (m.Modelo + ' / ' + m.Marca + ' / ' + m.Placa);";
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
                    data.Rows[i].Cells[1].Value = myReader.GetString(0);
                    data.Rows[i].Cells[3].Value = "Horas";
                    data.Rows[i].Cells[4].Value = myReader.GetValue(1).ToString();
                    data.Rows[i].Cells[2].Value = myReader.GetValue(2).ToString();
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


        public void cargarADF(DataGridView data, int orden, string adf)
        {
            while (data.Rows.Count != 0)
            {
                data.Rows.RemoveAt(0);
            }
            int i = 0;
            string query = "SELECT c.Modelo, c.Detalle, (i.Tipo + ' ' + i.Marca + ' ' + i.Modelo) FROM ControlOP As c INNER JOIN Maquinarias As i ON c.Modelo = i.ID WHERE c.Orden = " + orden + " AND ADF = '" + adf + "' GROUP BY c.Modelo, c.Detalle, (i.Tipo + ' ' + i.Marca + ' ' + i.Modelo)";
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
                    data.Rows.Add();
                    data.Rows[i].Cells[0].Value = i + 1;
                    data.Rows[i].Cells[1].Value = myReader.GetInt32(0);
                    data.Rows[i].Cells[2].Value = myReader.GetString(1);
                    data.Rows[i].Cells[3].Value = myReader.GetString(2);
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

        public void cargarADFCantidad(DataGridView data, int orden, string adf)
        {
            int index = 0;
            for (int i = index; i < data.Rows.Count; i++)
            {
                string query = "SELECT SUM(Lunes + Martes + Miercoles + Jueves + Viernes + Sabado) FROM ControlOP WHERE Orden = " + orden + " AND ADF = '" + adf + "' AND Modelo = " + data.Rows[i].Cells[1].Value.ToString() + " AND Detalle = '" + data.Rows[i].Cells[2].Value.ToString() + "'";
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
                        if (!myReader.GetValue(0).ToString().Equals(""))
                        {
                            data.Rows[i].Cells[4].Value = Int32.Parse(myReader.GetValue(0).ToString());
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

    }
}
