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
        List<string> productos = new List<string>();
        List<int> idProductos = new List<int>();
        List<double> volProductos = new List<double>();

        public frmResumenProcesamiento(int orden, int dia)
        {
            InitializeComponent();
            this.Text = "Resumen de Orden #: " + getNombreOP(orden);
            cargarInputs(dataGridView13, orden);
            dataGridView13.Columns[7].DefaultCellStyle.Font = new Font(dataGridView13.DefaultCellStyle.Font, FontStyle.Underline);
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
            label22.Text = "Total: " + getTotalADF020(dataGridView1, 3) + " horas.";
            getProductos(orden);
            getVolumenes();
            columnasInput(dataGridView14);
            formatoInputs(dataGridView14, dia);
            getTotales(dataGridView14);
        }

        public void getTotales(DataGridView data)
        {
            data.Columns[0].Width = 50;
            data.Rows.Add();
            data.Rows[data.Rows.Count - 1].Cells[0].Value = "Total";
            double volTotal = 0, volCol = 0, volFila = 0;
            double paquetesTotal = 0, paquetesCol = 0, paquetesFila = 0;
            for (int i = 0; i < idProductos.Count; i++)
            {
                volFila = 0;
                paquetesFila = 0;
                for (int j = 0; j < data.Rows.Count - 1; j++)
                {
                    data.Rows[j].Cells[2 + (i * 2)].Value = double.Parse(data.Rows[j].Cells[2 + (i * 2)].Value.ToString()) * double.Parse(data.Rows[j].Cells[1 + (i * 2)].Value.ToString());
                    volFila += double.Parse(data.Rows[j].Cells[2 + (i * 2)].Value.ToString());
                    paquetesFila += double.Parse(data.Rows[j].Cells[1 + (i * 2)].Value.ToString());
                }
                data.Rows[data.Rows.Count - 1].Cells[1 + (i * 2)].Value = paquetesFila;
                data.Rows[data.Rows.Count - 1].Cells[2 + (i * 2)].Value = volFila;
                paquetesTotal += paquetesFila;
                volTotal += volFila;
            }
            for (int i = 0; i < data.Rows.Count - 1; i++)
            {
                volCol = 0;
                paquetesCol = 0;
                for (int j = 0; j < idProductos.Count; j++)
                {
                    volCol += double.Parse(data.Rows[i].Cells[2 + (j * 2)].Value.ToString());
                    paquetesCol += double.Parse(data.Rows[i].Cells[1 + (j * 2)].Value.ToString());
                }
                data.Rows[i].Cells[data.Columns.Count - 2].Value = paquetesCol;
                data.Rows[i].Cells[data.Columns.Count - 1].Value = volCol;
            }
            data.Columns[0].DefaultCellStyle.Font = new Font(dataGridView14.DefaultCellStyle.Font, FontStyle.Bold);
            data.Columns[data.Columns.Count - 1].DefaultCellStyle.Font = new Font(dataGridView14.DefaultCellStyle.Font, FontStyle.Bold);
            data.Columns[data.Columns.Count - 2].DefaultCellStyle.Font = new Font(dataGridView14.DefaultCellStyle.Font, FontStyle.Bold);
            data.Rows[data.Rows.Count - 1].DefaultCellStyle.Font = new Font(dataGridView14.DefaultCellStyle.Font, FontStyle.Bold);
            data.Rows[data.Rows.Count - 1].Cells[data.Columns.Count - 2].Value = paquetesTotal;
            data.Rows[data.Rows.Count - 1].Cells[data.Columns.Count - 1].Value = volTotal;
            data.Columns[data.Columns.Count - 1].Width = 100;
            data.Columns[data.Columns.Count - 2].Width = 100;
        }

        public void formatoInputs(DataGridView data, int dia)
        {
            for (int i = 0; i < dia + 1; i++)
            {
                data.Rows.Add();
                data.Rows[i].Cells[0].Value = i + 1;
                for (int j = 0; j < volProductos.Count; j++)
			    {
                    data.Rows[i].Cells[2 + (j * 2)].Value = volProductos[j];
			    }                
            }
            //data.Rows.Add();
            //data.Rows[data.Rows.Count-1].Cells[0].Value = "Total";
            cargarPaquetesTotales(data);
        }

        public void cargarPaquetesTotales(DataGridView data)
        {
            for (int j = 0; j < idProductos.Count; j++)
            {
                for (int i = 0; i < data.Rows.Count; i++)
                {
                    string query = "SELECT SUM(porcentaje) FROM Paquete WHERE OP = " + op + " AND dia = " + (Int32.Parse(data.Rows[i].Cells[0].Value.ToString()) - 1) + " AND Parcial = 0 AND Producto = " + idProductos[j];
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
                            double value = 0;
                            if (myReader.GetValue(0).ToString().Equals(""))
                                value = 0;
                            else
                                value = double.Parse(myReader.GetValue(0).ToString());
                            data.Rows[i].Cells[1 + (j*2)].Value = value;
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
                cargarPaquetesTotalesParciales(data);   
            }
        }

        public void cargarPaquetesTotalesParciales(DataGridView data)
        {
            for (int j = 0; j < idProductos.Count; j++)
            {
                for (int i = 0; i < data.Rows.Count; i++)
                {
                    string query = "SELECT SUM(porcentaje) FROM paqueteParcial WHERE OP = " + op + " AND dia = " + (Int32.Parse(data.Rows[i].Cells[0].Value.ToString()) - 1) + " AND Producto = " + idProductos[j];
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
                            double value = 0;
                            if (myReader.GetValue(0).ToString().Equals(""))
                                value = 0;
                            else
                                value = double.Parse(myReader.GetValue(0).ToString());
                            if (data.Rows[i].Cells[1 + (j * 2)].Value != null)
                                data.Rows[i].Cells[1 + (j*2)].Value = double.Parse(data.Rows[i].Cells[1 + (j*2)].Value.ToString()) + value;
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

        public void getProductos(int id)
        {
            string query = "SELECT pp.Id, pp.Orden, pp.Producto, p.Codigo, pp.Cantidad, pp.Volumen FROM Productos AS p INNER JOIN produccionProducto AS pp ON p.ID = pp.Producto WHERE Orden = " + id;
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
                    productos.Add(myReader.GetString(3));
                    idProductos.Add(myReader.GetInt32(2));
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }

            productos.Add("Total");
        }

        public void getVolumenes()
        {
            for (int i = 0; i < idProductos.Count; i++)
            {
                volProductos.Add(getVolumen(idProductos[i]));
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

        public void columnasInput(DataGridView dataGridView14)
        {
            dataGridView14.Columns.Add("dia", "Día");
            for (int i = 0; i < productos.Count; i++)
            {
                dataGridView14.Columns.Add("numPaquetes" + i, "# Paquetes");
                dataGridView14.Columns.Add("vol" + i, "Volumen");
                //dataGridView1.Columns.Add(idProductos[i].ToString(), productos[i]);
            }
            //dataGridView1.Columns.Add("numPaquetes", "# Paquetes");
            //dataGridView1.Columns.Add("vol", "Volumen");

            //for (int j = 0; j < dataGridView14.ColumnCount; j++)
            //{
            //    dataGridView14.Columns[j].Width = 45;
            //}

            dataGridView14.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
            dataGridView14.ColumnHeadersHeight = dataGridView1.ColumnHeadersHeight * 2;
            dataGridView14.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
            dataGridView14.CellPainting += new DataGridViewCellPaintingEventHandler(dataGridView14_CellPainting);
            dataGridView14.Paint += new PaintEventHandler(dataGridView14_Paint);

            dataGridView14.Scroll += new ScrollEventHandler(dataGridView14_Scroll);
            dataGridView14.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridView14_ColumnWidthChanged);
        }

        void dataGridView14_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            Rectangle rtHeader = dataGridView14.DisplayRectangle;
            rtHeader.Height = dataGridView14.ColumnHeadersHeight / 2;
            dataGridView14.Invalidate(rtHeader);
        }

        void dataGridView14_Scroll(object sender, ScrollEventArgs e)
        {
            Rectangle rtHeader = dataGridView14.DisplayRectangle;
            rtHeader.Height = dataGridView14.ColumnHeadersHeight / 2;
            dataGridView14.Invalidate(rtHeader);
        }

        void dataGridView14_Paint(object sender, PaintEventArgs e)
        {
            for (int j = 1; j < dataGridView14.Columns.Count - 1; )
            {
                Rectangle r1 = dataGridView14.GetCellDisplayRectangle(j, -1, true);
                int w2 = dataGridView14.GetCellDisplayRectangle(j + 1, -1, true).Width;
                r1.X += 1;
                r1.Y += 1;
                r1.Width = r1.Width + w2 - 2;
                r1.Height = r1.Height / 2 - 2;
                e.Graphics.FillRectangle(new SolidBrush(dataGridView14.ColumnHeadersDefaultCellStyle.BackColor), r1);
                StringFormat format = new StringFormat();
                format.Alignment = StringAlignment.Center;
                format.LineAlignment = StringAlignment.Center;
                e.Graphics.DrawString(productos[j / 2],
                    dataGridView14.ColumnHeadersDefaultCellStyle.Font,
                    new SolidBrush(dataGridView14.ColumnHeadersDefaultCellStyle.ForeColor),
                    r1,
                    format);
                j += 2;
            }
        }

        void dataGridView14_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex == -1 && e.ColumnIndex > -1)
            {
                Rectangle r2 = e.CellBounds;
                r2.Y += e.CellBounds.Height / 2;
                r2.Height = e.CellBounds.Height / 2;

                e.PaintBackground(r2, true);

                e.PaintContent(r2);
                e.Handled = true;
            }
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
            double total = 0;
            string query = "SELECT recibosOrdenes.Dia, Recibo.numRecibo, Lotes.Lote, recibosOrdenes.Porcentaje, recibosOrdenes.volumen, Recibo.Modulo, h.OT, h.ID FROM ((recibosOrdenes INNER JOIN Recibo ON recibosOrdenes.Recibo = Recibo.Id) INNER JOIN Lotes ON Recibo.Lote = Lotes.Codigo) INNER JOIN historicoOrdenes AS h ON Recibo.Orden = h.ID WHERE (((recibosOrdenes.OP)= " + orden + ")) ORDER BY recibosOrdenes.Dia;";
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
                    data.Rows[i].Cells[1].Value = myReader.GetInt32(0) + 1;
                    data.Rows[i].Cells[2].Value = myReader.GetString(1);
                    data.Rows[i].Cells[3].Value = myReader.GetString(2);
                    data.Rows[i].Cells[4].Value = myReader.GetString(3);
                    data.Rows[i].Cells[5].Value = myReader.GetDouble(4);
                    data.Rows[i].Cells[6].Value = myReader.GetInt32(5);
                    data.Rows[i].Cells[7].Value = myReader.GetString(6);
                    data.Rows[i].Cells[8].Value = myReader.GetInt32(7);
                    total += myReader.GetDouble(4);
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
            label23.Text = "Volumen total seleccionado: " + total + " m3.";
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
                            data.Rows[i].Cells[7].Value = valor * double.Parse(myReader.GetValue(0).ToString());
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
                    data.Rows[i].Cells[7].Value = valor * double.Parse(myReader.GetValue(0).ToString());
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
            string query = "SELECT Detalle, SUM(Lunes+Martes+Miercoles+Jueves+Viernes+Sabado) FROM ControlOP WHERE ADF = '" + adf + "' AND Orden = " + orden + " GROUP BY Detalle";
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
                    data.Rows[i].Cells[2].Value = "Horas";
                    data.Rows[i].Cells[3].Value = myReader.GetValue(1).ToString();
                    //data.Rows[i].Cells[2].Value = myReader.GetValue(2).ToString();
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
            string query = "SELECT c.Modelo, c.Detalle, (i.Modelo + ' / ' + i.Marca + ' / ' + i.Placa) FROM ControlOP As c INNER JOIN Maquinarias As i ON c.Modelo = i.ID WHERE c.Orden = " + orden + " AND ADF = '" + adf + "' GROUP BY c.Modelo, c.Detalle, (i.Modelo + ' / ' + i.Marca + ' / ' + i.Placa)";
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
                            data.Rows[i].Cells[4].Value = double.Parse(myReader.GetValue(0).ToString());
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

        private void dataGridView14_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {

        }

        private void dataGridView13_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView13.CurrentCell.ColumnIndex == 7)
            {
                frmCrearOrden newFrm = new frmCrearOrden(dataGridView13.Rows[dataGridView13.CurrentCell.RowIndex].Cells[8].Value.ToString(), 1);
                newFrm.Show();
            }
        }

    }
}
