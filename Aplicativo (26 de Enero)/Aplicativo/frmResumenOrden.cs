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
    public partial class frmResumenOrden : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        bool sw2 = true;
        int diasTotales = 0;
        double costoEstipuladoJornal, costoEstipuladoInsumos, costoJornal, costoInsumos;

        public frmResumenOrden(int orden)
        {
            InitializeComponent();
            this.Text = "Resumen de Orden #: " + orden;
            hideFormatos();
            getFormatos(orden.ToString());
            eliminarFormatos();
            if (sw2)
            {
                formatoCerrado(dataGridView14, "SELECT ID, (Clase + ' ' + Marca + ' ' + Modelo) As Insumo FROM Insumos", "Insumo");
                cargarCerrado(orden, dataGridView14);
                dataGridView14.Columns[4].Visible = false;
                getOrden(orden);
                lblCostoEI.Text = "Costo Insumos: " + String.Format("{0:c}", (costoEstipuladoInsumos - costoEstipuladoJornal));
                lblCostoEJ.Text = "Costo Jornal: " + String.Format("{0:c}", costoEstipuladoJornal);
                lblCostoI.Text = "Costo Insumos: " + String.Format("{0:c}", costoInsumos);
                lblCostoJ.Text = "Costo Jornal: " + String.Format("{0:c}", costoJornal);
                lblCostoEF.Text = "Costo Final: " + String.Format("{0:c}", (costoEstipuladoInsumos));
                lblCostoF.Text = "Costo Final: " + String.Format("{0:c}", (costoInsumos + costoJornal));
                for (int i = 0; i < tabControl1.TabPages.Count; i++)
                {
                    if (tabControl1.TabPages[i].Text.Equals("ADF002"))
                    {
                        cargarEmpleados(dataGridView3, orden);
                        cargarADFJornal(dataGridView3, orden);
                        cargarCostosDiarios(dataGridView3);
                        dataGridView3.Columns[dataGridView3.Columns.Count - 1].DefaultCellStyle.Format = "c";
                        label2.Text ="Dias totales : " + diasTotales.ToString() + " dias    Costo: " + String.Format("{0:c}", (getCostoData(dataGridView3)));
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF003"))
                    {
                        cargarADF(dataGridView1, orden, "ADF003");
                        cargarADFCantidad(dataGridView1, orden, "ADF003");
                        dataGridView1.Columns[dataGridView1.Columns.Count - 1].DefaultCellStyle.Format = "c";
                        label4.Text = "Total: " + getTotal(dataGridView1);
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF004"))
                    {
                        cargarADF(dataGridView2, orden, "ADF004");
                        cargarADFCantidad(dataGridView2, orden, "ADF004");
                        dataGridView2.Columns[dataGridView2.Columns.Count - 1].DefaultCellStyle.Format = "c";
                        label5.Text = "Total: " + getTotal(dataGridView2);
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF005"))
                    {
                        cargarADF(dataGridView4, orden, "ADF005");
                        cargarADFCantidad(dataGridView4, orden, "ADF005");
                        cargarADF2(dataGridView5, orden, "ADF005-2");
                        cargarADF2Cantidad(dataGridView5, orden, "ADF005-2");
                        cargarADFDaños(dataGridView6, orden, "ADF005");
                        dataGridView4.Columns[dataGridView4.Columns.Count - 1].DefaultCellStyle.Format = "c";
                        label3.Text = "Total: " + getTotal(dataGridView4);
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF006"))
                    {
                        cargarADF(dataGridView9, orden, "ADF006");
                        cargarADFCantidad(dataGridView9, orden, "ADF006");
                        cargarADF2(dataGridView8, orden, "ADF006-2");
                        cargarADF2Cantidad(dataGridView8, orden, "ADF006-2");
                        cargarADFDaños(dataGridView7, orden, "ADF006");
                        dataGridView9.Columns[dataGridView9.Columns.Count - 1].DefaultCellStyle.Format = "c";
                        label9.Text = "Total: " + getTotal(dataGridView9);
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF007"))
                    {
                        cargarEquipoADF(dataGridView10, orden, "ADF007");
                        cargarCantidadEquipoADF2(dataGridView10, orden, "ADF007");
                        dataGridView10.Columns[dataGridView10.Columns.Count - 1].DefaultCellStyle.Format = "c";
                        dataGridView10.Columns[7].Visible = false;
                        dataGridView10.Columns[6].Visible = false;
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF008"))
                    {
                        cargarEquipoADF(dataGridView11, orden, "ADF008");
                        cargarCantidadEquipoADF2(dataGridView11, orden, "ADF008");
                        dataGridView11.Columns[7].Visible = false;
                        dataGridView11.Columns[6].Visible = false;
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF009"))
                    {
                        cargarADF(dataGridView12, orden, "ADF009");
                        cargarADFCantidad(dataGridView12, orden, "ADF009");
                        dataGridView12.Columns[dataGridView12.Columns.Count - 1].DefaultCellStyle.Format = "c";
                        label11.Text = "Total: " + getTotal(dataGridView12);
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF010"))
                    {
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF011"))
                    {
                        cargarADF(dataGridView13, orden, "ADF011");
                        cargarADFCantidad(dataGridView13, orden, "ADF011");
                        dataGridView13.Columns[dataGridView13.Columns.Count - 1].DefaultCellStyle.Format = "c";
                        label12.Text = "Total: " + getTotal(dataGridView13);
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF012"))
                    {
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF013"))
                    {
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF014"))
                    {
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF015"))
                    {
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF016"))
                    {
                        cargarADFMant(dataGridView15, orden, "ADF016");
                        cargarADFCantidadMant(dataGridView15, orden, "ADF016");
                        dataGridView15.Columns[dataGridView15.Columns.Count - 1].DefaultCellStyle.Format = "c";
                        dataGridView15.Columns[dataGridView15.Columns.Count - 2].DefaultCellStyle.Format = "c";
                        label10.Text = "Total: " + getTotalMant(dataGridView15);
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF017"))
                    {
                        cargarADFMant(dataGridView16, orden, "ADF017");
                        cargarADFCantidadMant(dataGridView16, orden, "ADF017");
                        dataGridView16.Columns[dataGridView16.Columns.Count - 1].DefaultCellStyle.Format = "c";
                        dataGridView16.Columns[dataGridView16.Columns.Count - 2].DefaultCellStyle.Format = "c";
                        label13.Text = "Total: " + getTotalMant(dataGridView16);
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF018"))
                    {
                        cargarADFMant(dataGridView17, orden, "ADF018");
                        cargarADFCantidadMant(dataGridView17, orden, "ADF018");
                        dataGridView17.Columns[dataGridView17.Columns.Count - 1].DefaultCellStyle.Format = "c";
                        dataGridView17.Columns[dataGridView17.Columns.Count - 2].DefaultCellStyle.Format = "c";
                        label14.Text = "Total: " + getTotalMant(dataGridView17);
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF019"))
                    {
                    }
                }
            }
        }


        public bool registroExiste(int orden)
        {
            string query = "SELECT * FROM registroAlmacenOrden WHERE Orden = " + orden;
            //Ejecutar el query y llenar el GridView.
            OleDbConnection conn = new OleDbConnection();
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

        public int getCostoData(DataGridView data)
        {
            int costo = 0;
            for (int i = 0; i < data.Rows.Count; i++)
            {
                costo += Int32.Parse(data.Rows[i].Cells[data.Columns.Count - 1].Value.ToString());
            }
            return costo;
        }

        public void cargarCerrado(int orden, DataGridView data)
        {
            OleDbConnection conn = new OleDbConnection();
            string query = "SELECT a.Detalle,a.Modelo,SUM(a.Lunes+a.Martes+a.Miercoles+a.Jueves+a.Viernes+a.Sabado),i.Tipo FROM almacenOrden As a INNER JOIN Insumos AS i ON a.Modelo = i.ID WHERE a.Orden = " + orden + " GROUP BY a.Modelo,a.Detalle,i.Tipo";
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
                    data.Rows[i].Cells[3].Value = myReader.GetInt32(1);
                    data.Rows[i].Cells[4].Value = myReader.GetValue(2).ToString();
                    //if (myReader.GetString(3).Equals("Prestable"))
                        data.Rows[i].Cells[5].Value = myReader.GetValue(2).ToString();
                    //else
                    //    data.Rows[i].Cells[5].Value = 0;
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

        public void cargarCerradoExiste(int orden, DataGridView data)
        {
            OleDbConnection conn = new OleDbConnection();
            string query = "SELECT a.*,i.Tipo FROM registroAlmacenOrden As a INNER JOIN Insumos As i ON a.Modelo = i.ID WHERE a.Orden = " + orden;
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
                    data.Rows[i].Cells[2].Value = myReader.GetString(2);
                    data.Rows[i].Cells[3].Value = myReader.GetInt32(3);
                    data.Rows[i].Cells[4].Value = myReader.GetValue(4).ToString();
                    if (myReader.GetString(6).Equals("Prestable"))
                        data.Rows[i].Cells[5].Value = myReader.GetValue(5).ToString();
                    else
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

        public void formatoCerrado(DataGridView data, string query, string display)
        {
            DataGridViewComboBoxColumn combo = new DataGridViewComboBoxColumn();
            combo.HeaderText = "Modelo";
            data.Columns.Add("Column1", "#");
            data.Columns[0].ReadOnly = true;
            data.Columns.Add("Column2", "ID");
            data.Columns[1].Visible = false;
            data.Columns.Add("Column3", "Detalle");
            data.Columns[2].ReadOnly = true;
            Variables.cargar(combo, query, display, data);
            data.Columns.Add(combo);
            data.Columns[3].ReadOnly = true;
            data.Columns.Add("Column4", "Cant. Entregados");
            data.Columns[4].ReadOnly = true;
            data.Columns.Add("Column4", "Cant. Recibidos");
            data.Columns[0].FillWeight = 40;
            data.Columns[2].FillWeight = 200;
            data.Columns[3].FillWeight = 300;
        }

        public void cargarADFDaños(DataGridView data, int orden, string adf)
        {
            while (data.Rows.Count != 0)
            {
                data.Rows.RemoveAt(0);
            }
            string query = "SELECT Detalle, Descripcion, SUM(Lunes+Martes+Miercoles+Jueves+Viernes+Sabado) FROM Daños WHERE ADF = '" + adf + "' AND Orden = " + orden + " GROUP BY Detalle, Descripcion";
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

        public void cargarADF2(DataGridView data, int orden, string adf)
        {
            while (data.Rows.Count != 0)
            {
                data.Rows.RemoveAt(0);
            }
            string query = "SELECT c.Modelo, c.Detalle, (i.Tipo + ' ' + i.Marca + ' ' + i.Modelo), c.Unidad FROM Control As c INNER JOIN Maquinarias As i ON c.Modelo = i.ID WHERE c.Orden = " + orden + " AND ADF = '" + adf + "' GROUP BY c.Modelo, c.Detalle, (i.Tipo + ' ' + i.Marca + ' ' + i.Modelo), c.Unidad";
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

        public void cargarADF2Cantidad(DataGridView data, int orden, string adf)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                string query = "SELECT SUM(Lunes + Martes + Miercoles + Jueves + Viernes + Sabado) FROM Control WHERE Orden = " + orden + " AND ADF = '" + adf + "' AND Modelo = " + data.Rows[i].Cells[1].Value.ToString() + " AND Detalle = '" + data.Rows[i].Cells[2].Value.ToString() + "'";
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

        public void cargarEmpleados(DataGridView data, int orden)
        {
            while (data.Rows.Count != 0)
            {
                data.Rows.RemoveAt(0);
            }
            string query = "SELECT t.ID, (t.Nombres + ' ' + t.Apellidos), t.Cedula, c.Cargo FROM CargoLaboral AS c INNER JOIN (Trabajadores AS t INNER JOIN ordenEmpleados AS s ON t.ID = s.Trabajador) ON c.ID = t.Cargo WHERE s.Orden = " + orden;
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
                    data.Rows[i].Cells[3].Value = myReader.GetInt32(2);
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

        public int getSalarioDiario(int id)
        {
            string query = "SELECT c.salario FROM CargoLaboral AS c INNER JOIN Trabajadores AS t ON c.ID = t.Cargo WHERE t.ID = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int salario = 0;
            try
            {
                while (myReader.Read())
                {
                    salario = myReader.GetInt32(0);
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            if (salario > 0)
                salario = salario / 30;
            return salario;
        }

        public void cargarADFJornal(DataGridView data, int orden)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                string query = "SELECT SUM(Lunes + Martes + Miercoles + Jueves + Viernes + Sabado + Domingo) FROM adf002 WHERE Orden = " + orden + " AND Trabajador = " + data.Rows[i].Cells[1].Value.ToString();
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
                        diasTotales += Int32.Parse(myReader.GetValue(0).ToString());
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

        public void cargarCostosDiarios(DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                data.Rows[i].Cells[6].Value = Int32.Parse(data.Rows[i].Cells[5].Value.ToString()) * getSalarioDiario(Int32.Parse(data.Rows[i].Cells[1].Value.ToString()));
            }
        }

        public void cargarCantidadEquipoADF(DataGridView data, int orden, string adf) 
        {
            int i = 0;
            string query = "SELECT SUM(Lunes + Martes + Miercoles + Jueves + Viernes + Sabado) FROM Control WHERE Orden = " + orden + " AND ADF = '" + adf + "' AND Modelo = " + data.Rows[i].Cells[1].Value.ToString() + " AND Detalle = '" + data.Rows[i].Cells[2].Value.ToString() + "'";
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

        public void cargarCantidadEquipoADF2(DataGridView data, int orden, string adf)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                string query = "SELECT SUM(Lunes + Martes + Miercoles + Jueves + Viernes + Sabado) FROM Control WHERE Orden = " + orden + " AND ADF = '" + adf + "' AND Modelo = " + data.Rows[i].Cells[1].Value.ToString() + " AND Detalle = '" + data.Rows[i].Cells[2].Value.ToString() + "'";
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

        public void cargarADFCantidad(DataGridView data, int orden, string adf)
        {
            int index = 0;
            if (adf.Equals("ADF006"))
            {
                cargarCantidadEquipoADF(data, orden, "ADF006-1");
                index++;
            }
            for (int i = index; i < data.Rows.Count; i++)
            {
                string query = "SELECT SUM(Lunes + Martes + Miercoles + Jueves + Viernes + Sabado) FROM Control WHERE Orden = " + orden + " AND ADF = '" + adf + "' AND Modelo = " + data.Rows[i].Cells[1].Value.ToString() + " AND Detalle = '" + data.Rows[i].Cells[2].Value.ToString() + "'";
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

        public void cargarADFCantidadMant(DataGridView data, int orden, string adf)
        {
            int index = 0;
            for (int i = index; i < data.Rows.Count; i++)
            {
                string query = "SELECT SUM(Cantidad) FROM formatoMantenimiento WHERE Orden = " + orden + " AND ADF = '" + adf + "' AND Insumo = " + data.Rows[i].Cells[1].Value.ToString() + " AND Detalle = '" + data.Rows[i].Cells[3].Value.ToString() + "'" + " AND Tipo = '" + data.Rows[i].Cells[2].Value.ToString() + "'";
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
                            data.Rows[i].Cells[6].Value = myReader.GetValue(0).ToString();
                            var cultureInfo = new System.Globalization.CultureInfo("en-US");
                            int valor = Int32.Parse(data.Rows[i].Cells[7].Value.ToString(), cultureInfo);
                            data.Rows[i].Cells[8].Value = valor * Int32.Parse(myReader.GetValue(0).ToString());
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


        public string getTotal(DataGridView data)
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

        public string getTotalMant(DataGridView data)
        {
            string total = "";
            int valor = 0;
            for (int i = 0; i < data.Rows.Count; i++)
            {
                valor += Int32.Parse(data.Rows[i].Cells[8].Value.ToString());
            }
            total = String.Format("{0:c}", valor);
            return total;
        }

        public void cargarEquipoADF(DataGridView data, int orden, string adf)
        {
            string query = "SELECT c.Modelo, c.Detalle, (i.Tipo + ' ' + i.Marca + ' ' + i.Modelo), c.Unidad FROM Control As c INNER JOIN Maquinarias As i ON c.Modelo = i.ID WHERE c.Orden = " + orden + " AND ADF = '" + adf + "' GROUP BY c.Modelo, c.Detalle, (i.Tipo + ' ' + i.Marca + ' ' + i.Modelo), c.Unidad";
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

        public void cargarADF(DataGridView data, int orden, string adf)
        {
            while (data.Rows.Count != 0)
            {
                data.Rows.RemoveAt(0);
            }
            int i = 0;
            if (adf.Equals("ADF006"))
            {
                cargarEquipoADF(data, orden, "ADF006-1");
                i++;
            }
            string query = "SELECT c.Modelo, c.Detalle, (i.Clase + ' ' + i.Marca + ' ' + i.Modelo), c.Unidad, i.Costo_Unitario FROM Control As c INNER JOIN Insumos As i ON c.Modelo = i.ID WHERE c.Orden = " + orden + " AND ADF = '" + adf + "' GROUP BY c.Modelo, c.Detalle, (i.Clase + ' ' + i.Marca + ' ' + i.Modelo), c.Unidad, i.Costo_Unitario";
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

        public void cargarADFMant(DataGridView data, int orden, string adf)
        {
            while (data.Rows.Count != 0)
            {
                data.Rows.RemoveAt(0);
            }
            int i = 0;
            string query = "SELECT c.Insumo, c.Tipo, c.Detalle, (i.Clase + ' ' + i.Marca + ' ' + i.Modelo), i.Unidad_Medida, i.Costo_Unitario FROM formatoMantenimiento As c INNER JOIN Insumos As i ON c.Insumo = i.ID WHERE c.Orden = " + orden + " AND ADF = '" + adf + "' GROUP BY c.Insumo, c.Tipo, c.Detalle, (i.Clase + ' ' + i.Marca + ' ' + i.Modelo), i.Unidad_Medida, i.Costo_Unitario";
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
                    if (myReader.GetInt32(0) != 0)
                        data.Rows[i].Cells[4].Value = myReader.GetValue(3).ToString();
                    else
                        data.Rows[i].Cells[4].Value = "";
                    data.Rows[i].Cells[5].Value = myReader.GetString(4);
                    //data.Rows[i].Cells[5].Value = myReader.GetValue(4).ToString();
                    data.Rows[i].Cells[7].Value = String.Format("{0:c}", myReader.GetDouble(5).ToString());
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

        public void hideFormatos()
        {
            for (int i = 1; i < tabControl1.TabPages.Count; i++)
            {
                ((Control)this.tabControl1.TabPages[i]).Enabled = false;
            }
        }

        public void eliminarFormatos()
        {
            for (int i = 1; i < tabControl1.TabPages.Count; i++)
            {
                if (((Control)this.tabControl1.TabPages[i]).Enabled == false)
                {
                    if (tabControl1.HasChildren)
                    {
                        tabControl1.TabPages.RemoveAt(i);
                        i--;
                    }
                    else
                    {
                        sw2 = false;
                    }
                }
            }
        }

        public void getFormatos(string orden)
        {
            string query = "SELECT f.Formato, f.Actividad FROM formatosActividad AS f INNER JOIN historicoOrdenes AS h ON f.Actividad = h.Actividad WHERE h.ID = " + orden + " ORDER BY h.ID desc";
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
                    for (int i = 0; i < tabControl1.TabPages.Count; i++)
                    {
                        if (tabControl1.TabPages[i].Text.Equals(myReader.GetString(0)))
                        {
                            ((Control)this.tabControl1.TabPages[i]).Enabled = true;
                        }
                    }
                }
                if (!myReader.HasRows)
                {
                    MessageBox.Show("La actividad no contiene formatos.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    this.Close();
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

            string query = "SELECT Costo, CostoJornal, costoFinal, costoJornalFinal FROM historicoOrdenes WHERE ID = " + orden;
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
                    costoEstipuladoInsumos = double.Parse(myReader.GetValue(0).ToString());
                    costoEstipuladoJornal = double.Parse(myReader.GetValue(1).ToString());
                    costoInsumos = double.Parse(myReader.GetValue(2).ToString());
                    costoJornal = double.Parse(myReader.GetValue(3).ToString());
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
