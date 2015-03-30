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
    public partial class frmResumenOrden : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        List<string> formatos = new List<string>();
        bool sw2 = true;
        double diasTotales = 0;
        double costoEstipuladoJornal, costoEstipuladoInsumos, costoJornal, costoInsumos;
        string nomOrden;

        public frmResumenOrden(int orden)
        {
            InitializeComponent();
            hideFormatos();
            getFormatos(orden.ToString());
            eliminarFormatos();
            if (sw2)
            {
                formatoCerrado(dataGridView14, "SELECT ID, (Clase + ' ' + Marca + ' ' + Modelo) As Insumo FROM Insumos", "Insumo");
                cargarCerrado(orden, dataGridView14);
                dataGridView14.Columns[4].Visible = false;
                getOrden(orden);
                this.Text = "Resumen de Orden #: " + nomOrden;
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
                        if(Variables.tipo == 1)
                            label2.Text ="Dias totales : " + diasTotales.ToString() + " dias    Costo: " + String.Format("{0:c}", (getCostoData(dataGridView3)));
                        else
                            label2.Text = "Dias totales : " + diasTotales.ToString();
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
                        label1.Text = "Total: " + getTotal(dataGridView10, 5);
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF008"))
                    {
                        cargarEquipoADF(dataGridView11, orden, "ADF008");
                        cargarCantidadEquipoADF2(dataGridView11, orden, "ADF008");
                        dataGridView11.Columns[7].Visible = false;
                        dataGridView11.Columns[6].Visible = false;
                        label15.Text = "Total: " + getTotal(dataGridView11,5);
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF009"))
                    {
                        cargarADF(dataGridView12, orden, "ADF009");
                        cargarADFCantidad(dataGridView12, orden, "ADF009");
                        dataGridView12.Columns[dataGridView12.Columns.Count - 1].DefaultCellStyle.Format = "c";
                        dataGridView12.Columns[7].Visible = false;
                        dataGridView12.Columns[6].Visible = false;
                        label11.Text = "Total: " + getTotal(dataGridView12, 5);
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF010"))
                    {
                        cargarADFSemilla(dataGridView18, orden);
                        cargarADFCantidadSemilla(dataGridView18, orden);
                        label16.Text = "Total: " + getTotal(dataGridView18, 6);
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
                        cargarADFTransferencia(dataGridView19, orden, "ADF013");
                        cargarADFCantidadTransferencia(dataGridView19, orden, "ADF013");
                        label17.Text = "Total: " + getTotal(dataGridView19, 9) + " m3";
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF014"))
                    {
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF015"))
                    {
                        cargarADFTransferencia(dataGridView20, orden, "ADF015");
                        cargarADFCantidadTransferencia(dataGridView20, orden, "ADF015");
                        label18.Text = "Total: " +getTotal(dataGridView20, 9) + " m3";
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
                    if (Variables.tipo != 1)
                    {
                        dataGridView3.Columns[dataGridView3.Columns.Count - 1].Visible = false;
                        groupBox3.Visible = false;
                        groupBox4.Visible = false;
                        ocultarPrecio(dataGridView1, label4);
                        ocultarPrecio(dataGridView2, label5);
                        ocultarPrecio(dataGridView4, label3);
                        ocultarPrecio(dataGridView9, label9);
                        ocultarPrecio(dataGridView13, label12);
                        ocultarPrecio(dataGridView15, label10);
                        ocultarPrecio(dataGridView16, label13);
                        ocultarPrecio(dataGridView17, label14);
                        linkLabel1.Enabled = false;
                    }
                }
            }
        }

        public void ocultarPrecio(DataGridView data, Label label){
            data.Columns[data.Columns.Count - 1].Visible = false;
            data.Columns[data.Columns.Count - 2].Visible = false;
            label.Visible = false;
        }

        public void cargarCostoEquipoADF(DataGridView data)
        {
            int i = 0;
            string query = "SELECT valorHora FROM Maquinarias WHERE ID = " + data.Rows[i].Cells[1].Value.ToString();
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
                    data.Rows[i].Cells[6].Value = myReader.GetValue(0);
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

        public double getCostoData(DataGridView data)
        {
            double costo = 0;
            for (int i = 0; i < data.Rows.Count; i++)
            {
                costo += double.Parse(data.Rows[i].Cells[data.Columns.Count - 1].Value.ToString());
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
                        diasTotales += double.Parse(myReader.GetValue(0).ToString());
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
                data.Rows[i].Cells[6].Value = double.Parse(data.Rows[i].Cells[5].Value.ToString()) * getSalarioDiario(Int32.Parse(data.Rows[i].Cells[1].Value.ToString()));
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
                    double valor = double.Parse(data.Rows[i].Cells[6].Value.ToString(), cultureInfo);
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
                        double valor = double.Parse(data.Rows[i].Cells[6].Value.ToString(), cultureInfo);
                        data.Rows[i].Cells[7].Value = valor * double.Parse(myReader.GetValue(0).ToString());
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
                cargarCostoEquipoADF(data);
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

        public void cargarADFCantidadSemilla(DataGridView data, int orden)
        {
            int index = 0;
            for (int i = index; i < data.Rows.Count; i++)
            {
                string query = "SELECT SUM(Lunes + Martes + Miercoles + Jueves + Viernes + Sabado) FROM formatoSemilla WHERE Orden = " + orden + " AND Trabajador = " + data.Rows[i].Cells[1].Value.ToString() + " AND Semilla = " + data.Rows[i].Cells[4].Value.ToString();
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

        public void cargarADFCantidadTransferencia(DataGridView data, int orden, string ADF)
        {
            int index = 0;
            for (int i = index; i < data.Rows.Count; i++)
            {
                string query = "SELECT COUNT(ID) FROM Transferencia WHERE Orden = " + orden + " AND Especie = '" + data.Rows[i].Cells[2].Value + "' AND motivoRaleo = '" + data.Rows[i].Cells[3].Value + "' AND Trailer = '" + data.Rows[i].Cells[4].Value + "' AND Diametro = " + data.Rows[i].Cells[5].Value + " AND Largo = " + data.Rows[i].Cells[6].Value + " AND Cantidad = " + data.Rows[i].Cells[7].Value + " AND Volumen = " + data.Rows[i].Cells[8].Value.ToString().Replace(",",".") + " AND ADF = '" + ADF + "'";
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
                            data.Rows[i].Cells[1].Value = myReader.GetValue(0).ToString();
                            data.Rows[i].Cells[9].Value = double.Parse(myReader.GetValue(0).ToString()) * double.Parse(data.Rows[i].Cells[8].Value.ToString());
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
                            double valor = double.Parse(data.Rows[i].Cells[7].Value.ToString(), cultureInfo);
                            data.Rows[i].Cells[8].Value = valor * double.Parse(myReader.GetValue(0).ToString());
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
            double valor = 0;
            for (int i = 0; i < data.Rows.Count; i++)
            {
                valor += double.Parse(data.Rows[i].Cells[7].Value.ToString());
            }
            total = String.Format("{0:c}", valor);
            return total;
        }

        public string getTotal(DataGridView data, int column)
        {
            string total = "";
            double valor = 0;
            for (int i = 0; i < data.Rows.Count; i++)
            {
                valor += double.Parse(data.Rows[i].Cells[column].Value.ToString());
            }
            total = valor.ToString();
            return total;
        }

        public string getTotalMant(DataGridView data)
        {
            string total = "";
            double valor = 0;
            for (int i = 0; i < data.Rows.Count; i++)
            {
                valor += double.Parse(data.Rows[i].Cells[8].Value.ToString());
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

        public void cargarADFSemilla(DataGridView data, int orden)
        {
            while (data.Rows.Count != 0)
            {
                data.Rows.RemoveAt(0);
            }
            int i = 0;
            string query = "SELECT c.Semilla, c.Trabajador, (t.Nombres + ' ' + t.Apellidos), t.cedula, (i.Clase + ' ' + i.Marca + ' ' + i.Modelo), i.Costo_Unitario FROM Trabajadores AS t INNER JOIN (Insumos AS i INNER JOIN formatoSemilla AS c ON i.ID = c.Semilla) ON t.ID = c.Trabajador WHERE c.Orden = " + orden + " GROUP BY t.Cedula, c.Semilla, c.Trabajador, (i.Clase + ' ' + i.Marca + ' ' + i.Modelo), i.Costo_Unitario, (t.Nombres + ' ' + t.Apellidos)";
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
                    data.Rows[i].Cells[1].Value = myReader.GetInt32(1);
                    data.Rows[i].Cells[2].Value = myReader.GetString(2);                    
                    data.Rows[i].Cells[3].Value = myReader.GetInt32(3);                                            
                    data.Rows[i].Cells[4].Value = myReader.GetValue(0).ToString();
                    data.Rows[i].Cells[5].Value = myReader.GetValue(4).ToString();
                    data.Rows[i].Cells[7].Value = String.Format("{0:c}", myReader.GetDouble(5));
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

        public void cargarADFTransferencia(DataGridView data, int orden, string adf)
        {
            while (data.Rows.Count != 0)
            {
                data.Rows.RemoveAt(0);
            }
            int i = 0;
            string query = "SELECT Especie, motivoRaleo, Trailer, Diametro, Largo, Cantidad, Volumen FROM Transferencia WHERE Orden = " + orden + " AND ADF = '" + adf + "' GROUP BY Especie, motivoRaleo, Trailer, Diametro, Largo, Cantidad, Volumen";
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
                    data.Rows[i].Cells[2].Value = myReader.GetString(0);
                    data.Rows[i].Cells[3].Value = myReader.GetString(1);
                    data.Rows[i].Cells[4].Value = myReader.GetString(2);                    
                    data.Rows[i].Cells[5].Value = myReader.GetInt32(3);                    
                    data.Rows[i].Cells[6].Value = myReader.GetInt32(4);
                    data.Rows[i].Cells[7].Value = myReader.GetInt32(5);
                    data.Rows[i].Cells[8].Value = myReader.GetDouble(6);
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
                        formatos.Add(tabControl1.TabPages[i].Text);
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

            string query = "SELECT Costo, CostoJornal, costoFinal, costoJornalFinal, OT FROM historicoOrdenes WHERE ID = " + orden;
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
                    nomOrden = myReader.GetString(4);
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

        public void imprimirFormato1(DataGridView data, string sheetName, Microsoft.Office.Interop.Excel.Application XcelApp, string label)
        {
            if (data.Rows.Count > 0)
            {
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                XcelApp.Worksheets.Add();
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XcelApp.Worksheets.get_Item(1);
                xlWorkSheet.Name = sheetName;
                Microsoft.Office.Interop.Excel.Range excelCellrange;
                for (int i = 1; i < data.Columns.Count + 1; i++)
                {
                    if (i == 1)
                    {
                        xlWorkSheet.Cells[2, i + 2] = data.Columns[i - 1].HeaderText;
                    }
                    if (i != 2 && i != 1)
                    {
                        xlWorkSheet.Cells[2, i + 1] = data.Columns[i - 1].HeaderText;
                    }
                }

                for (int i = 0; i < data.Rows.Count; i++)
                {
                    for (int j = 0; j < data.Columns.Count; j++)
                    {
                        if (j == 0)
                            xlWorkSheet.Cells[i + 3, j + 3] = data.Rows[i].Cells[j].Value.ToString();
                        if (j != 1 && j != 0 && j != data.Columns.Count-1)
                            xlWorkSheet.Cells[i + 3, j + 2] = data.Rows[i].Cells[j].Value.ToString();
                        if (j == data.Columns.Count-1)
                            xlWorkSheet.Cells[i + 3, j + 2] = String.Format("{0:c}",data.Rows[i].Cells[j].Value);

                        if (i == 0)
                        {
                            excelCellrange = xlWorkSheet.Range[xlWorkSheet.Cells[i + 2, 3], xlWorkSheet.Cells[i + 2, data.Columns.Count + 1]];
                            excelCellrange.Interior.Color = System.Drawing.Color.LightGreen;
                            excelCellrange.AutoFilter(1);
                            //excelCellrange.Interior.Color = System.Drawing.Color.Blue;
                            //excelCellrange.Font.Color = System.Drawing.Color.White;
                        }
                    }
                }
                xlWorkSheet.Cells[data.Rows.Count + 3, data.Columns.Count + 1] = label;
                excelCellrange = xlWorkSheet.Range[xlWorkSheet.Cells[2, 3], xlWorkSheet.Cells[data.Rows.Count + 2, data.Columns.Count + 1]];
                excelCellrange.EntireColumn.AutoFit();
                Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;
            }
        }

        public void imprimirFormato5(DataGridView data, string sheetName, Microsoft.Office.Interop.Excel.Application XcelApp, string label)
        {
            if (data.Rows.Count > 0)
            {
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                XcelApp.Worksheets.Add();
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XcelApp.Worksheets.get_Item(1);
                xlWorkSheet.Name = sheetName;
                Microsoft.Office.Interop.Excel.Range excelCellrange;
                for (int i = 1; i < data.Columns.Count + 1; i++)
                {
                        xlWorkSheet.Cells[2, i + 1] = data.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < data.Rows.Count; i++)
                {
                    for (int j = 0; j < data.Columns.Count; j++)
                    {
                            xlWorkSheet.Cells[i + 3, j + 2] = data.Rows[i].Cells[j].Value.ToString();
                        if (i == 0)
                        {
                            excelCellrange = xlWorkSheet.Range[xlWorkSheet.Cells[i + 2, 2], xlWorkSheet.Cells[i + 2, data.Columns.Count + 1]];
                            excelCellrange.Interior.Color = System.Drawing.Color.LightGreen;
                            excelCellrange.AutoFilter(1);
                            //excelCellrange.Interior.Color = System.Drawing.Color.Blue;
                            //excelCellrange.Font.Color = System.Drawing.Color.White;
                        }
                    }
                }
                xlWorkSheet.Cells[data.Rows.Count + 3, data.Columns.Count + 1] = label;
                excelCellrange = xlWorkSheet.Range[xlWorkSheet.Cells[2, 2], xlWorkSheet.Cells[data.Rows.Count + 2, data.Columns.Count + 1]];
                excelCellrange.EntireColumn.AutoFit();
                Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;
            }
        }


        public void imprimirFormato3(DataGridView data, string sheetName, Microsoft.Office.Interop.Excel.Application XcelApp, string label, int columns)
        {
            if (data.Rows.Count > 0)
            {
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                XcelApp.Worksheets.Add();
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XcelApp.Worksheets.get_Item(1);
                xlWorkSheet.Name = sheetName;
                Microsoft.Office.Interop.Excel.Range excelCellrange;
                for (int i = 1; i < columns; i++)
                {
                    if (i == 1)
                    {
                        xlWorkSheet.Cells[2, i + 2] = data.Columns[i - 1].HeaderText;
                    }
                    if (i != 2 && i != 1)
                    {
                        xlWorkSheet.Cells[2, i + 1] = data.Columns[i - 1].HeaderText;
                    }
                }

                for (int i = 0; i < data.Rows.Count; i++)
                {
                    for (int j = 0; j < columns - 1; j++)
                    {
                        if (j == 0)
                            xlWorkSheet.Cells[i + 3, j + 3] = data.Rows[i].Cells[j].Value.ToString();
                        if (j != 1 && j != 0)
                            xlWorkSheet.Cells[i + 3, j + 2] = data.Rows[i].Cells[j].Value.ToString();
                        if (i == 0)
                        {
                            excelCellrange = xlWorkSheet.Range[xlWorkSheet.Cells[i + 2, 3], xlWorkSheet.Cells[i + 2, columns]];
                            excelCellrange.Interior.Color = System.Drawing.Color.LightGreen;
                            excelCellrange.AutoFilter(1);
                            //excelCellrange.Interior.Color = System.Drawing.Color.Blue;
                            //excelCellrange.Font.Color = System.Drawing.Color.White;
                        }
                    }
                }
                xlWorkSheet.Cells[data.Rows.Count + 3, columns] = label;
                excelCellrange = xlWorkSheet.Range[xlWorkSheet.Cells[2, 3], xlWorkSheet.Cells[data.Rows.Count + 2, columns]];
                excelCellrange.EntireColumn.AutoFit();
                Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;
            }
        }

        public void imprimirFormato4(DataGridView data, string sheetName, Microsoft.Office.Interop.Excel.Application XcelApp, string label, int columns)
        {
            if (data.Rows.Count > 0)
            {
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                XcelApp.Worksheets.Add();
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XcelApp.Worksheets.get_Item(1);
                xlWorkSheet.Name = sheetName;
                Microsoft.Office.Interop.Excel.Range excelCellrange;
                for (int i = 1; i < columns; i++)
                {
                    if (i == 1)
                    {
                        xlWorkSheet.Cells[2, i + 2] = data.Columns[i - 1].HeaderText;
                    }
                    if (i != 2 && i != 1 && i != 5 && i !=6 && i != 7)
                    {
                        xlWorkSheet.Cells[2, i + 1] = data.Columns[i - 1].HeaderText;
                    }
                    if (i == 6 || i == 7)
                    {
                        xlWorkSheet.Cells[2, i] = data.Columns[i - 1].HeaderText;
                    }
                }

                for (int i = 0; i < data.Rows.Count; i++)
                {
                    for (int j = 0; j < columns - 1; j++)
                    {
                        if (j == 0 || j == 4)
                            xlWorkSheet.Cells[i + 3, j + 3] = data.Rows[i].Cells[j].Value.ToString();
                        if (j != 1 && j != 0 && j != 4 && j != 5 && j != 6)
                            xlWorkSheet.Cells[i + 3, j + 2] = data.Rows[i].Cells[j].Value.ToString();
                        if(j == 5 || j == 6)
                            xlWorkSheet.Cells[i + 3, j + 1] = data.Rows[i].Cells[j].Value.ToString();
                    }
                    if (i == 0)
                    {
                        excelCellrange = xlWorkSheet.Range[xlWorkSheet.Cells[i + 2, 3], xlWorkSheet.Cells[i + 2, columns-1]];
                        excelCellrange.Interior.Color = System.Drawing.Color.LightGreen;
                        excelCellrange.AutoFilter(1);
                        //excelCellrange.Interior.Color = System.Drawing.Color.Blue;
                        //excelCellrange.Font.Color = System.Drawing.Color.White;
                    }
                }
                xlWorkSheet.Cells[data.Rows.Count + 3, columns-1] = label;
                excelCellrange = xlWorkSheet.Range[xlWorkSheet.Cells[2, 3], xlWorkSheet.Cells[data.Rows.Count + 2, columns -1]];
                excelCellrange.EntireColumn.AutoFit();
                Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;
            }
        }

        
        public void imprimirFormato2(DataGridView data, DataGridView data2, DataGridView data3, string sheetName, Microsoft.Office.Interop.Excel.Application XcelApp, string label)
        {
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            XcelApp.Worksheets.Add();
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XcelApp.Worksheets.get_Item(1);
            xlWorkSheet.Name = sheetName;
            Microsoft.Office.Interop.Excel.Range excelCellrange;

            if (data.Rows.Count > 0)
            {
                xlWorkSheet.Cells[1, "D"] = "Control Insumos";
                for (int i = 1; i < data.Columns.Count + 1; i++)
                {
                    if (i == 1)
                    {
                        xlWorkSheet.Cells[2, i + 2] = data.Columns[i - 1].HeaderText;
                    }
                    if (i != 2 && i != 1)
                    {
                        xlWorkSheet.Cells[2, i + 1] = data.Columns[i - 1].HeaderText;
                    }
                }

                for (int i = 0; i < data.Rows.Count; i++)
                {
                    for (int j = 0; j < data.Columns.Count; j++)
                    {
                        if (j == 0)
                            xlWorkSheet.Cells[i + 3, j + 3] = data.Rows[i].Cells[j].Value.ToString();
                        if (j != 1 && j != 0 && j != data.Columns.Count - 1)
                            xlWorkSheet.Cells[i + 3, j + 2] = data.Rows[i].Cells[j].Value.ToString();
                        if (j == data.Columns.Count - 1)
                            xlWorkSheet.Cells[i + 3, j + 2] = String.Format("{0:c}", data.Rows[i].Cells[j].Value);

                        if (i == 0)
                        {
                            excelCellrange = xlWorkSheet.Range[xlWorkSheet.Cells[i + 2, 3], xlWorkSheet.Cells[i + 2, data.Columns.Count + 1]];
                            excelCellrange.Interior.Color = System.Drawing.Color.LightGreen;
                            excelCellrange.AutoFilter(1);
                            //excelCellrange.Interior.Color = System.Drawing.Color.Blue;
                            //excelCellrange.Font.Color = System.Drawing.Color.White;
                        }
                    }
                }
                xlWorkSheet.Cells[data.Rows.Count + 3, data.Columns.Count + 1] = label;
                excelCellrange = xlWorkSheet.Range[xlWorkSheet.Cells[2, 3], xlWorkSheet.Cells[data.Rows.Count + 2, data.Columns.Count + 1]];
                excelCellrange.EntireColumn.AutoFit();
                Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;
            }

            if (data2.Rows.Count > 0)
            {
                xlWorkSheet.Cells[data.Rows.Count + 4, "D"] = "Reporte de Produción";
                for (int i = 1; i < data2.Columns.Count + 1; i++)
                {
                    if (i == 1)
                    {
                        xlWorkSheet.Cells[data.Rows.Count + 5, i + 2] = data2.Columns[i - 1].HeaderText;
                    }
                    if (i != 2 && i != 1)
                    {
                        xlWorkSheet.Cells[data.Rows.Count + 5, i + 1] = data2.Columns[i - 1].HeaderText;                        
                    }
                }

                for (int i = 0; i < data2.Rows.Count; i++)
                {
                    for (int j = 0; j < data2.Columns.Count; j++)
                    {
                        if (j == 0)
                            xlWorkSheet.Cells[data.Rows.Count + i + 6, j + 3] = data2.Rows[i].Cells[j].Value.ToString();
                        if (j != 1 && j != 0)
                            xlWorkSheet.Cells[data.Rows.Count + i + 6, j + 2] = data2.Rows[i].Cells[j].Value.ToString();
                        if (i == 0)
                        {
                            excelCellrange = xlWorkSheet.Range[xlWorkSheet.Cells[data.Rows.Count + i + 5, 3], xlWorkSheet.Cells[data.Rows.Count + i + 5, data2.Columns.Count + 1]];
                            excelCellrange.Interior.Color = System.Drawing.Color.LightGreen;
                            excelCellrange.AutoFilter(1);
                            //excelCellrange.Interior.Color = System.Drawing.Color.Blue;
                            //excelCellrange.Font.Color = System.Drawing.Color.White;
                        }
                    }
                }
                excelCellrange = xlWorkSheet.Range[xlWorkSheet.Cells[data.Rows.Count + 5, 3], xlWorkSheet.Cells[data.Rows.Count + data2.Rows.Count + 5, data2.Columns.Count + 1]];
                excelCellrange.EntireColumn.AutoFit();
                Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;
            }

            if (data3.Rows.Count > 0)
            {
                xlWorkSheet.Cells[data2.Rows.Count + data.Rows.Count + 7, "D"] = "Reporte de Daños";
                for (int i = 1; i < data3.Columns.Count + 1; i++)
                {
                    if (i == 1)
                    {
                        xlWorkSheet.Cells[data2.Rows.Count + data.Rows.Count + 8, i + 2] = data3.Columns[i - 1].HeaderText;
                    }
                    if (i != 2 && i != 1)
                    {
                        xlWorkSheet.Cells[data2.Rows.Count + data.Rows.Count + 8, i + 1] = data3.Columns[i - 1].HeaderText;
                    }
                }

                for (int i = 0; i < data3.Rows.Count; i++)
                {
                    for (int j = 0; j < data3.Columns.Count; j++)
                    {
                        if (j == 0)
                            xlWorkSheet.Cells[data2.Rows.Count + data.Rows.Count + i + 9, j + 3] = data3.Rows[i].Cells[j].Value.ToString();
                        if (j != 1 && j != 0)
                            if (j != 1 && j != 0)
                                if (j != 1 && j != 0)
                            xlWorkSheet.Cells[data2.Rows.Count + data.Rows.Count + i + 9, j + 2] = data3.Rows[i].Cells[j].Value.ToString();
                        if (i == 0)
                        {
                            excelCellrange = xlWorkSheet.Range[xlWorkSheet.Cells[data2.Rows.Count + data.Rows.Count + i + 8, 3], xlWorkSheet.Cells[data2.Rows.Count + data.Rows.Count + i + 8, data3.Columns.Count + 1]];
                            excelCellrange.Interior.Color = System.Drawing.Color.LightGreen;
                            excelCellrange.AutoFilter(1);
                            //excelCellrange.Interior.Color = System.Drawing.Color.Blue;
                            //excelCellrange.Font.Color = System.Drawing.Color.White;
                        }
                    }
                }
                excelCellrange = xlWorkSheet.Range[xlWorkSheet.Cells[data2.Rows.Count + data.Rows.Count + 8, 3], xlWorkSheet.Cells[data3.Rows.Count + data2.Rows.Count + data.Rows.Count + 8, data3.Columns.Count + 1]];
                excelCellrange.EntireColumn.AutoFit();
                Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;
            }
        }
        
        public void imprimirResumen()
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Aplicativo\\Formatos");
            Microsoft.Office.Interop.Excel.Application XcelApp = new Microsoft.Office.Interop.Excel.Application();
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Aplicativo\\Formatos\\", "Resumen*");
            XcelApp.Application.Workbooks.Add(prueba[0]);
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XcelApp.Worksheets.get_Item(1);
            xlWorkSheet.Cells[6, "E"] = String.Format("{0:c}", (costoEstipuladoInsumos - costoEstipuladoJornal));
            xlWorkSheet.Cells[6, "I"] = String.Format("{0:c}", (costoEstipuladoJornal));
            xlWorkSheet.Cells[6, "M"] = String.Format("{0:c}", (costoEstipuladoInsumos));
            xlWorkSheet.Cells[11, "E"] = String.Format("{0:c}", costoInsumos);
            xlWorkSheet.Cells[11, "I"] = String.Format("{0:c}", costoJornal);
            xlWorkSheet.Cells[11, "M"] = String.Format("{0:c}", (costoInsumos + costoJornal));
            XcelApp.Visible = true;
            for (int i = 0; i < tabControl1.TabPages.Count; i++)
            {
                if (tabControl1.TabPages[i].Text.Equals("ADF002"))
                {
                    imprimirFormato1(dataGridView3, "ADF002", XcelApp,label2.Text);
                }
                else if (tabControl1.TabPages[i].Text.Equals("ADF003"))
                {
                    imprimirFormato1(dataGridView1, "ADF003", XcelApp, label4.Text);
                }
                else if (tabControl1.TabPages[i].Text.Equals("ADF004"))
                {
                    imprimirFormato1(dataGridView2, "ADF004", XcelApp, label5.Text);
                }
                else if (tabControl1.TabPages[i].Text.Equals("ADF005"))
                {
                    imprimirFormato2(dataGridView4, dataGridView5, dataGridView6, "ADF005", XcelApp, label3.Text);
                }
                else if (tabControl1.TabPages[i].Text.Equals("ADF006"))
                {
                    imprimirFormato2(dataGridView9, dataGridView8, dataGridView7, "ADF006", XcelApp, label9.Text);
                }
                else if (tabControl1.TabPages[i].Text.Equals("ADF007"))
                {
                    imprimirFormato3(dataGridView10, "ADF007", XcelApp, label1.Text, dataGridView10.Columns.Count - 1);
                }
                else if (tabControl1.TabPages[i].Text.Equals("ADF008"))
                {
                    imprimirFormato3(dataGridView11, "ADF008", XcelApp, label15.Text, dataGridView11.Columns.Count - 1);
                }
                else if (tabControl1.TabPages[i].Text.Equals("ADF009"))
                {
                    imprimirFormato3(dataGridView12, "ADF009", XcelApp, label11.Text, dataGridView12.Columns.Count - 1);
                }
                else if (tabControl1.TabPages[i].Text.Equals("ADF010"))
                {
                    imprimirFormato4(dataGridView18, "ADF010", XcelApp, label16.Text, dataGridView18.Columns.Count);
                }
                else if (tabControl1.TabPages[i].Text.Equals("ADF011"))
                {
                    imprimirFormato1(dataGridView13, "ADF011", XcelApp, label12.Text);
                }
                else if (tabControl1.TabPages[i].Text.Equals("ADF012"))
                {
                }
                else if (tabControl1.TabPages[i].Text.Equals("ADF013"))
                {
                    imprimirFormato5(dataGridView19, "ADF013", XcelApp, label17.Text);
                }
                else if (tabControl1.TabPages[i].Text.Equals("ADF014"))
                {
                }
                else if (tabControl1.TabPages[i].Text.Equals("ADF015"))
                {
                    imprimirFormato5(dataGridView20, "ADF015", XcelApp, label18.Text);
                }
                else if (tabControl1.TabPages[i].Text.Equals("ADF016"))
                {
                    imprimirFormato1(dataGridView15, "ADF016", XcelApp, label10.Text);
                }
                else if (tabControl1.TabPages[i].Text.Equals("ADF017"))
                {
                    imprimirFormato1(dataGridView16, "ADF017", XcelApp, label13.Text);
                }
                else if (tabControl1.TabPages[i].Text.Equals("ADF018"))
                {
                    imprimirFormato1(dataGridView17, "ADF018", XcelApp, label14.Text);
                }
                else if (tabControl1.TabPages[i].Text.Equals("ADF019"))
                {
                }
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            imprimirResumen();
        }


    }
}
