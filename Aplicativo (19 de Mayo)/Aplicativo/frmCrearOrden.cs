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
    public partial class frmCrearOrden : Form
    {
        String connectionString = Variables.connectionString;
                OleDbConnection conn = new OleDbConnection();
        List<string> maquinarias = new List<string>();
        List<string> empleados = new List<string>();
        int tip, lote, actividad, tipousuario;
        double area = 0;
        int p = 0, a = 0, b = 0, c = 0;
        string user,unidad;
        double rodilla,cintura,cabeza,costo,costoJornal;
        string modificar,estado;
        string fechaExpedicion,costoFinal, nomOrden;
        int OT, semana;

        public void cargarDepartamentos()
        {
            string query = "SELECT * FROM Departamentos";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            txtDepartamento.DataSource = ds.Tables[0];
            txtDepartamento.DisplayMember = "Departamento";
            txtDepartamento.ValueMember = "ID";
            txtDepartamento.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtDepartamento.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void cargarContratos()
        {
            string query = "SELECT * FROM Contratos";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            txtContrato.DataSource = ds.Tables[0];
            txtContrato.DisplayMember = "Contrato";
            txtContrato.ValueMember = "ID";
            txtContrato.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtContrato.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void cargarActividades()
        {
            string query = "SELECT * FROM Actividades";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            txtActividad.DataSource = ds.Tables[0];
            txtActividad.DisplayMember = "Actividad";
            txtActividad.ValueMember = "ID";
            txtActividad.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtActividad.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void cargarActividades(string tipo)
        {
            if (conn.State == ConnectionState.Closed)
            {
                string query = "SELECT * FROM Actividades WHERE Tipo_Actividad = '" + tipo + "'";
                //Ejecutar el query y llenar el ComboBox.
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand(query, conn);
                DataTable maquinaria = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                txtActividad.DataSource = ds.Tables[0];
                txtActividad.DisplayMember = "Actividad";
                txtActividad.ValueMember = "ID";
                txtActividad.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                txtActividad.AutoCompleteSource = AutoCompleteSource.ListItems;
            }
        }

        public void cargar(ComboBox combo, string query, string display)
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
            combo.ValueMember = "ID";
            combo.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            combo.AutoCompleteSource = AutoCompleteSource.ListItems;
        }


        public void cargarCentro()
        {
            string query = "SELECT * FROM CentroDeCostos";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            txtCentroDeCostos.DataSource = ds.Tables[0];
            txtCentroDeCostos.DisplayMember = "Centro";
            txtCentroDeCostos.ValueMember = "ID";
            txtCentroDeCostos.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtCentroDeCostos.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void cargarSupervisores()
        {
            string query = "SELECT Trabajadores.ID As ID, (Trabajadores.Nombres + Trabajadores.Apellidos) As Nombre FROM Trabajadores INNER JOIN Usuarios ON Trabajadores.ID = Usuarios.Trabajador;";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            txtSupervisor.DataSource = ds.Tables[0];
            txtSupervisor.DisplayMember = "Nombre";
            txtSupervisor.ValueMember = "ID";
            txtSupervisor.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtSupervisor.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void cargarOperador()
        {
            string query = "SELECT * FROM Operador";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            txtOperador.DataSource = ds.Tables[0];
            txtOperador.DisplayMember = "Operador";
            txtOperador.ValueMember = "ID";
            txtOperador.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtOperador.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void cargarLotes()
        {
            string query = "SELECT Codigo, Lote FROM Lotes Group By Codigo,Lote UNION ALL SELECT Codigo, Lote FROM Areas Group By Codigo,Lote UNION ALL SELECT Codigo, Lote FROM LoteGanadero Group By Codigo,Lote";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            txtCodLote.DataSource = ds.Tables[0];
            txtCodLote.DisplayMember = "Codigo";
            txtCodLote.ValueMember = "Codigo";
            txtCodLote.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtCodLote.AutoCompleteSource = AutoCompleteSource.ListItems;
            txtLote.DataSource = ds.Tables[0];
            txtLote.DisplayMember = "Lote";
            txtLote.ValueMember = "Codigo";
            txtLote.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtLote.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void cargarMaquinaria()
        {
            string query = "SELECT ID, (Tipo + ' / ' + Modelo + ' / ' + Placa) As Maquina FROM Maquinarias";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            txtMaquinaria.DataSource = ds.Tables[0];
            txtMaquinaria.DisplayMember = "Maquina";
            txtMaquinaria.ValueMember = "ID";
            txtMaquinaria.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtMaquinaria.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void cargarEmpleados()
        {
            string query = "SELECT ID, (Nombres + '  ' + Apellidos) As nombre FROM Trabajadores";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            txtEmpleado.DataSource = ds.Tables[0];
            txtEmpleado.DisplayMember = "nombre";
            txtEmpleado.ValueMember = "ID";
            txtEmpleado.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtEmpleado.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public int getMaxID()
        {
            string query = "SELECT MAX(ID) FROM historicoOrdenes";
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

        public void getOrden(string orden){

            string query = "SELECT * FROM historicoOrdenes WHERE ID = " + orden;
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
                    OT = myReader.GetInt32(0);
                    fechaExpedicion = myReader.GetString(1);
                    label1.Text = "Fecha de Expedición: " + myReader.GetString(1);
                    DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
                    DateTime date1 = DateTime.ParseExact(myReader.GetString(1), "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                    Calendar cal = dfi.Calendar;
                    semana = cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);
                    label8.Text = "Semana de Expidición: " + cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek).ToString();
                    txtContrato.SelectedValue = myReader.GetInt32(2);
                    txtOperador.SelectedValue = myReader.GetInt32(3);
                    txtDepartamento.SelectedValue = myReader.GetInt32(4);
                    txtSupervisor.SelectedValue = myReader.GetInt32(5);
                    dateTimePicker1.Value = DateTime.ParseExact(myReader.GetString(7),"dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                    dateTimePicker2.Value = DateTime.ParseExact(myReader.GetString(8), "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                    actividad = myReader.GetInt32(6);
                    estado = myReader.GetString(11);
                    txtCosto.Text = String.Format("{0:c}",myReader.GetInt32(13));
                    //txtEstado.Text = myReader.GetString(11);
                    txtCentroDeCostos.SelectedValue = myReader.GetInt32(12);
                    lote = myReader.GetInt32(9);
                    area = myReader.GetDouble(10);
                    txtCostoJornal.Text = myReader.GetInt32(15).ToString();
                    txtDescripcion.Text = myReader.GetString(16).ToString();
                    comboBox3.SelectedValue = myReader.GetInt32(17);
                    comboBox2.SelectedValue = myReader.GetInt32(18);
                    txtCostoFinal.Text = String.Format("{0:c}",myReader.GetInt32(19));
                    txtCostoJornalFinal.Text = String.Format("{0:c}",myReader.GetInt32(20));
                    costoFinal = String.Format("{0:c}", (double.Parse(txtCostoJornalFinal.Text, NumberStyles.Currency) + double.Parse(txtCostoFinal.Text, NumberStyles.Currency)));
                    label33.Text = "Total: " + String.Format("{0:c}", (double.Parse(txtCostoJornalFinal.Text,NumberStyles.Currency) + double.Parse(txtCostoFinal.Text,NumberStyles.Currency)));
                    nomOrden = myReader.GetString(21);
                    comboBox6.SelectedValue = myReader.GetInt32(22);
                    label2.Text = "Orden de Trabajo #: " + nomOrden;
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

        public frmCrearOrden(int orden,string usuario)
        {
            InitializeComponent();
            user = usuario;
            if (orden == 0)
            {
                radioButton2.Checked = true;
                cargarDepartamentos();
                cargarActividades();
                cargarCentro();
                cargarSupervisores();
                cargarContratos();
                cargarOperador();
                cargarLotes();
                cargarMaquinaria();
                cargarEmpleados();
                cargar(comboBox5, "SELECT ID,Nombre FROM Cuadrilla", "Nombre");
                cargar(comboBox4, "SELECT * FROM Proveedores", "Proveedor");
                cargar(comboBox3, "SELECT * FROM Clientes", "Cliente");
                cargar(comboBox2, "SELECT ID,(Nombres + ' ' + Apellidos) As Nombre FROM Transportadores", "Nombre");
                cargar(comboBox6, "SELECT historicoOrdenes.ID, (historicoOrdenes.OT+' / '+historicoOrdenes.fechaInicio+' / '+Lotes.Lote) AS OrdenT FROM (historicoOrdenes INNER JOIN Lotes ON historicoOrdenes.Lote = Lotes.Codigo) INNER JOIN Actividades ON historicoOrdenes.Actividad = Actividades.ID WHERE Actividades.Actividad LIKE '%Extraccion%'", "OrdenT");
                comboBox2.SelectedItem = null;
                comboBox3.SelectedItem = null;
                comboBox4.SelectedItem = null;
                comboBox5.SelectedItem = null;
                comboBox6.SelectedItem = null;
                label26.Visible = false;
                button6.Enabled = false;
                textBox1.Visible = false;
                txtSupervisor.SelectedItem = null;
                txtMaquinaria.SelectedItem = null;
                txtContrato.SelectedItem = null;
                txtEmpleado.SelectedItem = null;
                txtLote.SelectedItem = null;
                txtCodLote.SelectedItem = null;
                txtCentroDeCostos.SelectedItem = null;
                txtDepartamento.SelectedItem = null;
                txtActividad.SelectedItem = null;
                txtCentroDeCostos.SelectedItem = null;
                txtOperador.SelectedItem = null;
                comboBox1.Visible = false;
                label27.Visible = false;
                txtDescripcion.ReadOnly = false;
                comboBox2.Enabled = false;
                p = 1;
                txtCentroDeCostos.Text = "N/A";
                txtContrato.Text = "N/A";
                comboBox4.Text = "N/A";
                comboBox3.Text = "N/A";
                label1.Text += "  " + DateTime.Now.Day + "/" + DateTime.Now.Month + "/" + DateTime.Now.Year;
                DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
                DateTime date1 = DateTime.Now;
                Calendar cal = dfi.Calendar;
                label8.Text = "Semana de Expidición: " + cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek).ToString();
                //label2.Text += "  " + (getMaxID()+1) ;
                dateTimePicker1.Format = DateTimePickerFormat.Custom;
                dateTimePicker1.CustomFormat = "dd/MM/yyyy";
                dateTimePicker2.Format = DateTimePickerFormat.Custom;
                dateTimePicker2.CustomFormat = "dd/MM/yyyy";
            }
        }

        public frmCrearOrden(string orden, int tipo)
        {
            InitializeComponent();
            tipousuario = tipo;
            modificar = orden;
            cargarDepartamentos();
            cargarActividades();
            cargarCentro();
            cargarSupervisores();
            cargarContratos();
            cargarOperador();
            cargarLotes();
            cargarMaquinaria();
            cargarEmpleados();
            cargar(comboBox5, "SELECT ID,Nombre FROM Cuadrilla", "Nombre");
            cargar(comboBox4, "SELECT * FROM Proveedores", "Proveedor");
            cargar(comboBox3, "SELECT * FROM Clientes", "Cliente");
            cargar(comboBox2, "SELECT ID,(Nombres + ' ' + Apellidos) As Nombre FROM Transportadores", "Nombre");
            cargar(comboBox6, "SELECT historicoOrdenes.ID, (historicoOrdenes.OT+' / '+historicoOrdenes.fechaInicio+' / '+Lotes.Lote) AS OrdenT FROM (historicoOrdenes INNER JOIN Lotes ON historicoOrdenes.Lote = Lotes.Codigo) INNER JOIN Actividades ON historicoOrdenes.Actividad = Actividades.ID WHERE Actividades.Actividad LIKE '%Extraccion%'", "OrdenT");
            comboBox2.SelectedItem = null;
            comboBox3.SelectedItem = null;
            comboBox4.SelectedItem = null;
            comboBox5.SelectedItem = null;
            comboBox6.SelectedItem = null;
            txtLote.SelectedItem = null;
            txtCodLote.SelectedItem = null;
            txtActividad.SelectedItem = null;
            txtMaquinaria.SelectedItem = null;
            txtEmpleado.SelectedItem = null;
            comboBox1.Visible = false;
            label27.Visible = false;
            comboBox2.Enabled = false;
            getOrden(orden);
            txtCodLote.SelectedValue = lote;
            txtLote.SelectedValue = lote;
            txtActividad.SelectedValue = actividad;
            if (!txtArea.Text.Equals("")) { 
                if (area != Double.Parse(txtArea.Text))
                {
                    radioButton1.Checked = true;
                    txtAreaIntervenir.Text = area.ToString();
                }
                else
                {
                    radioButton2.Checked = true;
                }
            }
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = Application.CurrentCulture.DateTimeFormat.ShortDatePattern;
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = Application.CurrentCulture.DateTimeFormat.ShortDatePattern;
            getMaquinas(orden);
            getEmpleados(orden);
            this.Text = "Modificar Orden de Trabajo";
            button1.Text = "Modificar Formato";
            txtDescripcion.ReadOnly = false;
            tip = 5;            
            if (tipo == 1)
            {
                this.Text = "Orden de Trabajo";
                button1.Text = "Asignar Empleados a la Orden de Trabajo";
                txtOperador.Enabled = false;
                txtDepartamento.Enabled = false;
                txtSupervisor.Enabled = false;
                txtDescripcion.ReadOnly = true;
                comboBox3.Enabled = false;
                comboBox4.Enabled = false;
                label28.Visible = false;
                txtCostoJornal.Visible = false;
                txtCostoFinal.Visible = false;
                txtCostoJornalFinal.Visible = false;
                label33.Visible = false;
                label32.Visible = false;
                txtActividad.Enabled = false;
                txtCentroDeCostos.Enabled = false;
                txtTipoActividad.Enabled = false;
                dateTimePicker1.Enabled = false;
                dateTimePicker2.Enabled = false;
                txtMaquinaria.Enabled = false;
                txtCodLote.Enabled = false;
                txtCodLote2.Enabled = false;
                //button10.Visible = false;
                txtLote.Enabled = false;
                txtLote2.Enabled = false;
                txtPredio.Enabled = false;
                comboBox1.Enabled = false;
                textBox1.Enabled = false;
                txtArea.Enabled = false;
                txtAreaIntervenir.Enabled = false;
                radioButton2.Enabled = false;
                radioButton1.Enabled = false;
                txtEstado.Enabled = false;
                txtDescripcion.Enabled = false;
                txtCondiciones.Enabled = false;
                listBox1.Enabled = false;
                button2.Enabled = false;
                button3.Enabled = false;
                label25.Visible = false;
                label26.Visible = false;
                textBox1.Visible = false;
                txtCosto.Visible = false;
                txtContrato.Visible = false;
                button4.Enabled = false;
            }
            txtEstado.Text = estado;
            getActividad(txtActividad.SelectedValue.ToString());
            if (unidad.Equals("Ha"))
            {
                label26.Visible = false;
                textBox1.Visible = false;
                label23.Visible = true;
                radioButton1.Visible = true;
                radioButton2.Visible = true;
                label24.Visible = true;
                txtAreaIntervenir.Visible = true;
            }
            else
            {
                label26.Visible = true;
                label26.Text = "Cantidad de " + unidad + ":";
                textBox1.Visible = true;
                label23.Visible = false;
                radioButton1.Visible = false;
                radioButton2.Visible = false;
                label24.Visible = false;
                txtAreaIntervenir.Visible = false;
                textBox1.Text = area.ToString();
            }
            if (txtActividad.Text.Contains("Transporte"))
            {
                comboBox1.Text = estado;                
                label27.Visible = true;
                comboBox1.Visible = true;
                txtEstado.Visible = false;
            }
            if (txtActividad.Text.Contains("Mantenimiento"))
            {
                txtEstado.Visible = false;
                comboBox1.Visible = false;
                label27.Visible = false;
                label14.Visible = false;
            }
            if (txtActividad.Text.Contains("Transporte") || txtActividad.Text.Contains("Transporte"))
            {
                comboBox6.Visible = true;
                label35.Visible = true;
            }
            if (Variables.tipo == 1)
            {
                label25.Visible = true;
                label28.Visible = true;
                label32.Visible = true;
                label33.Visible = true;
                txtCosto.Visible = true;
                txtCostoJornalFinal.Visible = true;
                txtCostoFinal.Visible = true;
                txtCosto.ReadOnly = true;
                txtCostoJornalFinal.ReadOnly = true;
                txtCostoFinal.ReadOnly = true;
            }
        }

        public void agregarOrden(int id)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO historicoOrdenes (fechaExpedicion,Contrato,Operador,Departamento,Supervisor,Actividad,fechaInicio,fechaFinal,Lote,Area,estadoLote,Centro,Costo,estadoOrden,CostoJornal,Descripcion,Cliente,Transportador,costoFinal,costoJornalFinal,OT,transOT) VALUES (@fechaExpedicion,@Contrato,@Operador,@Departamento,@Supervisor,@Actividad,@fechaInicio,@fechaFinal,@Lote,@Area,@estadoLote,@Centro,@Costo,@estadoOrden,@CostoJornal,@Descripcion,@Cliente,@Transportador,@costoFinal,@costoJornalFinal,@OT,@transOT)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@fechaExpedicion", OleDbType.VarChar).Value = DateTime.Now.ToString("dd") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.Year;
                cmd.Parameters.Add("@Contrato", OleDbType.VarChar).Value = txtContrato.SelectedValue;
                cmd.Parameters.Add("@Operador", OleDbType.VarChar).Value = txtOperador.SelectedValue;
                cmd.Parameters.Add("@Departamento", OleDbType.VarChar).Value = txtDepartamento.SelectedValue;
                cmd.Parameters.Add("@Supervisor", OleDbType.VarChar).Value = txtSupervisor.SelectedValue;
                cmd.Parameters.Add("@Actividad", OleDbType.VarChar).Value = txtActividad.SelectedValue;
                cmd.Parameters.Add("@fechaInicio", OleDbType.VarChar).Value = dateTimePicker1.Value.ToString("dd") + "/" + dateTimePicker1.Value.ToString("MM") + "/" + dateTimePicker1.Value.Year;
                cmd.Parameters.Add("@fechaFinal", OleDbType.VarChar).Value = dateTimePicker2.Value.ToString("dd") + "/" + dateTimePicker2.Value.ToString("MM") + "/" + dateTimePicker2.Value.Year;
                cmd.Parameters.Add("@Lote", OleDbType.VarChar).Value = txtLote.SelectedValue;
                if (radioButton2.Checked)
                    if (textBox1.Visible == true)
                        cmd.Parameters.Add("@Area", OleDbType.VarChar).Value = textBox1.Text;
                    else
                        cmd.Parameters.Add("@Area", OleDbType.VarChar).Value = txtArea.Text;
                else
                    cmd.Parameters.Add("@Area", OleDbType.VarChar).Value = txtAreaIntervenir.Text;
                if (txtEstado.Visible)
                    cmd.Parameters.Add("@estadoLote", OleDbType.VarChar).Value = txtEstado.Text;
                else
                    cmd.Parameters.Add("@estadoLote", OleDbType.VarChar).Value = comboBox1.Text;
                cmd.Parameters.Add("@Centro", OleDbType.VarChar).Value = txtCentroDeCostos.SelectedValue;
                cmd.Parameters.Add("@Costo", OleDbType.VarChar).Value = txtCosto.Text;
                cmd.Parameters.Add("@estadoOrden", OleDbType.VarChar).Value = "Activa";
                cmd.Parameters.Add("@CostoJornal", OleDbType.VarChar).Value = txtCostoJornal.Text;
                cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = txtDescripcion.Text;
                cmd.Parameters.Add("@Cliente", OleDbType.VarChar).Value = comboBox3.SelectedValue;
                cmd.Parameters.Add("@Transportador", OleDbType.VarChar).Value = comboBox2.SelectedValue;
                cmd.Parameters.Add("@costoFinal", OleDbType.VarChar).Value = txtCostoFinal.Text;
                cmd.Parameters.Add("@costoJornalFinal", OleDbType.VarChar).Value = txtCostoJornalFinal.Text;
                cmd.Parameters.Add("@OT", OleDbType.VarChar).Value = "OT-" + id.ToString().PadLeft(4,'0') + "-" + DateTime.Now.Year;
                if (txtActividad.Text.Contains("Transporte") || txtActividad.Text.Contains("Transporte"))                
                    cmd.Parameters.Add("@transOT", OleDbType.VarChar).Value = comboBox6.SelectedValue;
                else
                    cmd.Parameters.Add("@transOT", OleDbType.VarChar).Value = 0;
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

        public void getEmpleados(string id)
        {
            listBox2.Items.Clear();
            empleados.Clear();
            string query = "SELECT i.ID, i.Orden, (m.Nombres + ' ' + m.Apellidos) As nombre, m.ID FROM Trabajadores AS m INNER JOIN ordenEmpleados AS i ON m.ID = i.Trabajador WHERE i.Orden = " + id;
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

        public void modificarOrden(string id)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE historicoOrdenes SET Contrato=@Contrato,Operador=@Operador,Departamento=@Departamento,Supervisor=@Supervisor,Actividad=@Actividad,fechaInicio=@fechaInicio,fechaFinal=@fechaFinal,Lote=@Lote,Area=@Area,estadoLote=@estadoLote,Centro=@Centro,Costo=@Costo,estadoOrden=@estadoOrden,CostoJornal=@CostoJornal,Descripcion=@Descripcion,Cliente=@Cliente,Transportador=@Transportador,costoFinal=@costoFinal,costoJornalFinal=@costoJornalFinal,transOT=@transOT WHERE ID = " + id);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Contrato", OleDbType.VarChar).Value = txtContrato.SelectedValue;
                cmd.Parameters.Add("@Operador", OleDbType.VarChar).Value = txtOperador.SelectedValue;
                cmd.Parameters.Add("@Departamento", OleDbType.VarChar).Value = txtDepartamento.SelectedValue;
                cmd.Parameters.Add("@Supervisor", OleDbType.VarChar).Value = txtSupervisor.SelectedValue;
                cmd.Parameters.Add("@Actividad", OleDbType.VarChar).Value = txtActividad.SelectedValue;
                cmd.Parameters.Add("@fechaInicio", OleDbType.VarChar).Value = dateTimePicker1.Value.ToString("dd") + "/" + dateTimePicker1.Value.ToString("MM") + "/" + dateTimePicker1.Value.Year;
                cmd.Parameters.Add("@fechaFinal", OleDbType.VarChar).Value = dateTimePicker2.Value.ToString("dd") + "/" + dateTimePicker2.Value.ToString("MM") + "/" + dateTimePicker2.Value.Year;
                cmd.Parameters.Add("@Lote", OleDbType.VarChar).Value = txtLote.SelectedValue;
                if (textBox1.Visible == true)
                        cmd.Parameters.Add("@Area", OleDbType.VarChar).Value = textBox1.Text;
                else
                {
                    if (radioButton2.Checked)
                        cmd.Parameters.Add("@Area", OleDbType.VarChar).Value = txtArea.Text;
                    else
                        cmd.Parameters.Add("@Area", OleDbType.VarChar).Value = txtAreaIntervenir.Text;
                }

                if (txtEstado.Visible)
                    cmd.Parameters.Add("@estadoLote", OleDbType.VarChar).Value = txtEstado.Text;
                else if(comboBox1.Visible)
                    cmd.Parameters.Add("@estadoLote", OleDbType.VarChar).Value = comboBox1.Text;
                else
                    cmd.Parameters.Add("@estadoLote", OleDbType.VarChar).Value = "N/A";
                cmd.Parameters.Add("@Centro", OleDbType.VarChar).Value = txtCentroDeCostos.SelectedValue;
                cmd.Parameters.Add("@Costo", OleDbType.VarChar).Value = txtCosto.Text;
                cmd.Parameters.Add("@estadoOrden", OleDbType.VarChar).Value = "Activa";
                cmd.Parameters.Add("@CostoJornal", OleDbType.VarChar).Value = txtCostoJornal.Text;
                cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = txtDescripcion.Text;
                cmd.Parameters.Add("@Cliente", OleDbType.VarChar).Value = comboBox3.SelectedValue;
                cmd.Parameters.Add("@Transportador", OleDbType.VarChar).Value = comboBox2.SelectedValue;
                cmd.Parameters.Add("@costoFinal", OleDbType.VarChar).Value = txtCostoFinal.Text;
                cmd.Parameters.Add("@costoJornalFinal", OleDbType.VarChar).Value = txtCostoJornalFinal.Text;
                if (txtActividad.Text.Contains("Transporte") || txtActividad.Text.Contains("Transporte"))
                    cmd.Parameters.Add("@transOT", OleDbType.VarChar).Value = comboBox6.SelectedValue;
                else
                    cmd.Parameters.Add("@transOT", OleDbType.VarChar).Value = 0;
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

        public void agregarOrdenMaquina(int id)
        {
            for (int i = 0; i < maquinarias.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("INSERT INTO ordenMaquinas(Orden,Maquina) VALUES (@Orden,@Maquina)");
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@Orden", OleDbType.VarChar).Value = id;
                    cmd.Parameters.Add("@Maquina", OleDbType.VarChar).Value = maquinarias[i];
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

        public void eliminarOrdenMaquina(string id)
        {
            for (int i = 0; i < maquinarias.Count+1; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM ordenMaquinas WHERE Orden = " + id);
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

        public void agregarOrdenEmpleados(int id)
        {
            for (int i = 0; i < empleados.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("INSERT INTO ordenEmpleados(Orden,Trabajador) VALUES (@Orden,@Trabajador)");
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

        public void eliminarOrdenEmpleados(string id)
        {
            for (int i = 0; i < empleados.Count+1; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM ordenEmpleados WHERE Orden = " + id);
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
            string query = "SELECT * FROM historicoOrdenes WHERE OT like '%" + DateTime.Now.Year + "%'";
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

        private void button1_Click(object sender, EventArgs e)
        {
            if (!txtActividad.Text.Equals(""))
            {
                if (!txtTipoActividad.Text.Equals(""))
                {
                    if (!txtOperador.Text.Equals(""))
                    {
                        if (!txtDepartamento.Text.Equals(""))
                        {
                            if (!txtSupervisor.Text.Equals(""))
                            {
                                if (!txtContrato.Text.Equals(""))
                                {
                                    if (!txtCodLote.Text.Equals("") || !txtLote.Text.Equals("")) {
                                        if(!comboBox3.Text.Equals("")){
                                            if (!comboBox2.Text.Equals(""))
                                            {
                                                if (button1.Text.Equals("Generar Orden"))
                                                {
                                                    int id = getMaxID();
                                                    int id2 = id;
                                                    if (!existeAñoActual())
                                                        id2 = 0;                                                    
                                                    agregarOrden(id2+1);
                                                    id++;
                                                    agregarOrdenMaquina(id);
                                                    agregarOrdenEmpleados(id);
                                                }
                                                else
                                                {
                                                    modificarOrden(modificar);
                                                    eliminarOrdenMaquina(modificar);
                                                    agregarOrdenMaquina(Int32.Parse(modificar));
                                                    eliminarOrdenEmpleados(modificar);
                                                    agregarOrdenEmpleados(Int32.Parse(modificar));
                                                }
                                                if (tipousuario == 0)
                                                {
                                                    frmHistoricoOrdenes newFrm = new frmHistoricoOrdenes(0, user);
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
                                                MessageBox.Show("Favor seleccionar el transportador, seleccionar Proveedor N/A si no aplica.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show("Favor seleccionar el cliente, seleccionar N/A si no aplica.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("Favor seleccionar lote", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Favor seleccionar contrato", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                }
                            }
                            else
                            {
                                MessageBox.Show("Favor seleccionar supervisor", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Favor seleccionar departamento.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Favor seleccionar operador", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
                else
                {
                    MessageBox.Show("Favor seleccionar tipo de actividad", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            else
            {
                MessageBox.Show("Favor seleccionar actividad.","Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (tipousuario == 0)
            {
                frmHistoricoOrdenes newFrm = new frmHistoricoOrdenes(0,user);
                this.Hide();
                newFrm.ShowDialog();
                this.Close();
            }
            else
            {
                this.Close();
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                txtAreaIntervenir.ReadOnly = false;
                if(costo != 0)
                    txtCosto.Text = String.Format("{0:c}",costo);
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                txtAreaIntervenir.ReadOnly = true;
                txtAreaIntervenir.Text = "";
                if (costo != 0 && !txtArea.Text.Equals(""))
                {
                    txtCosto.Text = String.Format("{0:c}",(costo * double.Parse(txtArea.Text)));
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (txtMaquinaria.Text.Equals(""))
            {
                MessageBox.Show("Favor seleccionar la maquinaria deseada.");
            }
            else
            {
                listBox1.Items.Add(txtMaquinaria.Text);
                maquinarias.Add(txtMaquinaria.SelectedValue.ToString());
                string idEmpleado = empleadoMaquina(txtMaquinaria.SelectedValue.ToString());
                if (!idEmpleado.Equals(""))
                {
                    txtEmpleado.SelectedValue = idEmpleado;
                    listBox2.Items.Add(txtEmpleado.Text);
                    empleados.Add(txtEmpleado.SelectedValue.ToString());
                    txtEmpleado.SelectedItem = null;
                }
            }
        }

        public string empleadoMaquina(string maquina) 
        {
            string idEmpleado = "";
            string query = "SELECT ID From Trabajadores WHERE Maquina = " + maquina;
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
                    idEmpleado = myReader.GetInt32(0).ToString();
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return idEmpleado;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (txtMaquinaria.Text.Equals(""))
            {
                MessageBox.Show("Favor seleccionar la maquinaria deseada.");
            }
            else
            {
                maquinarias.RemoveAt(listBox1.SelectedIndex);
                listBox1.Items.Remove(listBox1.SelectedItem);
            }
        }

        public void getLote(string codigo)
        {
                string query = "SELECT l.Codigo, l.Lote, l.areaEfectiva As Area, b.Predio FROM Lotes AS l INNER JOIN BancoTierras AS b ON l.Predio = b.ID WHERE l.Codigo = " + codigo + " Group By l.Codigo,l.Lote,l.areaEfectiva,b.Predio UNION ALL SELECT g.Codigo, g.Lote, g.Area, b.Predio FROM LoteGanadero AS g INNER JOIN BancoTierras AS b ON g.Predio = b.ID WHERE g.Codigo = " + codigo + " Group By g.Codigo,g.Lote,g.Area,b.Predio UNION ALL SELECT a.Codigo, a.Lote, a.Area, b.Predio FROM Areas AS a INNER JOIN BancoTierras AS b ON a.Predio = b.ID WHERE a.Codigo = " + codigo + " Group By a.Codigo,a.Lote,a.Area,b.Predio";
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
                        txtCodLote2.Text = myReader.GetInt32(0).ToString();
                        txtLote2.Text = myReader.GetString(1);
                        txtArea.Text = myReader.GetString(2);
                        txtPredio.Text = myReader.GetString(3);
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

        public void getActividad(string Actividad)
        {
            string query = "SELECT a.ID,a.Actividad, a.Unidad_de_Medida,SUM(i.Costo_Unitario*ia.Cantidad), d.Departamento,a.Tipo_Actividad,a.Condicion_Minima,a.Descripcion_Actividad,a.Rodilla,a.Cintura,a.Cabeza FROM (Actividades AS a INNER JOIN (Insumos AS i INNER JOIN insumoActividades AS ia ON i.ID = ia.Actividad) ON a.ID = ia.Insumo) INNER JOIN Departamentos AS d ON d.ID = a.Departamento WHERE a.ID = " + Actividad + " GROUP BY a.ID,a.Actividad, a.Unidad_de_Medida,d.Departamento,a.Tipo_Actividad,a.Condicion_Minima,a.Descripcion_Actividad,a.Rodilla,a.Cintura,a.Cabeza";
            //string query = "SELECT Tipo_Actividad,Condicion_Minima,Descripcion_Actividad FROM Actividades WHERE ID = " + Actividad;
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
                    unidad = myReader.GetString(2).ToString();
                    costo = double.Parse(myReader.GetValue(3).ToString());
                    txtTipoActividad.Text = myReader.GetString(5);                    
                    if(tip!=5)
                        txtDescripcion.Text = myReader.GetString(7);                    
                    tip++;
                    txtCondiciones.Text = myReader.GetString(6);
                    rodilla = Double.Parse(myReader.GetString(8));
                    cintura = Double.Parse(myReader.GetString(9));
                    cabeza = Double.Parse(myReader.GetString(10));
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

        public void getMaquinas(string id)
        {
            listBox1.Items.Clear();
            maquinarias.Clear();
            string query = "SELECT i.ID, i.Orden, (m.Tipo + ' /' + m.Modelo + '/' + m.Placa) As Maquinaria, m.ID FROM Maquinarias AS m INNER JOIN ordenMaquinas AS i ON m.ID = i.Maquina WHERE i.Orden = " + id;
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
                    listBox1.Items.Add(myReader.GetString(2));
                    maquinarias.Add(myReader.GetInt32(3).ToString());
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

        private void txtCodLote_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtCodLote.SelectedItem != null && !txtCodLote.SelectedValue.ToString().Equals("System.Data.DataRowView"))
            {                
                if (txtCodLote.Text.Equals("-1")){
                    txtLote.Text = "N/A";
                    txtPredio.Text = "N/A";
                    txtLote.Text = "N/A";
                    txtCodLote2.Text = "-1";
                    txtLote2.Text = "N/A";
                    txtArea.Text = "0";
                }                    
                else
                    getLote(txtCodLote.SelectedValue.ToString());
                if (costo != 0 && textBox1.Visible == false)
                {
                    txtCosto.Text = String.Format("{0:c}",(costo * double.Parse(txtArea.Text)));
                    txtCostoJornal.Text = (costoJornal * double.Parse(txtArea.Text)).ToString();
                    //label33.Text = "Total: " + String.Format("{0:c}", (Int32.Parse(txtCostoJornal.Text) + Int32.Parse(txtCostoFinal.Text)));
                    txtEstado.Refresh();
                }
            }
        }

        private void txtLote_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtLote.SelectedItem != null && !txtLote.SelectedValue.ToString().Equals("System.Data.DataRowView"))
                if (txtLote.Text.Equals("N/A"))
                {
                    txtCodLote.Text = "-1";
                    txtPredio.Text = "N/A";
                    txtLote.Text = "N/A";
                    txtCodLote2.Text = "-1";
                    txtLote2.Text = "N/A";
                    txtArea.Text = "0";
                }
                else
                    getLote(txtLote.SelectedValue.ToString());
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            DateTime date1 = dateTimePicker1.Value;
            Calendar cal = dfi.Calendar;
            label9.Text = "Semana de Inicio: " + cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek).ToString();
        }

        public void getCostoJornal(double tipo)
        {
            string query = "SELECT SUM(i.Costo_Unitario * s.Cantidad) As Costo FROM Insumos AS i INNER JOIN insumoActividades AS s ON i.ID = s.Actividad WHERE i.Marca = 'Jornal'  AND s.Insumo = " + txtActividad.SelectedValue + " GROUP BY i.Marca";
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
                    costoJornal = double.Parse(myReader.GetValue(0).ToString()) * tipo;
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


        private void txtActividad_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtActividad.SelectedItem != null && !txtActividad.SelectedValue.ToString().Equals("System.Data.DataRowView"))
            {
                if (p != 0)
                {
                    getActividad(txtActividad.SelectedValue.ToString());
                    txtCosto.Text = String.Format("{0:c}",costo);
                    getCostoJornal(1);
                    txtCostoJornal.Text = costoJornal.ToString();
                    if (txtActividad.Text.Contains("Transporte") || txtActividad.Text.Contains("Transporte"))
                    {
                        comboBox6.Visible = true;
                        label35.Visible = true;
                    }
                    else
                    {
                        comboBox6.Visible = false;
                        label35.Visible = false;
                    }
                    if (txtActividad.Text.Contains("Transporte"))
                    {
                        comboBox1.Visible = true;
                        label27.Visible = true;
                        label14.Visible = false;
                        txtEstado.Visible = false;
                    }
                    else
                    {
                        comboBox1.Visible = false;
                        label27.Visible = false;
                        label14.Visible = true;
                        txtEstado.Visible = true;
                    }
                    if (unidad.Equals("Ha"))
                    {
                        label26.Visible = false;
                        textBox1.Visible = false;
                        label23.Visible = true;
                        radioButton1.Visible = true;
                        radioButton2.Visible = true;
                        label24.Visible = true;
                        txtAreaIntervenir.Visible = true;
                    }
                    else
                    {
                        label26.Visible = true;
                        label26.Text = "Cantidad de " + unidad + ":";
                        textBox1.Visible = true;
                        label23.Visible = false;
                        radioButton1.Visible = false;
                        radioButton2.Visible = false;
                        label24.Visible = false;
                        txtAreaIntervenir.Visible = false;
                    }
                    if (txtActividad.Text.Contains("Mantenimiento"))
                    {
                        txtCodLote.Text = "-1";
                        txtLote.Text = "N/A";
                        txtEstado.Visible = false;
                        comboBox1.Visible = false;
                        label27.Visible = false;
                        label14.Visible = false;
                    }
                    //label33.Text = "Total: " + String.Format("{0:c}", (Int32.Parse(txtCostoJornal.Text) + Int32.Parse(txtCostoFinal.Text)));
                }
                else
                {
                    p++;
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            maquinarias.Clear();
            txtOperador.Text = "";
            txtDepartamento.Text = "";
            txtSupervisor.Text = "";
            txtActividad.Text = "";
            txtCentroDeCostos.Text = "";
            label9.Text = "Semana de Incio:";            
            txtTipoActividad.Text = "";
            dateTimePicker1.Value = DateTime.Now;
            dateTimePicker2.Value = DateTime.Now;
            txtMaquinaria.Text = "";
            txtCodLote.Text = "";
            txtCodLote2.Text = "";
            txtLote.Text = "";
            txtLote2.Text = "";
            txtPredio.Text = "";
            txtCosto.Text = "0";
            txtContrato.Text = "";
            comboBox1.Text = "";
            textBox1.Text = "1";
            txtArea.Text = "";
            txtAreaIntervenir.Text = "";
            radioButton2.Checked = true;
            txtEstado.Text = "";
            txtDescripcion.Text = "";
            txtCondiciones.Text = "";
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

        public void getCosto(double tipo)
        {
            string query = "SELECT SUM(Costo) FROM(SELECT SUM(i.Costo_Unitario * s.Cantidad) As Costo FROM Insumos AS i INNER JOIN insumoActividades AS s ON i.ID = s.Actividad WHERE i.Marca = 'Jornal'  AND s.Insumo = " + txtActividad.SelectedValue + " GROUP BY i.Marca  UNION ALL SELECT SUM(i.Costo_Unitario *  s.Cantidad) As Costo FROM Insumos AS i INNER JOIN insumoActividades AS s ON i.ID = s.Actividad WHERE i.Marca <> 'Jornal' AND s.Insumo = " + txtActividad.SelectedValue + " GROUP BY i.Marca)";
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
                    costo = double.Parse(myReader.GetValue(0).ToString()) * tipo;
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

        private void txtEstado_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (costo != 0)
            {
                double index = 1;
                if (txtEstado.SelectedItem.Equals("Rodilla"))
                {
                    getCostoJornal(rodilla);
                    getCosto(rodilla);
                    if (txtAreaIntervenir.Text.Equals("") && !txtArea.Text.Equals(""))
                        index = double.Parse(txtArea.Text);
                    else if (!txtAreaIntervenir.Text.Equals(""))
                        index = double.Parse(txtAreaIntervenir.Text);
                    else
                        index = double.Parse(textBox1.Text);
                    txtCosto.Text = String.Format("{0:c}",(costo * index));
                    txtCostoJornal.Text = (costoJornal * index).ToString();
                    //txtCosto.Text = (costo * rodilla).ToString();
                }
                else if (txtEstado.SelectedItem.Equals("Cintura"))
                {
                    getCostoJornal(cintura);
                    getCosto(cintura);
                    if (txtAreaIntervenir.Text.Equals("") && !txtArea.Text.Equals(""))
                        index = double.Parse(txtArea.Text);
                    else if (!txtAreaIntervenir.Text.Equals(""))
                        index = double.Parse(txtAreaIntervenir.Text);
                    else
                        index = double.Parse(textBox1.Text);
                    txtCosto.Text = String.Format("{0:c}",(costo * index));
                    txtCostoJornal.Text = (costoJornal * index).ToString();
                    //txtCosto.Text = (costo * cintura).ToString();
                }
                else if (txtEstado.SelectedItem.Equals("Cabeza"))
                {
                    getCostoJornal(cabeza);
                    getCosto(cabeza);
                    if (txtAreaIntervenir.Text.Equals("") && !txtArea.Text.Equals(""))
                        index = double.Parse(txtArea.Text);
                    else if (!txtAreaIntervenir.Text.Equals(""))
                        index = double.Parse(txtAreaIntervenir.Text);
                    else
                        index = double.Parse(textBox1.Text);
                    txtCosto.Text = String.Format("{0:c}",(costo * index));
                    txtCostoJornal.Text = (costoJornal * index).ToString();
                    //txtCosto.Text = (costo * cabeza).ToString();
                }
                //label33.Text = "Total: " + String.Format("{0:c}", (Int32.Parse(txtCostoJornal.Text) + Int32.Parse(txtCostoFinal.Text)));
            }
        }

        private void txtAreaIntervenir_TextChanged(object sender, EventArgs e)
        {
            if (!txtAreaIntervenir.Text.Equals(""))
            {
                if (b != 0)
                {
                    if (costo != 0)
                    {
                        txtCosto.Text = String.Format("{0:c}",(costo * double.Parse(txtAreaIntervenir.Text)));
                        txtCostoJornal.Text = (costoJornal * double.Parse(txtAreaIntervenir.Text)).ToString();
                        //label33.Text = "Total: " + String.Format("{0:c}", (Int32.Parse(txtCostoJornal.Text) + Int32.Parse(txtCostoFinal.Text)));
                    }
                }
                else
                {
                    b++;
                }
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (!textBox1.Text.Equals(""))
            {
                if (a != 0)
                {
                    if (costo != 0)
                    {
                        txtCosto.Text = String.Format("{0:c}",(costo * double.Parse(textBox1.Text)));
                        txtCostoJornal.Text = (costoJornal * double.Parse(textBox1.Text)).ToString();
                        //label33.Text = "Total: " + String.Format("{0:c}", (Int32.Parse(txtCostoJornal.Text) + Int32.Parse(txtCostoFinal.Text)));
                    }
                }
                else
                {
                    a++;
                }
            }
        }

        private void txtTipoActividad_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtTipoActividad.Text.Equals("Manual"))
            {
                cargarActividades("Manual");
            }
            else if (txtTipoActividad.Text.Equals("Mecanizada"))
            {
                cargarActividades("Mecanizada");
            }
            else
            {
                cargarActividades("Mantenimiento");
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            DateTime date1 = DateTime.Now;
            Calendar cal = dfi.Calendar;
            int semana = cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);
            frmOrdenFormatos newFrm = new frmOrdenFormatos(modificar,semana,tipousuario);
            if (!newFrm.IsDisposed) {
                this.Hide();
                newFrm.ShowDialog();
                this.Close();
            }  
        }

        private void frmCrearOrden_Load(object sender, EventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox4.SelectedItem != null && !comboBox4.SelectedValue.ToString().Equals("System.Data.DataRowView"))
            {
                comboBox2.SelectedItem = null;
                cargar(comboBox2, "SELECT ID, (Nombres + ' ' + Apellidos) As Nombre FROM Transportadores WHERE Proveedor = " + comboBox4.SelectedValue, "Nombre");
                if(comboBox2.Items.Count > 0)
                    comboBox2.Enabled = true;
                else
                    comboBox2.Enabled = false;
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {          
            
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

        private void button10_Click(object sender, EventArgs e)
        {
            frmResumenOrden newFrm = new frmResumenOrden(Int32.Parse(modificar));
            newFrm.Show();
        }

        public void imprimirOT()
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Aplicativo\\Formatos");
            Microsoft.Office.Interop.Excel.Application XcelApp = new Microsoft.Office.Interop.Excel.Application();
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Aplicativo\\Formatos\\", "OT*");
            XcelApp.Application.Workbooks.Add(prueba[0]);
            XcelApp.Cells[5, "C"] = fechaExpedicion;
            XcelApp.Cells[5, "L"] = OT;
            XcelApp.Cells[7, "C"] = txtOperador.Text;
            XcelApp.Cells[7, "H"] = txtDepartamento.Text;
            XcelApp.Cells[7, "L"] = txtSupervisor.Text;
            XcelApp.Cells[9, "C"] = dateTimePicker1.Text;
            XcelApp.Cells[11, "C"] = dateTimePicker2.Text;
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            DateTime date1 = dateTimePicker1.Value;
            Calendar cal = dfi.Calendar;
            XcelApp.Cells[9, "I"] = cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek).ToString();
            XcelApp.Cells[11, "I"] = semana;
            XcelApp.Cells[15, "C"] = comboBox4.Text;
            XcelApp.Cells[15, "H"] = comboBox2.Text;
            XcelApp.Cells[13, "L"] = txtContrato.Text;
            XcelApp.Cells[13, "F"] = comboBox6.Text;
            XcelApp.Cells[15, "L"] = comboBox3.Text;
            XcelApp.Cells[17, "C"] = txtTipoActividad.Text;
            XcelApp.Cells[17, "H"] = txtActividad.Text;
            XcelApp.Cells[17, "L"] = txtCentroDeCostos.Text;
            XcelApp.Cells[20, "C"] = txtDescripcion.Text;
            XcelApp.Cells[22, "C"] = txtCondiciones.Text;
            XcelApp.Cells[25, "C"] = txtCodLote.Text;
            XcelApp.Cells[27, "G"] = txtCodLote.Text;
            XcelApp.Cells[27, "C"] = txtLote.Text;
            XcelApp.Cells[25, "K"] = txtLote.Text;
            XcelApp.Cells[25, "G"] = txtPredio.Text;
            XcelApp.Cells[27, "K"] = txtArea.Text;
            if (label23.Visible)
            {
                if(radioButton1.Checked)
                    XcelApp.Cells[30, "E"] = "Si";
                else
                    XcelApp.Cells[30, "E"] = "No";
                XcelApp.Cells[30, "I"] = txtAreaIntervenir.Text;
            }
            else
            {
                XcelApp.Cells[30, "C"] = label26.Text;
                XcelApp.Cells[30, "E"] = textBox1.Text;
                XcelApp.Cells[30, "G"] = "";
            }
            if (comboBox1.Visible)
            {
                XcelApp.Cells[30, "L"] = comboBox1.Text;
                XcelApp.Cells[30, "K"] = label27.Text;
            }
            else
            {
                XcelApp.Cells[30, "L"] = txtEstado.Text;
                XcelApp.Cells[30, "K"] = label14.Text;
            }
            XcelApp.Cells[33, "C"] = txtCosto.Text;
            XcelApp.Cells[33, "F"] = txtCostoJornalFinal.Text;
            XcelApp.Cells[33, "I"] = txtCostoFinal.Text;
            XcelApp.Cells[33, "L"] = costoFinal;
            XcelApp.Cells[20, "J"] = getFormatoSeparado(listBox1);
            XcelApp.Cells[6, "P"] = getFormatoSeparado(listBox2);
            XcelApp.Visible = true;
        }

        public string getFormatoSeparado(ListBox lb)
        {
            string texto = "";
            for (int i = 0; i < lb.Items.Count; i++)
            {
                if(i == 0)
                    texto += lb.Items[i].ToString();
                else
                    texto += "\n" + lb.Items[i].ToString();                
            }
            return texto;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            imprimirOT();
        }
    }
}
