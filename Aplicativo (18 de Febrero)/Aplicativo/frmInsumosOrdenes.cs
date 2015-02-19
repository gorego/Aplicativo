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
    public partial class frmInsumosOrdenes : Form
    {
        String connectionString = Variables.connectionString;
        string fechaInicial = "";
        int OT,semanaActual,semanaOrden,semana;

        public frmInsumosOrdenes(int orden)
        {
            InitializeComponent();
            OT = orden;
            string estado = getEstado(orden);
            getOrden(orden);
            if (estado.Equals("Cerrada"))
            {
                tabControl1.TabPages.RemoveAt(0);
                formatoCerrado(dataGridView2, "SELECT ID, (Clase + ' ' + Marca + ' ' + Modelo) As Insumo FROM Insumos", "Insumo");
                if (!registroExiste(OT))
                    cargarCerrado(OT, dataGridView2);
                else
                    cargarCerradoExiste(OT, dataGridView2);
            }
            else
            {
                DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
                IFormatProvider culture = new System.Globalization.CultureInfo("es-CO", true);
                DateTime date = DateTime.Parse(fechaInicial, culture, System.Globalization.DateTimeStyles.AssumeLocal);
                Calendar cal = dfi.Calendar;
                DateTime date1 = DateTime.Now;
                semanaOrden = cal.GetWeekOfYear(date, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);
                semanaActual = cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);
                semana = semanaActual-semanaOrden;
                tabControl1.TabPages.RemoveAt(1);
                formato(dataGridView1, "SELECT ID, (Clase + ' ' + Marca + ' ' + Modelo) As Insumo FROM Insumos", "Insumo");
                //dataGridView1.Rows[0].Cells[0].Value = 1;
                cargarRegistro(semanaActual - semanaOrden, OT, dataGridView1);
                Variables.cargar2(comboBox1, "SELECT Clase From Insumos GROUP BY Clase", "Clase");
                Variables.cargar2(comboBox2, "SELECT Marca From Insumos GROUP BY Marca", "Marca");
                Variables.cargar(comboBox3, "SELECT * From Insumos", "Modelo");
            }
            this.Text = "Insumos OT # " + orden;
            linkLabel2.Visible = false;
            label1.Text = "Semana #: " + (semanaActual - semanaOrden + 1);
            if (semanaActual-semanaOrden <= 0)
                linkLabel1.Visible = false;
        }

        public void formato(DataGridView data, string query, string display)
        {
            DataGridViewComboBoxColumn combo = new DataGridViewComboBoxColumn();
            combo.HeaderText = "Modelo";
            data.Columns.Add("Column1", "#");
            data.Columns.Add("Column2", "ID");
            data.Columns[1].Visible = false;
            data.Columns.Add("Column3", "Detalle");
            Variables.cargar(combo,query,display,data);
            data.Columns.Add(combo);
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

        public string getUnidad(int id)
        {
            OleDbConnection conn = new OleDbConnection();
            string query = "SELECT Unidad_Medida FROM Insumos WHERE ID = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            string jornal = "";
            try
            {
                while (myReader.Read())
                {
                    jornal = myReader.GetString(0);
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return jornal;
        }

        public string getEstado(int id)
        {
            OleDbConnection conn = new OleDbConnection();
            string query = "SELECT estadoOrden FROM historicoOrdenes WHERE ID = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            string jornal = "";
            try
            {
                while (myReader.Read())
                {
                    jornal = myReader.GetString(0);
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return jornal;
        }

        public void getOrden(int orden)
        {
            string query = "SELECT h.Area, h.estadoLote, (t.Nombres+' '+t.Apellidos), b.Predio, Areas.Lote, a.Actividad, a.Unidad_de_Medida,h.fechaInicio,h.fechaFinal,Areas.Codigo,a.Descripcion_Actividad,h.Cliente,h.Transportador FROM ((historicoOrdenes AS h INNER JOIN Trabajadores AS t ON h.Supervisor = t.ID) INNER JOIN Actividades AS a ON h.Actividad = a.ID) INNER JOIN (BancoTierras AS b INNER JOIN Areas ON b.ID = Areas.Predio) ON h.Lote = Areas.Codigo WHERE h.ID = " + orden + " UNION ALL SELECT h.Area, h.estadoLote, (t.Nombres+' '+t.Apellidos), b.Predio, Lotes.Lote, a.Actividad, a.Unidad_de_Medida,h.fechaInicio,h.fechaFinal, Lotes.Codigo,a.Descripcion_Actividad,h.Cliente,h.Transportador FROM ((historicoOrdenes AS h INNER JOIN Trabajadores AS t ON h.Supervisor = t.ID) INNER JOIN Actividades AS a ON h.Actividad = a.ID) INNER JOIN (BancoTierras AS b INNER JOIN Lotes ON b.ID = Lotes.Predio) ON h.Lote = Lotes.Codigo WHERE h.ID = " + orden + " UNION ALL SELECT h.Area, h.estadoLote, (t.Nombres+' '+t.Apellidos), b.Predio, LoteGanadero.Lote, a.Actividad, a.Unidad_de_Medida,h.fechaInicio,h.fechaFinal, LoteGanadero.Codigo,a.Descripcion_Actividad,h.Cliente,h.Transportador FROM ((historicoOrdenes AS h INNER JOIN Trabajadores AS t ON h.Supervisor = t.ID) INNER JOIN Actividades AS a ON h.Actividad = a.ID) INNER JOIN (BancoTierras AS b INNER JOIN LoteGanadero ON b.ID = LoteGanadero.Predio) ON h.Lote = LoteGanadero.Codigo WHERE h.ID = " + orden;
            //Ejecutar el query y llenar el GridView.
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                if (myReader.Read())
                {
                    fechaInicial = myReader.GetString(7);
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

        public bool getExiste(int semana, int orden)
        {
            string query = "SELECT * FROM almacenOrden WHERE Semana = " + semana + " AND Orden = " + orden;
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

        public void formato(DataGridView data, string formato, int column)
        {
            for (int j = 4; j < 16; j = j + 2)
            {
                data.Rows[column].Cells[j].Value = formato;
            }
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

        public void crearRegistro(int semana, int orden, DataGridView data)
        {
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = connectionString;
            OleDbCommand cmd;
            for (int i = 0; i < data.Rows.Count; i++)
            {
                cmd = new OleDbCommand("INSERT INTO almacenOrden(Semana,Unidad,Orden,Detalle,Modelo,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado) VALUES (@Semana,@Unidad,@Orden,@Detalle,@Modelo,@Lunes,@Martes,@Miercoles,@Jueves,@Viernes,@Sabado)");
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
                    if (data.Rows[i].Cells[3].Value != null)
                        cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = data.Rows[i].Cells[3].Value.ToString();
                    else
                        cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[5].Value != null)
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = data.Rows[i].Cells[5].Value.ToString();
                    else
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[7].Value != null)
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = data.Rows[i].Cells[7].Value.ToString();
                    else
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[9].Value != null)
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = data.Rows[i].Cells[9].Value.ToString();
                    else
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[11].Value != null)
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = data.Rows[i].Cells[11].Value.ToString();
                    else
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[13].Value != null)
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = data.Rows[i].Cells[13].Value.ToString();
                    else
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[15].Value != null)
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = data.Rows[i].Cells[15].Value.ToString();
                    else
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = 0;
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

        public void crearCerrado(int orden, DataGridView data)
        {
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = connectionString;
            OleDbCommand cmd;
            for (int i = 0; i < data.Rows.Count; i++)
            {
                cmd = new OleDbCommand("INSERT INTO registroAlmacenOrden(Orden,Detalle,Modelo,CantEnt,CantRecibido) VALUES (@Orden,@Detalle,@Modelo,@CantEnt,@CantRecibido)");
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@Orden", OleDbType.VarChar).Value = orden;
                    if (data.Rows[i].Cells[2].Value != null)
                        cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = data.Rows[i].Cells[2].Value.ToString();
                    else
                        cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = "";
                    if (data.Rows[i].Cells[3].Value != null)
                        cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = data.Rows[i].Cells[3].Value.ToString();
                    else
                        cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@CantEnt", OleDbType.VarChar).Value = data.Rows[i].Cells[4].Value.ToString();
                    cmd.Parameters.Add("@CantRecibido", OleDbType.VarChar).Value = data.Rows[i].Cells[5].Value.ToString();
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

        public void modificaCerrado(int orden, DataGridView data)
        {
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = connectionString;
            OleDbCommand cmd;
            for (int i = 0; i < data.Rows.Count; i++)
            {                
                int k = 0;
                if (data.Rows[i].Cells[1].Value != null)
                {
                    cmd = new OleDbCommand("UPDATE registroAlmacenOrden SET CantEnt=@CantEnt,CantRecibido=@CantRecibido WHERE ID = " + data.Rows[i].Cells[1].Value.ToString());
                }
                else
                {
                    cmd = new OleDbCommand("UPDATE registroAlmacenOrden SET CantEnt=@CantEnt,CantRecibido=@CantRecibido WHERE Detalle = '" + data.Rows[i].Cells[2].Value.ToString() + "' AND Modelo = " + data.Rows[i].Cells[3].Value.ToString());
                }
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@CantEnt", OleDbType.VarChar).Value = data.Rows[i].Cells[4].Value.ToString();
                    cmd.Parameters.Add("@CantRecibido", OleDbType.VarChar).Value = data.Rows[i].Cells[5].Value.ToString();
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

        public void modificaRegistro(int semana, int orden, DataGridView data)
        {
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = connectionString;
            OleDbCommand cmd;
            for (int i = 0; i < data.Rows.Count; i++)
            {
                int j = 0;
                int k = 0;
                if (data.Rows[i].Cells[3].Value != null)
                {
                    if (data.Rows[i].Cells[1].Value != null)
                    {
                        cmd = new OleDbCommand("UPDATE almacenOrden SET Detalle=@Detalle,Modelo=@Modelo,Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado WHERE ID = " + data.Rows[i].Cells[1].Value.ToString());
                        k = 0;
                    }
                    else
                    {
                        cmd = new OleDbCommand("INSERT INTO almacenOrden(Semana,Unidad,Orden,Detalle,Modelo,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado) VALUES (@Semana,@Unidad,@Orden,@Detalle,@Modelo,@Lunes,@Martes,@Miercoles,@Jueves,@Viernes,@Sabado)");
                        k = 1;
                    }
                    j = 0;
                }
                else
                {
                    cmd = new OleDbCommand("UPDATE almacenOrden SET Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado WHERE ID = " + data.Rows[i].Cells[1].Value.ToString());
                    j = 1;
                }
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    if (k == 1)
                    {
                        cmd.Parameters.Add("@Semana", OleDbType.VarChar).Value = semana;
                        cmd.Parameters.Add("@Unidad", OleDbType.VarChar).Value = data.Rows[i].Cells[4].Value.ToString();
                        cmd.Parameters.Add("@Orden", OleDbType.VarChar).Value = orden;
                    }
                    if (data.Rows[i].Cells[2].Value != null)
                        cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = data.Rows[i].Cells[2].Value.ToString();
                    else
                        cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = ""; 
                    if (j == 0)
                        cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = data.Rows[i].Cells[3].Value.ToString();
                    else
                        cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[5].Value != null)
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = data.Rows[i].Cells[5].Value.ToString();
                    else
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[7].Value != null)
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = data.Rows[i].Cells[7].Value.ToString();
                    else
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[9].Value != null)
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = data.Rows[i].Cells[9].Value.ToString();
                    else
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[11].Value != null)
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = data.Rows[i].Cells[11].Value.ToString();
                    else
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[13].Value != null)
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = data.Rows[i].Cells[13].Value.ToString();
                    else
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[15].Value != null)
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = data.Rows[i].Cells[15].Value.ToString();
                    else
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = 0;
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

        public void cargarRegistro(int semana, int orden, DataGridView data)
        {
            OleDbConnection conn = new OleDbConnection();
            string query = "SELECT * FROM almacenOrden WHERE Semana = " + semana + " AND Orden = " + orden;
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
                    if (!myReader.IsDBNull(4))
                        if (myReader.GetInt32(4) != 0)
                            data.Rows[i].Cells[3].Value = myReader.GetInt32(4);
                    data.Rows[i].Cells[1].Value = myReader.GetInt32(0);
                    data.Rows[i].Cells[0].Value = i + 1;
                    data.Rows[i].Cells[2].Value = myReader.GetString(5);
                    for (int j = 6, k = 5; j < 12; j++, k = k + 2)
                    {
                        data.Rows[i].Cells[k].Value = myReader.GetInt32(j).ToString();
                    }
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
                    if(myReader.GetString(3).Equals("Prestable"))
                        data.Rows[i].Cells[5].Value = myReader.GetValue(2).ToString();
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
        
        public void modificarValores(int semana, int orden, string simbolo)
        {
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = connectionString;
            OleDbCommand cmd;            
            cmd = new OleDbCommand("UPDATE Insumos AS i INNER JOIN almacenOrden AS c ON i.ID = c.Modelo SET i.Cantidad_Stock = (i.Cantidad_Stock " + simbolo +"c.Lunes " + simbolo + "c.Martes " + simbolo + "c.Miercoles " + simbolo + "c.Jueves " + simbolo + "c.Viernes " + simbolo + "c.Sabado) WHERE c.Semana = @semana AND c.Orden = @orden");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@semana", OleDbType.VarChar).Value = semana;
                cmd.Parameters.Add("@orden", OleDbType.VarChar).Value = orden;                
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

        public void modificarValoresCerrado(int orden, string simbolo)
        {
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = connectionString;
            OleDbCommand cmd;
            cmd = new OleDbCommand("UPDATE Insumos AS i INNER JOIN registroAlmacenOrden AS c ON i.ID = c.Modelo SET i.Cantidad_Stock = (i.Cantidad_Stock " + simbolo + " c.CantRecibido) WHERE c.Orden = @orden");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@orden", OleDbType.VarChar).Value = orden;
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

        
        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            if (dataGridView1.Columns.Count > 5)
            {
                dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[0].Value = dataGridView1.Rows.Count;
            }
        }

        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            Variables.dirtyCell(dataGridView1);
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Columns.Count > 10)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1.Rows[i].Cells[3].FormattedValue != null && !dataGridView1.Rows[i].Cells[3].FormattedValue.Equals(""))
                    {
                        formato(dataGridView1, getUnidad((int)dataGridView1.Rows[i].Cells[3].Value), i);
                    }
                }
                Contador(dataGridView1, 5, 17);                
            }
        }

        public double getCantidadStock(int id)
        {
            double cantidad = 0;
            OleDbConnection conn = new OleDbConnection();
            string query = "SELECT Cantidad_Stock FROM Insumos WHERE ID = " + id;
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
                    cantidad = myReader.GetDouble(0);
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

        public bool existeMenosCantidad(DataGridView data) {            
            bool sw = false;
            for (int i = 0; i < data.Rows.Count; i++)
            {
                if (Int32.Parse(data.Rows[i].Cells[data.Columns.Count - 1].Value.ToString()) > getCantidadStock(Int32.Parse(data.Rows[i].Cells[3].Value.ToString())))
                {
                    sw = true;
                    MessageBox.Show("No se encuentra esa cantidad de " + data.Rows[i].Cells[3].Value + " en stock.", "Error");
                }
            }
            return sw;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (!existeMenosCantidad(dataGridView1))
            {
                if (getExiste(semanaActual - semanaOrden, OT) == false)
                {
                    crearRegistro(semanaActual - semanaOrden, OT, dataGridView1);
                    modificarValores(semanaActual - semanaOrden, OT, "-");
                    Variables.agregarLog("Usuario ha modificado el registro de insumos en la OT #" + OT, Variables.userName);
                }
                else
                {
                    modificarValores(semanaActual - semanaOrden, OT, "+");
                    modificaRegistro(semanaActual - semanaOrden, OT, dataGridView1);
                    modificarValores(semanaActual - semanaOrden, OT, "-");
                    Variables.agregarLog("Usuario ha modificado el registro de insumos en la OT #" + OT, Variables.userName);
                }
                MessageBox.Show("Entrega de Insumos Registrado.");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            semanaActual--;
            cargarRegistro(semanaActual - semanaOrden, OT, dataGridView1);
            linkLabel1.Visible = true;
            linkLabel2.Visible = true;
            if (semanaActual-semanaOrden <= 0)
            {
                linkLabel1.Visible = false;
            }
            label1.Text = "Semana #: " + (semanaActual - semanaOrden + 1);
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            semanaActual++;
            cargarRegistro(semanaActual - semanaOrden, OT, dataGridView1);
            linkLabel1.Visible = true;
            linkLabel2.Visible = true;
            if (semana == semanaActual-semanaOrden)
            {
                linkLabel2.Visible = false;
            }
            label1.Text = "Semana #: " + (semanaActual - semanaOrden + 1);
        }

        public bool puedeRetornar(DataGridView data)
        {
            bool sw = true;
            for (int i = 0; i < data.Rows.Count; i++)
            {
                if (Int32.Parse(data.Rows[i].Cells[4].Value.ToString()) < Int32.Parse(data.Rows[i].Cells[5].Value.ToString()))
                {
                    sw = false;
                }
            }
            return sw;
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (puedeRetornar(dataGridView2))
            {
                if (!registroExiste(OT))
                {
                    crearCerrado(OT, dataGridView2);
                    modificarValoresCerrado(OT, "+");
                    Variables.agregarLog("Usuario ha registrado el retorno de los insumos en la OT #" + OT, Variables.userName);
                }
                else
                {
                    modificarValoresCerrado(OT, "-");
                    modificaCerrado(OT, dataGridView2);
                    modificarValoresCerrado(OT, "+");
                    Variables.agregarLog("Usuario ha registrado el retorno de los insumos en la OT #" + OT, Variables.userName);
                }
                MessageBox.Show("Retorno de Insumos a Bodega Registrado.");
            }
            else
            {
                MessageBox.Show("No se puede retornar mas de lo entregado.", "Error");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox3.SelectedItem != null)
            {
                dataGridView1.Rows.Add();
                dataGridView1.Rows[dataGridView1.Rows.Count-1].Cells[3].Value = comboBox3.SelectedValue;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Variables.cargar(comboBox3, "SELECT * FROM Insumos WHERE Clase = '" + comboBox1.Text + "'", "Modelo");
            comboBox2.Items.Clear();
            Variables.cargar2(comboBox2, "SELECT Marca FROM Insumos WHERE Clase = '" + comboBox1.Text + "' GROUP BY Marca", "Marca");
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {           
            if(comboBox1.Text.Equals(""))
                Variables.cargar(comboBox3, "SELECT * FROM Insumos WHERE Marca = '" + comboBox2.Text + "'", "Modelo");
            else
                Variables.cargar(comboBox3, "SELECT * FROM Insumos WHERE Marca = '" + comboBox2.Text + "' AND Clase = '" + comboBox1.Text + "'", "Modelo");
            //comboBox1.Items.Clear();
            //Variables.cargar2(comboBox1, "SELECT Clase FROM Insumos WHERE Marca = '" + comboBox2.Text + "' GROUP BY Clase", "Clase");
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

    }
}
