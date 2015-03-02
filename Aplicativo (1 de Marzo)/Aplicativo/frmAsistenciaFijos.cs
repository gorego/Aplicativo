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
using Excel = Microsoft.Office.Interop.Excel; 

namespace Aplicativo
{
    public partial class frmAsistenciaFijos : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        int tip = 0;
        bool sw = false;
        bool sw2 = false;
        int supervisor;

        public void cargarEmpleados(string id)
        {
            while (dataGridView3.Rows.Count != 0)
            {
                dataGridView3.Rows.RemoveAt(0);
            }
            string query = "SELECT t.ID, (t.Nombres + ' ' + t.Apellidos), t.Cedula, c.Cargo FROM CargoLaboral AS c INNER JOIN (Trabajadores AS t INNER JOIN supervisorEmpleados AS s ON t.ID = s.Trabajador) ON c.ID = t.Cargo WHERE s.Supervisor = " + id;
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
                    dataGridView3.Rows[i].Cells[0].Value = i+1;                 
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

        public frmAsistenciaFijos(string id, int tipo)
        {
            InitializeComponent();            
            cargarEmpleados(id);
            tip = tipo;
            supervisor = Int32.Parse(id);
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            DateTime date1 = DateTime.Now;
            Calendar cal = dfi.Calendar;
            int semana = cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);
            label1.Text = "Semana #: " + semana;
            if (tipo == 1)
            {
                semana--;
                button4.Text = "Semana actual";
            }                
            cargarSemanal(semana,DateTime.Now.Year,supervisor);
            if (fijoExiste(semana, DateTime.Now.Year, supervisor) == false)
            {
                crearSemanal(semana, DateTime.Now.Year, supervisor);
            }
            //if (!sw2 && !DateTime.Now.DayOfWeek.ToString().Equals("Monday"))
            //{
            //    MessageBox.Show("Error");
            //    this.Close();
            //}
        }

        public void agregarLog(string accion,string usuario)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO historicoIngresos (Usuario,Accion,Fecha) VALUES (@Usuario,@Accion,@Fecha)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Usuario", OleDbType.VarChar).Value = usuario;
                cmd.Parameters.Add("@Accion", OleDbType.VarChar).Value = accion;
                cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = DateTime.Now.Day + "/" + DateTime.Now.Month + "/" + DateTime.Now.Year + " - " + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Minute;
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

        public string getUsuario(string id) 
        {
            string usuario = "";
            string query = "SELECT usuario FROM Usuarios WHERE trabajador = " + id;
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
                    usuario = myReader.GetString(0);
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return usuario;
        }

        public string getSupervisor(int id)
        {
            string usuario = "";
            string query = "SELECT (Nombres + ' ' + Apellidos) FROM Trabajadores WHERE ID = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int i = 0;
            try
            {
                if (myReader.Read())
                {
                    usuario = myReader.GetString(0);
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return usuario;
        }

        public bool fijoExiste(int semana, int ano, int supervisor)
        {
            string query = "SELECT * FROM fijoSemanal WHERE Semana = " + semana + " AND Ano = " + ano + " AND Supervisor = " + supervisor;
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

        public void crearSemanal(int semana, int ano, int supervisor)
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("INSERT INTO fijoSemanal (Semana,Ano,Supervisor,Trabajador,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,Domingo,Estado,Editable) VALUES (@Semana,@Ano,@Supervisor,@Trabajador,@Lunes,@Martes,@Miercoles,@Jueves,@Viernes,@Sabado,@Domingo,@Estado,@Editable)");
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@Semana", OleDbType.VarChar).Value = semana;
                    cmd.Parameters.Add("@Ano", OleDbType.VarChar).Value = ano;
                    cmd.Parameters.Add("@Supervisor", OleDbType.VarChar).Value = supervisor;
                    cmd.Parameters.Add("@Trabajador", OleDbType.VarChar).Value = dataGridView3.Rows[i].Cells[1].Value.ToString();
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

        public void cargarSemanal(int semana, int ano, int supervisor)
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                string query = "SELECT * FROM fijoSemanal WHERE Semana = " + semana + " AND Ano = " + ano + " AND Supervisor = " + supervisor + " AND Trabajador = " + dataGridView3.Rows[i].Cells[1].Value.ToString();
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
                    for (int j = 5; j < 11; j++)
                    {
                        if (myReader.GetInt32(j) == 1)
                            dataGridView3.Rows[i].Cells[j].Value = true;                        
                    }
                    if (myReader.GetInt32(13) == 1)
                        sw2 = true;
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

        public void modificarSemanal(int semana, int ano, int supervisor)
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("UPDATE fijoSemanal SET Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado,Domingo=@Domingo WHERE Semana = " + semana + " AND Ano = " + ano + " AND Supervisor = " + supervisor + " AND Trabajador = " + dataGridView3.Rows[i].Cells[1].Value.ToString());
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
        }

        public void agregarDias(int semana, int ano, int supervisor)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Trabajadores AS t INNER JOIN fijoSemanal AS f ON t.ID = f.Trabajador SET t.diasLaborados = (t.diasLaborados + f.Lunes + f.Martes + f.Miercoles + f.Jueves + f.Viernes + f.Sabado + f.Domingo),f.Estado=1, f.Editable=0 WHERE f.Semana = @semana AND f.Ano = @ano AND f.Supervisor = @supervisor");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@semana", OleDbType.VarChar).Value = semana;
                cmd.Parameters.Add("@ano", OleDbType.VarChar).Value = ano;
                cmd.Parameters.Add("@supervisor", OleDbType.VarChar).Value = supervisor;
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

        public void eliminarDias(int semana, int ano, int supervisor)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Trabajadores AS t INNER JOIN fijoSemanal AS f ON t.ID = f.Trabajador SET t.diasLaborados = (t.diasLaborados - f.Lunes - f.Martes - f.Miercoles - f.Jueves - f.Viernes - f.Sabado - f.Domingo),f.Estado=1, f.Editable=0 WHERE f.Semana = @semana AND f.Ano = @ano AND f.Supervisor = @supervisor");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@semana", OleDbType.VarChar).Value = semana;
                cmd.Parameters.Add("@ano", OleDbType.VarChar).Value = ano;
                cmd.Parameters.Add("@supervisor", OleDbType.VarChar).Value = supervisor;
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

        private void button1_Click(object sender, EventArgs e)
        {
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            DateTime date1 = DateTime.Now;
            Calendar cal = dfi.Calendar;
            int semana = cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);
            string usuario = getUsuario(supervisor.ToString());
            if (fijoExiste(semana, DateTime.Now.Year, supervisor) == false)
            {
                crearSemanal(semana, DateTime.Now.Year, supervisor);
                agregarLog("Usuario ha modificado su asistencia semanal.",usuario);
                agregarDias(semana, DateTime.Now.Year, supervisor);
            }
            else
            {
                if (tip == 1)
                {
                    eliminarDias(semana-1, DateTime.Now.Year, supervisor);
                    modificarSemanal(semana-1, DateTime.Now.Year, supervisor);
                    agregarDias(semana-1, DateTime.Now.Year, supervisor);
                    agregarLog("Usuario ha modificado su asistencia semanal de la semana anterior.", usuario);
                }
                else
                {
                    eliminarDias(semana, DateTime.Now.Year, supervisor);
                    modificarSemanal(semana, DateTime.Now.Year, supervisor);
                    agregarDias(semana, DateTime.Now.Year, supervisor);
                    agregarLog("Usuario ha modificado su asistencia semanal.", usuario);
                }

            }
            //if (sw2) 
            //{
            //    agregarDias(semana, DateTime.Now.Year, supervisor);                
            //    agregarLog("Permiso de modificación de asistencia revocado.", usuario);
            //}
            MessageBox.Show("Asistencia registrada.");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridView3_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            selectAll();
            Contador();
        }

        public void selectAll() { 
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
            if (dataGridView3.IsCurrentCellDirty)
            {
                dataGridView3.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            DateTime date1 = DateTime.Now;
            Calendar cal = dfi.Calendar;
            int semana = cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek) - 1;
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos");
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos\\","ADF001*");      
            Microsoft.Office.Interop.Excel.Application XcelApp = new Microsoft.Office.Interop.Excel.Application();
            XcelApp.Application.Workbooks.Add(prueba[0]);
            XcelApp.Cells[5, 2] = getSupervisor(supervisor);
            XcelApp.Cells[5, 8] = DateTime.Now.Day + " / " + DateTime.Now.Month + " / " + DateTime.Now.Year;
            XcelApp.Cells[5, 12] = semana;
            for (int i = 0; i < dataGridView3.Rows.Count; i++)            
                for (int j = 0; j < 3; j++)                
                    XcelApp.Cells[18 + i, 2 + j] = dataGridView3.Rows[i].Cells[2 + j].Value;                               
            XcelApp.Columns.AutoFit();
            XcelApp.Visible = true;
        }

        public bool esEditable(int semana, int ano, int supervisor)
        {
            string query = "SELECT * FROM fijoSemanal WHERE Editable = 1 AND Semana = " + semana + " AND Ano = " + ano + " AND Supervisor = " + supervisor;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                if (myReader.HasRows)
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


        private void button4_Click(object sender, EventArgs e)
        {
            if (button4.Text.Equals("Semana Anterior"))
            {
                DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
                DateTime date1 = DateTime.Now;
                Calendar cal = dfi.Calendar;
                int semana = cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);
                if (esEditable(semana - 1, DateTime.Now.Year, supervisor) || DateTime.Now.DayOfWeek.Equals("Monday"))
                {
                    frmAsistenciaFijos newFrm = new frmAsistenciaFijos(supervisor.ToString(), 1);
                    this.Hide();
                    newFrm.ShowDialog();
                    this.Close();
                }
                else
                {
                    MessageBox.Show("No se puede editar la semana anterior, favor contactar al administrador.");
                }
            }
            else
            {
                frmAsistenciaFijos newFrm = new frmAsistenciaFijos(supervisor.ToString(),0);
                this.Hide();
                newFrm.ShowDialog();
                this.Close();
            }
        }
    }
}
