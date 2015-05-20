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
    public partial class frmAsistenciaOrdenes : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        bool sw = false;
        bool sw2 = false;
        int semana;
        int semanaTemp;

        public frmAsistenciaOrdenes()
        {
            InitializeComponent();
            cargarEmpleados();
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            DateTime date1 = DateTime.Now;
            Calendar cal = dfi.Calendar;
            semana = cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);
            semanaTemp = semana;
            label1.Text = "Semana #: " + semana;
            cargarSemanal(semana, DateTime.Now.Year);
        }

        public void cargarEmpleados()
        {
            while (dataGridView3.Rows.Count != 0)
            {
                dataGridView3.Rows.RemoveAt(0);
            }
            string query = "SELECT t.ID,(t.Nombres + ' ' + t.Apellidos), t.Cedula, c.Cargo FROM ((CargoLaboral AS c INNER JOIN Trabajadores AS t ON c.ID = t.Cargo) INNER JOIN produccionEmpleados AS e ON t.ID = e.Trabajador) INNER JOIN historicoProduccion AS h ON e.Orden = h.ID WHERE h.Estado <> 'Cerrada' GROUP BY t.ID, (t.Nombres + ' ' + t.Apellidos), t.Cedula, c.Cargo";
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

        public void clearSemana()
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                for (int j = 5; j < 11; j++)
                {
                    dataGridView3.Rows[i].Cells[j].Value = false;
                }
            }
        }

        public void cargarSemanal(int semana, int ano)
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                string query = "SELECT * FROM adf002OP WHERE Semana = " + semana + " AND Ano = " + ano + " AND Trabajador = " + dataGridView3.Rows[i].Cells[1].Value.ToString();
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
                            if (myReader.GetInt32(j-1) == 1)
                                dataGridView3.Rows[i].Cells[j].Value = true;
                        }
                        if (myReader.GetInt32(12) == 1)
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
        
        public void crearSemanal(int semana, int ano)
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("INSERT INTO adf002OP (Semana,Ano,Trabajador,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,Domingo,Estado,Editable) VALUES (@Semana,@Ano,@Trabajador,@Lunes,@Martes,@Miercoles,@Jueves,@Viernes,@Sabado,@Domingo,@Estado,@Editable)");
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@Semana", OleDbType.VarChar).Value = semana;
                    cmd.Parameters.Add("@Ano", OleDbType.VarChar).Value = ano;
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

        public void modificarSemanal(int semana, int ano)
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("UPDATE adf002OP SET Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado,Domingo=@Domingo WHERE Semana = " + semana + " AND Ano = " + ano + " AND Trabajador = " + dataGridView3.Rows[i].Cells[1].Value.ToString());
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

        public bool fijoExiste(int semana, int ano)
        {
            string query = "SELECT * FROM adf002OP WHERE Semana = " + semana + " AND Ano = " + ano;
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

        private void button1_Click(object sender, EventArgs e)
        {
            if (fijoExiste(semanaTemp, DateTime.Now.Year) == false)
            {
                crearSemanal(semanaTemp, DateTime.Now.Year);
                //agregarDias(semana, DateTime.Now.Year, supervisor);
            }
            else
            {
                //eliminarDias(semana - 1, DateTime.Now.Year, supervisor);
                modificarSemanal(semanaTemp, DateTime.Now.Year);
                //agregarDias(semana - 1, DateTime.Now.Year, supervisor);
            }
            MessageBox.Show("Asistencia registrada.");
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

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            semanaTemp--;
            label1.Text = "Semana #: " + semanaTemp;
            clearSemana();
            cargarSemanal(semanaTemp, DateTime.Now.Year);
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            semanaTemp++;
            label1.Text = "Semana #: " + semanaTemp;
            clearSemana();
            cargarSemanal(semanaTemp, DateTime.Now.Year);
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

    }
}
