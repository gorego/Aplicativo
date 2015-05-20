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
using System.Net;
using System.Globalization;

namespace Aplicativo
{
    public partial class frmIngreso : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        int error = 0;

        public void Ingresar() { 
            string query = "SELECT * FROM Usuarios WHERE usuario = '"+txtUsuario.Text.ToLower()+"'";
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                if (myReader.Read())
                {
                    if (txtClave.Text.Equals(myReader.GetString(2)))
                    {
                        if (myReader.GetInt32(3).Equals(1))
                        {
                            Variables.userName = txtUsuario.Text;
                            Variables.tipo = myReader.GetInt32(3);
                            Variables.userID = myReader.GetInt32(4);
                            frmInicio newFrm = new frmInicio(myReader.GetInt32(4),txtUsuario.Text);
                            // always call Close when done reading.
                            myReader.Close();
                            // always call Close when done reading.
                            conn.Close();
                            this.Hide();
                            newFrm.ShowDialog();
                            this.Close();
                        }
                        else if (myReader.GetInt32(3).Equals(2))
                        {
                            Variables.userName = txtUsuario.Text;
                            Variables.tipo = myReader.GetInt32(3);
                            Variables.userID = myReader.GetInt32(4);
                            frmHistoricoOrdenes newFrm = new frmHistoricoOrdenes(myReader.GetInt32(4), 1, txtUsuario.Text);
                            // always call Close when done reading.
                            myReader.Close();
                            // always call Close when done reading.
                            conn.Close();
                            this.Hide();
                            newFrm.ShowDialog();
                            this.Close();
                        }
                        else if (myReader.GetInt32(3).Equals(3))
                        {
                            Variables.userName = txtUsuario.Text;
                            Variables.tipo = myReader.GetInt32(3);
                            Variables.userID = myReader.GetInt32(4);
                            frmOrdenesAlamacen newFrm = new frmOrdenesAlamacen(1);
                            // always call Close when done reading.
                            myReader.Close();
                            // always call Close when done reading.
                            conn.Close();
                            this.Hide();
                            newFrm.ShowDialog();
                            this.Close();
                        }
                        else if (myReader.GetInt32(3).Equals(4)) 
                        {
                            Variables.userName = txtUsuario.Text;
                            Variables.tipo = myReader.GetInt32(3);
                            Variables.userID = myReader.GetInt32(4);
                            frmHistoricoProduccion newFrm = new frmHistoricoProduccion(1);
                            myReader.Close();
                            conn.Close();
                            this.Hide();
                            newFrm.ShowDialog();
                            this.Close();
                        }
                        else if (myReader.GetInt32(3).Equals(5))
                        {
                            Variables.userName = txtUsuario.Text;
                            Variables.tipo = myReader.GetInt32(3);
                            Variables.userID = myReader.GetInt32(4);
                            frmOrdenesCuchilla newFrm = new frmOrdenesCuchilla(1);
                            myReader.Close();
                            conn.Close();
                            this.Hide();
                            newFrm.ShowDialog();
                            this.Close();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Usuario o contraseña incorrecta.");
                        error = 1;
                    }
                }
                else
                {
                    MessageBox.Show("Usuario o contraseña incorrecta.");
                    error = 1;
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

        public frmIngreso()
        {
            InitializeComponent();
            //if(!DateTime.Now.DayOfWeek.ToString().Equals("Monday"))
            //    agregarDias();
        }

        public void agregarDias()
        {            
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Trabajadores AS t INNER JOIN fijoSemanal AS f ON t.ID = f.Trabajador SET t.diasLaborados = (t.diasLaborados + f.Lunes + f.Martes + f.Miercoles + f.Jueves + f.Viernes + f.Sabado + f.Domingo),f.Estado=1 WHERE f.Estado = 0 AND f.Semana = @semana AND f.Ano = @ano");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
                DateTime date1 = DateTime.Now;
                Calendar cal = dfi.Calendar;
                int semana = cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek) - 1;
                cmd.Parameters.Add("@semana", OleDbType.VarChar).Value = semana;
                cmd.Parameters.Add("@ano", OleDbType.VarChar).Value = DateTime.Now.Year;
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

        private void btnIngresar_Click(object sender, EventArgs e)
        {
            //if (CheckForInternetConnection())
            //{
            //    Ingresar();
            //}
            //else
            //{
            //    MessageBox.Show("No");
            //}
            Ingresar();
        }

        public void agregarLog(string accion)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO historicoIngresos (Usuario,Accion,Fecha) VALUES (@Usuario,@Accion,@Fecha)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Usuario", OleDbType.VarChar).Value = txtUsuario.Text;
                cmd.Parameters.Add("@Accion", OleDbType.VarChar).Value = accion;
                cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = DateTime.Now.ToString("dd") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.Year + " - " + DateTime.Now.Hour + ":" + DateTime.Now.ToString("mm") + ":" + DateTime.Now.ToString("ss");
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

        public static bool CheckForInternetConnection()
        {
            try
            {
                using (var client = new WebClient())
                using (var stream = client.OpenRead("http://www.google.com"))
                {
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        private void txtClave_KeyUp(object sender, KeyEventArgs e)
        {
            //if (e.KeyValue == 13)
            //{
            //    e.Handled = true;
            //    e.SuppressKeyPress = true;
            //    if (CheckForInternetConnection())
            //    {                   
            //        if (error == 0)
            //            Ingresar();
            //        else
            //            error = 0;
            //    }
            //    else
            //    {
            //        MessageBox.Show("No hay conexion.");
            //    }
            //}
        }

        private void txtClave_Click(object sender, EventArgs e)
        {
            txtClave.Text = "";
        }

        private void txtClave_Enter(object sender, EventArgs e)
        {
            txtClave.Text = "";
        }

        private void txtUsuario_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                if (CheckForInternetConnection())
                {
                    e.Handled = true;                
                    if (error == 0)
                        Ingresar();
                    else
                        error = 0;
                }
                else
                {
                    MessageBox.Show("No hay conexion.");
                }
            }
        }

        private void txtClave_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                if (CheckForInternetConnection())
                {
                    e.Handled = true;
                    if (error == 0)
                        Ingresar();
                    else
                        error = 0;
                }
                else
                {
                    MessageBox.Show("No hay conexion.");
                }
            }
        }
    }
}
