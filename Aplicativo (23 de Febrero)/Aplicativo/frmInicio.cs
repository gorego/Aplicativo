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
    public partial class frmInicio : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        int usuario;
        string user;

        public frmInicio(int nombre,string username)
        {            
            InitializeComponent();
            usuario = nombre;
            user = username;
            agregarLog("Ingreso", user);            
        }

        private void btnOperador_Click(object sender, EventArgs e)
        {
            frmOperDepar newFrm = new frmOperDepar(1);
            newFrm.ShowDialog();
        }

        private void btnDepartamentos_Click(object sender, EventArgs e)
        {
            frmOperDepar newFrm = new frmOperDepar(2);
            newFrm.ShowDialog();
        }

        private void btnSupervisor_Click(object sender, EventArgs e)
        {
            frmSupervisor newFrm = new frmSupervisor();
            newFrm.Show();
        }

        private void btnActividad_Click(object sender, EventArgs e)
        {
            frmActividades newFrm = new frmActividades();
            newFrm.Show();
        }

        private void btnMaquinaria_Click(object sender, EventArgs e)
        {
            frmMaquinaria newFrm = new frmMaquinaria();
            newFrm.Show();
        }

        private void btnEmpleados_Click(object sender, EventArgs e)
        {
            frmEmpleados newFrm = new frmEmpleados();
            newFrm.Show();
        }

        private void btnBanco_Click(object sender, EventArgs e)
        {
            frmBanco newFrm = new frmBanco();
            newFrm.Show();
        }

        private void btnCentro_Click(object sender, EventArgs e)
        {
            frmCentro newFrm = new frmCentro();
            newFrm.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            frmInsumos newFrm = new frmInsumos(0);
            newFrm.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            frmHistoricoOrdenes newFrm = new frmHistoricoOrdenes(usuario,user);
            newFrm.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            frmContratos newFrm = new frmContratos();
            newFrm.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            frmUsuarios newFrm = new frmUsuarios();
            newFrm.Show();
        }

        public void agregarLog(string accion, string usuario)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO historicoIngresos (Usuario,Accion,Fecha) VALUES (@Usuario,@Accion,@Fecha)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Usuario", OleDbType.VarChar).Value = usuario;
                cmd.Parameters.Add("@Accion", OleDbType.VarChar).Value = accion;

                cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = DateTime.Now.Day + "/" + DateTime.Now.Month + "/" + DateTime.Now.Year + " - " + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second;
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

        private void button5_Click(object sender, EventArgs e)
        {
            frmLogs newFrm = new frmLogs();
            newFrm.Show();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmLotes newFrm = new frmLotes();
            newFrm.Show();
        }

        private void linkLabel6_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmSupervisor newFrm = new frmSupervisor();
            newFrm.Show();
        }

        private void linkLabel7_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmLogs newFrm = new frmLogs();
            newFrm.Show();
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmUnidadDeManejo newFrm = new frmUnidadDeManejo();
            newFrm.Show();
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmPropietarios newFrm = new frmPropietarios();
            newFrm.Show();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmCargosLaborales newFrm = new frmCargosLaborales();
            newFrm.Show();
        }

        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmFormatos newFrm = new frmFormatos();
            newFrm.Show();
        }

        private void linkLabel8_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmCrearOrden newFrm = new frmCrearOrden(0,user);
            newFrm.Show();
        }

        private void linkLabel9_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
        }

        private void frmInicio_Load(object sender, EventArgs e)
        {

        }

        private void linkLabel11_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmMunicipios newFrm = new frmMunicipios();
            newFrm.Show();
        }

        private void linkLabel12_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmMaquinaPropietarios newFrm = new frmMaquinaPropietarios();
            newFrm.Show();
        }

        private void linkLabel13_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmMaquinaEstado newFrm = new frmMaquinaEstado("Maquinarias");
            newFrm.Show();
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            frmBiblioteca newFrm = new frmBiblioteca();
            newFrm.Show();
        }

        private void linkLabel10_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmProveedor newFrm = new frmProveedor();
            newFrm.Show();
        }

        private void linkLabel9_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmClientes newFrm = new frmClientes();
            newFrm.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            frmProductos newFrm = new frmProductos();
            newFrm.Show();
        }

        private void linkLabel14_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmCuchillas newFrm = new frmCuchillas(0);
            newFrm.Show();
        }

        private void linkLabel15_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmCrearProduccion newFrm = new frmCrearProduccion(0);
            newFrm.Show();
        }

        private void linkLabel16_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmPluviometria newFrm = new frmPluviometria();
            newFrm.Show();
        }

        private void linkLabel17_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmOrdenesAlamacen newFrm = new frmOrdenesAlamacen(0);
            newFrm.Show();
        }

        private void linkLabel18_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmCuadrilla newFrm = new frmCuadrilla();
            newFrm.Show();
        }

        private void button8_Click(object sender, EventArgs e)
        {
        }

        private void button7_Click(object sender, EventArgs e)
        {
            frmHistoricoProduccion newFrm = new frmHistoricoProduccion(0);
            newFrm.Show();
        }

        private void linkLabel19_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmReciboMadera newFrm = new frmReciboMadera();
            newFrm.Show();
        }

        private void linkLabel20_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmOrdenesCuchilla newFrm = new frmOrdenesCuchilla(0);
            newFrm.Show();
        }

        private void linkLabel21_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmSemillas newFrm = new frmSemillas();
            newFrm.Show();
        }

        private void linkLabel22_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmRecibos newFrm = new frmRecibos();
            newFrm.Show();
        }

    }
}
