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
    public partial class frmOrdenes : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();        

        public void cargarOrdenes(string tipo, string id)
        {
            string query;
            if (tipo.Equals("Lote"))
            {
                query = "SELECT h.ID, h.OT, a.Actividad, h.Area, h.fechaInicio, h.fechaFinal, (t.Nombres + ' ' + t.Apellidos) As Supervisor, (h.CostoFinal + h.costoJornalFinal), h.estadoOrden FROM Areas AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote WHERE h.estadoOrden = 'Cerrada' AND h." + tipo + " = " + id + " UNION ALL SELECT h.ID, h.OT, a.Actividad, h.Area, h.fechaInicio, h.fechaFinal, (t.Nombres + ' ' + t.Apellidos) As Supervisor, (h.CostoFinal + h.costoJornalFinal), h.estadoOrden FROM LoteGanadero AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote WHERE h.estadoOrden = 'Cerrada' AND h." + tipo + " = " + id + " UNION ALL SELECT h.ID, h.OT, a.Actividad, h.Area, h.fechaInicio, h.fechaFinal,(t.Nombres + ' ' + t.Apellidos) As Supervisor, (h.CostoFinal + h.costoJornalFinal), h.estadoOrden FROM Lotes AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote WHERE h.estadoOrden = 'Cerrada' AND h." + tipo + " = " + id;
            }
            else
            {
                query = "SELECT h.ID, h.OT, a.Actividad, area.Lote, h.fechaInicio, h.fechaFinal, (t.Nombres + ' ' + t.Apellidos) As Supervisor, (h.CostoFinal + h.costoJornalFinal), h.estadoOrden FROM Areas AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote WHERE h.estadoOrden = 'Cerrada' AND h." + tipo + " = " + id + " UNION ALL SELECT h.ID, h.OT, a.Actividad, area.Lote, h.fechaInicio, h.fechaFinal, (t.Nombres + ' ' + t.Apellidos) As Supervisor, (h.CostoFinal + h.costoJornalFinal), h.estadoOrden FROM LoteGanadero AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote WHERE h.estadoOrden = 'Cerrada' AND h." + tipo + " = " + id + " UNION ALL SELECT h.ID, h.OT, a.Actividad, area.Lote, h.fechaInicio, h.fechaFinal,(t.Nombres + ' ' + t.Apellidos) As Supervisor, (h.CostoFinal + h.costoJornalFinal), h.estadoOrden FROM Lotes AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote WHERE h.estadoOrden = 'Cerrada' AND h." + tipo + " = " + id;
            }
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable banco = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(banco);
            dataGridView2.DataSource = banco;
            dataGridView2.Columns[0].Visible = false;
            dataGridView2.Columns[1].HeaderText = "Orden de Trabajo #";
            dataGridView2.Columns[4].HeaderText = "Fecha de Inicio";
            dataGridView2.Columns[5].HeaderText = "Fecha de Finalización";
            dataGridView2.Columns[7].DefaultCellStyle.Format = "c";
            dataGridView2.Columns[7].HeaderText = "Costo Final";
            dataGridView2.Columns[8].HeaderText = "Estado de la Orden";
        }

        public void cargarOrdenes(string tipo, string id, string año)
        {
            string query;
            if (tipo.Equals("Lote"))
            {
                query = "SELECT h.ID, h.OT, a.Actividad, h.Area, h.fechaInicio, h.fechaFinal, (t.Nombres + ' ' + t.Apellidos) As Supervisor, (h.CostoFinal + h.costoJornalFinal), h.estadoOrden FROM Areas AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote WHERE h.fechaInicio LIKE '%" + año + "%' AND h.estadoOrden = 'Cerrada' AND h." + tipo + " = " + id + " UNION ALL SELECT h.ID, h.OT, a.Actividad, h.Area, h.fechaInicio, h.fechaFinal, (t.Nombres + ' ' + t.Apellidos) As Supervisor, (h.CostoFinal + h.costoJornalFinal), h.estadoOrden FROM LoteGanadero AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote WHERE h.fechaInicio LIKE '%" + año + "%' AND h.estadoOrden = 'Cerrada' AND h." + tipo + " = " + id + " UNION ALL SELECT h.ID, h.OT, a.Actividad, h.Area, h.fechaInicio, h.fechaFinal,(t.Nombres + ' ' + t.Apellidos) As Supervisor, (h.CostoFinal + h.costoJornalFinal), h.estadoOrden FROM Lotes AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote WHERE h.fechaInicio LIKE '%" + año + "%' AND h.estadoOrden = 'Cerrada' AND h." + tipo + " = " + id;
            }
            else
            {
                query = "SELECT h.ID, h.OT, a.Actividad, area.Lote, h.fechaInicio, h.fechaFinal, (t.Nombres + ' ' + t.Apellidos) As Supervisor, (h.CostoFinal + h.costoJornalFinal), h.estadoOrden FROM Areas AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote WHERE h.fechaInicio LIKE '%" + año + "%' AND h.estadoOrden = 'Cerrada' AND h." + tipo + " = " + id + " UNION ALL SELECT h.ID, h.OT, a.Actividad, area.Lote, h.fechaInicio, h.fechaFinal, (t.Nombres + ' ' + t.Apellidos) As Supervisor, (h.CostoFinal + h.costoJornalFinal), h.estadoOrden FROM LoteGanadero AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote WHERE h.fechaInicio LIKE '%" + año + "%' AND  h.estadoOrden = 'Cerrada' AND h." + tipo + " = " + id + " UNION ALL SELECT h.ID, h.OT, a.Actividad, area.Lote, h.fechaInicio, h.fechaFinal,(t.Nombres + ' ' + t.Apellidos) As Supervisor, (h.CostoFinal + h.costoJornalFinal), h.estadoOrden FROM Lotes AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote WHERE h.fechaInicio LIKE '%" + año + "%'  AND h.estadoOrden = 'Cerrada' AND h." + tipo + " = " + id;
            }
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable banco = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(banco);
            if (año.Equals(DateTime.Now.Year.ToString()))
            {
                dataGridView3.DataSource = banco;
                dataGridView3.Columns[0].Visible = false;
                dataGridView3.Columns[1].HeaderText = "Orden de Trabajo #";
                dataGridView3.Columns[4].HeaderText = "Fecha de Inicio";
                dataGridView3.Columns[5].HeaderText = "Fecha de Finalización";
                dataGridView3.Columns[7].DefaultCellStyle.Format = "c";
                dataGridView3.Columns[7].HeaderText = "Costo Final";
                dataGridView3.Columns[8].HeaderText = "Estado de la Orden";
            }
            else
            {
                dataGridView1.DataSource = banco;
                dataGridView1.Columns[0].Visible = false;
                dataGridView1.Columns[1].HeaderText = "Orden de Trabajo #";
                dataGridView1.Columns[4].HeaderText = "Fecha de Inicio";
                dataGridView1.Columns[5].HeaderText = "Fecha de Finalización";
                dataGridView1.Columns[7].DefaultCellStyle.Format = "c";
                dataGridView1.Columns[7].HeaderText = "Costo Final";
                dataGridView1.Columns[8].HeaderText = "Estado de la Orden";
            }
        }

        public void totalReparacion(DataGridView data)
        {            
            int total = 0;
            foreach (DataGridViewRow row in data.Rows)
            {
                total += Int32.Parse(row.Cells[7].Value.ToString());
            }
            label1.Text = "Total: " + String.Format("{0:c}",total);
        }

        public frmOrdenes(string tipo,string id)
        {
            InitializeComponent();
            cargarOrdenes(tipo, id,DateTime.Now.Year.ToString());
            cargarOrdenes(tipo, id, (DateTime.Now.Year-1).ToString());
            cargarOrdenes(tipo,id);
            totalReparacion(dataGridView3);
            dataGridView1.Columns[7].HeaderText = "Costo Final";
            dataGridView2.Columns[7].HeaderText = "Costo Final";
            dataGridView3.Columns[7].HeaderText = "Costo Final";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            frmCrearOrden newFrm = new frmCrearOrden(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString(), 1);
            newFrm.Show();
        }

        private void frmOrdenes_Load(object sender, EventArgs e)
        {

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (tabControl1.SelectedIndex == 0)
            {
                totalReparacion(dataGridView3);
            }
            if (tabControl1.SelectedIndex == 1)
            {
                totalReparacion(dataGridView1);
            }
            if (tabControl1.SelectedIndex == 2)
            {
                totalReparacion(dataGridView2);
            }
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            frmCrearOrden newFrm = new frmCrearOrden(dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value.ToString(), 1);
            newFrm.Show();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            frmCrearOrden newFrm = new frmCrearOrden(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString(), 1);
            newFrm.Show();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
                Variables.imprimir(dataGridView3);
            else if (tabControl1.SelectedIndex == 1)
                Variables.imprimir(dataGridView1);
            else
                Variables.imprimir(dataGridView2);
        }
    }
}
