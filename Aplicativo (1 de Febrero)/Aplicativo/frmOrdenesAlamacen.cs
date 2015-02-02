using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Aplicativo
{
    public partial class frmOrdenesAlamacen : Form
    {
        int tipousuario;
        public frmOrdenesAlamacen(int tipo)
        {
            InitializeComponent();
            tipousuario = tipo;
            cargarOrdenesActivas();
            Variables.agregarLog("Ingreso",Variables.userName);
        }

        public void cargarOrdenesActivas()
        {
            Variables.cargar(dataGridView2, "SELECT h.ID, h.OT, a.Actividad, area.Lote, h.fechaInicio, h.fechaFinal, (t.Nombres + ' ' + t.Apellidos) As Supervisor, h.estadoOrden FROM Areas AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote ORDER BY h.ID DESC UNION ALL SELECT h.ID, h.OT, a.Actividad, area.Lote, h.fechaInicio, h.fechaFinal, (t.Nombres + ' ' + t.Apellidos) As Supervisor, h.estadoOrden FROM LoteGanadero AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote ORDER BY h.ID DESC UNION ALL SELECT h.ID, h.OT, a.Actividad, area.Lote, h.fechaInicio, h.fechaFinal,(t.Nombres + ' ' + t.Apellidos) As Supervisor, h.estadoOrden FROM Lotes AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote ORDER BY h.ID DESC");
            dataGridView2.Columns[0].Visible = false;
            dataGridView2.Columns[1].HeaderText = "Orden de Trabajo #";
            dataGridView2.Columns[4].HeaderText = "Fecha de Inicio";
            dataGridView2.Columns[5].HeaderText = "Fecha de Finalización";
            dataGridView2.Columns[7].HeaderText = "Estado de la Orden";
            dataGridView2.Columns[1].DefaultCellStyle.Font = new Font(dataGridView2.DefaultCellStyle.Font, FontStyle.Underline);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            frmInsumos newFrm = new frmInsumos(tipousuario);
            newFrm.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView2.CurrentCell.ColumnIndex == 1)
            {
                frmInsumosOrdenes newFrm = new frmInsumosOrdenes(Int32.Parse(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString()));
                newFrm.Show();
            }
        }
    }
}
