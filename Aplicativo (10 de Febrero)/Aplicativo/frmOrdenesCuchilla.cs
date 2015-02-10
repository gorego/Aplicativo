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
    public partial class frmOrdenesCuchilla : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        int tipousuario = 0;

        public frmOrdenesCuchilla(int tipo)
        {
            InitializeComponent();
            Variables.cargar(dataGridView2, "SELECT ID,OP,Tipo,Especie,fechaInicio,horaInicio,fechaFinal,Estado FROM historicoProduccion WHERE Estado <> 'Cerrada' ORDER BY ID Desc");
            formatoDataGridView(dataGridView2);
            dataGridView2.Columns[1].DefaultCellStyle.Font = new Font(dataGridView2.DefaultCellStyle.Font, FontStyle.Underline);
            estadoOrdenes();
            tipousuario = tipo;
        }

        public void formatoDataGridView(DataGridView data)
        {
            data.Columns[0].Visible = false;
            data.Columns[1].HeaderText = "OP #";
            data.Columns[4].HeaderText = "Fecha de Inicio";
            data.Columns[5].HeaderText = "Hora de Inicio";
            data.Columns[6].HeaderText = "Fecha de Terminación";
        }

        public void estadoOrdenes()
        {
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                DateTime fechaFinal = DateTime.ParseExact(dataGridView2.Rows[i].Cells[6].Value.ToString(), "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                if (DateTime.Compare(DateTime.Now.Date, fechaFinal.Date) > 0 && !dataGridView2.Rows[i].Cells[7].Value.Equals("Cerrada"))
                {
                    dataGridView2.Rows[i].Cells[7].Value = "Vencida";
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            frmCuchillas newFrm = new frmCuchillas(tipousuario);
            newFrm.Show();
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView2.CurrentCell.ColumnIndex == 1 && tipousuario == 0)
            {
                frmCuchillasOrden newFrm = new frmCuchillasOrden(Int32.Parse(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString()),0);
                this.Hide();
                newFrm.Show();
                this.Close();
            }
            else if (dataGridView2.CurrentCell.ColumnIndex == 1 && tipousuario != 0)
            {
                frmCuchillasOrden newFrm = new frmCuchillasOrden(Int32.Parse(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString()),1);
                newFrm.Show();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}
