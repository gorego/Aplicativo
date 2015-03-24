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
    public partial class frmHistoricoProduccion : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        int usuario = 0;
        public frmHistoricoProduccion(int tipo)
        {
            usuario = tipo;
            InitializeComponent();
            if (usuario == 1)
            {
                tabControl1.TabPages.Remove(tabPage2);
                button1.Enabled = false;
                button5.Enabled = true;
                button3.Visible = false;
                button4.Visible = false;
                button9.Enabled = true;
            }
            else
            {
                Variables.cargar(dataGridView1, "SELECT historicoProduccion.ID, historicoProduccion.OP, historicoProduccion.Tipo, historicoProduccion.Especie, historicoProduccion.fechaInicio, historicoProduccion.horaInicio, historicoProduccion.fechaFinal, SUM(produccionProducto.Volumen), historicoProduccion.Estado FROM historicoProduccion INNER JOIN produccionProducto ON historicoProduccion.Id = produccionProducto.Orden GROUP BY historicoProduccion.ID, historicoProduccion.OP, historicoProduccion.Tipo, historicoProduccion.Especie, historicoProduccion.fechaInicio, historicoProduccion.horaInicio, historicoProduccion.fechaFinal, historicoProduccion.Estado ORDER BY historicoProduccion.ID DESC;");
                formatoDataGridView(dataGridView1);
                roundpoint(dataGridView1);
                dataGridView1.Columns[1].DefaultCellStyle.Font = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Underline);
                estadoOrdenes(0);
                dataGridView1.Columns[0].Visible = false;
            }
            Variables.cargar(dataGridView2, "SELECT historicoProduccion.ID, historicoProduccion.OP, historicoProduccion.Tipo, historicoProduccion.Especie, historicoProduccion.fechaInicio, historicoProduccion.horaInicio, historicoProduccion.fechaFinal, SUM(produccionProducto.Volumen), historicoProduccion.Estado FROM historicoProduccion INNER JOIN produccionProducto ON historicoProduccion.Id = produccionProducto.Orden WHERE Estado <> 'Cerrada' GROUP BY historicoProduccion.ID, historicoProduccion.OP, historicoProduccion.Tipo, historicoProduccion.Especie, historicoProduccion.fechaInicio, historicoProduccion.horaInicio, historicoProduccion.fechaFinal, historicoProduccion.Estado ORDER BY historicoProduccion.ID DESC;");
            formatoDataGridView(dataGridView2);
            roundpoint(dataGridView2);
            dataGridView2.Columns[1].DefaultCellStyle.Font = new Font(dataGridView2.DefaultCellStyle.Font, FontStyle.Underline);
            estadoOrdenes(1);
            dataGridView2.Columns[0].Visible = false;
        }

        public void formatoDataGridView(DataGridView data)
        {
            data.Columns[0].Visible = false;
            data.Columns[1].HeaderText = "OP #";
            data.Columns[4].HeaderText = "Fecha de Inicio";
            data.Columns[5].HeaderText = "Hora de Inicio";
            data.Columns[6].HeaderText = "Fecha de Terminación";
            data.Columns[7].HeaderText = "Vol. a Producir";
        }
        private void button8_Click(object sender, EventArgs e)
        {
            frmReciboMadera newFrm = new frmReciboMadera();
            newFrm.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            frmCrearProduccion newFrm = new frmCrearProduccion(0);
            this.Hide();
            newFrm.Show();
            this.Close();
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView2.CurrentCell.ColumnIndex == 1 && usuario == 0)
            {
                frmCrearProduccion newFrm = new frmCrearProduccion(Int32.Parse(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString()),0);
                this.Hide();
                newFrm.Show();
                this.Close();
            }
            else if (dataGridView2.CurrentCell.ColumnIndex == 1 && usuario != 0)
            {
                frmCrearProduccion newFrm = new frmCrearProduccion(Int32.Parse(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString()),1);
                newFrm.Show();
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentCell.ColumnIndex == 1)
            {
                frmCrearProduccion newFrm = new frmCrearProduccion(Int32.Parse(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString()),0);
                this.Hide();
                newFrm.Show();
                this.Close();
            }
        }

        public void modificarOrden(string estado, DataGridView data)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE historicoProduccion SET Estado=@estadoOrden WHERE ID = " + data.Rows[data.CurrentCell.RowIndex].Cells[0].Value.ToString());
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@estadoOrden", OleDbType.VarChar).Value = estado;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Orden de trabajo " + estado + ".");
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

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de cerrar la " + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                modificarOrden("Cerrada",dataGridView2);
                Variables.cargar(dataGridView1, "SELECT ID,OP,Tipo,Especie,fechaInicio,horaInicio,fechaFinal,Estado FROM historicoProduccion ORDER BY ID Desc");
                formatoDataGridView(dataGridView1);
                dataGridView1.Columns[1].DefaultCellStyle.Font = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Underline);
                Variables.cargar(dataGridView2, "SELECT ID,OP,Tipo,Especie,fechaInicio,horaInicio,fechaFinal,Estado FROM historicoProduccion WHERE Estado <> 'Cerrada' ORDER BY ID Desc");
                formatoDataGridView(dataGridView2);
                dataGridView2.Columns[1].DefaultCellStyle.Font = new Font(dataGridView2.DefaultCellStyle.Font, FontStyle.Underline);
                estadoOrdenes(0);
                estadoOrdenes(1);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        public void estadoOrdenes(int tipo)
        {
            if (tipo == 0)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    DateTime fechaFinal = DateTime.ParseExact(dataGridView1.Rows[i].Cells[6].Value.ToString(), "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                    if (DateTime.Compare(DateTime.Now.Date, fechaFinal.Date) > 0 && !dataGridView1.Rows[i].Cells[8].Value.Equals("Cerrada"))
                    {
                        dataGridView1.Rows[i].Cells[8].Value = "Vencida";
                    }
                }
            }
            else
            {
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    DateTime fechaFinal = DateTime.ParseExact(dataGridView2.Rows[i].Cells[6].Value.ToString(), "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                    if (DateTime.Compare(DateTime.Now.Date, fechaFinal.Date) > 0 && !dataGridView2.Rows[i].Cells[8].Value.Equals("Cerrada"))
                    {
                        dataGridView2.Rows[i].Cells[8].Value = "Vencida";
                    }
                }
            }
        }

        public void roundpoint(DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                data.Rows[i].Cells[7].Value = Math.Round(Convert.ToDouble(data.Rows[i].Cells[7].Value.ToString()), 3, MidpointRounding.AwayFromZero);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de abrir la " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                modificarOrden("Activa",dataGridView1);
                Variables.cargar(dataGridView1, "SELECT ID,OP,Tipo,Especie,fechaInicio,horaInicio,fechaFinal,Estado FROM historicoProduccion ORDER BY ID Desc");
                formatoDataGridView(dataGridView1);
                dataGridView1.Columns[1].DefaultCellStyle.Font = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Underline);
                Variables.cargar(dataGridView2, "SELECT ID,OP,Tipo,Especie,fechaInicio,horaInicio,fechaFinal,Estado FROM historicoProduccion WHERE Estado <> 'Cerrada' ORDER BY ID Desc");
                formatoDataGridView(dataGridView2);
                dataGridView2.Columns[1].DefaultCellStyle.Font = new Font(dataGridView2.DefaultCellStyle.Font, FontStyle.Underline);
                estadoOrdenes(0);
                estadoOrdenes(1);
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
                Variables.imprimir(dataGridView2);
            else
                Variables.imprimir(dataGridView1);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            frmAsistenciaFijos newFrm = new frmAsistenciaFijos(Variables.userID.ToString(), 0);
            if (!newFrm.IsDisposed) { newFrm.Show(); }  
        }

        private void button9_Click(object sender, EventArgs e)
        {
            frmAsistenciaOrdenes newFrm = new frmAsistenciaOrdenes();
            if (!newFrm.IsDisposed) { newFrm.Show(); }  
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            frmReciboMaderCliente newFrm = new frmReciboMaderCliente();
            newFrm.Show();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            frmBodega newFrm = new frmBodega();
            newFrm.Show();
        }
    }
}
