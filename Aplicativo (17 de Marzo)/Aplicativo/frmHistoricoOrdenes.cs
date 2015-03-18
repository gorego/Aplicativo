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
    public partial class frmHistoricoOrdenes : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        int usuario;
        string user;
        int inicio = 0;

        public void cargarOrdenes()
        {
            string query = "SELECT h.ID, h.OT, a.Actividad, area.Lote, h.fechaInicio, h.fechaFinal, (t.Nombres + ' ' + t.Apellidos) As Supervisor, h.estadoOrden, h.costoFinal, h.costoJornalFinal, (h.costoFinal + h.costoJornalFinal) FROM Areas AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote ORDER BY h.ID desc UNION ALL SELECT h.ID, h.OT, a.Actividad, area.Lote, h.fechaInicio, h.fechaFinal, (t.Nombres + ' ' + t.Apellidos) As Supervisor, h.estadoOrden, h.costoFinal, h.costoJornalFinal, (h.costoFinal + h.costoJornalFinal) FROM LoteGanadero AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote ORDER BY h.ID desc UNION ALL SELECT h.ID, h.OT, a.Actividad, area.Lote, h.fechaInicio, h.fechaFinal,(t.Nombres + ' ' + t.Apellidos) As Supervisor, h.estadoOrden, h.costoFinal, h.costoJornalFinal, (h.costoFinal + h.costoJornalFinal) FROM Lotes AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote ORDER BY h.ID desc";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable banco = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(banco);
            dataGridView1.DataSource = banco;
            dataGridView1.Columns[1].HeaderText = "Orden de Trabajo #";
            dataGridView1.Columns[4].HeaderText = "Fecha de Inicio";
            dataGridView1.Columns[5].HeaderText = "Fecha de Finalización";
            dataGridView1.Columns[7].HeaderText = "Estado de la Orden";
            dataGridView1.Columns[8].HeaderText = "Costo Insumos";
            dataGridView1.Columns[9].HeaderText = "Costo Jornal";
            dataGridView1.Columns[10].HeaderText = "Costo Final";
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[8].DefaultCellStyle.Format = "c";
            dataGridView1.Columns[9].DefaultCellStyle.Format = "c";
            dataGridView1.Columns[10].DefaultCellStyle.Format = "c";
            dataGridView1.Columns[8].Visible = false;
            dataGridView1.Columns[9].Visible = false;
            dataGridView1.Columns[10].Visible = false;
        }

        public void cargarOrdenesActivas()
        {
            string query = "SELECT h.ID, h.OT, a.Actividad, area.Lote, h.fechaInicio, h.fechaFinal, (t.Nombres + ' ' + t.Apellidos) As Supervisor, h.estadoOrden, h.costoFinal, h.costoJornalFinal, (h.costoFinal + h.costoJornalFinal) FROM Areas AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote WHERE h.estadoOrden = 'Activa' OR h.estadoOrden = 'Vencida' ORDER BY h.ID desc UNION ALL SELECT h.ID, h.OT, a.Actividad, area.Lote, h.fechaInicio, h.fechaFinal, (t.Nombres + ' ' + t.Apellidos) As Supervisor, h.estadoOrden, h.costoFinal, h.costoJornalFinal, (h.costoFinal + h.costoJornalFinal) FROM LoteGanadero AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote WHERE h.estadoOrden = 'Activa' OR h.estadoOrden = 'Vencida' ORDER BY h.ID desc UNION ALL SELECT h.ID, h.OT, a.Actividad, area.Lote, h.fechaInicio, h.fechaFinal,(t.Nombres + ' ' + t.Apellidos) As Supervisor, h.estadoOrden, h.costoFinal, h.costoJornalFinal, (h.costoFinal + h.costoJornalFinal) FROM Lotes AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote WHERE h.estadoOrden = 'Activa' OR h.estadoOrden = 'Vencida' ORDER BY h.ID desc";       
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable banco = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(banco);
            dataGridView2.DataSource = banco;
            dataGridView2.Columns[1].HeaderText = "Orden de Trabajo #";
            dataGridView2.Columns[4].HeaderText = "Fecha de Inicio";
            dataGridView2.Columns[5].HeaderText = "Fecha de Finalización";
            dataGridView2.Columns[7].HeaderText = "Estado de la Orden";
            dataGridView2.Columns[8].HeaderText = "Costo Insumos";
            dataGridView2.Columns[9].HeaderText = "Costo Jornal";
            dataGridView2.Columns[10].HeaderText = "Costo Final";
            dataGridView2.Columns[0].Visible = false;
            dataGridView2.Columns[8].Visible = false;
            dataGridView2.Columns[9].Visible = false;
            dataGridView2.Columns[10].Visible = false;
        }

        public void cargarOrdenesActivas(int id)
        {
            string query = "SELECT h.ID, h.OT, a.Actividad, area.Lote, h.fechaInicio, h.fechaFinal, (t.Nombres + ' ' + t.Apellidos) As Supervisor, h.estadoOrden FROM Areas AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote WHERE (h.estadoOrden = 'Activa' OR h.estadoOrden = 'Vencida') AND h.Supervisor = " + id + " ORDER BY h.ID desc UNION ALL SELECT h.ID, h.OT, a.Actividad, area.Lote, h.fechaInicio, h.fechaFinal, (t.Nombres + ' ' + t.Apellidos) As Supervisor, h.estadoOrden FROM LoteGanadero AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote WHERE (h.estadoOrden = 'Activa' OR h.estadoOrden = 'Vencida') AND h.Supervisor = " + id + " ORDER BY h.ID desc UNION ALL SELECT h.ID, h.OT, a.Actividad, area.Lote, h.fechaInicio, h.fechaFinal,(t.Nombres + ' ' + t.Apellidos) As Supervisor, h.estadoOrden FROM Lotes AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote WHERE (h.estadoOrden = 'Activa' OR h.estadoOrden = 'Vencida') AND h.Supervisor = " + id + " ORDER BY h.ID desc";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable banco = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(banco);
            dataGridView2.DataSource = banco;
            dataGridView2.Columns[1].HeaderText = "Orden de Trabajo #";
            dataGridView2.Columns[4].HeaderText = "Fecha de Inicio";
            dataGridView2.Columns[5].HeaderText = "Fecha de Finalización";
            dataGridView2.Columns[7].HeaderText = "Estado de la Orden";
            dataGridView2.Columns[0].Visible = false;
        }

        public void estadoOrdenes()
        {
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                DateTime fechaFinal = DateTime.ParseExact(dataGridView2.Rows[i].Cells[5].Value.ToString(), "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                if(DateTime.Compare(DateTime.Now.Date,fechaFinal.Date) > 0 && !dataGridView2.Rows[i].Cells[7].Value.Equals("Cerrada")){
                    dataGridView2.Rows[i].Cells[7].Value = "Vencida";
                }
            }
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                DateTime fechaFinal = DateTime.ParseExact(dataGridView1.Rows[i].Cells[5].Value.ToString(), "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                if (DateTime.Compare(DateTime.Now.Date, fechaFinal.Date) > 0 && !dataGridView1.Rows[i].Cells[7].Value.Equals("Cerrada"))
                {
                    dataGridView1.Rows[i].Cells[7].Value = "Vencida";
                }
            }
        }

        public frmHistoricoOrdenes(int nombre,string username)
        {
            InitializeComponent();
            usuario = nombre;
            cargarOrdenes();
            cargarOrdenesActivas();
            estadoOrdenes();
            user = username;
            dataGridView1.Columns[1].DefaultCellStyle.Font = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Underline);
            dataGridView2.Columns[1].DefaultCellStyle.Font = new Font(dataGridView2.DefaultCellStyle.Font, FontStyle.Underline);
            button5.Enabled = false;
            dataGridView1.Columns[0].Visible = false;
            dataGridView2.Columns[0].Visible = false;
            dataGridView2.Sort(dataGridView2.Columns[0], ListSortDirection.Descending);
            dataGridView1.Sort(dataGridView1.Columns[0], ListSortDirection.Descending);
        }

        public frmHistoricoOrdenes(int nombre,int admin, string username)
        {
            InitializeComponent();            
            usuario = nombre;
            user = username;
            tabControl1.TabPages.Remove(tabPage2);
            cargarOrdenesActivas(nombre);
            button1.Enabled = false;
            button3.Visible = false;
            button4.Visible = false;
            dataGridView2.Height = 318;
            estadoOrdenes();
            dataGridView2.Columns[0].Visible = false;
            dataGridView2.Columns[6].Visible = false;
            inicio = admin;
            dataGridView2.Columns[1].DefaultCellStyle.Font = new Font(dataGridView2.DefaultCellStyle.Font, FontStyle.Underline);
            //dataGridView2.Sort(dataGridView2.Columns[0], ListSortDirection.Ascending);
            agregarLog("Ingreso", username);
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


        private void button1_Click(object sender, EventArgs e)
        {
            frmCrearOrden newFrm = new frmCrearOrden(0,user);
            this.Hide();
            newFrm.Show();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridView2_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //for (int i = 0; i < dataGridView2.Rows.Count; i++)
            //{
            //    if (dataGridView2.Rows[i].Cells[6].Value.Equals("Vencida"))
            //    {
            //        dataGridView2.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
            //    }
            //}
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //for (int i = 0; i < dataGridView1.Rows.Count; i++)
            //{
            //    if (dataGridView1.Rows[i].Cells[6].Value.Equals("Vencida"))
            //        dataGridView1.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
            //}                  
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView2.CurrentCell.ColumnIndex == 1 && inicio == 0)
            {
                frmCrearOrden newFrm = new frmCrearOrden(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString(),0);
                this.Hide();
                newFrm.Show();
                this.Close();
            }
            else if (dataGridView2.CurrentCell.ColumnIndex == 1 && inicio != 0)
            {
                frmCrearOrden newFrm = new frmCrearOrden(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString(),1);
                newFrm.Show();
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentCell.ColumnIndex == 1)
            {
                frmCrearOrden newFrm = new frmCrearOrden(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString(),0);
                this.Hide();
                newFrm.Show();
                this.Close();
            }
        }

        public void cerrarOrden()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE historicoOrdenes SET estadoOrden=@estadoOrden WHERE ID = " + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString());
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@estadoOrden", OleDbType.VarChar).Value = "Cerrada";
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Orden de trabajo cerrada.");
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

        public void abrirOrden()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE historicoOrdenes SET estadoOrden=@estadoOrden WHERE ID = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@estadoOrden", OleDbType.VarChar).Value = "Activa";
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Orden de trabajo abierta.");
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
                cerrarOrden();
                cargarOrdenes();
                cargarOrdenesActivas();
                estadoOrdenes();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el orden de trabajo # " + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {                
                string id = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM historicoOrdenes WHERE id = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Orden de trabajo eliminada.");
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
                cargarOrdenesActivas();
                cargarOrdenes();
                estadoOrdenes();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            frmAsistenciaFijos newFrm = new frmAsistenciaFijos(usuario.ToString(),0);
            if (!newFrm.IsDisposed) { newFrm.Show(); }            
        }

        public void imprimirOrdenes(DataGridView data)
        {
            if (data.Rows.Count > 0)
            {
                Microsoft.Office.Interop.Excel.Application XcelApp = new Microsoft.Office.Interop.Excel.Application();
                XcelApp.Application.Workbooks.Add(Type.Missing);
                Microsoft.Office.Interop.Excel.Range excelCellrange;
                for (int i = 2; i < data.Columns.Count + 1; i++)
                {
                    XcelApp.Cells[2, i + 1] = data.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < data.Rows.Count; i++)
                {
                    for (int j = 1; j < data.Columns.Count; j++)
                    {
                        if (j == data.Columns.Count - 1 || j == data.Columns.Count - 2 || j == data.Columns.Count - 3)
                        {
                            XcelApp.Cells[i + 3, j + 2] = String.Format("{0:c}",data.Rows[i].Cells[j].Value);
                        }
                        else
                        {
                            XcelApp.Cells[i + 3, j + 2] = data.Rows[i].Cells[j].Value.ToString();
                        }
                        if (i == 0)
                        {
                            excelCellrange = XcelApp.Range[XcelApp.Cells[i + 2, 3], XcelApp.Cells[i + 2, data.Columns.Count + 1]];
                            excelCellrange.Interior.Color = System.Drawing.Color.LightGreen;
                            excelCellrange.AutoFilter(1);
                            //excelCellrange.Font.Color = System.Drawing.Color.White;
                        }
                    }
                }
                excelCellrange = XcelApp.Range[XcelApp.Cells[2, 3], XcelApp.Cells[data.Rows.Count + 2, data.Columns.Count + 1]];
                excelCellrange.EntireColumn.AutoFit();
                Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;
                XcelApp.Columns.AutoFit();
                XcelApp.Visible = true;
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
                imprimirOrdenes(dataGridView2);
            else
                imprimirOrdenes(dataGridView1);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el orden de trabajo # " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {
                string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM historicoOrdenes WHERE id = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Orden de trabajo eliminada.");
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
                cargarOrdenesActivas();
                cargarOrdenes();
                estadoOrdenes();
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de abrir la " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                abrirOrden();
                cargarOrdenes();
                cargarOrdenesActivas();
                estadoOrdenes();
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            cargarOrdenes();
            cargarOrdenesActivas();
            estadoOrdenes();
        }
    }
}
