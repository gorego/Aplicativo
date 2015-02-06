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
    public partial class frmOperDepar : Form
    {
        //Base de datos.
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        int opc = 0;
        public frmOperDepar(int number)
        {
            InitializeComponent();
            opc = number;
            string query = "";
            //Si number es igual a uno entra en el select del operador.
            if (number == 1) 
                query = "SELECT * FROM Operador";
            else 
            {
                //Sino en departamentos y se modifica el tamano del form.
                this.Text = "Departamentos"; 
                query = "SELECT d.ID, d.Departamento, (t.Nombres + ' ' + t.Apellidos) As Supervisor, t.Celular FROM Departamentos AS d INNER JOIN Trabajadores AS t ON d.ID = t.Departamento WHERE t.Supervisor = 'Si'";
            }
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable operadores = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(operadores);
            gridOperadores.DataSource = operadores;
            gridOperadores.Columns[1].DefaultCellStyle.Font = new Font(gridOperadores.DefaultCellStyle.Font, FontStyle.Underline);
            if(number != 1)
                gridOperadores.Columns[0].Visible = false;
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            //Cerrar ventana actual.
            this.Close();
        }

        private void gridOperadores_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            string tipo = "";
            if (opc == 1)
                tipo = "Operador";
            else
                tipo = "Departamento";
            if (gridOperadores.CurrentCell.ColumnIndex == 1)
            {
                frmOrdenes newFrm = new frmOrdenes(tipo,gridOperadores.Rows[gridOperadores.CurrentCell.RowIndex].Cells[0].Value.ToString());
                newFrm.Show();
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Variables.imprimir(gridOperadores);
        }
    }
}
