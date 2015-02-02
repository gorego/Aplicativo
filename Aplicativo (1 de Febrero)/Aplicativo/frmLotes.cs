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
    public partial class frmLotes : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();

        public void cargarLotes() {
            while (dataGridView1.Rows.Count != 0)
            {
                dataGridView1.Rows.RemoveAt(0);
            }
            string query = "SELECT l.Codigo, l.Lote, b.Predio, l.Longitud, l.Latitud, l.Figura, p.Nombre, u.Unidad, m.Municipio, l.Especie, l.Ano, l.areaPlantacion, l.areaEfectiva,l.estadoFSC ,l.FSC, l.sumidero, l.CIF,l.restAmbiental,l.registroPlantacion,l.Ubicacion,l.semillaOrigen, l.extraidoEntresaca, l.extraidoTalaRaza, l.extraidoRecuperacion FROM UnidadDeManejo AS u INNER JOIN ((Municipio AS m INNER JOIN (Propietarios AS p INNER JOIN BancoTierras AS b ON p.ID = b.Propietario) ON m.ID = b.Municipio) INNER JOIN Lotes AS l ON b.ID = l.Predio) ON u.ID = l.Unidad ORDER BY l.Codigo";
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
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[i].Cells[0].Value = myReader.GetInt32(0);
                    dataGridView1.Rows[i].Cells[1].Value = myReader.GetString(1);
                    dataGridView1.Rows[i].Cells[2].Value = myReader.GetString(2);
                    dataGridView1.Rows[i].Cells[3].Value = myReader.GetValue(3);
                    dataGridView1.Rows[i].Cells[4].Value = myReader.GetValue(4);
                    dataGridView1.Rows[i].Cells[5].Value = myReader.GetString(5);
                    dataGridView1.Rows[i].Cells[6].Value = myReader.GetString(6);
                    dataGridView1.Rows[i].Cells[7].Value = myReader.GetString(7);
                    dataGridView1.Rows[i].Cells[8].Value = myReader.GetString(8);
                    dataGridView1.Rows[i].Cells[9].Value = myReader.GetString(9);
                    dataGridView1.Rows[i].Cells[10].Value = myReader.GetInt32(10);
                    dataGridView1.Rows[i].Cells[11].Value = myReader.GetString(11);
                    dataGridView1.Rows[i].Cells[12].Value = myReader.GetString(12);
                    dataGridView1.Rows[i].Cells[13].Value = myReader.GetString(13);
                    dataGridView1.Rows[i].Cells[14].Value = myReader.GetString(14);
                    dataGridView1.Rows[i].Cells[15].Value = myReader.GetString(15);
                    dataGridView1.Rows[i].Cells[16].Value = myReader.GetString(16);
                    dataGridView1.Rows[i].Cells[17].Value = myReader.GetInt32(17);
                    dataGridView1.Rows[i].Cells[18].Value = myReader.GetInt32(18);
                    dataGridView1.Rows[i].Cells[19].Value = myReader.GetString(19);
                    dataGridView1.Rows[i].Cells[20].Value = myReader.GetString(20);
                    dataGridView1.Rows[i].Cells[21].Value = Math.Round(myReader.GetDouble(21), 4, MidpointRounding.AwayFromZero);
                    dataGridView1.Rows[i].Cells[22].Value = Math.Round(myReader.GetDouble(22), 4, MidpointRounding.AwayFromZero);
                    dataGridView1.Rows[i].Cells[23].Value = Math.Round(myReader.GetDouble(23), 4, MidpointRounding.AwayFromZero);
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

        public void cargarCampamento()
        {
            while (dataGridView5.Rows.Count != 0)
            {
                dataGridView5.Rows.RemoveAt(0);
            }
            string query = "SELECT c.ID, b.Predio, c.Longitud, c.Latitud, c.codCamp, c.Ocupacion, c.tipoMaterial, c.materialPred, c.Energia, c.agua, c.aguaAseo, c.otroServicio FROM Campamentos AS c INNER JOIN BancoTierras AS b ON c.Predio = b.ID;";
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
                    dataGridView5.Rows.Add();
                    dataGridView5.Rows[i].Cells[0].Value = myReader.GetInt32(0);
                    dataGridView5.Rows[i].Cells[1].Value = myReader.GetString(1);
                    dataGridView5.Rows[i].Cells[2].Value = myReader.GetString(2);
                    dataGridView5.Rows[i].Cells[3].Value = myReader.GetString(3);
                    dataGridView5.Rows[i].Cells[4].Value = myReader.GetString(4);
                    dataGridView5.Rows[i].Cells[5].Value = myReader.GetString(5);
                    dataGridView5.Rows[i].Cells[6].Value = myReader.GetString(6);
                    dataGridView5.Rows[i].Cells[7].Value = myReader.GetString(7);
                    dataGridView5.Rows[i].Cells[8].Value = myReader.GetString(8);
                    dataGridView5.Rows[i].Cells[9].Value = myReader.GetString(9);
                    dataGridView5.Rows[i].Cells[10].Value = myReader.GetString(10);
                    dataGridView5.Rows[i].Cells[11].Value = myReader.GetString(11);
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

        public void cargarLotesGanadero()
        {
            while (dataGridView3.Rows.Count != 0)
            {
                dataGridView3.Rows.RemoveAt(0);
            }
            string query = "SELECT l.Codigo, l.Lote, b.Predio, l.Longitud, l.Latitud, p.Nombre, u.Unidad, l.Pasto, m.Municipio, l.Area, l.Renovacion, l.Riesgo, l.Henificable,l.Ubicacion FROM (Propietarios AS p INNER JOIN (Municipio AS m INNER JOIN BancoTierras AS b ON m.ID = b.Municipio) ON p.ID = b.Propietario) INNER JOIN (UnidadDeManejo AS u INNER JOIN LoteGanadero AS l ON u.ID = l.Unidad) ON b.ID = l.Predio;";
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
                    dataGridView3.Rows[i].Cells[0].Value = myReader.GetInt32(0);
                    dataGridView3.Rows[i].Cells[1].Value = myReader.GetString(1);
                    dataGridView3.Rows[i].Cells[2].Value = myReader.GetString(2);
                    dataGridView3.Rows[i].Cells[3].Value = myReader.GetValue(3);
                    dataGridView3.Rows[i].Cells[4].Value = myReader.GetValue(4);
                    dataGridView3.Rows[i].Cells[5].Value = myReader.GetString(5);
                    dataGridView3.Rows[i].Cells[6].Value = myReader.GetString(6);
                    dataGridView3.Rows[i].Cells[7].Value = myReader.GetString(7);
                    dataGridView3.Rows[i].Cells[8].Value = myReader.GetString(8);
                    dataGridView3.Rows[i].Cells[9].Value = myReader.GetValue(9);
                    dataGridView3.Rows[i].Cells[10].Value = myReader.GetInt32(10);
                    dataGridView3.Rows[i].Cells[11].Value = myReader.GetString(11);
                    dataGridView3.Rows[i].Cells[12].Value = myReader.GetString(12);
                    dataGridView3.Rows[i].Cells[13].Value = myReader.GetString(13);                    
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

        public void cargarAreas(int clase)
        {
            if (clase == 1)
            {
                while (dataGridView2.Rows.Count != 0)
                {
                    dataGridView2.Rows.RemoveAt(0);
                }
            }
            else
            {
                while (dataGridView4.Rows.Count != 0)
                {
                    dataGridView4.Rows.RemoveAt(0);
                }
            }
            string query = "SELECT a.Codigo, a.Lote, b.Predio, a.Longitud, a.Latitud, p.Nombre, u.Unidad, m.Municipio, a.Area,a.Ubicacion FROM Municipio AS m INNER JOIN (Propietarios AS p INNER JOIN (UnidadDeManejo AS u INNER JOIN (BancoTierras AS b INNER JOIN Areas AS a ON b.ID = a.Predio) ON u.ID = a.Unidad) ON p.ID = b.Propietario) ON m.ID = b.Municipio WHERE a.Clase = " + clase;
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
                    if (clase == 1)
                    {
                        dataGridView2.Rows.Add();
                        dataGridView2.Rows[i].Cells[0].Value = myReader.GetInt32(0);
                        dataGridView2.Rows[i].Cells[1].Value = myReader.GetString(1);
                        dataGridView2.Rows[i].Cells[2].Value = myReader.GetString(2);
                        dataGridView2.Rows[i].Cells[3].Value = myReader.GetValue(3);
                        dataGridView2.Rows[i].Cells[4].Value = myReader.GetValue(4);
                        dataGridView2.Rows[i].Cells[5].Value = myReader.GetString(5);
                        dataGridView2.Rows[i].Cells[6].Value = myReader.GetString(6);
                        dataGridView2.Rows[i].Cells[7].Value = myReader.GetString(7);
                        dataGridView2.Rows[i].Cells[8].Value = myReader.GetValue(8);
                        dataGridView2.Rows[i].Cells[9].Value = myReader.GetString(9);
                    }
                    else
                    {
                        dataGridView4.Rows.Add();
                        dataGridView4.Rows[i].Cells[0].Value = myReader.GetInt32(0);
                        dataGridView4.Rows[i].Cells[1].Value = myReader.GetString(1);
                        dataGridView4.Rows[i].Cells[2].Value = myReader.GetString(2);
                        dataGridView4.Rows[i].Cells[3].Value = myReader.GetValue(3);
                        dataGridView4.Rows[i].Cells[4].Value = myReader.GetValue(4);
                        dataGridView4.Rows[i].Cells[5].Value = myReader.GetString(5);
                        dataGridView4.Rows[i].Cells[6].Value = myReader.GetString(6);
                        dataGridView4.Rows[i].Cells[7].Value = myReader.GetString(7);
                        dataGridView4.Rows[i].Cells[8].Value = myReader.GetValue(8);
                        dataGridView4.Rows[i].Cells[9].Value = myReader.GetString(9);
                    }
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

        public void modificarArea(int clase)
        {
            string codigo = "";
            if (clase == 1) 
            {
                codigo = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString();
            }
            else
            {
                codigo = dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[0].Value.ToString();
            }
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Areas SET Codigo=@Codigo,Lote=@Lote,Predio=@Predio,Longitud=@Longitud,Latitud=@Latitud,Unidad=@Unidad,Area=@Area,Ubicacion=@Ubicacion WHERE Codigo = " + codigo);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                if (clase == 1)
                {
                    cmd.Parameters.Add("@Codigo", OleDbType.VarChar).Value = txtCodProt.Text;
                    cmd.Parameters.Add("@Lote", OleDbType.VarChar).Value = txtLoteProt.Text;
                    cmd.Parameters.Add("@Predio", OleDbType.VarChar).Value = txtPredProt.SelectedValue;
                    cmd.Parameters.Add("@Longitud", OleDbType.VarChar).Value = txtLongProt.Text;
                    cmd.Parameters.Add("@Latitud", OleDbType.VarChar).Value = txtLatProt.Text;
                    cmd.Parameters.Add("@Unidad", OleDbType.VarChar).Value = txtUnidadProt.SelectedValue;
                    cmd.Parameters.Add("@Area", OleDbType.VarChar).Value = txtAreaProt.Text;
                    cmd.Parameters.Add("@Ubicacion", OleDbType.VarChar).Value = textBox2.Text;
                }
                else
                {
                    cmd.Parameters.Add("@Codigo", OleDbType.VarChar).Value = txtCodAgua.Text;
                    cmd.Parameters.Add("@Lote", OleDbType.VarChar).Value = txtLoteAgua.Text;
                    cmd.Parameters.Add("@Predio", OleDbType.VarChar).Value = txtPredAgua.SelectedValue;
                    cmd.Parameters.Add("@Longitud", OleDbType.VarChar).Value = txtLongAgua.Text;
                    cmd.Parameters.Add("@Latitud", OleDbType.VarChar).Value = txtLatAgua.Text;
                    cmd.Parameters.Add("@Unidad", OleDbType.VarChar).Value = txtUnidadAgua.SelectedValue;
                    cmd.Parameters.Add("@Area", OleDbType.VarChar).Value = txtAreaAgua.Text;
                    cmd.Parameters.Add("@Ubicacion", OleDbType.VarChar).Value = textBox3.Text;
                }
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Lote modificado.");
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

        public void cargarUnidades() { 
            string query = "SELECT * FROM UnidadDeManejo";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            txtUnidad.DataSource = ds.Tables[0];
            txtUnidad.DisplayMember = "Unidad";
            txtUnidad.ValueMember = "ID";
            txtUnidad.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtUnidad.AutoCompleteSource = AutoCompleteSource.ListItems;
            txtUnidadAgua.DataSource = ds.Tables[0];
            txtUnidadAgua.DisplayMember = "Unidad";
            txtUnidadAgua.ValueMember = "ID";
            txtUnidadAgua.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtUnidadAgua.AutoCompleteSource = AutoCompleteSource.ListItems;
            txtUnidadProt.DataSource = ds.Tables[0];
            txtUnidadProt.DisplayMember = "Unidad";
            txtUnidadProt.ValueMember = "ID";
            txtUnidadProt.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtUnidadProt.AutoCompleteSource = AutoCompleteSource.ListItems;
            txtUnidadGan.DataSource = ds.Tables[0];
            txtUnidadGan.DisplayMember = "Unidad";
            txtUnidadGan.ValueMember = "ID";
            txtUnidadGan.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtUnidadGan.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void cargarPredios()
        {
            string query = "SELECT * FROM BancoTierras";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            txtPredio.DataSource = ds.Tables[0];
            txtPredio.DisplayMember = "Predio";
            txtPredio.ValueMember = "ID";
            txtPredio.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtPredio.AutoCompleteSource = AutoCompleteSource.ListItems;
            txtPredGan.DataSource = ds.Tables[0];
            txtPredGan.DisplayMember = "Predio";
            txtPredGan.ValueMember = "ID";
            txtPredGan.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtPredGan.AutoCompleteSource = AutoCompleteSource.ListItems;
            txtPredProt.DataSource = ds.Tables[0];
            txtPredProt.DisplayMember = "Predio";
            txtPredProt.ValueMember = "ID";
            txtPredProt.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtPredProt.AutoCompleteSource = AutoCompleteSource.ListItems;
            txtPredAgua.DataSource = ds.Tables[0];
            txtPredAgua.DisplayMember = "Predio";
            txtPredAgua.ValueMember = "ID";
            txtPredAgua.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtPredAgua.AutoCompleteSource = AutoCompleteSource.ListItems;
            Camp1.DataSource = ds.Tables[0];
            Camp1.DisplayMember = "Predio";
            Camp1.ValueMember = "ID";
            Camp1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            Camp1.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void agregarLote() {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Lotes (Codigo,Lote,Predio,Longitud,Latitud,Figura,Unidad,Especie,Ano,areaPlantacion,areaEfectiva,estadoFSC,FSC,sumidero,CIF,restAmbiental,registroPlantacion,Ubicacion,semillaOrigen,extraidoEntresaca,extraidoTalaRaza,extraidoRecuperacion) VALUES (@Codigo,@Lote,@Predio,@Longitud,@Latitud,@Figura,@Unidad,@Especie,@Ano,@areaPlantacion,@areaEfectiva,@estadoFSC,@FSC,@sumidero,@CIF,@restAmbiental,@registroPlantacion,@Ubicacion,@semillaOrigen,@extraidoEntresaca,@extraidoTalaRaza,@extraidoRecuperacion)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Codigo", OleDbType.VarChar).Value = txtCodigo.Text;
                cmd.Parameters.Add("@Lote", OleDbType.VarChar).Value = txtLote.Text;
                cmd.Parameters.Add("@Predio", OleDbType.VarChar).Value = txtPredio.SelectedValue;
                cmd.Parameters.Add("@Longitud", OleDbType.VarChar).Value = txtLong.Text;
                cmd.Parameters.Add("@Latitud", OleDbType.VarChar).Value = txtLat.Text;
                cmd.Parameters.Add("@Figura", OleDbType.VarChar).Value = txtFigura.Text;
                cmd.Parameters.Add("@Unidad", OleDbType.VarChar).Value = txtUnidad.SelectedValue;
                cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = txtEspecie.Text;
                cmd.Parameters.Add("@Ano", OleDbType.VarChar).Value = txtAno.Text;
                cmd.Parameters.Add("@areaPlantacion", OleDbType.VarChar).Value = txtAreaPlant.Text;
                cmd.Parameters.Add("@areaEfectiva", OleDbType.VarChar).Value = txtAreaEfe.Text;
                cmd.Parameters.Add("@estadoFSC", OleDbType.VarChar).Value = txtFSC.Text;
                cmd.Parameters.Add("@FSC", OleDbType.VarChar).Value = txtNumFSC.Text;
                cmd.Parameters.Add("@sumidero", OleDbType.VarChar).Value = txtNumSum.Text;
                cmd.Parameters.Add("@CIF", OleDbType.VarChar).Value = txtNumCif.Text;
                cmd.Parameters.Add("@restAmbiental", OleDbType.VarChar).Value = txtAmbiental.Text;
                cmd.Parameters.Add("@registroPlantacion", OleDbType.VarChar).Value = txtPlantacion.Text;
                cmd.Parameters.Add("@Ubicacion", OleDbType.VarChar).Value = txtUbicacion.Text;
                cmd.Parameters.Add("@semillaOrigen", OleDbType.VarChar).Value = textBox7.Text;
                cmd.Parameters.Add("@extraidoEntresaca", OleDbType.VarChar).Value = textBox4.Text;
                cmd.Parameters.Add("@extraidoTalaRaza", OleDbType.VarChar).Value = textBox5.Text;
                cmd.Parameters.Add("@extraidoRecuperacion", OleDbType.VarChar).Value = textBox6.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Lote agregado.");
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

        public void agregarCampamento()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Campamentos (Predio,Longitud,Latitud,codCamp,Ocupacion,tipoMaterial,materialPred,Energia,agua,aguaAseo,otroServicio) VALUES (@Predio,@Longitud,@Latitud,@codCamp,@Ocupacion,@tipoMaterial,@materialPred,@Energia,@agua,@aguaAseo,@otroServicio)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Predio", OleDbType.VarChar).Value = Camp1.SelectedValue;                
                cmd.Parameters.Add("@Longitud", OleDbType.VarChar).Value = Camp4.Text;
                cmd.Parameters.Add("@Latitud", OleDbType.VarChar).Value = Camp5.Text;
                cmd.Parameters.Add("@codCamp", OleDbType.VarChar).Value = Camp6.Text;
                cmd.Parameters.Add("@tipoMaterial", OleDbType.VarChar).Value = Camp2.Text;
                cmd.Parameters.Add("@materialPred", OleDbType.VarChar).Value = Camp3.Text;
                cmd.Parameters.Add("@Ocupacion", OleDbType.VarChar).Value = Camp7.Text;                
                cmd.Parameters.Add("@Energia", OleDbType.VarChar).Value = Camp8.Text;
                cmd.Parameters.Add("@agua", OleDbType.VarChar).Value = Camp9.Text;
                cmd.Parameters.Add("@aguaAseo", OleDbType.VarChar).Value = Camp10.Text;
                cmd.Parameters.Add("@otroServicio", OleDbType.VarChar).Value = Camp11.Text;                
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Campamento agregado.");
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

        public void modificarCampamento()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Campamentos SET Predio=@Predio,Longitud=@Longitud,Latitud=@Latitud,codCamp=@codCamp,Ocupacion=@Ocupacion,tipoMaterial=@tipoMaterial,materialPred=@materialPred,Energia=@Energia,agua=@agua,aguaAseo=@aguaAseo,otroServicio=@otroServicio WHERE ID = " + dataGridView5.Rows[dataGridView5.CurrentCell.RowIndex].Cells[0].Value.ToString());
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Predio", OleDbType.VarChar).Value = Camp1.SelectedValue;
                cmd.Parameters.Add("@Longitud", OleDbType.VarChar).Value = Camp4.Text;
                cmd.Parameters.Add("@Latitud", OleDbType.VarChar).Value = Camp5.Text;
                cmd.Parameters.Add("@codCamp", OleDbType.VarChar).Value = Camp6.Text;
                cmd.Parameters.Add("@Ocupacion", OleDbType.VarChar).Value = Camp7.Text;
                cmd.Parameters.Add("@tipoMaterial", OleDbType.VarChar).Value = Camp2.Text;
                cmd.Parameters.Add("@materialPred", OleDbType.VarChar).Value = Camp3.Text;
                cmd.Parameters.Add("@Energia", OleDbType.VarChar).Value = Camp8.Text;
                cmd.Parameters.Add("@agua", OleDbType.VarChar).Value = Camp9.Text;
                cmd.Parameters.Add("@aguaAseo", OleDbType.VarChar).Value = Camp10.Text;
                cmd.Parameters.Add("@otroServicio", OleDbType.VarChar).Value = Camp11.Text; 
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Lote modificado.");
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

        public void agregarLoteGanadero()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO LoteGanadero(Codigo,Lote,Predio,Longitud,Latitud,Unidad,Pasto,Area,Renovacion,Riesgo,Henificable,Ubicacion) VALUES (@Codigo,@Lote,@Predio,@Longitud,@Latitud,@Unidad,@Pasto,@Area,@Renovacion,@Riesgo,@Henificable,@Ubicacion)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Codigo", OleDbType.VarChar).Value = txtCodGan.Text;
                cmd.Parameters.Add("@Lote", OleDbType.VarChar).Value = txtLotGan.Text;
                cmd.Parameters.Add("@Predio", OleDbType.VarChar).Value = txtPredGan.SelectedValue;
                cmd.Parameters.Add("@Longitud", OleDbType.VarChar).Value = txtLongGan.Text;
                cmd.Parameters.Add("@Latitud", OleDbType.VarChar).Value = txtLatGan.Text;                
                cmd.Parameters.Add("@Unidad", OleDbType.VarChar).Value = txtUnidadGan.SelectedValue;                
                cmd.Parameters.Add("@Pasto", OleDbType.VarChar).Value = txtPastoGan.Text;
                cmd.Parameters.Add("@Area", OleDbType.VarChar).Value = txtAreaGan.Text;
                cmd.Parameters.Add("@Renovacion", OleDbType.VarChar).Value = txtRenoGan.Text;
                cmd.Parameters.Add("@Riesgo", OleDbType.VarChar).Value = txtRiesgoGan.Text;
                cmd.Parameters.Add("@Henificable", OleDbType.VarChar).Value = txtHenGan.Text;
                cmd.Parameters.Add("@Ubicacion", OleDbType.VarChar).Value = textBox1.Text;                                       
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Lote agregado.");
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

        public void agregarArea(int clase)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Areas(Codigo,Lote,Predio,Longitud,Latitud,Unidad,Area,Clase,Ubicacion) VALUES (@Codigo,@Lote,@Predio,@Longitud,@Latitud,@Unidad,@Area,@Clase,@Ubicacion)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                if (clase == 1)
                {
                    cmd.Parameters.Add("@Codigo", OleDbType.VarChar).Value = txtCodProt.Text;
                    cmd.Parameters.Add("@Lote", OleDbType.VarChar).Value = txtLoteProt.Text;
                    cmd.Parameters.Add("@Predio", OleDbType.VarChar).Value = txtPredProt.SelectedValue;
                    cmd.Parameters.Add("@Longitud", OleDbType.VarChar).Value = txtLongProt.Text;
                    cmd.Parameters.Add("@Latitud", OleDbType.VarChar).Value = txtLatProt.Text;
                    cmd.Parameters.Add("@Unidad", OleDbType.VarChar).Value = txtUnidadProt.SelectedValue;
                    cmd.Parameters.Add("@Area", OleDbType.VarChar).Value = txtAreaProt.Text;
                    cmd.Parameters.Add("@Unidad", OleDbType.VarChar).Value = "1";
                    cmd.Parameters.Add("@Ubicacion", OleDbType.VarChar).Value = textBox2.Text;
                }
                else
                {
                    cmd.Parameters.Add("@Codigo", OleDbType.VarChar).Value = txtCodAgua.Text;
                    cmd.Parameters.Add("@Lote", OleDbType.VarChar).Value = txtLoteAgua.Text;
                    cmd.Parameters.Add("@Predio", OleDbType.VarChar).Value = txtPredAgua.SelectedValue;
                    cmd.Parameters.Add("@Longitud", OleDbType.VarChar).Value = txtLongAgua.Text;
                    cmd.Parameters.Add("@Latitud", OleDbType.VarChar).Value = txtLatAgua.Text;
                    cmd.Parameters.Add("@Unidad", OleDbType.VarChar).Value = txtUnidadAgua.SelectedValue;
                    cmd.Parameters.Add("@Area", OleDbType.VarChar).Value = txtAreaAgua.Text;
                    cmd.Parameters.Add("@Unidad", OleDbType.VarChar).Value = "2";
                    cmd.Parameters.Add("@Ubicacion", OleDbType.VarChar).Value = textBox3.Text;
                }
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Lote agregado.");
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

        public void modificarLote()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Lotes SET Codigo=@Codigo,Lote=@Lote,Predio=@Predio,Longitud=@Longitud,Latitud=@Latitud,Figura=@Figura,Unidad=@Unidad,Especie=@Especie,Ano=@Ano,areaPlantacion=@areaPlantacion,areaEfectiva=@areaEfectiva,estadoFSC=@estadoFSC,FSC=@FSC,sumidero=@sumidero,CIF=@CIF, restAmbiental=@restAmbiental,registroPlantacion=@registroPlantacion,Ubicacion=@Ubicacion,semillaOrigen=@semillaOrigen, extraidoEntresaca=@extraidoEntresaca,extraidoTalaRaza=@extraidoTalaRaza,extraidoRecuperacion=@extraidoRecuperacion WHERE Codigo = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Codigo", OleDbType.VarChar).Value = txtCodigo.Text;
                cmd.Parameters.Add("@Lote", OleDbType.VarChar).Value = txtLote.Text;
                cmd.Parameters.Add("@Predio", OleDbType.VarChar).Value = txtPredio.SelectedValue;
                cmd.Parameters.Add("@Longitud", OleDbType.VarChar).Value = txtLong.Text;
                cmd.Parameters.Add("@Latitud", OleDbType.VarChar).Value = txtLat.Text;
                cmd.Parameters.Add("@Figura", OleDbType.VarChar).Value = txtFigura.Text;
                cmd.Parameters.Add("@Unidad", OleDbType.VarChar).Value = txtUnidad.SelectedValue;
                cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = txtEspecie.Text;
                cmd.Parameters.Add("@Ano", OleDbType.VarChar).Value = txtAno.Text;
                cmd.Parameters.Add("@areaPlantacion", OleDbType.VarChar).Value = txtAreaPlant.Text;
                cmd.Parameters.Add("@areaEfectiva", OleDbType.VarChar).Value = txtAreaEfe.Text;
                cmd.Parameters.Add("@estadoFSC", OleDbType.VarChar).Value = txtFSC.Text;
                cmd.Parameters.Add("@FSC", OleDbType.VarChar).Value = txtNumFSC.Text;
                cmd.Parameters.Add("@sumidero", OleDbType.VarChar).Value = txtNumSum.Text;
                cmd.Parameters.Add("@CIF", OleDbType.VarChar).Value = txtNumCif.Text;
                cmd.Parameters.Add("@restAmbiental", OleDbType.VarChar).Value = txtAmbiental.Text;
                cmd.Parameters.Add("@registroPlantacion", OleDbType.VarChar).Value = txtPlantacion.Text;
                cmd.Parameters.Add("@Ubicacion", OleDbType.VarChar).Value = txtUbicacion.Text;
                cmd.Parameters.Add("@semillaOrigen", OleDbType.VarChar).Value = textBox7.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Lote modificado.");
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

        public void modificarLoteGanadero()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE LoteGanadero SET Codigo=@Codigo,Lote=@Lote,Predio=@Predio,Longitud=@Longitud,Latitud=@Latitud,Unidad=@Unidad,Pasto=@Pasto,Area=@Area,Renovacion=@Renovacion,Riesgo=@Riesgo,Henificable=@Henificable,Ubicacion=@Ubicacion WHERE Codigo = " + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value.ToString());
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Codigo", OleDbType.VarChar).Value = txtCodGan.Text;
                cmd.Parameters.Add("@Lote", OleDbType.VarChar).Value = txtLotGan.Text;
                cmd.Parameters.Add("@Predio", OleDbType.VarChar).Value = txtPredGan.SelectedValue;
                cmd.Parameters.Add("@Longitud", OleDbType.VarChar).Value = txtLongGan.Text;
                cmd.Parameters.Add("@Latitud", OleDbType.VarChar).Value = txtLatGan.Text;
                cmd.Parameters.Add("@Unidad", OleDbType.VarChar).Value = txtUnidadGan.SelectedValue;
                cmd.Parameters.Add("@Pasto", OleDbType.VarChar).Value = txtPastoGan.Text;
                cmd.Parameters.Add("@Area", OleDbType.VarChar).Value = txtAreaGan.Text;
                cmd.Parameters.Add("@Renovacion", OleDbType.VarChar).Value = txtRenoGan.Text;
                cmd.Parameters.Add("@Riesgo", OleDbType.VarChar).Value = txtRiesgoGan.Text;
                cmd.Parameters.Add("@Henificable", OleDbType.VarChar).Value = txtHenGan.Text;
                cmd.Parameters.Add("@Ubicacion", OleDbType.VarChar).Value = textBox1.Text;                                
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Lote modificado.");
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

        public void reiniciarTablero() {
            txtCodigo.Text = "";
            txtLote.Text = "";
            txtPredio.Text = "";
            txtLong.Text = "";
            txtLat.Text = "";
            txtFigura.Text = "";
            txtUnidad.Text = "";
            txtEspecie.Text = "";
            txtAno.Text = "";
            txtAreaPlant.Text = "";
            txtAreaEfe.Text = "";
            txtFSC.Text = "";
            txtNumFSC.Text = "";
            txtNumSum.Text = "";
            txtNumCif.Text = "";
            txtCif.Text = "";
            txtCarbono.Text = "";
            txtPlantacion.Text = "";
            txtUbicacion.Text = "";
            txtAmbiental.Text = "";
            textBox7.Text = "";
        }

        public void buscarLote()
        {
            while (dataGridView1.Rows.Count != 0)
            {
                dataGridView1.Rows.RemoveAt(0);
            }
            string query = "SELECT l.Codigo, l.Lote, b.Predio, l.Longitud, l.Latitud, l.Figura, p.Nombre, u.Unidad, m.Municipio, l.Especie, l.Ano, l.areaPlantacion, l.areaEfectiva,l.estadoFSC ,l.FSC, l.sumidero, l.CIF FROM UnidadDeManejo AS u INNER JOIN ((Municipio AS m INNER JOIN (Propietarios AS p INNER JOIN BancoTierras AS b ON p.ID = b.Propietario) ON m.ID = b.Municipio) INNER JOIN Lotes AS l ON b.ID = l.Predio) ON u.ID = l.Unidad ";
            int i = 0;
            if (!txtCodigo.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "l.Codigo LIKE '%" + txtCodigo.Text + "%'";
            }
            if (!txtLote.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "l.Lote LIKE '%" + txtLote.Text + "%'";
            }
            if (!txtPredio.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "b.Predio LIKE '" + txtPredio.Text + "'";
            }
            if (!txtLong.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "l.Longitud LIKE '%" + txtLong.Text + "%'";
            }
            if (!txtLat.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "l.Latitud LIKE '%" + txtLat.Text + "%'";
            }
            if (!txtFigura.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "l.Figura LIKE '%" + txtFigura.Text + "%'";
            }
            if (!txtUnidad.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "u.Unidad LIKE '%" + txtUnidad.Text + "%'";
            }
            if (!txtEspecie.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "l.Especie LIKE '%" + txtEspecie.Text + "%'";
            }
            if (!txtAno.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "l.Ano LIKE '%" + txtAno.Text + "%'";
            }
            if (!txtAreaPlant.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "l.areaPlantacion LIKE '%" + txtAreaPlant.Text + "%'";
            }
            if (!txtAreaEfe.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "l.areaEfectiva LIKE '%" + txtAreaEfe.Text + "%'";
            }
            if (!txtFSC.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "l.estadoFSC LIKE '%" + txtFSC.Text + "%'";
            }
            if (!txtNumFSC.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "l.FSC LIKE '%" + txtNumFSC.Text + "%'";
            }
            if (!txtNumSum.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "l.sumidero LIKE '%" + txtNumSum.Text + "%'";
            }
            if (!txtNumCif.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "l.CIF LIKE '%" + txtNumCif.Text + "%'";
            }
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int j = 0;
            try
            {
                while (myReader.Read())
                {
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[j].Cells[0].Value = myReader.GetInt32(0);
                    dataGridView1.Rows[j].Cells[1].Value = myReader.GetString(1);
                    dataGridView1.Rows[j].Cells[2].Value = myReader.GetString(2);
                    dataGridView1.Rows[j].Cells[3].Value = myReader.GetInt32(3);
                    dataGridView1.Rows[j].Cells[4].Value = myReader.GetInt32(4);
                    dataGridView1.Rows[j].Cells[5].Value = myReader.GetString(5);
                    dataGridView1.Rows[j].Cells[6].Value = myReader.GetString(6);
                    dataGridView1.Rows[j].Cells[7].Value = myReader.GetString(7);
                    dataGridView1.Rows[j].Cells[8].Value = myReader.GetString(8);
                    dataGridView1.Rows[j].Cells[9].Value = myReader.GetString(9);
                    dataGridView1.Rows[j].Cells[10].Value = myReader.GetInt32(10);
                    dataGridView1.Rows[j].Cells[11].Value = myReader.GetString(11);
                    dataGridView1.Rows[j].Cells[12].Value = myReader.GetString(12);
                    dataGridView1.Rows[j].Cells[13].Value = myReader.GetString(13);
                    dataGridView1.Rows[j].Cells[14].Value = myReader.GetString(14);
                    dataGridView1.Rows[j].Cells[15].Value = myReader.GetString(15);
                    dataGridView1.Rows[j].Cells[16].Value = myReader.GetString(16);
                    j++;
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

        public int getMaxID() {
            string query = "SELECT MAX(id) FROM Lotes";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int id = 0;
            try
            {
                while (myReader.Read())
                {
                    id = myReader.GetInt32(0);
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();               
            }
            return id;
        }

        public bool existeCodigo(string codigo)
        {
            string query = "SELECT ID FROM Lotes WHERE Codigo = " + codigo + " UNION ALL SELECT ID FROM Areas WHERE Codigo = " + codigo + " UNION ALL SELECT ID FROM LoteGanadero WHERE Codigo = " + codigo;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            bool existe = true;
            try
            {
                if (myReader.HasRows) 
                {
                    existe = true;
                }
                else
                {
                    existe = false;
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return existe;
        }


        public void agregarVias(int lote)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO ViasDeMovilizacion (distanciaAserradero,arroyo,criticos,portones,tiempo,distanciaBarranquilla,puentes,Lote) VALUES (@distanciaAserradero,@arroyo,@criticos,@portones,@tiempo,@distanciaBarranquilla,@puentes,@Lote)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@distanciaAserradero", OleDbType.VarChar).Value = 0;
                cmd.Parameters.Add("@arroyo", OleDbType.VarChar).Value = 0;
                cmd.Parameters.Add("@criticos", OleDbType.VarChar).Value = 0;
                cmd.Parameters.Add("@portones", OleDbType.VarChar).Value = 0;
                cmd.Parameters.Add("@tiempo", OleDbType.VarChar).Value = 0;
                cmd.Parameters.Add("@distanciaBarranquilla", OleDbType.VarChar).Value = 0;
                cmd.Parameters.Add("@puentes", OleDbType.VarChar).Value = 0;
                cmd.Parameters.Add("@Lote", OleDbType.VarChar).Value = lote;
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

        public void buscarLoteGanadero()
        {
            while (dataGridView3.Rows.Count != 0)
            {
                dataGridView3.Rows.RemoveAt(0);
            }
            string query = "SELECT l.Codigo, l.Lote, b.Predio, l.Longitud, l.Latitud, p.Nombre, u.Unidad, l.Pasto, m.Municipio, l.Area, l.Renovacion, l.Riesgo, l.Henificable FROM Propietarios AS p INNER JOIN (Municipio AS m INNER JOIN BancoTierras AS b ON m.ID = b.Municipio) ON p.ID = b.Propietario, UnidadDeManejo AS u INNER JOIN LoteGanadero AS l ON u.ID = l.Unidad ";
            int i = 0;
            if (!txtCodGan.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "l.Codigo LIKE '%" + txtCodGan.Text + "%'";
            }
            if (!txtLotGan.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "l.Lote LIKE '%" + txtLotGan.Text + "%'";
            }
            if (!txtPredGan.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "b.Predio LIKE '" + txtPredGan.Text + "'";
            }
            if (!txtLongGan.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "l.Longitud LIKE '%" + txtLongGan.Text + "%'";
            }
            if (!txtLatGan.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "l.Latitud LIKE '%" + txtLatGan.Text + "%'";
            }
            if (!txtUnidadGan.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "u.Unidad LIKE '%" + txtUnidadGan.Text + "%'";
            }
            if (!txtAreaGan.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "l.Area LIKE '%" + txtAreaGan.Text + "%'";
            }
            if (!txtRenoGan.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "l.Renovacion LIKE '%" + txtRenoGan.Text + "%'";
            }
            if (!txtRiesgoGan.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "l.Riesgo LIKE '%" + txtRiesgoGan.Text + "%'";
            }
            if (!txtHenGan.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "l.Henificable LIKE '%" + txtHenGan.Text + "%'";
            }
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int j = 0;
            try
            {
                while (myReader.Read())
                {
                    dataGridView3.Rows.Add();
                    dataGridView3.Rows[j].Cells[0].Value = myReader.GetInt32(0);
                    dataGridView3.Rows[j].Cells[1].Value = myReader.GetString(1);
                    dataGridView3.Rows[j].Cells[2].Value = myReader.GetString(2);
                    dataGridView3.Rows[j].Cells[3].Value = myReader.GetInt32(3);
                    dataGridView3.Rows[j].Cells[4].Value = myReader.GetInt32(4);
                    dataGridView3.Rows[j].Cells[5].Value = myReader.GetString(5);
                    dataGridView3.Rows[j].Cells[6].Value = myReader.GetString(6);
                    dataGridView3.Rows[j].Cells[7].Value = myReader.GetString(7);
                    dataGridView3.Rows[j].Cells[8].Value = myReader.GetString(8);
                    dataGridView3.Rows[j].Cells[9].Value = myReader.GetInt32(9);
                    dataGridView3.Rows[j].Cells[10].Value = myReader.GetInt32(10);
                    dataGridView3.Rows[j].Cells[11].Value = myReader.GetString(11);
                    dataGridView3.Rows[j].Cells[12].Value = myReader.GetString(12);                    
                    j++;
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

        public void buscarCampamento()
        {
            while (dataGridView5.Rows.Count != 0)
            {
                dataGridView5.Rows.RemoveAt(0);
            }
            string query = "SELECT c.ID, b.Predio, c.Longitud, c.Latitud, c.codCamp, c.Ocupacion, c.tipoMaterial, c.materialPred, c.Energia, c.agua, c.aguaAseo, c.otroServicio FROM Campamentos AS c INNER JOIN BancoTierras AS b ON c.Predio = b.ID ";
            int i = 0;
            if (!Camp1.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "b.Predio LIKE '%" + Camp1.Text + "%'";
            }
            if (!Camp3.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "c.materialPred LIKE '%" + Camp3.Text + "%'";
            }
            if (!Camp2.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "c.tipoMaterial LIKE '" + Camp2.Text + "'";
            }
            if (!Camp4.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "c.Longitud LIKE '%" + Camp4.Text + "%'";
            }
            if (!Camp5.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "c.Latitud LIKE '%" + Camp5.Text + "%'";
            }
            if (!Camp6.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "c.codCamp LIKE '%" + Camp6.Text + "%'";
            }
            if (!Camp7.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "c.Ocupacion LIKE '%" + Camp7.Text + "%'";
            }
            if (!Camp8.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "c.Energia LIKE '%" + Camp8.Text + "%'";
            }
            if (!Camp9.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "c.agua LIKE '%" + txtRiesgoGan.Text + "%'";
            }
            if (!Camp10.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "c.aguaAseo LIKE '%" + Camp10.Text + "%'";
            }
            if (!Camp11.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "c.otroServicio LIKE '%" + Camp11.Text + "%'";
            }
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int j = 0;
            try
            {
                while (myReader.Read())
                {
                    dataGridView5.Rows.Add();
                    dataGridView5.Rows[j].Cells[0].Value = myReader.GetInt32(0);
                    dataGridView5.Rows[j].Cells[1].Value = myReader.GetString(1);
                    dataGridView5.Rows[j].Cells[2].Value = myReader.GetString(2);
                    dataGridView5.Rows[j].Cells[3].Value = myReader.GetString(3);
                    dataGridView5.Rows[j].Cells[4].Value = myReader.GetString(4);
                    dataGridView5.Rows[j].Cells[5].Value = myReader.GetString(5);
                    dataGridView5.Rows[j].Cells[6].Value = myReader.GetString(6);
                    dataGridView5.Rows[j].Cells[7].Value = myReader.GetString(7);
                    dataGridView5.Rows[j].Cells[8].Value = myReader.GetString(8);
                    dataGridView5.Rows[j].Cells[9].Value = myReader.GetString(9);
                    dataGridView5.Rows[j].Cells[10].Value = myReader.GetString(10);
                    dataGridView5.Rows[j].Cells[11].Value = myReader.GetString(11);                    
                    j++;
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

        public void buscarLoteProt()
        {
            while (dataGridView2.Rows.Count != 0)
            {
                dataGridView2.Rows.RemoveAt(0);
            }
            string query = "SELECT a.Codigo, a.Lote, b.Predio, a.Longitud, a.Latitud, p.Nombre, u.Unidad, m.Municipio, a.Area FROM Municipio AS m INNER JOIN (Propietarios AS p INNER JOIN (UnidadDeManejo AS u INNER JOIN (BancoTierras AS b INNER JOIN Areas AS a ON b.ID = a.Predio) ON u.ID = a.Unidad) ON p.ID = b.Propietario) ON m.ID = b.Municipio WHERE a.Clase = 1";
            if (!txtCodProt.Text.Equals(""))
            {            
                query += " AND ";
                query += "a.Codigo LIKE '%" + txtCodProt.Text + "%'";
            }
            if (!txtLoteProt.Text.Equals(""))
            {
                query += " AND ";
                query += "a.Lote LIKE '%" + txtLoteProt.Text + "%'";
            }
            if (!txtPredProt.Text.Equals(""))
            {
                query += " AND ";
                query += "b.Predio LIKE '" + txtPredProt.Text + "'";
            }
            if (!txtLongProt.Text.Equals(""))
            {
                query += " AND ";
                query += "a.Longitud LIKE '%" + txtLongProt.Text + "%'";
            }
            if (!txtLatProt.Text.Equals(""))
            {
                query += " AND ";
                query += "a.Latitud LIKE '%" + txtLatProt.Text + "%'";
            }
            if (!txtUnidadProt.Text.Equals(""))
            {
                query += " AND ";
                query += "u.Unidad LIKE '%" + txtUnidadProt.Text + "%'";
            }
            if (!txtAreaProt.Text.Equals(""))
            {
                 query += " AND ";
                query += "a.Area LIKE '%" + txtAreaProt.Text + "%'";
            }
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
                    dataGridView2.Rows.Add();
                    dataGridView2.Rows[i].Cells[0].Value = myReader.GetInt32(0);
                    dataGridView2.Rows[i].Cells[1].Value = myReader.GetString(1);
                    dataGridView2.Rows[i].Cells[2].Value = myReader.GetString(2);
                    dataGridView2.Rows[i].Cells[3].Value = myReader.GetInt32(3);
                    dataGridView2.Rows[i].Cells[4].Value = myReader.GetInt32(4);
                    dataGridView2.Rows[i].Cells[5].Value = myReader.GetString(5);
                    dataGridView2.Rows[i].Cells[6].Value = myReader.GetString(6);
                    dataGridView2.Rows[i].Cells[7].Value = myReader.GetString(7);
                    dataGridView2.Rows[i].Cells[8].Value = myReader.GetInt32(8);
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

        public void buscarLoteAgua()
        {
            while (dataGridView4.Rows.Count != 0)
            {
                dataGridView4.Rows.RemoveAt(0);
            }
            string query = "SELECT a.Codigo, a.Lote, b.Predio, a.Longitud, a.Latitud, p.Nombre, u.Unidad, m.Municipio, a.Area FROM Municipio AS m INNER JOIN (Propietarios AS p INNER JOIN (UnidadDeManejo AS u INNER JOIN (BancoTierras AS b INNER JOIN Areas AS a ON b.ID = a.Predio) ON u.ID = a.Unidad) ON p.ID = b.Propietario) ON m.ID = b.Municipio WHERE a.Clase = 2";
            if (!txtCodAgua.Text.Equals(""))
            {
                query += " AND ";
                query += "a.Codigo LIKE '%" + txtCodAgua.Text + "%'";
            }
            if (!txtLoteAgua.Text.Equals(""))
            {
                query += " AND ";
                query += "a.Lote LIKE '%" + txtLoteAgua.Text + "%'";
            }
            if (!txtPredAgua.Text.Equals(""))
            {
                query += " AND ";
                query += "b.Predio LIKE '" + txtPredAgua.Text + "'";
            }
            if (!txtLongAgua.Text.Equals(""))
            {
                query += " AND ";
                query += "a.Longitud LIKE '%" + txtLongAgua.Text + "%'";
            }
            if (!txtLatAgua.Text.Equals(""))
            {
                query += " AND ";
                query += "a.Latitud LIKE '%" + txtLatAgua.Text + "%'";
            }
            if (!txtUnidadAgua.Text.Equals(""))
            {
                query += " AND ";
                query += "u.Unidad LIKE '%" + txtUnidadAgua.Text + "%'";
            }
            if (!txtAreaAgua.Text.Equals(""))
            {
                query += " AND ";
                query += "a.Area LIKE '%" + txtAreaAgua.Text + "%'";
            }
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
                    dataGridView4.Rows.Add();
                    dataGridView4.Rows[i].Cells[0].Value = myReader.GetInt32(0);
                    dataGridView4.Rows[i].Cells[1].Value = myReader.GetString(1);
                    dataGridView4.Rows[i].Cells[2].Value = myReader.GetString(2);
                    dataGridView4.Rows[i].Cells[3].Value = myReader.GetInt32(3);
                    dataGridView4.Rows[i].Cells[4].Value = myReader.GetInt32(4);
                    dataGridView4.Rows[i].Cells[5].Value = myReader.GetString(5);
                    dataGridView4.Rows[i].Cells[6].Value = myReader.GetString(6);
                    dataGridView4.Rows[i].Cells[7].Value = myReader.GetString(7);
                    dataGridView4.Rows[i].Cells[8].Value = myReader.GetInt32(8);
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

        public frmLotes()
        {
            InitializeComponent();
            cargarLotes();
            cargarPredios();
            cargarUnidades();
            cargarAreas(1);
            cargarAreas(2);
            cargarLotesGanadero();
            cargarCampamento();
            txtCif.Text = "No";
            txtFSC.Text = "No";
            txtCarbono.Text = "No";
            txtPredio.SelectedItem = null;
            txtUnidad.SelectedItem = null;
            txtPredProt.SelectedItem = null;
            txtUnidadProt.SelectedItem = null;
            txtPredAgua.SelectedItem = null;
            txtUnidadAgua.SelectedItem = null;
            txtPredGan.SelectedItem = null;
            txtUnidadGan.SelectedItem = null;
            Camp1.SelectedItem = null;
            dataGridView1.Columns[1].DefaultCellStyle.Font = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Underline);
            dataGridView2.Columns[1].DefaultCellStyle.Font = new Font(dataGridView2.DefaultCellStyle.Font, FontStyle.Underline);
            dataGridView3.Columns[1].DefaultCellStyle.Font = new Font(dataGridView3.DefaultCellStyle.Font, FontStyle.Underline);
            dataGridView4.Columns[1].DefaultCellStyle.Font = new Font(dataGridView4.DefaultCellStyle.Font, FontStyle.Underline);            
        }

        private void btnUnidad_Click(object sender, EventArgs e)
        {
            frmUnidadDeManejo newFrm = new frmUnidadDeManejo();
            newFrm.Show();
        }

        private void btnBanco_Click(object sender, EventArgs e)
        {
            if (txtCodigo.Text.Equals(""))
            {
                MessageBox.Show("Favor ingresar el codigo del lote forestal.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }
            else if(txtLote.Text.Equals(""))
            {
                MessageBox.Show("Favor ingresar el nombre del lote forestal.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }
            else if(txtPredio.Text.Equals(""))
            {
                MessageBox.Show("Favor seleccionar el predio.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }
            else if (txtUnidad.Text.Equals(""))
            {
                MessageBox.Show("Favor seleccionar la unidad de manejo.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else if (existeCodigo(txtCodigo.Text))
	        {
                MessageBox.Show("Codigo ya existe, favor ingresar otro.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
	        }
            else if(txtAreaEfe.Text.Contains("."))
                MessageBox.Show("Usar coma en vez de punto en la area efectiva.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else if (txtAreaPlant.Text.Contains("."))
                MessageBox.Show("Usar coma en vez de punto en la area de plantación.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else
            {
                agregarLote();
                int lote = getMaxID();
                agregarVias(lote);
                cargarLotes();
                reiniciarTablero();
            }
        }

        private void btnModificar_Click(object sender, EventArgs e)
        {
            modificarLote();
            cargarLotes();
            reiniciarTablero();
        }

        private void btnBuscar_Click(object sender, EventArgs e)
        {
            buscarLote();
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el lote " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {

                string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM Lotes WHERE Codigo = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Lote eliminado.");
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
                cargarLotes();
                reiniciarTablero();
            }
        }

        private void btnReiniciar_Click(object sender, EventArgs e)
        {
            reiniciarTablero();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentCell.ColumnIndex == 0)
            {
                frmOrdenes newFrm = new frmOrdenes("Lote",dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
                newFrm.Show();
            }
            if (dataGridView1.CurrentCell.ColumnIndex == 1)
            {                
                frmEstudio newFrm = new frmEstudio(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
                newFrm.Show();
            }
            txtCodigo.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
            txtLote.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();
            txtPredio.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString();
            txtLong.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString();
            txtLat.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[4].Value.ToString();
            txtFigura.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Value.ToString();
            txtUnidad.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[7].Value.ToString();
            txtEspecie.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[9].Value.ToString();
            txtAno.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[10].Value.ToString();
            txtAreaPlant.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[11].Value.ToString();
            txtAreaEfe.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[12].Value.ToString();
            txtFSC.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[13].Value.ToString();
            txtNumFSC.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[14].Value.ToString();
            txtNumSum.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[15].Value.ToString();
            txtNumCif.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[16].Value.ToString();
            txtAmbiental.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[17].Value.ToString();
            txtPlantacion.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[18].Value.ToString();
            txtUbicacion.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[19].Value.ToString();
            textBox7.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[20].Value.ToString();
            textBox4.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[21].Value.ToString();
            textBox5.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[22].Value.ToString();
            textBox6.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[23].Value.ToString();
            if (txtNumCif.Equals(""))
                txtCif.Text = "No";
            else
                txtCif.Text = "Si";
            if (txtNumSum.Equals(""))
                txtCarbono.Text = "No";
            else
                txtCarbono.Text = "Si";
        }

        private void txtCif_TextChanged(object sender, EventArgs e)
        {
            if (txtCif.Text.Equals("Si"))
            {
                label17.Visible = true;
                txtNumCif.Visible = true;
            }
            else 
            {
                label17.Visible = false;
                txtNumCif.Visible = false;
                txtNumCif.Text = "";
            }
        }

        private void txtFSC_TextChanged(object sender, EventArgs e)
        {
            if (txtFSC.Text.Equals("No"))
            {
                label11.Visible = false;
                txtNumFSC.Visible = false;
                txtNumFSC.Text = "";
            }
            else
            {
                label11.Visible = true;
                txtNumFSC.Visible = true;
            }
        }

        private void txtCarbono_TextChanged(object sender, EventArgs e)
        {
            if (txtCarbono.Text.Equals("Si"))
            {
                label12.Visible = true;
                txtNumSum.Visible = true;
            }
            else
            {
                label12.Visible = false;
                txtNumSum.Visible = false;
                txtNumSum.Text = "";
            }
        }

        public void reiniciarArea(int clase) {
            if (clase == 2)
            {
                txtCodAgua.Text = "";
                txtLoteAgua.Text = "";
                txtPredAgua.Text = "";
                txtLongAgua.Text = "";
                txtLatAgua.Text = "";
                txtUnidadAgua.Text = "";
                txtAreaAgua.Text = "";
                textBox2.Text = "";
            }
            else 
            {
                txtCodProt.Text = "";
                txtLoteProt.Text = "";
                txtPredProt.Text = "";
                txtLongProt.Text = "";
                txtLatProt.Text = "";
                txtUnidadProt.Text = "";
                txtAreaProt.Text = "";
                textBox3.Text = "";
            }
        }

        public void reiniciarGanado() { 
            txtCodGan.Text = "";
            txtLotGan.Text = "";
            textBox1.Text = "";
            txtPredGan.Text = "";
            txtLongGan.Text = "";
            txtLatGan.Text = "";
            txtUnidadGan.Text = "";
            txtAreaGan.Text = "";
            txtRiesgoGan.Text = "";
            txtHenGan.Text = "";
            txtRenoGan.Text = "";
            txtPastoGan.Text = "";
            txtUnidadGan.Text = "";
            textBox1.Text = "";
        }

        public void reiniciarCampamento()
        {
            Camp1.Text = "";
            Camp2.Text = "";
            Camp3.Text = "";
            Camp4.Text = "";
            Camp5.Text = "";
            Camp6.Text = "";
            Camp7.Text = "";
            Camp8.Text = "";
            Camp9.Text = "";
            Camp10.Text = "";
            Camp11.Text = "";
        }

        private void txtAgregarGan_Click(object sender, EventArgs e)
        {
            if (txtCodGan.Text.Equals(""))
            {
                MessageBox.Show("Favor ingresar el codigo del lote ganadero.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }
            else if (txtAreaGan.Text.Contains("."))
                MessageBox.Show("Usar coma en vez de punto en la area del lote ganadero.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else if(txtLotGan.Text.Equals(""))
            {
                MessageBox.Show("Favor ingresar el nombre del lote ganadero.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }
            else if(txtPredGan.Text.Equals(""))
            {
                MessageBox.Show("Favor seleccionar el predio.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }
            else if (txtUnidadGan.Text.Equals(""))
            {
                MessageBox.Show("Favor seleccionar la unidad de manejo.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else if (existeCodigo(txtCodGan.Text))
            {
                MessageBox.Show("Codigo ya existe, favor ingresar otro.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                agregarLoteGanadero();
                cargarLotesGanadero();
                reiniciarGanado();
            }
        }

        private void txtModificarGan_Click(object sender, EventArgs e)
        {
            modificarLoteGanadero();
            cargarLotesGanadero();
            reiniciarGanado();
        }

        private void txtBuscarGan_Click(object sender, EventArgs e)
        {
            buscarLoteGanadero();
        }

        private void txtEliminarGan_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el lote " + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {

                string id = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM LoteGanadero WHERE Codigo = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Lote eliminado.");
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
                cargarLotesGanadero();
                reiniciarGanado();
            }
        }

        private void txtReniciarGan_Click(object sender, EventArgs e)
        {
            reiniciarGanado();
        }

        private void btnAgregarProt_Click(object sender, EventArgs e)
        {
            if (txtCodProt.Text.Equals(""))
            {
                MessageBox.Show("Favor ingresar el codigo del area de conservación.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }
            else if (txtAreaProt.Text.Contains("."))
                MessageBox.Show("Usar coma en vez de punto en la area de la area de conservación.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else if(txtLoteProt.Text.Equals(""))
            {
                MessageBox.Show("Favor ingresar el nombre del area de conservación.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }
            else if(txtPredProt.Text.Equals(""))
            {
                MessageBox.Show("Favor seleccionar el predio.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }
            else if (txtUnidadProt.Text.Equals(""))
            {
                MessageBox.Show("Favor seleccionar la unidad de manejo.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else if (existeCodigo(txtCodProt.Text))
            {
                MessageBox.Show("Codigo ya existe, favor ingresar otro.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                agregarArea(1);
                cargarAreas(1);
                reiniciarArea(1);
            }
        }

        private void txtModificarProt_Click(object sender, EventArgs e)
        {
            modificarArea(1);
            cargarAreas(1);
            reiniciarArea(1);
        }

        private void txtBuscarProt_Click(object sender, EventArgs e)
        {
            buscarLoteProt();
        }

        private void txtEliminarProt_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el lote " + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {

                string id = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM Areas WHERE Codigo = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Lote eliminado.");
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
                cargarAreas(1);
                reiniciarArea(1);
            }
        }

        private void txtReiniciarProt_Click(object sender, EventArgs e)
        {
            reiniciarArea(1);
        }

        private void btnAgregarAgua_Click(object sender, EventArgs e)
        {
            if (txtCodAgua.Text.Equals(""))
            {
                MessageBox.Show("Favor ingresar el codigo del cuerpo de agua.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }
            else if (txtAreaAgua.Text.Contains("."))
                MessageBox.Show("Usar coma en vez de punto en la area del cuerpo de agua.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else if(txtLoteAgua.Text.Equals(""))
            {
                MessageBox.Show("Favor ingresar el nombre del cuerpo de agua.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }
            else if(txtPredAgua.Text.Equals(""))
            {
                MessageBox.Show("Favor seleccionar el predio.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }
            else if(txtUnidadAgua.Text.Equals(""))
            {
                MessageBox.Show("Favor seleccionar la unidad de manejo.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else if (existeCodigo(txtCodAgua.Text))
            {
                MessageBox.Show("Codigo ya existe, favor ingresar otro.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                agregarArea(2);
                cargarAreas(2);
                reiniciarArea(2);
            }
        }

        private void btnModificarAgua_Click(object sender, EventArgs e)
        {
            modificarArea(2);
            cargarAreas(2);
            reiniciarArea(2);
        }

        private void btnBuscarAgua_Click(object sender, EventArgs e)
        {
            buscarLoteAgua();
        }

        private void btnEliminarAgua_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el lote " + dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {

                string id = dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM Areas WHERE Codigo = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Lote eliminado.");
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
                cargarAreas(2);
                reiniciarArea(2);
            }
        }

        private void btnReiniciarAgua_Click(object sender, EventArgs e)
        {
            reiniciarArea(2);
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView2.CurrentCell.ColumnIndex == 0)
            {
                frmOrdenes newFrm = new frmOrdenes("Lote", dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString());
                newFrm.Show();
            }
            if (dataGridView2.CurrentCell.ColumnIndex == 1)
            {
                frmInfoLote newFrm = new frmInfoLote(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString(), "Areas");
                newFrm.Show();
            }
            txtCodProt.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString();
            txtLoteProt.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[1].Value.ToString();
            txtPredProt.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[2].Value.ToString();
            txtLongProt.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[3].Value.ToString();
            txtLatProt.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[4].Value.ToString();            
            txtUnidadProt.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[6].Value.ToString();
            txtAreaProt.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[8].Value.ToString();
            textBox2.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[9].Value.ToString();
        }

        private void dataGridView4_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView4.CurrentCell.ColumnIndex == 0)
            {
                frmOrdenes newFrm = new frmOrdenes("Lote", dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[0].Value.ToString());
                newFrm.Show();
            }
            if (dataGridView4.CurrentCell.ColumnIndex == 1)
            {
                frmInfoLote newFrm = new frmInfoLote(dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[0].Value.ToString(), "Areas");
                newFrm.Show();
            }
            txtCodAgua.Text = dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[0].Value.ToString();
            txtLoteAgua.Text = dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[1].Value.ToString();
            txtPredAgua.Text = dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[2].Value.ToString();
            txtLongAgua.Text = dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[3].Value.ToString();
            txtLatAgua.Text = dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[4].Value.ToString();
            txtUnidadAgua.Text = dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[6].Value.ToString();
            txtAreaAgua.Text = dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[8].Value.ToString();
            textBox3.Text = dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[9].Value.ToString();
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                cargarLotes();
                reiniciarTablero();
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                cargarAreas(1);
                reiniciarArea(1);
            }
            else if (tabControl1.SelectedIndex == 3)
            {
                cargarAreas(2);
                reiniciarArea(2);
            }
            else if(tabControl1.SelectedIndex == 4)
            {
                cargarLotesGanadero();
                reiniciarGanado();
            }
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView3.CurrentCell.ColumnIndex == 0)
            {
                frmOrdenes newFrm = new frmOrdenes("Lote", dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value.ToString());
                newFrm.Show();
            }
            if (dataGridView3.CurrentCell.ColumnIndex == 1)
            {
                frmInfoLote newFrm = new frmInfoLote(dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value.ToString(),"LoteGanadero");
                newFrm.Show();
            }
            txtCodGan.Text = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value.ToString();
            txtLotGan.Text = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[1].Value.ToString();
            txtPredGan.Text = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[2].Value.ToString();
            txtLongGan.Text = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[3].Value.ToString();
            txtLatGan.Text = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[4].Value.ToString();
            txtUnidadGan.Text = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[6].Value.ToString();
            txtPastoGan.Text = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[7].Value.ToString();
            txtAreaGan.Text = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[9].Value.ToString();
            txtRiesgoGan.Text = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[11].Value.ToString();
            txtHenGan.Text = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[12].Value.ToString();
            textBox1.Text = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[13].Value.ToString();
            txtRenoGan.Text = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[10].Value.ToString();
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            reiniciarCampamento();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el lote " + dataGridView5.Rows[dataGridView5.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {

                string id = dataGridView5.Rows[dataGridView5.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM Campamentos WHERE id = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Campamento eliminado.");
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
                cargarCampamento();
                reiniciarCampamento();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            buscarCampamento();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (Camp1.Text.Equals("")) 
            {
                MessageBox.Show("Favor seleccionar el predio.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                if (Camp6.Text.Equals(""))
                {
                    MessageBox.Show("Favor ingresar el codigo del Campamento.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                {
                    if (existeCodigo(Camp6.Text))
	                {
                        MessageBox.Show("Codigo ya existe, favor ingresar otro.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
	                }
                    else
                    {
                        agregarCampamento();
                        cargarCampamento();
                        reiniciarCampamento();
                    }                    
                }
            }            
        }

        private void dataGridView5_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            Camp1.Text = dataGridView5.Rows[dataGridView5.CurrentCell.RowIndex].Cells[1].Value.ToString();
            Camp2.Text = dataGridView5.Rows[dataGridView5.CurrentCell.RowIndex].Cells[6].Value.ToString();
            Camp3.Text = dataGridView5.Rows[dataGridView5.CurrentCell.RowIndex].Cells[7].Value.ToString();
            Camp4.Text = dataGridView5.Rows[dataGridView5.CurrentCell.RowIndex].Cells[2].Value.ToString();
            Camp5.Text = dataGridView5.Rows[dataGridView5.CurrentCell.RowIndex].Cells[3].Value.ToString();
            Camp6.Text = dataGridView5.Rows[dataGridView5.CurrentCell.RowIndex].Cells[4].Value.ToString();
            Camp7.Text = dataGridView5.Rows[dataGridView5.CurrentCell.RowIndex].Cells[5].Value.ToString();
            Camp8.Text = dataGridView5.Rows[dataGridView5.CurrentCell.RowIndex].Cells[8].Value.ToString();
            Camp9.Text = dataGridView5.Rows[dataGridView5.CurrentCell.RowIndex].Cells[9].Value.ToString();
            Camp10.Text = dataGridView5.Rows[dataGridView5.CurrentCell.RowIndex].Cells[10].Value.ToString();
            Camp11.Text = dataGridView5.Rows[dataGridView5.CurrentCell.RowIndex].Cells[11].Value.ToString();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            modificarCampamento();
            cargarCampamento();
            reiniciarCampamento();
        }

        public void imprimirLotes(DataGridView data)
        {
            if (data.Rows.Count > 0)
            {                
                Microsoft.Office.Interop.Excel.Application XcelApp = new Microsoft.Office.Interop.Excel.Application();
                XcelApp.Application.Workbooks.Add(Type.Missing);
                Microsoft.Office.Interop.Excel.Range excelCellrange;
                for (int i = 1; i < data.Columns.Count + 1; i++)
                {
                    XcelApp.Cells[2, i+1] = data.Columns[i - 1].HeaderText;
                }
                for (int i = 0; i < data.Rows.Count; i++)
                {
                    for (int j = 0; j < data.Columns.Count; j++)
                    {
                        XcelApp.Cells[i + 3, j + 2] = data.Rows[i].Cells[j].Value.ToString();
                        if (i == 0)
                        {
                            excelCellrange = XcelApp.Range[XcelApp.Cells[i + 2, 2], XcelApp.Cells[i + 2, data.Columns.Count+1]];
                            excelCellrange.Interior.Color = System.Drawing.Color.LightGreen;
                            excelCellrange.AutoFilter(1);
                            //excelCellrange.Interior.Color = System.Drawing.Color.Blue;
                            //excelCellrange.Font.Color = System.Drawing.Color.White;
                        }
                    }
                }                
                excelCellrange = XcelApp.Range[XcelApp.Cells[2, 2], XcelApp.Cells[data.Rows.Count + 2, data.Columns.Count+1]];
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
                imprimirLotes(dataGridView1);
            else if (tabControl1.SelectedIndex == 1)
                imprimirLotes(dataGridView3);
            else if (tabControl1.SelectedIndex == 2)
                imprimirLotes(dataGridView2);
            else if (tabControl1.SelectedIndex == 3)
                imprimirLotes(dataGridView4);
            else
                imprimirLotes(dataGridView5);
	{
		 
	}
        }

        private void label61_Click(object sender, EventArgs e)
        {

        }

        private void label60_Click(object sender, EventArgs e)
        {

        }

        private void label63_Click(object sender, EventArgs e)
        {

        }
    }
}
