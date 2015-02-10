namespace Aplicativo
{
    partial class frmRegistrarPaquete
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox9 = new System.Windows.Forms.GroupBox();
            this.textBox13 = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.textBox22 = new System.Windows.Forms.TextBox();
            this.label23 = new System.Windows.Forms.Label();
            this.textBox23 = new System.Windows.Forms.TextBox();
            this.label24 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.button2 = new System.Windows.Forms.Button();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox9.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(102, 27);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(162, 21);
            this.comboBox1.TabIndex = 0;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(43, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Producto:";
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Checked = true;
            this.radioButton1.Location = new System.Drawing.Point(46, 171);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(112, 17);
            this.radioButton1.TabIndex = 2;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "Paquete Completo";
            this.radioButton1.UseVisualStyleBackColor = true;
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Location = new System.Drawing.Point(164, 171);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(100, 17);
            this.radioButton2.TabIndex = 3;
            this.radioButton2.Text = "Paquete Parcial";
            this.radioButton2.UseVisualStyleBackColor = true;
            this.radioButton2.CheckedChanged += new System.EventHandler(this.radioButton2_CheckedChanged);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(147, 206);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(73, 20);
            this.textBox1.TabIndex = 4;
            this.textBox1.Text = "0";
            this.textBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(76, 209);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(66, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "# de Piezas:";
            // 
            // groupBox9
            // 
            this.groupBox9.Controls.Add(this.textBox13);
            this.groupBox9.Controls.Add(this.label15);
            this.groupBox9.Controls.Add(this.textBox22);
            this.groupBox9.Controls.Add(this.label23);
            this.groupBox9.Controls.Add(this.textBox23);
            this.groupBox9.Controls.Add(this.label24);
            this.groupBox9.Location = new System.Drawing.Point(46, 68);
            this.groupBox9.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox9.Name = "groupBox9";
            this.groupBox9.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox9.Size = new System.Drawing.Size(218, 89);
            this.groupBox9.TabIndex = 39;
            this.groupBox9.TabStop = false;
            this.groupBox9.Text = "Cantidad en Empaque del Producto";
            // 
            // textBox13
            // 
            this.textBox13.Location = new System.Drawing.Point(93, 65);
            this.textBox13.Margin = new System.Windows.Forms.Padding(2);
            this.textBox13.Name = "textBox13";
            this.textBox13.ReadOnly = true;
            this.textBox13.Size = new System.Drawing.Size(81, 20);
            this.textBox13.TabIndex = 37;
            this.textBox13.Text = "0";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(52, 67);
            this.label15.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(31, 13);
            this.label15.TabIndex = 36;
            this.label15.Text = "Total";
            // 
            // textBox22
            // 
            this.textBox22.Location = new System.Drawing.Point(93, 42);
            this.textBox22.Margin = new System.Windows.Forms.Padding(2);
            this.textBox22.Name = "textBox22";
            this.textBox22.ReadOnly = true;
            this.textBox22.Size = new System.Drawing.Size(81, 20);
            this.textBox22.TabIndex = 35;
            this.textBox22.Text = "0";
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.Location = new System.Drawing.Point(62, 44);
            this.label23.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(25, 13);
            this.label23.TabIndex = 34;
            this.label23.Text = "Alto";
            // 
            // textBox23
            // 
            this.textBox23.Location = new System.Drawing.Point(93, 19);
            this.textBox23.Margin = new System.Windows.Forms.Padding(2);
            this.textBox23.Name = "textBox23";
            this.textBox23.ReadOnly = true;
            this.textBox23.Size = new System.Drawing.Size(81, 20);
            this.textBox23.TabIndex = 33;
            this.textBox23.Text = "0";
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Location = new System.Drawing.Point(50, 22);
            this.label24.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(38, 13);
            this.label24.TabIndex = 32;
            this.label24.Text = "Ancho";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.Window;
            this.dataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2,
            this.Column3,
            this.Column4});
            this.dataGridView1.Location = new System.Drawing.Point(298, 27);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.Size = new System.Drawing.Size(672, 195);
            this.dataGridView1.TabIndex = 40;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(309, 235);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(50, 13);
            this.label3.TabIndex = 42;
            this.label3.Text = "Ancho 1:";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(365, 232);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(70, 20);
            this.textBox2.TabIndex = 41;
            this.textBox2.Text = "0";
            this.textBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(457, 235);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(50, 13);
            this.label4.TabIndex = 44;
            this.label4.Text = "Ancho 2:";
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(513, 232);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(70, 20);
            this.textBox3.TabIndex = 43;
            this.textBox3.Text = "0";
            this.textBox3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(747, 235);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(37, 13);
            this.label5.TabIndex = 48;
            this.label5.Text = "Alto 2:";
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(648, 232);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(70, 20);
            this.textBox4.TabIndex = 47;
            this.textBox4.Text = "0";
            this.textBox4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(605, 235);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(37, 13);
            this.label6.TabIndex = 46;
            this.label6.Text = "Alto 1:";
            // 
            // textBox5
            // 
            this.textBox5.Location = new System.Drawing.Point(790, 232);
            this.textBox5.Name = "textBox5";
            this.textBox5.Size = new System.Drawing.Size(70, 20);
            this.textBox5.TabIndex = 45;
            this.textBox5.Text = "0";
            this.textBox5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(879, 230);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(91, 23);
            this.button1.TabIndex = 49;
            this.button1.Text = "Registrar";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(70, 235);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(72, 13);
            this.label7.TabIndex = 51;
            this.label7.Text = "# de Bodega:";
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Items.AddRange(new object[] {
            "10",
            "9",
            "8",
            "7",
            "6",
            "5",
            "4",
            "3",
            "2",
            "1"});
            this.comboBox2.Location = new System.Drawing.Point(147, 232);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(73, 21);
            this.comboBox2.TabIndex = 50;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(12, 282);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(958, 29);
            this.button2.TabIndex = 52;
            this.button2.Text = "Registrar Paquete";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Column1
            // 
            this.Column1.HeaderText = "Ancho 1";
            this.Column1.Name = "Column1";
            this.Column1.ReadOnly = true;
            // 
            // Column2
            // 
            this.Column2.HeaderText = "Ancho 2";
            this.Column2.Name = "Column2";
            this.Column2.ReadOnly = true;
            // 
            // Column3
            // 
            this.Column3.HeaderText = "Alto 1";
            this.Column3.Name = "Column3";
            this.Column3.ReadOnly = true;
            // 
            // Column4
            // 
            this.Column4.HeaderText = "Alto 2";
            this.Column4.Name = "Column4";
            this.Column4.ReadOnly = true;
            // 
            // frmRegistrarPaquete
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(982, 316);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.comboBox2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.textBox5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.groupBox9);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.radioButton2);
            this.Controls.Add(this.radioButton1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBox1);
            this.Name = "frmRegistrarPaquete";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Registrar Paquete OP #";
            this.groupBox9.ResumeLayout(false);
            this.groupBox9.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton radioButton1;
        private System.Windows.Forms.RadioButton radioButton2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox groupBox9;
        private System.Windows.Forms.TextBox textBox13;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.TextBox textBox22;
        private System.Windows.Forms.Label label23;
        private System.Windows.Forms.TextBox textBox23;
        private System.Windows.Forms.Label label24;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox textBox5;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column4;
    }
}