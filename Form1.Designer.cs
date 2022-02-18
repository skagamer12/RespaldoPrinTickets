
namespace Printickets
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.cant = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.code = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.desc1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.desc2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.linea = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.origin = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.imp = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.pie = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.upb = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.spcn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.print = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(29, 72);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(163, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Ingrese Referencia del Acreedor ";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(28, 88);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(164, 20);
            this.textBox1.TabIndex = 1;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(16, 131);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 41);
            this.button1.TabIndex = 2;
            this.button1.Text = "Entrada de Mercancias";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(101, 131);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 41);
            this.button2.TabIndex = 3;
            this.button2.Text = "Pedidos";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(182, 131);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 41);
            this.button3.TabIndex = 4;
            this.button3.Text = "Facturas Provedores";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.panel1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("panel1.BackgroundImage")));
            this.panel1.Location = new System.Drawing.Point(732, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(282, 103);
            this.panel1.TabIndex = 5;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
            this.dataGridView1.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.dataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.cant,
            this.code,
            this.desc1,
            this.desc2,
            this.linea,
            this.origin,
            this.imp,
            this.pie,
            this.upb,
            this.spcn});
            this.dataGridView1.Location = new System.Drawing.Point(12, 212);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.Size = new System.Drawing.Size(894, 264);
            this.dataGridView1.TabIndex = 6;
            this.dataGridView1.Visible = false;
            // 
            // cant
            // 
            this.cant.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.cant.FillWeight = 1000F;
            this.cant.HeaderText = "Cantidad";
            this.cant.Name = "cant";
            this.cant.ReadOnly = true;
            this.cant.Width = 74;
            // 
            // code
            // 
            this.code.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.code.HeaderText = "Código";
            this.code.Name = "code";
            this.code.ReadOnly = true;
            this.code.Width = 65;
            // 
            // desc1
            // 
            this.desc1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.desc1.HeaderText = "Descripción";
            this.desc1.Name = "desc1";
            this.desc1.Width = 88;
            // 
            // desc2
            // 
            this.desc2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.desc2.HeaderText = "Descripción 2";
            this.desc2.Name = "desc2";
            this.desc2.Width = 97;
            // 
            // linea
            // 
            this.linea.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.linea.HeaderText = "Linea";
            this.linea.Name = "linea";
            this.linea.ReadOnly = true;
            this.linea.Width = 58;
            // 
            // origin
            // 
            this.origin.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.origin.HeaderText = "Origen";
            this.origin.Name = "origin";
            this.origin.ReadOnly = true;
            this.origin.Width = 63;
            // 
            // imp
            // 
            this.imp.HeaderText = "Imprimir";
            this.imp.Name = "imp";
            this.imp.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.imp.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.imp.Width = 67;
            // 
            // pie
            // 
            this.pie.HeaderText = "Pin";
            this.pie.Name = "pie";
            this.pie.ReadOnly = true;
            this.pie.Visible = false;
            this.pie.Width = 47;
            // 
            // upb
            // 
            this.upb.HeaderText = "Unit Per Box";
            this.upb.Name = "upb";
            this.upb.ReadOnly = true;
            this.upb.Visible = false;
            this.upb.Width = 91;
            // 
            // spcn
            // 
            this.spcn.HeaderText = "SuppCatNum";
            this.spcn.Name = "spcn";
            this.spcn.ReadOnly = true;
            this.spcn.Visible = false;
            this.spcn.Width = 95;
            // 
            // print
            // 
            this.print.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.print.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.print.Image = ((System.Drawing.Image)(resources.GetObject("print.Image")));
            this.print.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.print.Location = new System.Drawing.Point(550, 71);
            this.print.Name = "print";
            this.print.Size = new System.Drawing.Size(102, 91);
            this.print.TabIndex = 7;
            this.print.Text = "Imprimir";
            this.print.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.print.UseVisualStyleBackColor = false;
            this.print.Visible = false;
            this.print.Click += new System.EventHandler(this.print_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(1048, 521);
            this.Controls.Add(this.print);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label1);
            this.DoubleBuffered = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "PrinTikets";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button print;
        private System.Windows.Forms.DataGridViewTextBoxColumn cant;
        private System.Windows.Forms.DataGridViewTextBoxColumn code;
        private System.Windows.Forms.DataGridViewTextBoxColumn desc1;
        private System.Windows.Forms.DataGridViewTextBoxColumn desc2;
        private System.Windows.Forms.DataGridViewTextBoxColumn linea;
        private System.Windows.Forms.DataGridViewTextBoxColumn origin;
        private System.Windows.Forms.DataGridViewCheckBoxColumn imp;
        private System.Windows.Forms.DataGridViewTextBoxColumn pie;
        private System.Windows.Forms.DataGridViewTextBoxColumn upb;
        private System.Windows.Forms.DataGridViewTextBoxColumn spcn;
    }
}

