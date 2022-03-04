namespace QRProgMed227
{
    partial class frmFiles
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmFiles));
            this.dgvProcesar = new System.Windows.Forms.DataGridView();
            this.btnProcesar = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvProcesar)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvProcesar
            // 
            this.dgvProcesar.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvProcesar.Location = new System.Drawing.Point(12, 12);
            this.dgvProcesar.Name = "dgvProcesar";
            this.dgvProcesar.RowHeadersWidth = 51;
            this.dgvProcesar.RowTemplate.Height = 24;
            this.dgvProcesar.Size = new System.Drawing.Size(636, 274);
            this.dgvProcesar.TabIndex = 0;
            this.dgvProcesar.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvProcesar_CellClick);
            // 
            // btnProcesar
            // 
            this.btnProcesar.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnProcesar.Location = new System.Drawing.Point(507, 292);
            this.btnProcesar.Name = "btnProcesar";
            this.btnProcesar.Size = new System.Drawing.Size(141, 38);
            this.btnProcesar.TabIndex = 2;
            this.btnProcesar.Text = "PROCESAR";
            this.btnProcesar.UseVisualStyleBackColor = true;
            this.btnProcesar.Click += new System.EventHandler(this.btnProcesar_Click);
            // 
            // frmFiles
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(660, 343);
            this.Controls.Add(this.btnProcesar);
            this.Controls.Add(this.dgvProcesar);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmFiles";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Configure las Fechas y Nombres";
            this.Load += new System.EventHandler(this.frmFiles_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmFiles_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.dgvProcesar)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dgvProcesar;
        private System.Windows.Forms.Button btnProcesar;
    }
}