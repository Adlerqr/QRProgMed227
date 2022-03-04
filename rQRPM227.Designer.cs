namespace QRProgMed227
{
    partial class rQRPM227 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public rQRPM227()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de componentes

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnExp = this.Factory.CreateRibbonButton();
            this.btnFile = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "QR-P227";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnExp);
            this.group1.Items.Add(this.btnFile);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // btnExp
            // 
            this.btnExp.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnExp.Image = global::QRProgMed227.Properties.Resources.importar;
            this.btnExp.Label = "Importar Archivo";
            this.btnExp.Name = "btnExp";
            this.btnExp.ShowImage = true;
            this.btnExp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExp_Click);
            // 
            // btnFile
            // 
            this.btnFile.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnFile.Image = global::QRProgMed227.Properties.Resources.construir;
            this.btnFile.Label = "Construir Archivo";
            this.btnFile.Name = "btnFile";
            this.btnFile.ShowImage = true;
            this.btnFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFile_Click);
            // 
            // rQRPM227
            // 
            this.Name = "rQRPM227";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.rQRPM227_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFile;
    }

    partial class ThisRibbonCollection
    {
        internal rQRPM227 rQRPM227
        {
            get { return this.GetRibbon<rQRPM227>(); }
        }
    }
}
