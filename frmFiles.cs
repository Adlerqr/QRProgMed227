using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SpreadsheetLight;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace QRProgMed227
{
    public partial class frmFiles : Form
    {
        DateTimePicker dtp;
        public frmFiles()
        {
            InitializeComponent();
        }

        private void frmFiles_Load(object sender, EventArgs e)
        {
            dgvProcesar.DataSource = rQRPM227.dtProg;
            dgvProcesar.AllowUserToResizeColumns = false;
            dgvProcesar.AllowUserToResizeRows = false;
            dgvProcesar.AllowUserToAddRows = false;
            dgvProcesar.Columns[0].Width = 360;            
            dgvProcesar.RowHeadersVisible = false;
            dgvProcesar.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable; //para que no se reordene al hacer clic
            dgvProcesar.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dgvProcesar.Font, FontStyle.Bold); //titulos en negrita

            DataGridViewComboBoxColumn dgvCbb = new DataGridViewComboBoxColumn();
            dgvCbb.Name = "Técnico";
            string fileTxt = File.ReadAllText(@"D:\QRProg\tecnicos.txt");
            int i = 0;
            foreach (string fila in fileTxt.Split('\n'))
            {
                if (!string.IsNullOrEmpty(fila)) //para terminar cuando este vacio
                {
                    dgvCbb.Items.Add(fila);
                }
                i++;
            }
            
            dgvProcesar.Columns.Add(dgvCbb); //agrega el combo a la columna
            dgvProcesar.Columns[1].Width = 170;
            dgvProcesar.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable; //para que no se reordene al hacer clic
            dgvProcesar.Columns.Add("Fecha","Fecha"); //agrega una columna mas para las 3 funciones posteriores
            dgvProcesar.Columns[2].Width = 100;
            dgvProcesar.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable; //para que no se reordene al hacer clic

            //configurar texto a la derecha, combos con su datos del item 0 y fecha actual colocada
            for (int j = 0; j < dgvProcesar.Rows.Count; j++)
            {
                dgvProcesar.Rows[j].Cells[1].Value = dgvCbb.Items[0];
                dgvProcesar.Rows[j].Cells[2].Value = DateTime.Today.ToString("dd/MM/yyyy");                
            }
            dgvProcesar.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvProcesar.Columns[0].ReadOnly = true;
        }

        private void dgvProcesar_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                if (e.ColumnIndex == 2)
                {
                    // Initialize the dateTimePicker1.
                    dtp = new DateTimePicker();
                    // Adding the dateTimePicker1 into DataGridView.   
                    dgvProcesar.Controls.Add(dtp);
                    // Setting the format i.e. mm/dd/yyyy)
                    dtp.Format = DateTimePickerFormat.Custom;
                    dtp.CustomFormat = "dd/MM/yyy";
                    // Create retangular area that represents the display area for a cell.
                    System.Drawing.Rectangle oRectangle = dgvProcesar.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);
                    // Setting area for dateTimePicker1.
                    dtp.Size = new Size(oRectangle.Width, oRectangle.Height);
                    // Setting location for dateTimePicker1.
                    dtp.Location = new System.Drawing.Point(oRectangle.X, oRectangle.Y);
                    // An event attached to dateTimePicker1 which is fired when any date is selected.
                    dtp.TextChanged += new EventHandler(DateTimePickerChange);
                    // An event attached to dateTimePicker1 which is fired when DateTimeControl is closed.
                    dtp.CloseUp += new EventHandler(DateTimePickerClose);
                }
            }
        }

        private void DateTimePickerChange(object sender, EventArgs e)
        {
            dgvProcesar.CurrentCell.Value = dtp.Text.ToString();
            //MessageBox.Show(string.Format("Date changed to {0}", dtp.Text.ToString()));
        }

        private void DateTimePickerClose(object sender, EventArgs e)
        {
            dtp.Visible = false;
        }

        private void btnProcesar_Click(object sender, EventArgs e)
        {
            Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;
            //para colocar las cabeceras
            List<string> cabec = new List<string>()
                {"contrato","sed","codruta","direccion","nombre","marca","serie","año","tarifa","promedio","fases","tecnico","fecha"};
            for (int i = 1; i <= 13; i++)
            {
                ws.Cells[1,i].Value = cabec[i-1];
            }

            int fila = 2;
            ws.Columns[13].NumberFormat = "@"; //para que la fecha no se modifique
            for (int i = 0; i < dgvProcesar.Rows.Count; i++)
            {
                SLDocument sl = new SLDocument(dgvProcesar.Rows[i].Cells[0].Value.ToString());
                int iRow = 2;               
                while (!string.IsNullOrEmpty(sl.GetCellValueAsString(iRow, 1)))
                {
                    for (int j = 1; j <= 11; j++)
                    {
                        Range rng = ws.Cells[fila, j];
                        rng.Value = sl.GetCellValueAsString(iRow, j);
                    }
                    ws.Cells[fila, 12].Value = dgvProcesar.Rows[i].Cells[1].Value;
                    ws.Cells[fila, 13].Value = dgvProcesar.Rows[i].Cells[2].Value;
                    iRow++;
                    fila++;
                }
            }

            Close();
        }

        private void frmFiles_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) Close();
        }
    }
}
