using Microsoft.Office.Tools.Ribbon;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QRProgMed227
{
    public partial class rQRPM227
    {
        public static DataTable dtProg = new DataTable();
        private void rQRPM227_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnExp_Click(object sender, RibbonControlEventArgs e)
        {
            string rOrigen = "";
            OpenFileDialog ofdOrigen = new OpenFileDialog();
            ofdOrigen.Filter = "Archivo de texto|*.txt";
            DialogResult drOrigen = ofdOrigen.ShowDialog();
            if (drOrigen == DialogResult.OK)
            {
                rOrigen = ofdOrigen.FileName;
            }

            //agrega el archivo al datatable
            string consulta = "";
            string fileTxt = File.ReadAllText(rOrigen);
            int i = 0;
            foreach (string fila in fileTxt.Split('\n'))
            {                
                if (!string.IsNullOrEmpty(fila) & i > 0) //para empezar de la 2da fila
                {
                    int j = 0;
                    foreach (string colum in fila.Split('	'))
                    {
                        if (j == 2) //para que solo agarre la tercera columna (contrato)
                        {
                            consulta += colum + ",";
                        }                        
                        j++;
                    }                    
                }
                i++;
            }

            //para cargar el excel principal a un datatable
            string path = @"D:\QRProg\SuminSanMartin.xlsx";
            SLDocument sl = new SLDocument(path);
            int iRow = 2; //linea de inicio en excel
            DataTable dtAll = new DataTable();
            dtAll.Columns.Add("contrato");
            dtAll.Columns.Add("sed");
            dtAll.Columns.Add("codruta");
            dtAll.Columns.Add("direccion");
            dtAll.Columns.Add("nombre");
            dtAll.Columns.Add("marca");
            dtAll.Columns.Add("serie");
            dtAll.Columns.Add("anio");
            dtAll.Columns.Add("tarifa");
            dtAll.Columns.Add("promedio");
            dtAll.Columns.Add("fases");
            dtAll.Columns.Add("latitud");
            dtAll.Columns.Add("longitud");
            while (!string.IsNullOrEmpty(sl.GetCellValueAsString(iRow, 1)))
            {
                DataRow dr = dtAll.NewRow();
                dr[0] = sl.GetCellValueAsString(iRow, 1);
                dr[1] = sl.GetCellValueAsString(iRow, 2);
                dr[2] = sl.GetCellValueAsString(iRow, 3);
                dr[3] = sl.GetCellValueAsString(iRow, 4);
                dr[4] = sl.GetCellValueAsString(iRow, 5);
                dr[5] = sl.GetCellValueAsString(iRow, 6);
                dr[6] = sl.GetCellValueAsString(iRow, 7);
                dr[7] = sl.GetCellValueAsString(iRow, 8);
                dr[8] = sl.GetCellValueAsString(iRow, 9);
                dr[9] = sl.GetCellValueAsString(iRow, 10);
                dr[10] = sl.GetCellValueAsString(iRow, 11);
                dr[11] = sl.GetCellValueAsDouble(iRow, 14);
                dr[12] = sl.GetCellValueAsDouble(iRow, 15);
                dtAll.Rows.Add(dr);
                iRow++;
            }

            consulta = consulta.TrimEnd(','); //eliminamos la ultima coma
            DataTable dtTec = new DataTable();
            dtTec = dtAll.Clone(); //copiando la estructura del datatable
            DataRow[] drTec = dtAll.Select("contrato in(" + consulta + ")");
            dtTec = drTec.CopyToDataTable();

            //creando un archivo excel
            var excelApplication = new Microsoft.Office.Interop.Excel.Application();
            var excelWorkBook = excelApplication.Application.Workbooks.Add(Type.Missing);
            DataColumnCollection dataColumnCollection = dtTec.Columns;
            for (int m = 1; m <= dtTec.Rows.Count + 1; m++)
            {
                for (int n = 1; n <= dtTec.Columns.Count; n++)
                {
                    if (m == 1)
                        excelApplication.Cells[m, n] = dataColumnCollection[n - 1].ToString();
                    else
                        excelApplication.Cells[m, n] = dtTec.Rows[m - 2][n - 1].ToString();
                }
            }
            rOrigen = rOrigen.Replace(".txt", ".xlsx");
            excelApplication.ActiveWorkbook.SaveCopyAs(rOrigen);
            excelApplication.ActiveWorkbook.Saved = true;
            excelApplication.Quit();
        }

        private void btnFile_Click(object sender, RibbonControlEventArgs e)
        {
            dtProg = new DataTable(); //para reiniciar el datatable
            dtProg.Columns.Add("File");
            OpenFileDialog ofdOrigen = new OpenFileDialog();
            ofdOrigen.Multiselect = true;
            ofdOrigen.Filter = "Archivos Excel|*.xlsx";
            DialogResult drOrigen = ofdOrigen.ShowDialog();
            if (drOrigen == DialogResult.OK)
            {
                int i = 0;
                foreach (string s in ofdOrigen.FileNames)
                {
                    dtProg.Rows.Add(s);
                    i++;
                }
            }
            else
            {
                return;
            }
            frmFiles frm = new frmFiles();
            frm.ShowDialog();
        }
    }
}
