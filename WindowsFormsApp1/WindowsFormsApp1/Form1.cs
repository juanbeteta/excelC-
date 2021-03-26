using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Text.RegularExpressions;
using System.IO;
using System.Diagnostics;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        public Form1()
        {
            InitializeComponent();
        }

        string obtenerFormato(string letra)
        {
            string salida = "";
            switch (letra)
            {
                case "C":
                    {
                        salida = "@";
                        break;
                    }
                case "N":
                    {
                        salida = "0";
                        break;
                    }
                case "F":
                    {
                        salida = "MM / DD / YYYY";
                        break;
                    }
                default: {
                        salida = "@";
                        break;
                    }
            }

            return salida;
        }

        bool comprobarTipo(string cadena, string tipo)
        {
            bool salida = false;
            switch (tipo)
            { 
                case "C":
                    {
                        int aux = 1;
                        salida = !Int32.TryParse(cadena, out aux);
                        break;
                    }
                case "N":
                    {
                        int aux = 1;
                        salida = Int32.TryParse(cadena, out aux);
                        break;
                    }
                case "F":
                    {
                        DateTime aux;
                        salida = DateTime.TryParse(cadena, out aux);
                        break;
                    }
                default:
                    {
                        salida = false;
                        break;
                    }
            }

            return salida;
        }

        string[] posicionCoordenadas (string[] columnas, Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            string[] salida = new string [columnas.Length];

            for (int i = 0; i < columnas.Length; i++)
            {
                Microsoft.Office.Interop.Excel.Range searchedRange = worksheet.get_Range("A1", "Z100");
                Microsoft.Office.Interop.Excel.Range currentFind = searchedRange.Find(columnas[i]);
                if (currentFind != null)
                {
                    int col = currentFind.Column;
                    int fil = currentFind.Row;

                    var obtenerTipoColumna = new String(columnas[i].Where(Char.IsLetter).ToArray());//Regex.Match(columnas[i], @"\d+").Value;
                    string tipoColumna1= obtenerFormato(obtenerTipoColumna);
                    string tipoColumna2 = obtenerTipoColumna;
                    salida[i] = fil + ";" + col + ";" + tipoColumna1 + ";" + tipoColumna2;
                }
            }

            return salida;
        }

        void GetExcelProcess(Microsoft.Office.Interop.Excel.Application excelApp)
        {
            int id;
            GetWindowThreadProcessId(excelApp.Hwnd, out id);
            Process.GetProcessById(id).Kill();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<string> errores = new List<string>();
            string fileName = @"d:\jmbeteta\Desktop\excel\Libro1.xlsx";
            
            //Create an excel application object
            Microsoft.Office.Interop.Excel.Application excelAppObj = new Microsoft.Office.Interop.Excel.Application();
            excelAppObj.DisplayAlerts = false;
            
            //Open the excel work book
            Microsoft.Office.Interop.Excel.Workbook workBook = excelAppObj.Workbooks.Open(fileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, false, false);

            //Get the first sheet of the selected work book
            //Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(1);

            //Datos 
            string fila0 =      "C1;C2;C3;N4;C5";
            string datosfila1 = "1;Miguel;S/N;0;hola";
            string datosfila2 = "Miguel;Juan;S/N;0;1";
            string datosfila3 = "2;Miguel;SA;s;a";
            
            //BD
            List<string> datosBD = new List<string> { datosfila1, datosfila2, datosfila3, datosfila3, datosfila2, datosfila1 };

            //obtengo posicion de la fila de plantilla
            string[] columbasExcel = fila0.Split(';');

            foreach (Microsoft.Office.Interop.Excel.Worksheet worksheet in workBook.Worksheets)
            {
                int col = 0, fil = 1;
                string[] columnasExcelCoordenadas = posicionCoordenadas(columbasExcel, worksheet);
                
                //itera las filas
                for (int i = 0; i < datosBD.LongCount(); i++)
                {
                    //split fila 
                    var fila = datosBD[i].Split(';').ToList();
                    
                    //pinta las filas
                    for (int j = 0; j < columnasExcelCoordenadas.Length; j++)
                    {
                        if(columnasExcelCoordenadas[j] != null) {
                            var coordenadas = columnasExcelCoordenadas[j].Split(';').ToList();

                            //obtiene el formato de la plantilla
                            var format = coordenadas[2];
                           
                            var x = int.Parse(coordenadas[0]);
                            var y = int.Parse(coordenadas[1]);

                            //filtra las columnas de cada fila
                            var fileCelda = fila.Where((c, index) => index == j).First();

                            if (!comprobarTipo(fileCelda, coordenadas[3]))
                            {
                                errores.Add("Error celda: " + (char)(y-1 + 'A') + (fil + x).ToString());
                               
                            }

                            worksheet.Cells[fil + x, col + y].NumberFormat = format;
                            worksheet.Cells[fil + x, col + y] = fileCelda;
                        }
                    }
                    fil++;
                }

                //contar las filas
                for (int j = 0; j < columnasExcelCoordenadas.Length; j++)
                {
                    if (columnasExcelCoordenadas[j] != null)
                    {
                        var coordenadas = columnasExcelCoordenadas[j].Split(';').ToList();

                        //obtiene el formato de la plantilla
                        var format = coordenadas[2];
                        var x = int.Parse(coordenadas[0]);
                        var y = int.Parse(coordenadas[1]);
                        worksheet.Cells[x + fil, y] = "=ROWS(" + (char)(y - 1 + 'A') + (x + 1).ToString() + ":" + (char)(y - 1 + 'A') + (fil + x - 1).ToString() + ")";
                    }
                }
            }
            
            //Save work book (.xlsx format)
            workBook.SaveAs(fileName,
                Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, 
                Type.Missing, 
                Type.Missing, 
                false, 
                false, 
                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, 
                Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, 
                Type.Missing, 
                Type.Missing);

            //kill process
            GetExcelProcess(excelAppObj);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DateTime aux;
            bool salida = DateTime.TryParse("5/1/x", out aux);
        }
    }
}
