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

namespace ExcelCsharp
{
    public partial class Form1 : Form
    {
        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);
        static string dbtxt = "db.txt";
        static string excelName = "Libro11.xlsx";
        static string excelNamePlantilla = "Libro1.xlsx";
        static char[] config = new char[] { '#', '$', '%', '&'};
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



        void generarDatos(List<Customer> customers)
        {
            string createText = "C1#$%;C2#$%;C3#$%;C4#$%;N5#$%\n";

            foreach (Customer customer in customers)
            {
                createText += customer.Name + ";" + customer.Email + ";" + customer.City + ";" + customer.Phone + ";" + customer.Country + "\n";
            }

            File.WriteAllText(obtenerRuta(dbtxt), createText);
        }

        string [] Obtenertipos(string columna)
        {
            List<string> tipo = new List<string>();

            var regexItem = new Regex("^[a-zA-Z0-9 ]*$");
            
            for(int i = 0; i < columna.Length; i++)
            {
              
                if (!regexItem.IsMatch(columna[i].ToString()))
                {
                    tipo.Add(columna[i].ToString());
                    i++;
                    string txto = "";
                    for (int k = i; k < columna.Length && regexItem.IsMatch(columna[k].ToString()) ; k++)
                    {
                        txto += columna[k];
                    }
                    tipo.Add(txto);
                }
            }
            return tipo.ToArray();
        }

        string[] posicionCoordenadas (string[] columnas, Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            string[] salida = new string [columnas.Length];

            for (int i = 0; i < columnas.Length; i++)
            {
                Microsoft.Office.Interop.Excel.Range searchedRange = worksheet.get_Range("A1", "Z100");
                

                string obtenerColumna = new string(columnas[i].Split(config)[0].ToArray());
                
                Microsoft.Office.Interop.Excel.Range currentFind = searchedRange.Find(obtenerColumna);
                if (currentFind != null)
                {
                    int col = currentFind.Column;
                    int fil = currentFind.Row;

                    var obtenerColorColumna = "";
                    var obtenerAlignColumna = "";
                    var obtenerFontColumna = "";

                    var obtenerTipoColumna = new string(columnas[i].Split(config)[0].Where(Char.IsLetter).ToArray());
                    string[] tipos = Obtenertipos(columnas[i]);

                    for (int j = 0; j < tipos.Length; j++)
                    {
                        switch (tipos[j])
                        {
                            case "#":
                                {
                                    obtenerColorColumna = tipos[++j];
                                    break;
                                }
                            case "$":
                                {
                                    obtenerAlignColumna = tipos[++j];
                                    break;
                                }
                            case "%":
                                {
                                    obtenerFontColumna = tipos[++j];
                                    break;
                                }
                        }
                    }

                    
                    string tipoColumna1= obtenerFormato(obtenerTipoColumna);
                    string tipoColumna2 = obtenerTipoColumna;
                    salida[i] = fil + ";" + col + ";" + tipoColumna1 + ";" + tipoColumna2 + ";" + obtenerColorColumna + ";" + obtenerAlignColumna + ";" + obtenerFontColumna;
                }
            }

            return salida;
        }
        
        void MostrarErrores(List<string> errores)
        {
            string txto = "";

            foreach(string error in errores)
            {
                txto += error + "\n";
            }
           
            using (StreamWriter outfile = new StreamWriter(obtenerRuta("errores.txt"), true))
            {
                outfile.WriteLine(DateTime.Now + " " + "\n" + txto);
            }
        }

        Microsoft.Office.Interop.Excel.XlHAlign Alinear(string alinear)
        {
            Microsoft.Office.Interop.Excel.XlHAlign salida;

            switch (alinear)
            {
                case "derecha":
                    {
                        salida = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                        break;
                    }
                case "izquierda":
                    {
                        salida = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                        break;
                    }
                case "centro":
                    {
                        salida = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        break;
                    }
                default:
                    {
                        salida = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignGeneral;
                        break;
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

        string obtenerRuta(string file = null)
        {
            string exeFile = new System.Uri(Assembly.GetEntryAssembly().CodeBase).AbsolutePath;
            string Dir = Path.GetDirectoryName(exeFile);
            string path = Path.GetFullPath(Path.Combine(Dir, @"..\..\..\..\" + file));
            
            return path;
        }

        Color obtenerColor(string color)
        {
            Color salida;

            switch(color){
                case "rojo":
                    {
                        salida = System.Drawing.Color.Red;
                        break;
                    }
                case "verde":
                    {
                        salida = System.Drawing.Color.Green;
                        break;
                    }
                case "azul":
                    {
                        salida = System.Drawing.Color.Blue;
                        break;
                    }
                case "amarillo":
                    {
                        salida = System.Drawing.Color.Yellow;
                        
                        break;
                    }
                default:
                    {
                        salida = System.Drawing.Color.Transparent;
                        break;
                    }
            }

            return salida;
        }

        void obtenerFont(Microsoft.Office.Interop.Excel.Worksheet worksheet, string[] range, string font)
        {
            switch (font)
            {     
                case "Strikethrough":
                    {
                        worksheet.get_Range(range[0], range[1]).Cells.Font.Strikethrough = true;
                        break;
                    }
                case "under":
                    {
                        worksheet.get_Range(range[0], range[1]).Cells.Font.Underline = true;
                        break;
                    }
                case "boldUnder":
                    {
                        worksheet.get_Range(range[0], range[1]).Cells.Font.Bold = true;
                        worksheet.get_Range(range[0], range[1]).Cells.Font.Underline = true;
                        break;
                    }
            }
        }

        string[] comprobarFilaPrincipal(string filaPrincipal)
        {
            string esValido = "1";
            string columna = "";

            var regexItem = new Regex("^[a-zA-Z0-9 ]*$");

            for (int i = 0; i < filaPrincipal.Length; i++)
            {
                if ((filaPrincipal[i].ToString() != ";" && !regexItem.IsMatch(filaPrincipal[i].ToString())) && !filaPrincipal[i].ToString().Any(item => config.Any(item1 => item == item1)))
                {
                    columna += filaPrincipal[i].ToString() + " ";
                    esValido = "0";
                }
            }

            return new string[]{ esValido, columna };
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                // lista de errores
                List<string> errores = new List<string>();

                //copiar la plantilla original
                FileInfo fi = new FileInfo(obtenerRuta(excelNamePlantilla));
                fi.CopyTo(obtenerRuta(excelName), true);
                
                //Obtener Ruta 
                string fileName = obtenerRuta(excelName);
                
                //Create an excel application object
                Microsoft.Office.Interop.Excel.Application excelAppObj = new Microsoft.Office.Interop.Excel.Application();
                excelAppObj.DisplayAlerts = false;

                //Open the excel work book
                Microsoft.Office.Interop.Excel.Workbook workBook = excelAppObj.Workbooks.Open(fileName,
                                                                                                     0,
                                                                                                 false,
                                                                                                     5,
                                                                                                    "",
                                                                                                    "",
                                                                                                 false,
                                                                                                 Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,
                                                                                                    "",
                                                                                                  true,
                                                                                                 false,
                                                                                                     0,
                                                                                                 false,
                                                                                                 false);

                //vuelco los datos del .txt a un array
                List<string> db = System.IO.File.ReadAllLines(obtenerRuta(dbtxt)).ToList();

                //filtro la primera fila y vuelco las demás filas al array
                List<string> filas = db.Where((item, indexer) => indexer > 0).ToList();

                //obtengo la primera fila 
                string[] columbasExcel = db[0].ToString().Split(';');
                
                string[] comprobarFilaP = comprobarFilaPrincipal(db[0].ToString());
                if (comprobarFilaP[0] == "0")
                {
                    errores.Add("Error en la configuración de la columna, estas configuraciones no existen: " + comprobarFilaP[1] + "\n" + "solo se pueden configurar con las: " + new String(config));
                }

                //recorro cada hoja
                foreach (Microsoft.Office.Interop.Excel.Worksheet worksheet in workBook.Worksheets)
                {
                    int col = 0, fil = 1;
                    string[] columnasExcelCoordenadas = posicionCoordenadas(columbasExcel, worksheet);

                    //itera las filas
                    for (int i = 0; i < filas.LongCount(); i++)
                    {
                        //split fila 
                        var fila = filas[i].Split(';').ToList();

                        //pinta las filas
                        for (int j = 0; j < columnasExcelCoordenadas.Length; j++)
                        {
                            if (columnasExcelCoordenadas[j] != null)
                            {
                                var coordenadas = columnasExcelCoordenadas[j].Split(';').ToList();

                                //obtiene el formato de la plantilla
                                var format = coordenadas[2];

                                var x = int.Parse(coordenadas[0]);
                                var y = int.Parse(coordenadas[1]);

                                //filtra las columnas de cada fila
                                var fileCelda = fila.Where((c, index) => index == j).First();

                                //listar errores de cada celda guardando su ubicación
                                if (!comprobarTipo(fileCelda, coordenadas[3]))
                                {
                                    worksheet.Cells[fil + x, col + y].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                    errores.Add("Error celda: " + (char)(y - 1 + 'A') + (fil + x).ToString() + " Hoja: " + worksheet.Name);
                                }

                                if (fileCelda != "")
                                {
                                    worksheet.Cells[fil + x, col + y].NumberFormat = format;
                                    worksheet.Cells[fil + x, col + y] = fileCelda;
                                }
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
                            var letraInicio = (char)(y - 1 + 'A');
                            var numeroInicio = (x + 1).ToString();
                            var numeroFinal = (fil + x - 1).ToString();
                            worksheet.Cells[x + fil, y].NumberFormat = "";
                            worksheet.Cells[x + fil, y] = "=CountA(" + letraInicio + numeroInicio + ":" + letraInicio + numeroFinal + ")";
                         
                            //#
                            if (coordenadas[4] != "")
                            {
                                worksheet.get_Range(letraInicio + numeroInicio, letraInicio + numeroFinal).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromName(coordenadas[4]));
                            }

                            //$
                            if(coordenadas[5] != "")
                            {
                                worksheet.get_Range(letraInicio + numeroInicio, letraInicio + numeroFinal).Cells.HorizontalAlignment = Alinear(coordenadas[5]);
                            }

                            //%
                            if (coordenadas[6] != "")
                            {
                                obtenerFont(worksheet, new string[] { letraInicio + numeroInicio, letraInicio + numeroFinal }, coordenadas[6]);
                            }
                        }
                    }
                    // autoajusto el tamaño de todas las celdas
                    worksheet.Columns["A:Z"].Autofit();
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

                //listo los errores de las celdas y las muestro
                MostrarErrores(errores.Distinct().ToList());

                MessageBox.Show("Volcado terminado");

            }catch(Exception ex)
            {
                MessageBox.Show("Error: " +ex.Message.ToString());
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            int numeros = textBox1.Text == "" ? 5 : int.Parse(textBox1.Text);
            var repository = new SampleCustomerRepository();
            List<Customer> customers = (List<Customer>)repository.GetCustomers(numeros);

            try
            {
                generarDatos(customers);
                MessageBox.Show("Archivos volcados a la ruta: " + obtenerRuta(dbtxt));
            }
            catch
            {
                MessageBox.Show("Error al volcar los archivos.");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(obtenerRuta(excelName));
        }

        private void button4_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(obtenerRuta(dbtxt));
        }
    }
}
