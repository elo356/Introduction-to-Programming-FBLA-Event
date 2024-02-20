using System;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace AverageProgram.Classes
{
    /// <summary>
    /// Clase encargada de manejar la exportación de datos a Excel y archivos de texto.
    /// </summary>
    internal class ExportManager
    {
        DataGridView dataGridView;
        Label labelWeightedAverage;
        Label labelUnweightedAverage;

        /// <summary>
        /// Constructor de la clase ExportManager.
        /// </summary>
        public ExportManager(DataGridView dataGV, Label labelWA, Label labelUA)
        {
            dataGridView = dataGV;
            labelWeightedAverage = labelWA;
            labelUnweightedAverage = labelUA;
        }

        /// <summary>
        /// Exporta los datos del DataGridView a un archivo de Excel.
        /// </summary>
        public void ExportToExcel(string fileName)
        {
            try
            {
                string downloadsPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads";
                string filePath = Path.Combine(downloadsPath, fileName + ".xlsx");

                // Crear una nueva aplicación Excel
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = true;

                // Crear una nueva hoja de trabajo
                Excel.Workbook workbook = excelApp.Workbooks.Add(Type.Missing);
                Excel.Worksheet worksheet = workbook.ActiveSheet as Excel.Worksheet;

                // Escribir los encabezados de columna en Excel
                int columnCount = dataGridView.Columns.Count;
                for (int i = 1; i <= columnCount - 1; i++) // Excluir la última columna
                {
                    worksheet.Cells[1, i] = dataGridView.Columns[i - 1].HeaderText;
                }

                // Escribir los datos de las celdas en Excel, excluyendo la última columna
                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    for (int j = 0; j < columnCount - 1; j++) // Excluir la última columna
                    {
                        worksheet.Cells[i + 2, j + 1] = dataGridView.Rows[i].Cells[j].Value?.ToString();
                    }
                }

                // Escribir los resultados al final del archivo de Excel
                int ultimaFila = dataGridView.Rows.Count + 3;
                worksheet.Cells[ultimaFila, 1] = "Promedio ponderado:";
                worksheet.Cells[ultimaFila, 2] = labelWeightedAverage.Text;
                worksheet.Cells[ultimaFila + 1, 1] = "Promedio no ponderado:";
                worksheet.Cells[ultimaFila + 1, 2] = labelUnweightedAverage.Text;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Exporta los datos del DataGridView a un archivo de texto.
        /// </summary>
        public void ExportToTXT(string fileName)
        {
            try
            {
                string downloadsPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads";
                string filePath = Path.Combine(downloadsPath, fileName + ".txt");

                // Crear o sobrescribir el archivo de texto
                using (StreamWriter sw = new StreamWriter(filePath))
                {
                    // Escribir los encabezados de columna
                    int columnCount = dataGridView.Columns.Count - 1; // Excluir la última columna
                    string columnHeaderLine = "+";
                    string columnHeaderValues = "|";
                    for (int i = 0; i < columnCount; i++)
                    {
                        columnHeaderLine += new string('-', 17) + "+";
                        columnHeaderValues += dataGridView.Columns[i].HeaderText.PadRight(17) + "|";
                    }
                    sw.WriteLine(columnHeaderLine);
                    sw.WriteLine(columnHeaderValues);
                    sw.WriteLine(columnHeaderLine);

                    // Escribir los datos de las celdas, excluyendo la última columna
                    for (int i = 0; i < dataGridView.Rows.Count; i++)
                    {
                        string rowDataLine = "|";
                        for (int j = 0; j < columnCount; j++)
                        {
                            string cellValue = dataGridView.Rows[i].Cells[j].Value?.ToString().PadRight(17);
                            rowDataLine += cellValue + "|";
                        }
                        sw.WriteLine(rowDataLine);
                        sw.WriteLine(columnHeaderLine);
                    }

                    // Escribir los promedios al final del archivo
                    sw.WriteLine("+---------------------------------------+");
                    sw.WriteLine($"| Promedio ponderado: {labelWeightedAverage.Text.PadRight(34)}");
                    sw.WriteLine($"| Promedio no ponderado: {labelUnweightedAverage.Text.PadRight(32)}");
                    sw.WriteLine("+---------------------------------------+");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
