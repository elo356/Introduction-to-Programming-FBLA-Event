using System;
using System.Windows.Forms;

namespace AverageProgram.Classes
{
    internal class CalculateAverage
    {
        public double CalculateOverallAverage(DataGridView dataGridView, bool weighted)
        {
            string creditColumnName = "creditColumn";
            string gradeColumnName = "averageColumn";

            double totalGrades = 0;
            double totalCredits = 0;

            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                double grade = 0;
                double credit = 0;

                // Verifica si las celdas no son nulas y si los valores pueden ser convertidos a double
                if (row.Cells[gradeColumnName].Value != null && row.Cells[creditColumnName].Value != null &&
                    double.TryParse(row.Cells[gradeColumnName].Value.ToString(), out grade) &&
                    double.TryParse(row.Cells[creditColumnName].Value.ToString(), out credit))
                {
                    if (weighted)
                    {
                        totalGrades += grade * credit;
                        totalCredits += credit;
                    }
                    else
                    {
                        totalGrades += grade;
                    }
                }
            }

            double average = 0;
            if (weighted)
            {
                if (totalCredits != 0)
                {
                   average =  totalGrades / totalCredits;
                }
            }
            else
            {
                int rowCount = dataGridView.Rows.Count;
                if (rowCount != 0)
                {
                    average = totalGrades / rowCount;
                }
            }

            return Math.Round(average, 2);
        }

        public char CalcularResultadoAEscala(double average)
        {
            if (average >= 90)
            {
                return 'A';
            }
            else if (average >= 80)
            {
                return 'B';
            }
            else if (average >= 70)
            {
                return 'C';
            }
            else if (average >= 60)
            {
                return 'D';
            }
            else return 'F';
        }
    }
}