using AverageProgram.Classes;
using AverageProgram.Language;
using System;
using System.Drawing;
using System.Globalization;
using System.Threading;
using System.Windows.Forms;

namespace AverageProgram
{
    /// <summary>
    /// Clase principal del formulario.
    /// </summary>
    public partial class Form1 : Form
    {
        private bool isUpdating = false;
        private Color textBoxForeColor = Color.FromArgb(224, 224, 224);
        private CalculateAverage calculateAverage = new CalculateAverage();

        /// <summary>
        /// Constructor
        /// </summary>
        public Form1()
        {
            InitializeComponent();
            InitializeUI();
            SetNavigationBorder(btnHome);
        }

        #region Initialization

        private void Form1_Load(object sender, EventArgs e)
        {
            AssignTextBoxEvents();
        }

        private void InitializeUI()
        {
            unweightedAverageBar.Value = 0;
            weightedAverageBar.Value = 0;
            LoadLanguage("en");
        }

        /// <summary>
        /// Carga el idioma de la interfaz de usuario
        /// </summary>
        /// <param name="code"></param>
        public void LoadLanguage(string code)
        {
            // Configura la cultura actual
            Thread.CurrentThread.CurrentUICulture = new CultureInfo(code);

            // Asigna textos a los elementos de la interfaz de usuario
            AssignUITexts();
        }

        #endregion

        #region UI Update Methods

        //Asginar textos segun el idioma
        private void AssignUITexts()
        {
            label1.Text = Resources.StringYourSub;
            label2.Text = Resources.StringResults;
            label4.Text = Resources.StringExport;
            labelAddSub.Text = Resources.StringAddSub;

            label12.Text = Resources.StringName;
            label11.Text = Resources.StringCredit;
            label10.Text = Resources.StringAverage;
            labelFN.Text = Resources.StringFileName;
            label6.Text = Resources.StringFormat;
            label3.Text = Resources.StringWA;
            label8.Text = Resources.StringUA;

            subjectNameTb.Text = Resources.StringEnterName;
            creditTb.Text = Resources.StringEnterCredit;
            averageTb.Text = Resources.StringEnterAverage;
            fileNameTb.Text = Resources.StringEnterFileName;
            userNameTB.Text = Resources.StringUserEnterName;

            subjectNameTb.Tag = Resources.StringEnterName;
            creditTb.Tag = Resources.StringEnterCredit;
            averageTb.Tag = Resources.StringEnterAverage;
            fileNameTb.Tag = Resources.StringEnterFileName;
            userNameTB.Tag = Resources.StringUserEnterName;

            addBtn.Text = Resources.StringAdd;
            deleteBtn.Text = Resources.StringDelete;
            cancelBtn.Text = Resources.StringCancel;
            btnExport.Text = Resources.StringExport;

            dataGridView.Columns[0].HeaderText = Resources.StringSubject;
            dataGridView.Columns[1].HeaderText = Resources.StringCredit;
            dataGridView.Columns[2].HeaderText = Resources.StringAverage;
            dataGridView.Columns[3].HeaderText = Resources.StringGrade;

            formatComboBox.Text = Resources.StringFormat;
        }

        // Establece el borde de navegación
        private void SetNavigationBorder(Button selectedButton)
        {
            pnlNav.Height = selectedButton.Height;
            pnlNav.Top = selectedButton.Top;
            pnlNav.Left = selectedButton.Left;
        }

        // Asigna eventos a los cuadros de texto
        private void AssignTextBoxEvents()
        {
            subjectNameTb.Enter += TextBox_Enter;
            subjectNameTb.Leave += TextBox_Leave;

            creditTb.Enter += TextBox_Enter;
            creditTb.Leave += TextBox_Leave;

            averageTb.Enter += TextBox_Enter;
            averageTb.Leave += TextBox_Leave;

            fileNameTb.Enter += TextBox_Enter;
            fileNameTb.Leave += TextBox_Leave;

            userNameTB.Enter += TextBox_Enter;
            userNameTB.Leave += TextBox_Leave;
        }

        // Maneja el evento Enter de los cuadros de texto
        private void TextBox_Enter(object sender, EventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            if (textBox.Text == textBox.Tag.ToString())
            {
                textBox.Text = "";
                textBox.ForeColor = textBoxForeColor;
            }
        }

        // Maneja el evento Leave de los cuadros de texto
        private void TextBox_Leave(object sender, EventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            if (string.IsNullOrWhiteSpace(textBox.Text))
            {
                textBox.Text = textBox.Tag.ToString();
                textBox.ForeColor = SystemColors.GrayText;
            }
        }

        #endregion

        #region Data Grid

        // Maneja el evento Click del botón de agregar
        private void addBtn_Click(object sender, EventArgs e)
        {
            if (IsValidSubjectName() && IsValidCredit() && IsValidAverage())
            {
                // Calcula la nota para imprimir la escala
                char grade = calculateAverage.CalcularResultadoAEscala(int.Parse(averageTb.Text));

                string credit = creditTb.Text.StartsWith(".") ? "0" + creditTb.Text : creditTb.Text;

                if (!isUpdating)
                {
                    // Agrega una nueva fila al DataGridView
                    dataGridView.Rows.Add(subjectNameTb.Text, credit, averageTb.Text, grade);

                    // Actualiza los promedios
                    UpdateAverages();
                    ResetFormState();
                }
                else
                {
                    // Obtiene la fila seleccionada y la actualiza
                    DataGridViewRow selectedRow = dataGridView.SelectedRows[0];
                    UpdateSelectedRow(selectedRow, grade, credit);
                    UpdateAverages();
                }
            }
        }

        // Maneja el evento Click del botón de eliminar
        private void deleteBtn_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show(Resources.Avert1, "",MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

            if (result == DialogResult.OK)
            {
                if (dataGridView.SelectedRows.Count > 0)
                {
                    // Obtiene la fila seleccionada y la elimina
                    DataGridViewRow selectedRow = dataGridView.SelectedRows[0];
                    dataGridView.Rows.Remove(selectedRow);

                    // Restablece el estado del formulario y actualiza los promedios
                    ResetFormState();
                    UpdateAverages();
                }
            }
        }

        // Maneja el evento CellClick del DataGridView
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && dataGridView.Columns[e.ColumnIndex].Name == "btnColumn")
            {
                DataGridViewRow row = dataGridView.Rows[e.RowIndex];

                // Actualiza los cuadros de texto con los datos de la fila seleccionada
                UpdateTextBoxesWithRowData(row);
                UpdateUIForEditing();

                subjectNameTb.ForeColor = textBoxForeColor;
                creditTb.ForeColor = textBoxForeColor;
                averageTb.ForeColor = textBoxForeColor;

                isUpdating = true;
            }
        }

        // Actualiza la fila seleccionada con los datos de los cuadros de texto
        private void UpdateSelectedRow(DataGridViewRow selectedRow, char grade, string credit)
        {
            selectedRow.Cells[0].Value = subjectNameTb.Text;
            selectedRow.Cells[1].Value = credit;
            selectedRow.Cells[2].Value = averageTb.Text;
            selectedRow.Cells[3].Value = grade;
            ResetFormState();
        }

        // Actualiza los cuadros de texto con los datos de la fila seleccionada
        private void UpdateTextBoxesWithRowData(DataGridViewRow row)
        {
            subjectNameTb.Text = row.Cells["SubjectColumn"].Value.ToString();
            averageTb.Text = row.Cells["AverageColumn"].Value.ToString();
            creditTb.Text = row.Cells["CreditColumn"].Value.ToString();
        }

        #endregion

        #region Exception Verification

        // Verifica si el nombre de la asignatura es válido
        private bool IsValidSubjectName()
        {
            if (subjectNameTb.Text != subjectNameTb.Tag.ToString())
                return true;
            else
            {
                MessageBox.Show(Resources.Expt1, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        // Verifica si el crédito es válido
        private bool IsValidCredit()
        {
            double credit;
            if (double.TryParse(creditTb.Text, out credit) && credit > 0 && credit <= 5)
                return true;
            else
            {
                MessageBox.Show(Resources.Expt2, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        // Verifica si el promedio es válido
        private bool IsValidAverage()
        {
            int average;
            if (int.TryParse(averageTb.Text, out average) && average >= 0 && average <= 100)
                return true;
            else
            {
                MessageBox.Show(Resources.Expt3, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        #endregion

        #region Navigation Buttons

        // Maneja el evento Click del botón de inicio
        private void btnHome_Click(object sender, EventArgs e)
        {
            SetNavigationBorder(btnHome);
        }

        // Maneja el evento Click del botón de configuración
        private void btnSetting_Click(object sender, EventArgs e)
        {
            SettingForm settingForm = new SettingForm();
            settingForm.Show();
        }

        #endregion

        #region Averages Calculation

        // Actualiza los promedios
        private void UpdateAverages()
        {
            double weightedAverage = calculateAverage.CalculateOverallAverage(dataGridView, true);
            double unweightedAverage = calculateAverage.CalculateOverallAverage(dataGridView, false);

            labelUnweightedAverage.Text = unweightedAverage.ToString("0.00");
            labelWeightedAverage.Text = weightedAverage.ToString("0.00");

            
            UpdateProgressBar(weightedAverage, weightedAverageBar, labelWeightedAverage);
            UpdateProgressBar(unweightedAverage, unweightedAverageBar, labelUnweightedAverage);
        }

        // Actualiza las barras de progreso
        private void UpdateProgressBar(double average, CircularProgressBar.CircularProgressBar bar, Label averageLabel)
        {
            char grade = calculateAverage.CalcularResultadoAEscala(average);

            Color color = GetProgressColor(average);
            averageLabel.ForeColor = color;

            bar.ProgressColor = color;
           
            bar.Text = grade.ToString();
            bar.SubscriptText = average.ToString("0.00");
            bar.Value = Convert.ToInt32(average);
        }

        // Obtiene el color de la barra de progreso según el promedio
        private Color GetProgressColor(double average)
        {
            if (average >= 90)
                return Color.Lime;
            else if (average >= 80)
                return Color.Yellow;
            else if (average >= 70)
                return Color.Orange;
            else if (average >= 60)
                return Color.Red;
            else
                return Color.DarkRed;
        }

        #endregion

        #region Data Input Panel Methods

        // Maneja el evento Click del botón de cancelar
        private void cancelBtn_Click(object sender, EventArgs e)
        {
            ResetFormState();
        }

        // Actualiza la interfaz de usuario para la edición
        private void UpdateUIForEditing()
        {
            labelAddSub.Text = Resources.StringEditSub;
            addBtn.Text = Resources.StringUpdate;
            deleteBtn.Visible = true;
            cancelBtn.Visible = true;
        }

        // Restablece el estado del formulario
        private void ResetFormState()
        {
            subjectNameTb.Text = subjectNameTb.Tag.ToString();
            creditTb.Text = creditTb.Tag.ToString();
            averageTb.Text = averageTb.Tag.ToString();

            subjectNameTb.ForeColor = SystemColors.GrayText;
            creditTb.ForeColor = SystemColors.GrayText;
            averageTb.ForeColor = SystemColors.GrayText;

            deleteBtn.Visible = false;
            cancelBtn.Visible = false;
            labelAddSub.Text = Resources.StringAddSub;
            addBtn.Text = Resources.StringAdd;
            isUpdating = false;
        }

        #endregion

        #region Export Sec
        // Exporta los datos del DataGridView
        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                // Verifica si hay filas en el DataGridView
                if (dataGridView.Rows.Count == 0)
                {
                    throw new ArgumentException(Resources.Expt6, nameof(dataGridView.Rows.Count));
                }

                // Verifica si se ha proporcionado un nombre de archivo
                if (string.IsNullOrWhiteSpace(fileNameTb.Text) || fileNameTb.Text == fileNameTb.Tag.ToString())
                {
                    throw new ArgumentException(Resources.Expt4, nameof(fileNameTb.Text));
                }

                // Verifica si se ha seleccionado un formato de archivo
                if (formatComboBox.SelectedIndex == -1)
                {
                    throw new InvalidOperationException(Resources.Expt5);
                }

                // Exporta los datos según el formato seleccionado
                ExportData(fileNameTb.Text, formatComboBox.SelectedIndex);

                // Restablece los controles de entrada de datos
                fileNameTb.Text = Resources.StringEnterFileName;
                fileNameTb.ForeColor = SystemColors.GrayText;
                formatComboBox.SelectedIndex = -1;
                formatComboBox.Text = Resources.StringFormat;

                MessageBox.Show(Resources.Success1, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (ArgumentException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Exporta los datos según el formato seleccionado
        private void ExportData(string fileName, int formatIndex)
        {
            ExportManager exportManager = new ExportManager(dataGridView, labelWeightedAverage, labelUnweightedAverage);

            if (formatIndex == 0)
            {
                exportManager.ExportToExcel(fileName);
            }
            else if (formatIndex == 1)
            {
                exportManager.ExportToTXT(fileName);
            }
        }
        #endregion
    }
}
