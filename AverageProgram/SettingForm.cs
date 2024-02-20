using AverageProgram.Language;
using System;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace AverageProgram
{
    public partial class SettingForm : Form
    {
        /// <summary>
        /// Constructor de SettingForm
        /// </summary>
        public SettingForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Este método se ejecuta cuando el formulario de configuración se carga.
        /// </summary>
        private void SettingForm_Load(object sender, EventArgs e)
        {
            // Asignar el texto del label de idioma utilizando los recursos de lenguaje
            label6.Text = Resources.StringLanguage;

            // Obtener el código de idioma actual del hilo actual
            string currentLanguageCode = Thread.CurrentThread.CurrentUICulture.Name;
        
            // Establecer la selección del ComboBox de idioma según el idioma actual
            if (currentLanguageCode == "es")
            {
                langComboBox.SelectedIndex = 1; // Español
            }
            else
            {
                langComboBox.SelectedIndex = 0; // Inglés (predeterminado)
            }
        }

        /// <summary>
        /// Este método se ejecuta cuando se hace clic en el botón de aceptar.
        /// </summary>
        private void okBtn_Click(object sender, EventArgs e)
        {
            // Establecer la cultura actual del hilo según la selección del ComboBox de idioma
            if (langComboBox.SelectedIndex == 0) 
            {
                Thread.CurrentThread.CurrentUICulture = new CultureInfo(""); // Inglés
            }
            else
            {
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("es"); // Español
            }

            // Obtener una instancia del formulario principal (Form1)
            Form1 form1 = Application.OpenForms.OfType<Form1>().FirstOrDefault();

            // Verificar si el formulario principal existe
            if (form1 != null)
            {
                // Cargar el idioma en el formulario principal
                form1.LoadLanguage(Thread.CurrentThread.CurrentUICulture.Name);
            }

            // Cerrar el formulario de configuración
            this.Close();
        }
    }
}
