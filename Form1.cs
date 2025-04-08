using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace ProcesadorTextoExcel
{
    public partial class Form1 : Form
    {
        // Quitar required y hacer private
        private TextBox txtRutaArchivo;
        private Button btnCargarArchivo;
        private Button btnExportarExcel;
        private Label lblInstruccion;
        private DataGridView dgvPreview;
        private Label lblEstado;

        public Form1()
        {
            InitializeComponent();
            ConfigurarControles();
        }

        private void ConfigurarControles()
        {
            // Crear instancias de los controles
            txtRutaArchivo = new TextBox();
            btnCargarArchivo = new Button();
            btnExportarExcel = new Button();
            lblInstruccion = new Label();
            dgvPreview = new DataGridView();
            lblEstado = new Label();

            // Configurar formulario
            this.Text = "Procesador Texto a Excel";
            this.Width = 600;
            this.Height = 450;
            this.StartPosition = FormStartPosition.CenterScreen;

            // Etiqueta de instrucción
            lblInstruccion.Text = "Seleccione un archivo de texto:";
            lblInstruccion.Location = new System.Drawing.Point(10, 15);
            lblInstruccion.AutoSize = true;

            // Campo para mostrar la ruta del archivo
            txtRutaArchivo.Width = 400;
            txtRutaArchivo.Location = new System.Drawing.Point(10, 40);
            txtRutaArchivo.ReadOnly = true;

            // Botón para cargar archivo
            btnCargarArchivo.Text = "Cargar TXT";
            btnCargarArchivo.Location = new System.Drawing.Point(420, 38);
            btnCargarArchivo.Size = new System.Drawing.Size(90, 25);

            // Botón para exportar a Excel
            btnExportarExcel.Text = "Exportar Excel";
            btnExportarExcel.Location = new System.Drawing.Point(10, 80);
            btnExportarExcel.Size = new System.Drawing.Size(120, 30);
            btnExportarExcel.Enabled = false;

            // Etiqueta para mostrar estado
            lblEstado.Text = "Estado: Esperando archivo...";
            lblEstado.Location = new System.Drawing.Point(140, 85);
            lblEstado.Width = 400;
            lblEstado.AutoSize = true;

            // DataGridView para previsualizar los datos
            dgvPreview.Location = new System.Drawing.Point(10, 120);
            dgvPreview.Size = new System.Drawing.Size(560, 280);
            dgvPreview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvPreview.AllowUserToAddRows = false;
            dgvPreview.AllowUserToDeleteRows = false;
            dgvPreview.ReadOnly = true;
            dgvPreview.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvPreview.RowHeadersVisible = false;

            // Asignar eventos
            btnCargarArchivo.Click += btnCargarArchivo_Click;
            btnExportarExcel.Click += btnExportarExcel_Click;

            // Agregar controles al formulario
            Controls.Add(lblInstruccion);
            Controls.Add(txtRutaArchivo);
            Controls.Add(btnCargarArchivo);
            Controls.Add(btnExportarExcel);
            Controls.Add(lblEstado);
            Controls.Add(dgvPreview);
        }

        // Corregido para manejar posibles valores nulos
        private void btnCargarArchivo_Click(object? sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Archivos de texto|*.txt";
                openFileDialog.Title = "Seleccionar archivo de texto";
                
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtRutaArchivo.Text = openFileDialog.FileName;
                    
                    try
                    {
                        ProcesarYMostrarDatos(openFileDialog.FileName);
                        btnExportarExcel.Enabled = true;
                        lblEstado.Text = "Estado: Archivo procesado correctamente.";
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error al procesar el archivo: {ex.Message}", 
                                      "Error", 
                                      MessageBoxButtons.OK, 
                                      MessageBoxIcon.Error);
                        lblEstado.Text = "Estado: Error al procesar el archivo.";
                    }
                }
            }
        }

        // Para manejar las advertencias de nulos, agregar verificaciones
        private void ProcesarYMostrarDatos(string filePath)
        {
            string texto = File.ReadAllText(filePath);
            var datos = ExtraerDatos(texto);
            
            if (dgvPreview != null)
            {
                dgvPreview.DataSource = null;
                dgvPreview.Columns.Clear();
                dgvPreview.DataSource = datos;
                
                // Verificar que las columnas existen antes de acceder a ellas
                if (dgvPreview.Columns["Nombre"] != null)
                    dgvPreview.Columns["Nombre"].HeaderText = "NOMBRE";
                
                if (dgvPreview.Columns["Apellido"] != null)
                    dgvPreview.Columns["Apellido"].HeaderText = "APELLIDO";
                
                if (dgvPreview.Columns["Email"] != null)
                    dgvPreview.Columns["Email"].HeaderText = "CASILLA-ELECTRONICA";
            }
        }

        private List<PersonaData> ExtraerDatos(string texto)
        {
            var personas = new List<PersonaData>();
            
            // Dividir el texto en líneas
            var lineas = texto.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            
            foreach (var linea in lineas)
            {
                // Buscar el email en la línea
                var emailMatch = Regex.Match(linea, @"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b");
                
                if (emailMatch.Success)
                {
                    string email = emailMatch.Value;
                    
                    // Extraer el nombre que está antes del email
                    string nombreCompleto = linea.Substring(0, emailMatch.Index).Trim();
                    
                    // Dividir en nombre y apellido
                    var partes = nombreCompleto.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    
                    string nombre = "";
                    string apellido = "";
                    
                    if (partes.Length >= 2)
                    {
                        // El último elemento es el apellido
                        apellido = partes[partes.Length - 1];
                        
                        // El resto es el nombre
                        nombre = string.Join(" ", partes.Take(partes.Length - 1));
                    }
                    else if (partes.Length == 1)
                    {
                        nombre = partes[0];
                    }
                    
                    personas.Add(new PersonaData
                    {
                        Nombre = nombre,
                        Apellido = apellido,
                        Email = email
                    });
                }
            }
            
            return personas;
        }

        // Corregido para manejar posibles valores nulos
        private void btnExportarExcel_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtRutaArchivo.Text))
            {
                MessageBox.Show("Debe seleccionar un archivo de texto primero.", 
                              "Advertencia", 
                              MessageBoxButtons.OK, 
                              MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // Obtener los datos que ya están en el DataGridView
                if (dgvPreview.DataSource is not List<PersonaData> datos || datos.Count == 0)
                {
                    MessageBox.Show("No hay datos para exportar.", 
                                  "Advertencia", 
                                  MessageBoxButtons.OK, 
                                  MessageBoxIcon.Warning);
                    return;
                }

                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Archivos Excel|*.xlsx";
                    saveFileDialog.Title = "Guardar archivo Excel";
                    saveFileDialog.FileName = "Datos.xlsx";
                    
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        GenerarExcel(datos, saveFileDialog.FileName);
                        MessageBox.Show("Excel generado exitosamente!\nUbicación: " + saveFileDialog.FileName, 
                                      "Éxito", 
                                      MessageBoxButtons.OK, 
                                      MessageBoxIcon.Information);
                        lblEstado.Text = "Estado: Excel generado correctamente.";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al generar Excel: {ex.Message}", 
                              "Error", 
                              MessageBoxButtons.OK, 
                              MessageBoxIcon.Error);
                lblEstado.Text = "Estado: Error al generar Excel.";
            }
        }

        private void GenerarExcel(List<PersonaData> datos, string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Datos");
                
                // Configurar encabezados
                worksheet.Cell(1, 1).Value = "NOMBRE";
                worksheet.Cell(1, 2).Value = "APELLIDO";
                worksheet.Cell(1, 3).Value = "CASILLA-ELECTRONICA";
                
                // Dar formato a los encabezados
                var headerRange = worksheet.Range("A1:C1");
                headerRange.Style
                    .Font.SetBold(true)
                    .Fill.SetBackgroundColor(XLColor.LightGray)
                    .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                
                // Agregar datos
                for (int i = 0; i < datos.Count; i++)
                {
                    worksheet.Cell(i + 2, 1).Value = datos[i].Nombre;
                    worksheet.Cell(i + 2, 2).Value = datos[i].Apellido;
                    worksheet.Cell(i + 2, 3).Value = datos[i].Email;
                }
                
                // Ajustar columnas al contenido
                worksheet.Columns().AdjustToContents();
                
                // Guardar el archivo
                workbook.SaveAs(filePath);
            }
        }

        // Clase para almacenar los datos de cada persona
        // Corregida para aceptar valores nulos o inicializar propiedades
        public class PersonaData
        {
            public string Nombre { get; set; } = string.Empty;
            public string Apellido { get; set; } = string.Empty; 
            public string Email { get; set; } = string.Empty;
        }

        // Eliminado el método Main para evitar el conflicto de punto de entrada
    }
}