using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace tp1_lab3
{
    public partial class Form1: Form
    {
        // Cadena de conexión a la base de datos (asume una base Access)
        private string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Datos\Empleados.accdb";
        private OleDbConnection connection;
        private OleDbDataAdapter adapter;
        private DataTable dataTable;
        public Form1()
        {
            InitializeComponent();
        }

        private void FormEmpleados_Load(object sender, EventArgs e)
        {
            // Inicializar la conexión y cargar datos al inicio
            CargarDatos();
        }

        private void CargarDatos()
        {
            try
            {
                // Crear conexión y adaptador
                connection = new OleDbConnection(connectionString);
                string consulta = "SELECT * FROM Empleados";
                adapter = new OleDbDataAdapter(consulta, connection);

                // Llenar DataTable y asignar al DataGridView
                dataTable = new DataTable();
                adapter.Fill(dataTable);
                dataGridViewEmpleados.DataSource = dataTable;

                // Mostrar mensaje de éxito
                lblEstado.Text = "Datos cargados correctamente";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al cargar datos: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblEstado.Text = "Error al cargar datos";
            }
        }

        private void btnAgregar_Click(object sender, EventArgs e)
        {
            try
            {
                // Validar datos de entrada
                if (string.IsNullOrWhiteSpace(txtNombre.Text) || string.IsNullOrWhiteSpace(txtApellido.Text))
                {
                    MessageBox.Show("Por favor ingrese nombre y apellido", "Datos incompletos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Abrir conexión
                connection.Open();

                // Preparar consulta de inserción
                string insertQuery = "INSERT INTO Empleados (Nombre, Apellido, Cargo, Salario) VALUES (?, ?, ?, ?)";
                OleDbCommand cmd = new OleDbCommand(insertQuery, connection);

                // Agregar parámetros
                cmd.Parameters.AddWithValue("@Nombre", txtNombre.Text);
                cmd.Parameters.AddWithValue("@Apellido", txtApellido.Text);
                cmd.Parameters.AddWithValue("@Cargo", txtCargo.Text);

                // Convertir salario si se proporciona
                decimal salario = 0;
                if (!string.IsNullOrEmpty(txtSalario.Text) && decimal.TryParse(txtSalario.Text, out salario))
                {
                    cmd.Parameters.AddWithValue("@Salario", salario);
                }
                else
                {
                    cmd.Parameters.AddWithValue("@Salario", DBNull.Value);
                }

                // Ejecutar comando
                cmd.ExecuteNonQuery();

                // Cerrar conexión
                connection.Close();

                // Limpiar campos
                LimpiarCampos();

                // Recargar datos
                CargarDatos();

                lblEstado.Text = "Empleado agregado correctamente";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al agregar empleado: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblEstado.Text = "Error al agregar empleado";

                // Asegurar que la conexión se cierra en caso de error
                if (connection.State == ConnectionState.Open)
                    connection.Close();
            }
        }

        private void LimpiarCampos()
        {
            txtNombre.Text = "";
            txtApellido.Text = "";
            txtCargo.Text = "";
            txtSalario.Text = "";
            txtNombre.Focus();
        }

        private void btnGuardarArchivo_Click(object sender, EventArgs e)
        {
            try
            {
                // Verificar si hay datos para guardar
                if (dataTable == null || dataTable.Rows.Count == 0)
                {
                    MessageBox.Show("No hay datos para guardar", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Mostrar diálogo para seleccionar ubicación del archivo
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Archivos de texto (*.txt)|*.txt";
                saveDialog.Title = "Guardar lista de empleados";
                saveDialog.FileName = "Empleados_" + DateTime.Now.ToString("yyyyMMdd");

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    // Crear archivo de texto
                    using (StreamWriter writer = new StreamWriter(saveDialog.FileName))
                    {
                        // Escribir encabezado
                        writer.WriteLine("LISTA DE EMPLEADOS - " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                        writer.WriteLine(new string('-', 80));

                        // Escribir títulos de columnas
                        writer.WriteLine(String.Format("{0,-5} {1,-20} {2,-20} {3,-15} {4,-10}",
                            "ID", "Nombre", "Apellido", "Cargo", "Salario"));
                        writer.WriteLine(new string('-', 80));

                        // Escribir datos de cada empleado
                        foreach (DataRow row in dataTable.Rows)
                        {
                            writer.WriteLine(String.Format("{0,-5} {1,-20} {2,-20} {3,-15} {4,-10}",
                                row["ID"],
                                row["Nombre"],
                                row["Apellido"],
                                row["Cargo"] ?? "N/A",
                                row["Salario"] != DBNull.Value ? string.Format("{0:C}", row["Salario"]) : "N/A"));
                        }

                        // Escribir pie de página
                        writer.WriteLine(new string('-', 80));
                        writer.WriteLine("Total de empleados: " + dataTable.Rows.Count);
                    }

                    MessageBox.Show("Archivo guardado exitosamente en: " + saveDialog.FileName,
                        "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    lblEstado.Text = "Archivo guardado correctamente";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al guardar archivo: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblEstado.Text = "Error al guardar archivo";
            }
        }

        private void btnLimpiar_Click(object sender, EventArgs e)
        {
            LimpiarCampos();
        }

        private void FormEmpleados_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Cerrar recursos si están abiertos
            if (connection != null && connection.State == ConnectionState.Open)
                connection.Close();
        }
    }
}
