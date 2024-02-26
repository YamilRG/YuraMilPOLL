using System.Data.SqlClient;
using System;
using System.Windows.Forms;
using System.IO;
using System.Linq;
using System.Drawing;

namespace test3
{
    partial class Form1
    {
        private Label FechaInicio;
        private Label FechaFinal;
        private TextBox startDateTimeTextBox;
        private TextBox endDateTimeTextBox;
        private CheckBox intervalCheckBox;
        private TextBox serverTextBox;
        private TextBox databaseTextBox;
        private TextBox usernameTextBox;
        private TextBox passwordTextBox;
        private TextBox tableTextBox;
        private Button getDataButton;
        private DataGridView dataGridView;
        private Button exportDataGridView;
        private Button cleanDataGridButton;
        private PictureBox logoImage;
        private PictureBox logoImage2;
        private Label serverText;
        private Label databaseText;
        private Label usernameText;
        private Label passwordText;
        private Label tableText;
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Text = "YuraMIL Poll";

            intervalCheckBox = new CheckBox();
            intervalCheckBox.Text = "Activar Intervalo de Fecha";
            intervalCheckBox.CheckedChanged += ToggleIntervalFields;
            intervalCheckBox.Location = new System.Drawing.Point(10, 85);
            intervalCheckBox.Size = new System.Drawing.Size(155, 20);
            this.Controls.Add(intervalCheckBox);


            FechaInicio = new Label();
            FechaInicio.Text = "Fecha Inicio:";
            FechaInicio.Location = new System.Drawing.Point(10, 115);
            FechaInicio.Size = new System.Drawing.Size(70, 15);
            this.Controls.Add(FechaInicio);

            FechaFinal = new Label();
            FechaFinal.Text = "Fecha Final:";
            FechaFinal.Location = new System.Drawing.Point(220, 115);
            FechaFinal.Size = new System.Drawing.Size(65, 15);
            this.Controls.Add(FechaFinal);

            startDateTimeTextBox = new TextBox();
            startDateTimeTextBox.Enabled = false;
            startDateTimeTextBox.Location = new System.Drawing.Point(80, 112);
            this.Controls.Add(startDateTimeTextBox);


            endDateTimeTextBox = new TextBox();
            endDateTimeTextBox.Enabled = false;
            endDateTimeTextBox.Location = new System.Drawing.Point(290, 112);
            this.Controls.Add(endDateTimeTextBox);

            cleanDataGridButton = new Button();
            cleanDataGridButton.Text = "Limpiar";
            cleanDataGridButton.Click += cleanDataGridView;
            cleanDataGridButton.Location = new System.Drawing.Point(470, 140);
            cleanDataGridButton.Size = new System.Drawing.Size(100, 25);
            this.Controls.Add(cleanDataGridButton);

            logoImage = new PictureBox();
            logoImage.Image = Image.FromFile("img\\logo.png");
            logoImage.Size = new System.Drawing.Size(100, 40);
            logoImage.Location = new System.Drawing.Point(470, 15);
            logoImage.SizeMode = PictureBoxSizeMode.StretchImage;
            this.Controls.Add(logoImage);

            logoImage2 = new PictureBox();
            logoImage2.Image = Image.FromFile("img\\logo22.png");
            logoImage2.Size = new System.Drawing.Size(50, 20);
            logoImage2.Location = new System.Drawing.Point(495, 50);
            logoImage2.SizeMode = PictureBoxSizeMode.StretchImage;
            this.Controls.Add(logoImage2);


            // Grid View
            getDataButton = new Button();
            getDataButton.Text = "Consultar Registros";
            getDataButton.Click += QueryAndGenerateCsv;
            getDataButton.Location = new System.Drawing.Point(10, 140);
            getDataButton.Size = new System.Drawing.Size(150, 25);
            this.Controls.Add(getDataButton);

            dataGridView = new DataGridView();
            dataGridView.Location = new System.Drawing.Point(10, 170);
            dataGridView.Size = new System.Drawing.Size(560, 400);
            dataGridView.ReadOnly = true;
            this.Controls.Add(dataGridView);

            exportDataGridView = new Button();
            exportDataGridView.Text = "Crear CSV";
            exportDataGridView.Location = new System.Drawing.Point(165, 140);
            exportDataGridView.Size = new System.Drawing.Size(150, 25);
            exportDataGridView.Click += ExportarCSV;
            this.Controls.Add(exportDataGridView);

            // ServerConnection

            // SERVER TEXT BOX & LABEL
            serverText = new Label();
            serverText.Text = "Server: ";
            serverText.Location = new System.Drawing.Point(50, 13);
            serverText.Size = new System.Drawing.Size(45, 20);
            this.Controls.Add(serverText);

            serverTextBox = new TextBox();
            serverTextBox.Text = "172.18.3.143";
            serverTextBox.Location = new System.Drawing.Point(100, 10);
            serverTextBox.Enabled = false;


            // DATABASE TEXT BOX & LABEL
            databaseText = new Label();
            databaseText.Text = "DataBase: ";
            databaseText.Location = new System.Drawing.Point(40, 32);
            databaseText.Size = new System.Drawing.Size(60, 20);
            this.Controls.Add(databaseText);

            databaseTextBox = new TextBox();
            databaseTextBox.Text = "IVMS2";
            databaseTextBox.Location = new System.Drawing.Point(100, 32);
            databaseTextBox.Enabled = false;


            // USERNAME TEXT BOX & LABEL
            usernameText = new Label();
            usernameText.Text = "Username: ";
            usernameText.Location = new System.Drawing.Point(40, 55);
            usernameText.Size = new System.Drawing.Size(60, 20);
            this.Controls.Add(usernameText);

            usernameTextBox = new TextBox();
            usernameTextBox.Text = "yura";
            usernameTextBox.Location = new System.Drawing.Point(100, 55);
            usernameTextBox.Enabled = false;


            // PASSWORD TEXT BOX & LABEL
            passwordText = new Label();
            passwordText.Text = "Password: ";
            passwordText.Location = new System.Drawing.Point(233, 13);
            passwordText.Size = new System.Drawing.Size(56, 20);
            this.Controls.Add(passwordText);

            passwordTextBox = new TextBox();
            passwordTextBox.Text = "mxyura871.";
            passwordTextBox.Location = new System.Drawing.Point(290, 10);
            passwordTextBox.PasswordChar = '*';
            passwordTextBox.Enabled = false;


            // TABLE TEXT BOX & LABEL
            tableText = new Label();
            tableText.Text = "Table: ";
            tableText.Location = new System.Drawing.Point(252, 33);
            tableText.Size = new System.Drawing.Size(37, 20);
            this.Controls.Add(tableText);

            tableTextBox = new TextBox();
            tableTextBox.Text = "lerdo";
            tableTextBox.Location = new System.Drawing.Point(290, 32);
            tableTextBox.Enabled = false;


            this.Controls.Add(serverTextBox);
            this.Controls.Add(databaseTextBox);
            this.Controls.Add(usernameTextBox);
            this.Controls.Add(passwordTextBox);
            this.Controls.Add(tableTextBox);


            // Generacion de Componenete
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Size = new System.Drawing.Size(595, 620);
            this.Icon = new System.Drawing.Icon("img\\1.ico");
        }

        public void cleanDataGridView(object sender, EventArgs e)
        {
            dataGridView.Rows.Clear();
            dataGridView.Columns.Clear();
            dataGridView.Refresh();
        }

        private void ToggleIntervalFields(object sender, EventArgs e)
        {
            startDateTimeTextBox.Enabled = intervalCheckBox.Checked;
            endDateTimeTextBox.Enabled = intervalCheckBox.Checked;
        }

        private void QueryAndGenerateCsv(object sender, EventArgs e)
        {

            cleanDataGridView(sender, e);

            string server = serverTextBox.Text;
            string database = databaseTextBox.Text;
            string username = usernameTextBox.Text;
            string passsword = passwordTextBox.Text;
            string table = tableTextBox.Text;
            string sqlQuery;

            string connectionString = $"Data Source={server};Initial Catalog={database}; User ID={username};Password={passsword}";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {

                connection.Open();

                if (intervalCheckBox.Checked)
                {
                    DateTime startDate = DateTime.Parse(startDateTimeTextBox.Text);
                    DateTime endDate = DateTime.Parse(endDateTimeTextBox.Text);

                    sqlQuery = $@"
                        SELECT ID_Empleado, ID_Lector, CONVERT(VARCHAR, FechaRegistro, 120) as FechaRegistro, 
                               CONVERT(VARCHAR, DATEADD(SECOND, 1, FechaRegistro), 120) as hora_sumada, 
                               ClaveNomina = 2
                        FROM {table}
                        WHERE FechaRegistro BETWEEN '{startDate:yyyy-MM-dd HH:mm:ss.fff}' AND '{endDate:yyyy-MM-dd HH:mm:ss.fff}'
                        ORDER BY FechaRegistro ASC";
                }
                else
                {
                    sqlQuery = $@"
                         SELECT ID_Empleado, ID_Lector, FORMAT(FechaRegistro, 'yyyy-MM-dd HH:mm:ss.fff') as FechaRegistro, FORMAT(DATEADD(SECOND, 1, FechaRegistro), 'yyyy-MM-dd HH:mm:ss.fff') as hora_sumada, ClaveNomina = 2 
                FROM {table} 
                WHERE FechaRegistro >= DATEADD(HOUR, -24, GETDATE()) ORDER BY FechaRegistro ASC";
                }

                using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                {
                    dataGridView.Columns.Add("ID_Empleado", "ID Empleado");
                    dataGridView.Columns.Add("ID_Lector", "ID Lector");
                    dataGridView.Columns.Add("FechaRegistro", "Fecha de Registro");
                    dataGridView.Columns.Add("hora_sumada", "Hora Sumada");
                    dataGridView.Columns.Add("ClaveNomina", "Clave de Nómina");
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            dataGridView.Rows.Add(reader["ID_Empleado"], reader["ID_Lector"], reader["FechaRegistro"], reader["hora_sumada"], reader["ClaveNomina"]);
                        }
                    }
                }

                connection.Close();
            }

        }

        private void ExportarCSV(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Archivos CSV (*.csv)|*.csv";
            saveFileDialog.Title = "Guardar CSV";


            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filepath = saveFileDialog.FileName;

                using (StreamWriter sw = new StreamWriter(filepath))
                {
                    foreach (DataGridViewRow row in dataGridView.Rows)
                    {

                        bool rowEnBlanco = row.Cells.OfType<DataGridViewCell>().Any(cell => cell.Value != null && !string.IsNullOrWhiteSpace(cell.Value.ToString()));

                        if (rowEnBlanco)
                        {
                            bool quitarComa = true;
                            foreach (DataGridViewCell cell in row.Cells)
                            {
                                if (!quitarComa)
                                {
                                    sw.Write(",");

                                }
                                sw.Write(cell.Value);
                                quitarComa = false;
                            }
                            sw.WriteLine();
                        }


                    }
                }
                MessageBox.Show("Datos Exportados a: " + filepath, "Exportacion Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        #endregion
    }
}

