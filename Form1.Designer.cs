using System;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace test3
{
    partial class Form1
    {
        private Label FechaInicio;
        private Label FechaFinal;
        //private TextBox startDateTimeTextBox;
        //private TextBox endDateTimeTextBox;

        private DateTimePicker startDateTimePicker;
        private DateTimePicker endDateTimePicker;


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
            intervalCheckBox.Text = "Activate Date Range";
            intervalCheckBox.CheckedChanged += ToggleIntervalFields;
            intervalCheckBox.Location = new System.Drawing.Point(10, 85);
            intervalCheckBox.Size = new System.Drawing.Size(155, 20);
            this.Controls.Add(intervalCheckBox);


            FechaInicio = new Label();
            FechaInicio.Text = "Start Date:";
            FechaInicio.Location = new System.Drawing.Point(40, 115);
            FechaInicio.Size = new System.Drawing.Size(70, 15);
            this.Controls.Add(FechaInicio);

            FechaFinal = new Label();
            FechaFinal.Text = "End Date:";
            FechaFinal.Location = new System.Drawing.Point(285, 115);
            FechaFinal.Size = new System.Drawing.Size(65, 15);
            this.Controls.Add(FechaFinal);
            


            /*
            startDateTimeTextBox = new TextBox();
            startDateTimeTextBox.Enabled = false;
            startDateTimeTextBox.Location = new System.Drawing.Point(80, 112);
            this.Controls.Add(startDateTimeTextBox);


            endDateTimeTextBox = new TextBox();
            endDateTimeTextBox.Enabled = false;
            endDateTimeTextBox.Location = new System.Drawing.Point(290, 112);
            this.Controls.Add(endDateTimeTextBox);
            */

            startDateTimePicker = new DateTimePicker();
            startDateTimePicker.Format = DateTimePickerFormat.Custom;
            startDateTimePicker.CustomFormat = "yyyy-MM-dd HH:mm:ss";
            startDateTimePicker.Enabled = false;
            startDateTimePicker.Location = new System.Drawing.Point(110, 112);
            startDateTimePicker.Size = new System.Drawing.Size(150, 40);
            this.Controls.Add(startDateTimePicker);

            endDateTimePicker = new DateTimePicker();
            endDateTimePicker.Format = DateTimePickerFormat.Custom;
            endDateTimePicker.CustomFormat = "yyyy-MM-dd HH:mm:ss";
            endDateTimePicker.Enabled = false;
            endDateTimePicker.Location = new System.Drawing.Point(350, 112);
            endDateTimePicker.Size = new System.Drawing.Size(150, 40);
            this.Controls.Add(endDateTimePicker);
            


            cleanDataGridButton = new Button();
            cleanDataGridButton.Text = "Clean";
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


            // ----------- Grid View -----------------------------------
            getDataButton = new Button();
            getDataButton.Text = "Get Data";
            getDataButton.Click += QueryAndGenerateCsv;
            getDataButton.Location = new System.Drawing.Point(10, 140);
            getDataButton.Size = new System.Drawing.Size(150, 25);
            this.Controls.Add(getDataButton);

            dataGridView = new DataGridView();
            dataGridView.Location = new System.Drawing.Point(10, 170);
            dataGridView.Size = new System.Drawing.Size(560, 400);

            ContextMenuStrip contextMenu = new ContextMenuStrip();
            ToolStripMenuItem deleteMenuItem = new ToolStripMenuItem("Eliminar fila");
            deleteMenuItem.Click += DeleteRowMenuItem_Click;
            contextMenu.Items.Add(deleteMenuItem);
            dataGridView.ContextMenuStrip = contextMenu;

            //              DESING DATAGRIDVIEW
            dataGridView.RowsDefaultCellStyle.BackColor = Color.White;
            dataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;
           

            dataGridView.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkGray;
            dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView.ColumnHeadersDefaultCellStyle.Font = new Font(dataGridView.Font, FontStyle.Bold);

            dataGridView.DefaultCellStyle.SelectionBackColor = Color.SteelBlue;
            dataGridView.DefaultCellStyle.SelectionForeColor = Color.White;

            dataGridView.BorderStyle = BorderStyle.None;
            dataGridView.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            dataGridView.RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            dataGridView.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;

            dataGridView.ScrollBars = ScrollBars.Both;

            dataGridView.ReadOnly = true;
            this.Controls.Add(dataGridView);

            // ----------- Export Button --------------------------------

            exportDataGridView = new Button();
            exportDataGridView.Text = "Create CSV";
            exportDataGridView.Location = new System.Drawing.Point(165, 140);
            exportDataGridView.Size = new System.Drawing.Size(150, 25);
            exportDataGridView.Click += ExportarCSV;
            this.Controls.Add(exportDataGridView);

            // ---------- ServerConnection ----------------------------

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

        private void DeleteRowMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView.SelectedRows.Count > 0)
            {
                DialogResult result = MessageBox.Show("Are You Sure Delete This Item?", "Confim Your Answer", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                
                if(result == DialogResult.Yes)
                {
                    foreach (DataGridViewRow row in dataGridView.SelectedRows)
                    {
                        if (!row.IsNewRow)
                        {
                            dataGridView.Rows.Remove(row);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Select 1 Row: ", "Atention!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void ToggleIntervalFields(object sender, EventArgs e)
        {
            startDateTimePicker.Enabled = intervalCheckBox.Checked;
            endDateTimePicker.Enabled = intervalCheckBox.Checked;
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

                try
                {
                    connection.Open();

                    if (intervalCheckBox.Checked)
                    {
                        DateTime startDate = DateTime.Parse(startDateTimePicker.Text);
                        DateTime endDate = DateTime.Parse(endDateTimePicker.Text);

                        sqlQuery = $@"
                              SELECT 
                                ID_Empleado, 
                                ID_Lector, 
                                FORMAT(FechaRegistro, 'yyyy-MM-dd HH:mm:ss.fff') as FechaRegistro, 
                                FORMAT(DATEADD(SECOND, 1, FechaRegistro), 'yyyy-MM-dd HH:mm:ss.fff') as hora_sumada, 
                                ClaveNomina = 2
                              FROM {table}
                                WHERE FechaRegistro BETWEEN '{startDate.ToString("yyyy-MM-dd HH:mm:ss")}' AND '{endDate.ToString("yyyy-MM-dd HH:mm:ss")}'
                                ORDER BY [FechaRegistro] ASC
                              ";


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
                        dataGridView.Columns["ID_Empleado"].Width = 85;
                        dataGridView.Columns["ID_Empleado"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dataGridView.Columns.Add("ID_Lector", "ID Lector");
                        dataGridView.Columns["ID_Lector"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dataGridView.Columns["ID_Lector"].Width = 67;
                        dataGridView.Columns.Add("FechaRegistro", "Fecha de Registro");
                        dataGridView.Columns["FechaRegistro"].Width = 130;
                        dataGridView.Columns.Add("hora_sumada", "Hora Sumada");
                        dataGridView.Columns["hora_sumada"].Width = 130;
                        dataGridView.Columns.Add("ClaveNomina", "Clave de Nómina");
                        dataGridView.Columns["ClaveNomina"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dataGridView.Columns["ClaveNomina"].Width = 110;
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
                catch(Exception ex)
                {
                    MessageBox.Show($"Error en el servidor: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
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

