using System.Windows.Forms;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System.IO;
using System;
using PCInfoParser_DB_Viewer_NET;
using System.Collections.Generic;
using System.Linq;

namespace PCInfoParser_DB_Viewer
{
    public class ExcelExporter
    {
        public void ExportToExcel(string fileName, string sheet1Name, string sheet2Name, string[,] data1, string[,] data2)
        {
            string[] headers1 = new string[] { "Кабинет", "LAN", "ФИО", "Монитор", "Диагональ", "Тип принтера", "Модель принтера", "ПК", "Материнская плата", "Процессор", "Частота процессора", "Баллы Passmark", "Дата выпуска", "Тип ОЗУ", "ОЗУ, 1 Планка", "ОЗУ, 2 Планка", "ОЗУ, 3 Планка", "ОЗУ, 4 Планка", "Сокет", "Диск 1", "Состояние диска 1", "Диск 2", "Состояние диска 2", "Диск 3", "Состояние диска 3", "Диск 4", "Состояние диска 4", "Операционная система", "Антивирус", "CPU Под замену", "Все CPU под сокет", "Дата создания" };
            string[] headers2 = new string[] { "Кабинет", "LAN", "ФИО", "Диск", "Наименование", "Прошивка", "Размер", "Время работы", "Включён", "Состояние", "Температура", "Дата создания" };
            string filePath;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //var folderBrowserDialog = new FolderBrowserDialog();
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Title = "Сохранить файл как...";
            saveFileDialog1.Filter = "Таблица Excel (*.xlsx)|*.xlsx";
            saveFileDialog1.InitialDirectory = Path.GetDirectoryName(Application.ExecutablePath);
            saveFileDialog1.FileName = fileName;
            //DialogResult result = folderBrowserDialog.ShowDialog();

            DialogResult result = saveFileDialog1.ShowDialog();

            if (result == DialogResult.OK)
            {
                // пользователь выбрал файл для сохранения
                filePath = saveFileDialog1.FileName;
                // выполните дальнейшую обработку сохранения файла
            }
            else
            {
                return;// пользователь отменил диалоговое окно
            }

            // Собираем полный путь к файлу
            //var filePath = Path.Combine(folderBrowserDialog.SelectedPath, fileName);

            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            using (var package = new ExcelPackage(filePath))
            {
                // создаем первую таблицу
                var sheet1 = package.Workbook.Worksheets.Add(sheet1Name);
                AddHeaders(sheet1, headers1);
                AddData(sheet1, data1);

                // создаем вторую таблицу
                var sheet2 = package.Workbook.Worksheets.Add(sheet2Name);
                AddHeaders(sheet2, headers2);
                AddData(sheet2, data2);

                // сохраняем файл
                package.Save();
            }
        }

        private void AddHeaders(ExcelWorksheet sheet, string[] headers)
        {
            for (int i = 0; i < headers.Length; i++)
            {
                sheet.Cells[1, i + 1].Value = headers[i];
                sheet.Cells[1, i + 1].Style.Font.Bold = true;
            }
        }

        private void AddData(ExcelWorksheet sheet, string[,] data)
        {
            int rows = data.GetLength(0);
            int columns = data.GetLength(1);

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    sheet.Cells[i + 2, j + 1].Value = data[i, j];
                }
            }
        }
    }

    public class Config
    {
        private Dictionary<string, string> values = new Dictionary<string, string>();

        public Config(string filename)
        {
            if (!File.Exists(filename))
            {
                SetValue("mysql_host", "localhost");
                SetValue("mysql_port", "3306");
                SetValue("mysql_user", "root");
                SetValue("mysql_password", "Не задано");
                Save(filename);
            }
            Load(filename);
        }

        public void Load(string filename)
        {
            using (var reader = new StreamReader(filename))
            {
                bool error = false;
                string errorstr = "Укажите ";
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine().Trim();
                    if (line.StartsWith("#") || string.IsNullOrWhiteSpace(line))
                        continue;
                    var parts = line.Split(new char[] { '=' }, 2);
                    if (parts.Length == 2)
                    {
                        var key = parts[0].Trim();
                        var value = parts[1].Trim();
                        if (value == "Не задано")
                        {
                            errorstr = errorstr + key + ", ";
                            error = true;
                        }
                        values[key] = value;
                    }
                }
                if (error == true)
                {
                    errorstr = errorstr[..^2] + " в dbviewer.cfg.";
                    Application.Run(new Form2(errorstr));
                    Environment.Exit(1);
                }
            }
        }

        public void Save(string filename)
        {
            using (var writer = new StreamWriter(filename))
            {
                foreach (var kvp in values)
                {
                    writer.WriteLine("{0} = {1}", kvp.Key, kvp.Value);
                }
            }
        }

        public string GetValue(string key, string defaultValue = "")
        {
            if (values.ContainsKey(key))
                return values[key];
            return defaultValue;
        }

        public void SetValue(string key, string value)
        {
            values[key] = value;
        }
    }

    public class MySQLConnect
    {
        private MySqlConnection connection;
        private string server;
        private string port;
        private string uid;
        private string password;

        public MySQLConnect(string server, string port, string uid, string password)
        {
            this.server = server;
            this.port = port;
            this.uid = uid;
            this.password = password;

            Initialize();
        }

        private void Initialize()
        {
            string connectionString = $"server={server};port={port};uid={uid};password={password};";
            connection = new MySqlConnection(connectionString);
        }

        public bool OpenConnection()
        {
            try
            {
                connection.Open();
                return true;
            }
            catch (MySqlException ex)
            {
                // Handle exception
                return false;
            }
            catch (InvalidOperationException ex)
            {
                return true;
            }
        }

        public bool CloseConnection()
        {
            try
            {
                connection.Close();
                return true;
            }
            catch (MySqlException ex)
            {
                // Handle exception
                return false;
            }
        }

        public void ExecuteQuery(string query)
        {
            if (OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.ExecuteNonQuery();
                CloseConnection();
            }
        }
        private string DateReverse(string date)
        {
            string[] dateSplit = date.Split('.');
            string dateReversed = dateSplit[2] + "-" + dateSplit[1] + "-" + dateSplit[0];
            return dateReversed;
        }
        public string[,] ParseTables(string database, string table, string ymd)
        {
            if (OpenConnection() == true)
            {
                ymd = DateReverse(ymd);
                MySqlCommand cmd = new MySqlCommand($"SELECT * FROM `{database}`.`{table}` WHERE `Дата создания` BETWEEN '{ymd} 00:00:00' AND '{ymd} 23:59:59'", connection);
                MySqlDataReader dataReader = cmd.ExecuteReader();

                int numCols = dataReader.FieldCount;
                List<string[]> rows = new List<string[]>();

                while (dataReader.Read())
                {
                    string[] row = new string[numCols];
                    for (int i = 0; i < numCols; i++)
                    {
                        row[i] = dataReader[i].ToString();
                    }
                    rows.Add(row);
                }

                dataReader.Close();

                string[,] result = new string[rows.Count, numCols];

                for (int i = 0; i < rows.Count; i++)
                {
                    for (int j = 0; j < numCols; j++)
                    {
                        result[i, j] = rows[i][j];
                    }
                }
                return result;
            }
            else
            {
                return null;
            }
        }
        public List<string> ParseTime(string table, string database)
        {
            List<string> result = new List<string>() { "" };
            if (OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand($"SELECT `Дата создания` FROM `{database}`.`{table}`", connection);
                MySqlDataReader dataReader = cmd.ExecuteReader();
                while (dataReader.Read())
                {
                    string time = dataReader["Дата создания"].ToString();
                    time = time.Split()[0];
                    result.Add(time);
                }
                dataReader.Close();
                List<string> retresult = result.Distinct().ToList();
                return retresult;
            }
            else
            {
                return null;
            }
        }

        public List<string> GetDatabases()
        {
            List<string> result = new List<string>() { "" };

            if (OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand("SHOW DATABASES", connection);
                MySqlDataReader dataReader = cmd.ExecuteReader();

                while (dataReader.Read())
                {
                    string databaseName = dataReader["Database"].ToString();

                    if (!databaseName.StartsWith("mysql") && !databaseName.StartsWith("information_schema") && !databaseName.StartsWith("performance_schema") && !databaseName.StartsWith("sys"))
                    {
                        result.Add(databaseName);
                    }
                }

                dataReader.Close();
            }

            return result;
        }
    }


    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static int Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            var config = new Config("dbviewer.cfg");
            MySQLConnect dbview = new MySQLConnect(config.GetValue("mysql_host"), config.GetValue("mysql_port"), config.GetValue("mysql_user"), config.GetValue("mysql_password"));
            List<string> databases = dbview.GetDatabases();
            Application.Run(new Form1(databases, dbview));
            return 0;
        }
    }
}