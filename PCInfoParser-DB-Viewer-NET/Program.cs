using System.Windows.Forms;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System.IO;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PCInfoParser_DB_Viewer_NET
{
    public class ExcelExporter
    {
        public void ExportToExcel(string fileName, string sheet1Name, string sheet2Name, string[,] data1, string[,] data2)
        {
            string[] headers1 = new string[] { "Кабинет", "LAN", "ФИО", "Монитор", "Диагональ", "Тип принтера", "Модель принтера", "ПК", "Материнская плата", "Процессор", "Частота процессора", "Баллы Passmark", "Дата выпуска", "Тип ОЗУ", "ОЗУ, 1 Планка", "ОЗУ, 2 Планка", "ОЗУ, 3 Планка", "ОЗУ, 4 Планка", "Сокет", "Диск 1", "Состояние диска 1", "Диск 2", "Состояние диска 2", "Диск 3", "Состояние диска 3", "Диск 4", "Состояние диска 4", "Операционная система", "Антивирус", "CPU Под замену", "Все CPU под сокет", "Дата создания" };
            string[] headers2 = new string[] { "Кабинет", "LAN", "ФИО", "Диск", "Наименование", "Прошивка", "Размер", "Время работы", "Включён", "Состояние", "Температура", "Дата создания" };
            string filePath;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            SaveFileDialog saveFileDialog1 = new()
            {
                Title = "Сохранить файл как...",
                Filter = "Таблица Excel (*.xlsx)|*.xlsx",
                InitialDirectory = Path.GetDirectoryName(Application.ExecutablePath),
                FileName = fileName
            };

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

            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            using var package = new ExcelPackage(filePath);
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

    public class IniFile
    {
        private readonly Dictionary<string, Dictionary<string, string>> data = new Dictionary<string, Dictionary<string, string>>();
        private readonly string fileName;

        public IniFile(string fileName)
        {
            this.fileName = fileName;

            if (File.Exists(fileName))
            {
                Load();
            }
            else
            {
                SetValue("MySQL", "IP", "127.0.0.1");
                SetValue("MySQL", "Port", "3306");
                SetValue("MySQL", "Database", "");
                SetValue("MySQL", "User", "");
                SetValue("MySQL", "Password", "");
                SetValue("Server", "Port", "");
                SetValue("Server", "Password", "");
                SetValue("App", "Minimaze", "false");
                SetValue("App", "ConnectMySQL", "false");
                SetValue("App", "ServerStart", "false");
                Save();
            }
        }

        public string GetValue(string section, string key)
        {
            if (data.TryGetValue(section, out Dictionary<string, string> sectionData))
            {
                if (sectionData.TryGetValue(key, out string value))
                {
                    return value;
                }
            }

            return null;
        }

        public void SetValue(string section, string key, string value)
        {
            if (!data.TryGetValue(section, out Dictionary<string, string> sectionData))
            {
                sectionData = new Dictionary<string, string>();
                data[section] = sectionData;
            }

            sectionData[key] = value;
        }

        public void Load()
        {
            data.Clear();

            string currentSection = null;

            foreach (string line in File.ReadAllLines(fileName))
            {
                string trimmedLine = line.Trim();
                if (trimmedLine.StartsWith("[") && trimmedLine.EndsWith("]"))
                {
                    currentSection = trimmedLine.Substring(1, trimmedLine.Length - 2);
                    if (!data.ContainsKey(currentSection))
                    {
                        data[currentSection] = new Dictionary<string, string>();
                    }
                }
                else if (!string.IsNullOrEmpty(trimmedLine))
                {
                    string[] parts = trimmedLine.Split(new char[] { '=' }, 2);
                    if (parts.Length > 1)
                    {
                        string currentKey = parts[0].Trim();
                        string currentValue = parts[1].Trim();
                        if (data.TryGetValue(currentSection, out Dictionary<string, string> sectionData))
                        {
                            sectionData[currentKey] = currentValue;
                        }
                    }
                }
            }
        }

        public void Save()
        {
            List<string> lines = new();
            foreach (KeyValuePair<string, Dictionary<string, string>> section in data)
            {
                lines.Add("[" + section.Key + "]");
                foreach (KeyValuePair<string, string> keyValuePair in section.Value)
                {
                    lines.Add(keyValuePair.Key + "=" + keyValuePair.Value);
                }
                lines.Add("");
            }

            File.WriteAllLines(fileName, lines.ToArray());
        }
    }

    public class MySQLConnect
    {
        private MySqlConnection connection;
        private readonly IniFile ini;
        private string connectionString;
        bool connection_status;

        public MySQLConnect(IniFile ini)
        {
            this.ini = ini;
        }

        public bool Connect()
        {
            connectionString = $"Server={ini.GetValue("MySQL", "IP")};Port={ini.GetValue("MySQL", "Port")};Database={ini.GetValue("MySQL", "Database")};Uid={ini.GetValue("MySQL", "User")};Pwd={ini.GetValue("MySQL", "Password")};";
            connection_status = false;
            try
            {
                connection = new(connectionString);
                connection.Open();
                connection_status = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return connection_status;
        }
        public string LastID(string table)
        {
            string query = $"SELECT ID FROM {table}_General"; // Замените на ваш запрос SELECT

            try
            {
                MySqlCommand command = new MySqlCommand(query, connection);
                MySqlDataReader reader = command.ExecuteReader();

                List<int> columnValues = new();


                while (reader.Read())
                {
                    string columnValue = reader.GetString(0); // Получение значения столбца по индексу (0)
                    columnValues.Add(Convert.ToInt32(columnValue));
                }

                columnValues = columnValues.Distinct().ToList();
                columnValues.Sort();

                int lastvalue = 0;

                foreach (int value in columnValues)
                {
                    lastvalue++;
                    if (value != lastvalue - 1)
                    {
                        lastvalue--;
                        break;
                    }
                }

                reader.Close();
                return lastvalue.ToString();
            }
            catch (Exception)
            {
                return "0";
            }
        }
        public void Disconnect()
        {
            if (!connection_status) connection.Close();
        }
        public bool ExecuteCommand(string commandText)
        {
            bool success = false;
            try
            {
                MySqlCommand command = new(commandText, connection);
                int rowsAffected = command.ExecuteNonQuery();

                // Если хотя бы одна строка была затронута, считаем операцию успешной
                success = rowsAffected > 0;
            }
            catch (MySqlException ex)
            {
                // Обработка ошибок подключения к базе данных
                if (ex.ErrorCode != -2147467259) Console.WriteLine($"An error occurred: {ex.Message}");
            }
            return success;
        }
    private string DateReverse(string date)
        {
            string[] dateSplit = date.Split('.');
            string dateReversed = dateSplit[2] + "-" + dateSplit[1] + "-" + dateSplit[0];
            return dateReversed;
        }
        public string[,] ParseTables(string database, string table, string ymd)
        {
            if (connection_status == true)
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
            List<string> result = new() { "" };
            if (connection_status == true)
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
            List<string> result = new() { "" };

            if (connection_status == true)
            {
                MySqlCommand cmd = new("SHOW DATABASES", connection);
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
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            IniFile ini = new("PCInfoParser-Server.ini");
            MySQLConnect dbview = new(ini);
            if (dbview.Connect())
            {
                List<string> databases = dbview.GetDatabases();
                Application.Run(new Form1(databases, dbview));
            }
            else
            {
                Error("Не удалось подключиться к MySQL! Проверьте настройки.");
            }
        }
        static void Error(string errorstr)
        {
            Application.Run(new Form2(errorstr));
            Environment.Exit(1);
        }
    }
}