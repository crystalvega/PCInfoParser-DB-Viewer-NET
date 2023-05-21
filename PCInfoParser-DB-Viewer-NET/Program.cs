using System.Windows.Forms;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System.IO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;

namespace PCInfoParser_DB_Viewer_NET
{
    public class ExcelExporter
    {
        public void ExportToExcel(string organization, string sheet1Name, string sheet2Name, MySQLConnect connect)
        {
            string[] headers1 = new string[] { "ID", "Кабинет", "LAN", "ФИО", "Монитор", "Диагональ", "Тип принтера", "Модель принтера", "ПК", "Материнская плата", "Процессор", "Частота процессора", "Баллы Passmark", "Дата выпуска", "Тип ОЗУ", "ОЗУ, 1 Планка", "ОЗУ, 2 Планка", "ОЗУ, 3 Планка", "ОЗУ, 4 Планка", "Сокет", "Диск 1", "Состояние диска 1", "Диск 2", "Состояние диска 2", "Диск 3", "Состояние диска 3", "Диск 4", "Состояние диска 4", "Операционная система", "Антивирус", "CPU Под замену", "Все CPU под сокет", "Дата создания" };
            string[] headers2 = new string[] { "ID", "Кабинет", "LAN", "ФИО", "Диск", "Наименование", "Прошивка", "Размер", "Время работы", "Включён", "Состояние", "Температура", "Дата создания" };
            string filePath;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            SaveFileDialog saveFileDialog1 = new()
            {
                Title = "Сохранить файл как...",
                Filter = "Таблица Excel (*.xlsx)|*.xlsx",
                InitialDirectory = Path.GetDirectoryName(Application.ExecutablePath),
                FileName = organization + ".xlsx"
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

            string[,] users = connect.ParseTables("Users", organization);
			string[,] data1raw = connect.ParseTablesExport("General", organization);
			string[,] data2raw = connect.ParseTablesExport("Disk", organization);

            for (int i=0; i < data1raw.GetLength(0); i++)
            {
                for (int i2 = 0; i2 < users.GetLength(0); i2++)
                {
                    if (users[i2,0] == data1raw[i, 0])
                    {
                        for(int i3 = 0; i3<4; i3++)
                        {
                            data1raw[i, i3] = users[i2, i3];
                        }
                    }
                }
            }

            for (int i = 0; i < data2raw.GetLength(0); i++)
            {
                for (int i2 = 0; i2 < users.GetLength(0); i2++)
                {
                    if (users[i2, 0] == data2raw[i, 0])
                    {
                        for (int i3 = 0; i3 < 4; i3++)
                        {
                            data2raw[i, i3] = users[i2, i3];
                        }
                    }
                }
            }
            // создаем первую таблицу
            var sheet1 = package.Workbook.Worksheets.Add(sheet1Name);
            AddHeaders(sheet1, headers1);

            AddData(sheet1, data1raw);

            // создаем вторую таблицу
            var sheet2 = package.Workbook.Worksheets.Add(sheet2Name);
            AddHeaders(sheet2, headers2);
            
            AddData(sheet2, data2raw);

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
        public string[,] ParseTables(string table, string organization, string ymd)
        {
            if (connection_status == true)
            {
                ymd = DateReverse(ymd);
                MySqlCommand cmd = new($"SELECT * FROM `{organization}_{table}` WHERE `Дата создания` BETWEEN '{ymd} 00:00:00' AND '{ymd} 23:59:59'", connection);
                MySqlDataReader dataReader = cmd.ExecuteReader();

                int numCols = dataReader.FieldCount;
                List<string[]> rows = new();

                while (dataReader.Read())
                {
                    string[] row = new string[numCols];
                    for (int i = 1; i < numCols; i++)
                    {
                         row[i-1] = dataReader[i].ToString();
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
        public string[,] ParseTables(string table, string organization)
        {
            if (connection_status == true)
            {
                MySqlCommand cmd = new($"SELECT * FROM `{organization}_{table}`", connection);
                MySqlDataReader dataReader = cmd.ExecuteReader();

                int numCols = dataReader.FieldCount;
                List<string[]> rows = new();

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
		public string[,] ParseTablesExport(string table, string organization)
		{
			if (connection_status == true)
			{
				MySqlCommand cmd = new($"SELECT * FROM `{organization}_{table}`", connection);
				MySqlDataReader dataReader = cmd.ExecuteReader();

				int numCols = dataReader.FieldCount;
				List<string[]> rows = new();

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

				string[,] result = new string[rows.Count, numCols + 3];

				for (int i = 0; i < rows.Count; i++)
				{
					result[i, 0] = rows[i][0];
					for (int j = 1; j < numCols; j++)
					{
						result[i, j + 3] = rows[i][j];
					}
				}
				return result;
			}
			else
			{
				return null;
			}
		}
		public List<string> ParseTime(string table, string organization, string id)
        {
            List<string> result = new();
            if (connection_status == true)
            {
                MySqlCommand cmd = new MySqlCommand($"SELECT `Дата создания` FROM `{organization}_{table}` WHERE `ID` like {id}", connection);
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

        public List<string> GetTables()
        {
            List<string> result = new() { "" };

            if (connection_status == true)
            {
                MySqlCommand cmd = new("SHOW TABLES", connection);
                MySqlDataReader dataReader = cmd.ExecuteReader();
                string database = ini.GetValue("MySQL", "Database");

                while (dataReader.Read())
                {
                    string tableName = dataReader[$"Tables_in_{database}"].ToString();

                    if (tableName.EndsWith("_users"))
                    {
                        result.Add(tableName.Replace("_users", ""));
                    }
                }

                dataReader.Close();
            }

            return result;
        }
    }


    internal static class Program
    {
		static bool IsProcessOpen(string processName)
		{
			Process[] processes = Process.GetProcessesByName(processName);
			return processes.Length > 1;
		}
		/// <summary>
		///  The main entry point for the application.
		/// </summary>
		[STAThread]
        static int Main()
        {
			if (IsProcessOpen("PCInfoParser-DB-Viewer-NET"))
			{
				MessageBox.Show("Приложение уже открыто!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return 1;
			}
			Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            IniFile ini = new("PCInfoParser-Server.ini");
            MySQLConnect dbview = new(ini);
            if (dbview.Connect()) Application.Run(new Form1(dbview));
            else
            {
                MessageBox.Show("Не удалось подключиться к MySQL! Проверьте настройки.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 2;
            }
            return 0;
        }
    }
}