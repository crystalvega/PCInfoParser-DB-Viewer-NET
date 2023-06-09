using MySqlX.XDevAPI.Relational;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace PCInfoParser_DB_Viewer_NET
{
    public partial class Form1 : Form
    {
        List<string> generalList = new List<string>() {"Монитор", "Диагональ", "Тип принтера", "Модель принтера", "ПК", "Материнская плата", "Процессор", "Частота процессора", "Баллы Passmark", "Дата выпуска", "Температура процессора", "Тип ОЗУ", "ОЗУ, 1 Планка", "ОЗУ, 2 Планка", "ОЗУ, 3 Планка", "ОЗУ, 4 Планка", "Сокет", "Диск 1", "Состояние диска 1", "Диск 2", "Состояние диска 2", "Диск 3", "Состояние диска 3", "Диск 4", "Состояние диска 4", "Операционная система", "Антивирус", "CPU Под замену", "Все CPU под сокет", "Дата создания" };
        List<string> diskList = new() {"Диск", "Наименование", "Прошивка", "Размер", "Время работы", "Включён", "Состояние", "Температура", "Дата создания" };
        List<string> userList = new() { "ID", "Кабинет", "LAN", "ФИО" };
        string organization;
        string[,] general;
        string[,] disk;
        string[,] users;
        string datetime;
        readonly MySQLConnect dbview;

        public Form1(MySQLConnect dbview)
        {
            this.dbview = dbview;
            InitializeComponent();
        }

        private void Combobox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "")
            {
                comboBox2.Enabled = false;
                comboBox2.SelectedIndex = -1;
                comboBox3.Enabled = false;
                comboBox3.SelectedIndex = -1;
                button3.Enabled = false;
                listView1.Columns.Clear();
                listView1.Enabled = false;
                listView2.Columns.Clear();
                listView2.Enabled = false;

            }
            else
            {
                button3.Enabled = true;
                comboBox2.Enabled = true;
                listView2.Enabled = true;
                organization = comboBox1.Text;
                users = dbview.ParseTables("Users", organization);
                UsersInput();
                //comboBox2.DataSource = dbview.ParseTime("General", table);

            }
        }
        private void Combobox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text == "" || comboBox2.Text == "Выбрать дату...")
            {
            }
            else
            {
                listView1.Enabled = true;
                button3.Enabled = true;
                datetime = comboBox2.Text;
                comboBox3_SelectedIndexChanged(sender, e);
            }
        }

		private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
		{
			if (comboBox3.Text == "Выбрать характеристики..." || comboBox3.Text == "")
			{

			}
			else if (comboBox3.Text == "Общие характеристики")
			{
				GeneralInput();
			}
			else if (comboBox3.Text == "S.M.A.R.T.")
			{
				DiskInput();
			}
		}

		private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.DataSource = dbview.GetTables();
            comboBox3.DataSource = new List<string>() { "Выбрать характеристики..." };

            this.listView1
                .GetType()
                .GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                .SetValue(listView1, true, null);

            this.listView2
                .GetType()
                .GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                .SetValue(listView2, true, null);
        }


        private void inputListView2(string[,] table)
        {
            for (int i = 0; i < table.GetLength(0); i++)
            {
                ListViewItem item = new(table[i, 0]);
                for (int j = 1; j < table.GetLength(1); j++)
                    item.SubItems.Add(table[i, j]);
                listView2.Items.Add(item);
            }
            listView2.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);

            int minColumnWidth = 60;
            int maxColumnWidth = 120;
            foreach (ColumnHeader column in listView2.Columns)
            {
                column.AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent);
                if (column.Width < minColumnWidth)
                {
                    column.Width = minColumnWidth;
                }
                else if (column.Width > maxColumnWidth)
                {
                    column.Width = maxColumnWidth;
                }
            }
        }

        private void inputListView1(string[,] table, string[] columns)
        {
            for (int i = 0; i < columns.Length; i++)
            {
                ListViewItem item = new(columns[i]);
                item.SubItems.Add(table[0, i]);
                listView1.Items.Add(item);
            }
            listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            foreach (ColumnHeader column in listView1.Columns)
            {
                column.AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent);
            }
        }

        private void UsersInput()
        {
            listView2.Columns.Clear();
            listView2.Items.Clear();
            listView2.Columns.Add("ID");
            listView2.Columns.Add("Кабинет");
            listView2.Columns.Add("LAN");
            listView2.Columns.Add("ФИО");
            inputListView2(this.users);
        }
            private void GeneralInput()
        {
            listView1.Columns.Clear();
            listView1.Items.Clear();
            this.general = dbview.ParseTables("General", organization, datetime);
            listView1.Columns.Add("Элемент");
            listView1.Columns.Add("Значение");
            inputListView1(this.general, generalList.ToArray());
            
        }
        private void DiskInput()
        {
            listView1.Columns.Clear();
            listView1.Items.Clear();
            this.disk = dbview.ParseTables("Disk", organization, datetime);
            listView1.Columns.Add("Элемент");
            listView1.Columns.Add("Значение");
            inputListView1(this.disk, diskList.ToArray());
        }
        private void Button3_click(object sender, EventArgs e)
        {
            ExcelExporter xlsxFile = new ExcelExporter();
            xlsxFile.ExportToExcel(organization, "Общие характеристики", "S.M.A.R.T.", dbview);
        }

        private void ListView2_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (listView2.SelectedItems[0].Text != "")
            { 
			comboBox2.Enabled = true;
            comboBox3.Enabled = true;
            listView1.Enabled = true;
            string id = listView2.SelectedItems[0].Text;
            string[] time = dbview.ParseTime("General", organization, id).ToArray();
            datetime = time[^1];
            if (comboBox3.Text != "Общие характеристики" && comboBox3.Text != "S.M.A.R.T." )
                comboBox3.DataSource = new List<string>() { "Общие характеристики", "S.M.A.R.T." };
            comboBox2.DataSource = time;
            comboBox2.Text = "Выбрать дату...";
            groupBox1.Text = $"Пользователь ID {id}";
			}
		}
    }
}