using PCInfoParser_DB_Viewer;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PCInfoParser_DB_Viewer_NET
{
    public partial class Form1 : Form
    {
        List<string> generalList = new List<string>() { "Кабинет", "LAN", "ФИО", "Монитор", "Диагональ", "Тип принтера", "Модель принтера", "ПК", "Материнская плата", "Процессор", "Частота процессора", "Баллы Passmark", "Дата выпуска", "Тип ОЗУ", "ОЗУ, 1 Планка", "ОЗУ, 2 Планка", "ОЗУ, 3 Планка", "ОЗУ, 4 Планка", "Сокет", "Диск 1", "Состояние диска 1", "Диск 2", "Состояние диска 2", "Диск 3", "Состояние диска 3", "Диск 4", "Состояние диска 4", "Операционная система", "Антивирус", "CPU Под замену", "Все CPU под сокет", "Дата создания" };
        List<string> diskList = new List<string>() { "Кабинет", "LAN", "ФИО", "Диск", "Наименование", "Прошивка", "Размер", "Время работы", "Включён", "Состояние", "Температура", "Дата создания" };
        List<string> databases = new List<string>();
        private System.Windows.Forms.Timer tooltipTimer = new System.Windows.Forms.Timer();
        private ListViewItem lastItem;
        string database;
        string[,] general;
        string[,] disk;
        string datetime;
        MySQLConnect dbview;

        public Form1(List<string> databases, MySQLConnect dbview)
        {
            this.databases = databases;
            this.dbview = dbview;
            InitializeComponent();
        }

        private void Combobox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "")
            {
                comboBox2.Enabled = false;
                comboBox2.SelectedIndex = -1;
                button1.Enabled = false;
                button2.Enabled = false;
                button3.Enabled = false;
                listView1.Columns.Clear();
                listView1.Enabled = false;

            }
            else
            {
                comboBox2.Enabled = true;
                this.database = comboBox1.Text;
                comboBox2.DataSource = dbview.ParseTime("all configuration", database);

            }
        }
        private void Combobox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text == "")
            {
                button1.Enabled = false;
                button2.Enabled = false;
                button3.Enabled = false;
                listView1.Columns.Clear();
                listView1.Enabled = false;
            }
            else
            {
                listView1.Enabled = true;
                button3.Enabled = true;
                datetime = comboBox2.Text;
                this.general = dbview.ParseTables(this.database, "all configuration", datetime);
                this.disk = dbview.ParseTables(this.database, "disk configuration", datetime);
                Button1_click(sender, e);
            }
        }

        private void inputListView(string[,] table)
        {
            for (int i = 0; i < table.GetLength(0); i++)
            {
                ListViewItem item = new ListViewItem(table[i, 0]);
                for (int j = 1; j < table.GetLength(1); j++)
                    item.SubItems.Add(table[i, j]);
                listView1.Items.Add(item);
            }
            listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);

            int minColumnWidth = 120;
            int maxColumnWidth = 400;
            foreach (ColumnHeader column in listView1.Columns)
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
        private void Button1_click(object sender, EventArgs e)
        {
            listView1.Columns.Clear();
            listView1.Items.Clear();
            button1.Enabled = false;
            button2.Enabled = true;
            generalList.ForEach(name => listView1.Columns.Add(name));
            inputListView(this.general);
        }
        private void Button2_click(object sender, EventArgs e)
        {
            listView1.Columns.Clear();
            listView1.Items.Clear();
            button2.Enabled = false;
            button1.Enabled = true;
            diskList.ForEach(name => listView1.Columns.Add(name));
            inputListView(this.disk);
        }
        private void Button3_click(object sender, EventArgs e)
        {
            string filePath = $"{this.database} {this.datetime}.xlsx";
            ExcelExporter xlsxFile = new ExcelExporter();
            xlsxFile.ExportToExcel(filePath, "Общие характеристики", "S.M.A.R.T.", this.general, this.disk);
        }
    }
}