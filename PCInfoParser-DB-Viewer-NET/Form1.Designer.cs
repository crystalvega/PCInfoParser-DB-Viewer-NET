using System.Drawing;
using System.Windows.Forms;

namespace PCInfoParser_DB_Viewer_NET
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            comboBox1 = new ComboBox();
            comboBox2 = new ComboBox();
            button1 = new Button();
            button2 = new Button();
            button3 = new Button();
            listView1 = new ListView();
            splitContainer1 = new SplitContainer();
            ((System.ComponentModel.ISupportInitialize)splitContainer1).BeginInit();
            splitContainer1.Panel1.SuspendLayout();
            splitContainer1.Panel2.SuspendLayout();
            splitContainer1.SuspendLayout();
            SuspendLayout();
            // 
            // comboBox1
            // 
            comboBox1.FormattingEnabled = true;
            comboBox1.Location = new Point(12, 12);
            comboBox1.Name = "comboBox1";
            comboBox1.Size = new Size(761, 23);
            comboBox1.TabIndex = 0;
            comboBox1.DataSource = databases;
            comboBox1.SelectedIndexChanged += Combobox1_SelectedIndexChanged;
            // 
            // comboBox2
            // 
            comboBox2.FormattingEnabled = true;
            comboBox2.Location = new Point(12, 41);
            comboBox2.Name = "comboBox2";
            comboBox2.Size = new Size(761, 23);
            comboBox2.TabIndex = 1;
            comboBox2.Enabled = false;
            comboBox2.SelectedIndexChanged += Combobox2_SelectedIndexChanged;
            // 
            // button1
            // 
            button1.Location = new Point(-12, -8);
            button1.Name = "button1";
            button1.Size = new Size(401, 42);
            button1.TabIndex = 3;
            button1.Text = "Общие характеристики";
            button1.UseVisualStyleBackColor = true;
            button1.Click += Button1_click;
            button1.Enabled = false;
            // 
            // button2
            // 
            button2.Location = new Point(-10, -8);
            button2.Name = "button2";
            button2.Size = new Size(401, 42);
            button2.TabIndex = 4;
            button2.Text = "S.M.A.R.T.";
            button2.UseVisualStyleBackColor = true;
            button2.Click += Button2_click;
            button2.Enabled = false;
            // 
            // button3
            // 
            button3.Location = new Point(12, 555);
            button3.Name = "button3";
            button3.Size = new Size(761, 23);
            button3.TabIndex = 6;
            button3.Text = "Сохранить в XLSX";
            button3.UseVisualStyleBackColor = true;
            button3.Click += Button3_click;
            button3.Enabled = false;
            // 
            // listView1
            // 
            listView1.Location = new Point(12, 99);
            listView1.Name = "listView1";
            listView1.Size = new Size(761, 450);
            listView1.TabIndex = 2;
            listView1.UseCompatibleStateImageBehavior = false;
            listView1.View = View.Details;
            listView1.FullRowSelect = true;
            listView1.Enabled = false;
            // 
            // splitContainer1
            // 
            splitContainer1.Location = new Point(12, 70);
            splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            splitContainer1.Panel1.Controls.Add(button1);
            // 
            // splitContainer1.Panel2
            // 
            splitContainer1.Panel2.Controls.Add(button2);
            splitContainer1.Size = new Size(761, 23);
            splitContainer1.SplitterDistance = 377;
            splitContainer1.TabIndex = 7;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(785, 590);
            Controls.Add(splitContainer1);
            Controls.Add(listView1);
            Controls.Add(button3);
            Controls.Add(comboBox2);
            Controls.Add(comboBox1);
            FormBorderStyle = FormBorderStyle.Fixed3D;
            Name = "Form1";
            Text = "PCInfoParser DB Viewer";
            splitContainer1.Panel1.ResumeLayout(false);
            splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)splitContainer1).EndInit();
            splitContainer1.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private ComboBox comboBox1;
        private ComboBox comboBox2;
        private Button button1;
        private Button button2;
        private Button button3;
        private ListView listView1;
        private SplitContainer splitContainer1;
    }
}