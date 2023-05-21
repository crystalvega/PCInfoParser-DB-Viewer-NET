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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
			this.comboBox1 = new System.Windows.Forms.ComboBox();
			this.comboBox2 = new System.Windows.Forms.ComboBox();
			this.button3 = new System.Windows.Forms.Button();
			this.listView1 = new System.Windows.Forms.ListView();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.comboBox3 = new System.Windows.Forms.ComboBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.listView2 = new System.Windows.Forms.ListView();
			this.splitContainer1 = new System.Windows.Forms.SplitContainer();
			this.groupBox1.SuspendLayout();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
			this.splitContainer1.Panel1.SuspendLayout();
			this.splitContainer1.Panel2.SuspendLayout();
			this.splitContainer1.SuspendLayout();
			this.SuspendLayout();
			// 
			// comboBox1
			// 
			this.comboBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.comboBox1.FormattingEnabled = true;
			this.comboBox1.Location = new System.Drawing.Point(10, 10);
			this.comboBox1.Name = "comboBox1";
			this.comboBox1.Size = new System.Drawing.Size(821, 21);
			this.comboBox1.TabIndex = 0;
			this.comboBox1.Text = "Выбрать организацию...";
			this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.Combobox1_SelectedIndexChanged);
			// 
			// comboBox2
			// 
			this.comboBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.comboBox2.Enabled = false;
			this.comboBox2.FormattingEnabled = true;
			this.comboBox2.Location = new System.Drawing.Point(6, 389);
			this.comboBox2.Name = "comboBox2";
			this.comboBox2.Size = new System.Drawing.Size(518, 21);
			this.comboBox2.TabIndex = 1;
			this.comboBox2.Text = "Выбрать дату...";
			this.comboBox2.SelectedIndexChanged += new System.EventHandler(this.Combobox2_SelectedIndexChanged);
			// 
			// button3
			// 
			this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button3.Enabled = false;
			this.button3.FlatAppearance.BorderSize = 0;
			this.button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.button3.Image = global::PCInfoParser_DB_Viewer_NET.Properties.Resources.diskette2;
			this.button3.Location = new System.Drawing.Point(837, 10);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(21, 21);
			this.button3.TabIndex = 6;
			this.button3.UseVisualStyleBackColor = true;
			this.button3.Click += new System.EventHandler(this.Button3_click);
			// 
			// listView1
			// 
			this.listView1.Alignment = System.Windows.Forms.ListViewAlignment.Left;
			this.listView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.listView1.Enabled = false;
			this.listView1.FullRowSelect = true;
			this.listView1.HideSelection = false;
			this.listView1.Location = new System.Drawing.Point(6, 46);
			this.listView1.Name = "listView1";
			this.listView1.Size = new System.Drawing.Size(518, 337);
			this.listView1.TabIndex = 2;
			this.listView1.UseCompatibleStateImageBehavior = false;
			this.listView1.View = System.Windows.Forms.View.Details;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Controls.Add(this.comboBox3);
			this.groupBox1.Controls.Add(this.listView1);
			this.groupBox1.Controls.Add(this.comboBox2);
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(530, 416);
			this.groupBox1.TabIndex = 8;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Пользователь {client_id}";
			// 
			// comboBox3
			// 
			this.comboBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.comboBox3.Enabled = false;
			this.comboBox3.FormattingEnabled = true;
			this.comboBox3.Location = new System.Drawing.Point(6, 19);
			this.comboBox3.Name = "comboBox3";
			this.comboBox3.Size = new System.Drawing.Size(518, 21);
			this.comboBox3.TabIndex = 3;
			this.comboBox3.Text = "Выбрать характеристики...";
			this.comboBox3.SelectedIndexChanged += new System.EventHandler(this.comboBox3_SelectedIndexChanged);
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.listView2);
			this.groupBox2.Location = new System.Drawing.Point(-2, 0);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(316, 416);
			this.groupBox2.TabIndex = 9;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Пользователи";
			// 
			// listView2
			// 
			this.listView2.Alignment = System.Windows.Forms.ListViewAlignment.Left;
			this.listView2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.listView2.Enabled = false;
			this.listView2.FullRowSelect = true;
			this.listView2.HideSelection = false;
			this.listView2.Location = new System.Drawing.Point(6, 19);
			this.listView2.MultiSelect = false;
			this.listView2.Name = "listView2";
			this.listView2.Size = new System.Drawing.Size(304, 391);
			this.listView2.TabIndex = 3;
			this.listView2.UseCompatibleStateImageBehavior = false;
			this.listView2.View = System.Windows.Forms.View.Details;
			this.listView2.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.ListView2_MouseDoubleClick);
			// 
			// splitContainer1
			// 
			this.splitContainer1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.splitContainer1.Location = new System.Drawing.Point(10, 37);
			this.splitContainer1.Name = "splitContainer1";
			// 
			// splitContainer1.Panel1
			// 
			this.splitContainer1.Panel1.Controls.Add(this.groupBox1);
			// 
			// splitContainer1.Panel2
			// 
			this.splitContainer1.Panel2.Controls.Add(this.groupBox2);
			this.splitContainer1.Size = new System.Drawing.Size(848, 416);
			this.splitContainer1.SplitterDistance = 530;
			this.splitContainer1.TabIndex = 10;
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(868, 465);
			this.Controls.Add(this.splitContainer1);
			this.Controls.Add(this.button3);
			this.Controls.Add(this.comboBox1);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MinimumSize = new System.Drawing.Size(500, 300);
			this.Name = "Form1";
			this.Text = "PCInfoParser DB Viewer";
			this.Load += new System.EventHandler(this.Form1_Load);
			this.groupBox1.ResumeLayout(false);
			this.groupBox2.ResumeLayout(false);
			this.splitContainer1.Panel1.ResumeLayout(false);
			this.splitContainer1.Panel2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
			this.splitContainer1.ResumeLayout(false);
			this.ResumeLayout(false);

        }

        #endregion

        private ComboBox comboBox1;
        private ComboBox comboBox2;
        private Button button3;
        private ListView listView1;
        private GroupBox groupBox1;
        private GroupBox groupBox2;
        private ComboBox comboBox3;
        private ListView listView2;
        private SplitContainer splitContainer1;
    }
}