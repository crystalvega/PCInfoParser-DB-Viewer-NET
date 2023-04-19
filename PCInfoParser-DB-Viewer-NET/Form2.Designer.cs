using System.Drawing;
using System;
using System.Windows.Forms;

namespace PCInfoParser_DB_Viewer_NET
{
    partial class Form2
    {
        string error = "123";
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
        /// 
        private void SetError(string error)
        {
            this.error = error;
        }
        private void Button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void InitializeComponent()
        {
            button1 = new Button();
            label1 = new Label();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(191, 77); // координата, в.
            button1.Name = "button1"; // имя button1
            button1.Size = new Size(127, 32); // Размер будет, пикс
            button1.TabIndex = 1; // для дальнейшего использования
            button1.Text = "OK";
            button1.UseVisualStyleBackColor = true; // стандартный цвет, используется по-умолчанию
            button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // label1
            // 
            label1.Location = new Point(12, 9);
            label1.Name = "label1";
            label1.RightToLeft = RightToLeft.No;
            label1.Size = new Size(491, 65);
            label1.TabIndex = 0;
            label1.Text = this.error;
            label1.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // Form2
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(515, 121);
            Controls.Add(label1);
            Controls.Add(button1);
            FormBorderStyle = FormBorderStyle.Fixed3D;
            MaximizeBox = false;
            Name = "Form2";
            Text = "Error";
            ResumeLayout(false);
        }

        #endregion

        private Label label1;
        private Button button1;
    }
}
