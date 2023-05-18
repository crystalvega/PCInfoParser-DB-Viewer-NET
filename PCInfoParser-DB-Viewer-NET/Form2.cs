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
    public partial class Form2 : Form
    {
        string error = "";
        public Form2(string error)
        {
            this.error = error;
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            label1.Text = error;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}