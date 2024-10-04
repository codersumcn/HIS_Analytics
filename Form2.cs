using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HisAnalytics
{
    public partial class Form2 : Form
    {

        public Form2()
        {
            InitializeComponent();
        }


        private void button3_Click(object sender, EventArgs e)
        {
            Form1 myForm = new Form1();
            myForm.ShowDialog(); // 启动模态对话框

        }
    }
}
