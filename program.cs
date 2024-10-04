using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Windows.Forms;

namespace HisAnalytics
{
    static class Program
    {
        public static string DataSource { get; internal set; }

        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Form2 f2 = new Form2();
            //f2.Show();
            if (f2.ShowDialog() == DialogResult.OK)
            {
                Application.Run(new Form1());
            }

            //Application.Run();
        }
    }
}
