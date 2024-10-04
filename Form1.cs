using System;
using System.Configuration;
using System.Data;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using ICSharpCode.SharpZipLib.Zip;
using Neusoft.HISFC.Management.Manager;
using Neusoft.NFC.Management;
using Oracle.DataAccess.Client;
using System.Net.Http;
using System.Threading.Tasks;
using System.Xml;

using System.Text.RegularExpressions;

namespace HisAnalytics
{

    public partial class Form1 : Form
    {
        private string _currentPath;



        private OracleConnection _con;

        public Form1()
        {
            InitializeComponent();
            textstart.Text = DateTime.Now.ToString("yyyy-MM-dd").ToString();
            textend.Text = DateTime.Now.ToString("yyyy-MM-dd").ToString();


        }

        public static string Decrypt(string strDataSource)
        {
            if (string.IsNullOrEmpty(strDataSource.Trim()))
                return (string)null;
            string str1 = strDataSource.Trim().Substring(0, strDataSource.LastIndexOf("A"));
            string str2 = "";
            for (int startIndex = 0; startIndex < str1.Length; ++startIndex)
            {
                string str3 = "";
                switch (str1.Substring(startIndex, 1))
                {
                    case "E":
                        int int16 = (int)Convert.ToInt16(str1.Substring(startIndex + 1, 1));
                        for (int index = 0; index < int16; ++index)
                            str3 += str1.Substring(startIndex + 2, 1);
                        startIndex += 2;
                        break;
                    case "D":
                        string str4 = str1.Substring(startIndex + 1, 1);
                        str3 = str3 + str4 + str4 + str4 + str4;
                        ++startIndex;
                        break;
                    case "C":
                        string str5 = str1.Substring(startIndex + 1, 1);
                        str3 = str3 + str5 + str5 + str5;
                        ++startIndex;
                        break;
                    case "B":
                        string str6 = str1.Substring(startIndex + 1, 1);
                        str3 = str3 + str6 + str6;
                        ++startIndex;
                        break;
                    default:
                        str3 = str1.Substring(startIndex, 1);
                        break;
                }
                str2 += str3;
            }
            string str7 = str2;
            int length1 = str7.Length;
            string str8 = "";
            for (int startIndex = 0; startIndex < str7.Length; ++startIndex)
            {
                string str9 = str7.Substring(startIndex, 1);
                str8 += Convert.ToString(Convert.ToInt32(str9, 16), 2).PadLeft(4, '0');
            }
            string str10 = str8;
            int num1 = str10.Length / 16;
            int num2 = 16;
            string[] strArray = new string[1600];
            for (int index = 0; index < str10.Length; ++index)
            {
                int num3 = index / num2;
                int num4 = index % num2;
                strArray[num4 * num1 + num3] = (num4 * num1 + num3) % 2 != 1 ? str10.Substring(num3 * num2 + num4, 1) : (str10.Substring(num3 * num2 + num4, 1) == "1" ? "0" : "1");
            }
            string input = string.Join("", strArray);
            string str11 = input.Length.ToString();
            int num5 = 0;
            for (int startIndex = 0; startIndex < str11.Length; ++startIndex)
                num5 += (int)Convert.ToInt16(str11.Substring(startIndex, 1));
            for (int index = num5 - 1; index < input.Length; index += num5)
            {
                int num6 = 0;
                for (int startIndex = index - 1; startIndex > index - num5; --startIndex)
                    num6 += (int)Convert.ToInt16(input.Substring(startIndex, 1));
                if (num6 % 2 == 1)
                {
                    string str12 = input.Substring(index, 1) == "0" ? "1" : "0";
                    input = input.Substring(0, index) + str12 + input.Substring(index + 1);
                }
            }
            string str13 = input.Replace("1", "");
            int num7 = input.Length - str13.Length;
            int length2 = (int)Convert.ToInt16(num7.ToString().Substring(num7.ToString().Length - 1));
            switch (length2)
            {
                case 0:
                    length2 = 7;
                    break;
                case 1:
                    length2 = 13;
                    break;
            }
            for (int index1 = length2 - 1; index1 < input.Length; index1 += length2)
            {
                string str14 = input.Substring(index1 - (length2 - 1), length2);
                string str15 = "";
                for (int index2 = 0; index2 < length2; ++index2)
                    str15 += str14.Substring(length2 - 1 - index2, 1);
                input = input.Substring(0, index1 - (length2 - 1)) + str15 + input.Substring(index1 + 1);
            }
            CaptureCollection captures = Regex.Match(input, "([01]{8})+").Groups[1].Captures;
            byte[] bytes = new byte[captures.Count];
            for (int i = 0; i < captures.Count; ++i)
                bytes[i] = Convert.ToByte(captures[i].Value, 2);
            string str16 = Encoding.Unicode.GetString(bytes, 0, bytes.Length);
            return str16.Substring(0, str16.LastIndexOf("A"));
        }

        internal static async Task<int> GetSettingNewAsync()
        {
            XmlDocument xmlDocument = new XmlDocument();
            try
            {
                // 直接从指定的URL加载HisProfile.xml
                string profilePath = "http://172.17.0.6/his/HisProfile.xml";

                // 使用HttpClient从URL获取XML文件内容
                using (HttpClient client = new HttpClient())
                {
                    string xmlContent = await client.GetStringAsync(profilePath);
                    xmlDocument.LoadXml(xmlContent);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("装载HisProfile.xml失败！\n" + ex.Message);
                return -1;
            }

            // 获取数据库设置节点
            XmlNode xmlNode2 = xmlDocument.SelectSingleNode("/设置/数据库设置");
            if (xmlNode2 == null)
            {
                MessageBox.Show("没有找到数据库设置!");
                return -1;
            }

            string encryptedPassword = xmlNode2.Attributes[0].Value;

            // 解密数据库密码
            string decryptedPassword;
            if (encryptedPassword.Contains("PWD"))
            {
                decryptedPassword = Decrypt(encryptedPassword);
            }
            else
            {
                decryptedPassword = encryptedPassword;  // 如果不需要解密则直接赋值
            }

            // 设置解密后的数据库连接字符串
            Program.DataSource = decryptedPassword;

            

            return 0;
        }


        private async void button1_Click(object sender, EventArgs e)
        {
            _currentPath = Application.StartupPath.Substring(Application.StartupPath.Length - 1, 1) != "\\"
                ? Application.StartupPath + "\\"
                : Application.StartupPath;

            // 禁用按钮以避免重复点击
            button1.Enabled = false;
            button1.Text = "正在连接...";

                // 异步获取设置并解密密码
                var result = await GetSettingNewAsync();
                if (result == -1)
                {
                    button1.Text = "连接失败";
                    button1.Enabled = true;
                    return;
                }
          

            // 异步处理设置和数据库连接
            var result2 = await Task.Run(() =>
            {
                return ConnectDb();
            });

            // 处理连接结果
            if (result2 == -1)
            {
                button1.Text = "连接失败";
                button1.Enabled = true;
                return;
            }

            // 连接成功
            button1.Text = "连接成功！";
            button1.Enabled = false;
        }


        int ConnectDb()
        {
            _con = new OracleConnection(Program.DataSource);
            try
            {
                _con.Open();
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show(@"无法连接数据库！\n" + ex.Message);
                return -1;
            }

            return 0;
        }


        int ExeSql(string sql, string filename)
        {
            DataTable dt = new DataTable();
            using (OracleCommand cmd = _con.CreateCommand())
            {
                cmd.CommandText = sql;
                OracleDataAdapter adap = new OracleDataAdapter(cmd);
                adap.Fill(dt);
            }

            FileHelpers.CsvEngine.DataTableToCsv(dt, ".\\data\\" + filename + ".temp");
            ZipFile(filename);
            MessageBox.Show(@"文件生成" + filename + @".dat");

            return 0;
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            //if (!checkdate()) return;

            //string sqlDetail = @"select empl_code,empl_name,dept_code,nurse_cell_code,empl_type,expert_flag from com_employee";
            string sqlDetail = ConfigurationManager.AppSettings["doctor"];

            ExeSql(sqlDetail, "doctor");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //if (!checkdate()) return;

            //string sqlDetail = @"select dept_code,dept_name from com_department";
            string sqlDetail = ConfigurationManager.AppSettings["department"];

            ExeSql(sqlDetail, "department");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string sqlDetail = ConfigurationManager.AppSettings["baseinfo"];
            ExeSql(sqlDetail, "baseinfo");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (!checkdate()) return;

            string sqlDetail =
                   string.Format(ConfigurationManager.AppSettings["feedetail"], textstart.Text, textend.Text);
            ExeSql(sqlDetail, "feedetail");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (!checkdate()) return;
            string sqlDetail = string.Format(ConfigurationManager.AppSettings["medicinelist"], textstart.Text,
                   textend.Text);
            ExeSql(sqlDetail, "medicinelist");
        }

        private bool checkdate()
        {
            DateTime dtStart = DateTime.ParseExact(textstart.Text, "yyyy-mm-dd",
                System.Globalization.CultureInfo.CurrentCulture);
            DateTime dtEnd = DateTime.ParseExact(textend.Text, "yyyy-mm-dd",
                System.Globalization.CultureInfo.CurrentCulture);
            var dt = (dtEnd - dtStart).Days;
            if (dt > 40 || dt <= 0)
            {
                MessageBox.Show(@"时间太长或者太短，要在1-40天！");
                return false;

            }

            return true;
        }

        private void ZipFile(string file)
        {
            //加密压缩1A1F713ABFE4D8DE6673DA9421E83DF
            //string password = "%$@#SY%%$$RTYUUUH";
            ZipOutputStream s = new ZipOutputStream(System.IO.File.Create(".\\data\\" + file + ".dat"));
            s.SetLevel(6); // 0 - store only to 9 - means best compression
            //s.Password = md5.encrypt(password);
            //打开压缩文件
            FileStream fs = System.IO.File.OpenRead(".\\data\\" + file + ".temp");

            byte[] buffer = new byte[fs.Length];
            fs.Read(buffer, 0, buffer.Length);

            Array arr = file.Split('\\');
            string le = arr.GetValue(arr.Length - 1).ToString();
            ZipEntry entry = new ZipEntry(le);
            entry.DateTime = DateTime.Now;
            entry.Size = fs.Length;
            fs.Close();
            s.PutNextEntry(entry);
            s.Write(buffer, 0, buffer.Length);
            buffer = null;
            s.Finish();
            s.Close();
            System.IO.File.Delete(".\\data\\" + file + ".temp");
        }




    }
}
