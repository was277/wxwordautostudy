using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace EnglishTool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }



        Thread Th = null;
        #region //初始化环境
        private void InitBtn_Click(object sender, EventArgs e)
        {
            MainFC.Init();
        }
        #endregion
        List<string> acclist = new List<string>();

        #region //开始学习
        private void StartLineBtn_Click(object sender, EventArgs e)
        {
            acclist = AccListText.Lines.ToList();
            Th = new Thread(new ThreadStart(Study));
            Th.IsBackground = true;
            Th.Start();
            StartLineBtn.Enabled = false;
        }

        private void Study()
        {
            foreach (string i in acclist)
            {
                string[] arr = i.Split(' ');
                if (arr.Length != 3)
                {
                    continue;
                }
                InfoText.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 账户：" + arr[0] + "即将开始登录\r\n" + InfoText.Text;
                string statu = MainFC.Login(arr[2], arr[0], arr[1]);
                InfoText.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 账户：" + arr[0] + "登录状态为：" + statu + "\r\n" + InfoText.Text;
                if (statu.IndexOf("成功") == -1)
                {
                    InfoText.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 账户：" + arr[0] + "登录不成功即将跳转下一个账户：\r\n" + InfoText.Text;
                    continue;
                }
                InfoText.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 账户：" + arr[0] + "开始进行教师测试\r\n" + InfoText.Text;
                Random irt = new Random();
                int timert = Int32.Parse(textBox1.Text) * 1000;
                int iScore = 100;
                if (radioButton12.Checked == true) { iScore = irt.Next(90, 100); }
                if (radioButton13.Checked == true) { iScore = 100; }
                string teastatu = MainFC.TeachTest(iScore, timert, arr[0]);
                InfoText.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 账户：" + arr[0] + "教师测试状态为：" + teastatu + "\r\n" + InfoText.Text;
                // List<string> perlist = MainFC.GetUrl();
                string str = "";
                if (radioButton1.Checked == true) { str = "/student/space.html?programName=CN-Level41"; }
                if (radioButton2.Checked == true) { str = "/student/space.html?programName=CN-Level42"; }
                if (radioButton3.Checked == true) { str = "/student/space.html?programName=CN-Level43"; }
                if (radioButton4.Checked == true) { str = "/student/space.html?programName=CN-Level44"; }
                if (radioButton5.Checked == true) { str = "/student/space.html?programName=CN-Level63"; }
                if (radioButton6.Checked == true) { str = "/student/space.html?programName=CN-Level64"; }
                if (radioButton7.Checked == true) { str = "/student/space.html?programName=CN-Level65"; }
                if (radioButton8.Checked == true) { str = "/student/space.html?programName=CN-Level66"; }
                if (radioButton9.Checked == true) { str = "/student/space.html?programName=CN-Level67"; }
                if (radioButton10.Checked == true) { str = "/student/space.html?programName=CN-Level61"; }
                if (radioButton11.Checked == true) { str = "/student/space.html?programName=CN-Level62"; }

                //foreach(string str in perlist)
                //{
                InfoText.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 账户：" + arr[0] + "开始进行学前测试\r\n" + InfoText.Text;
                string perstatu = MainFC.PerStudy(str, arr[0]);
                InfoText.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 账户：" + arr[0] + "学前测试状态为：" + perstatu + "\r\n" + InfoText.Text;
                InfoText.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 账户：" + arr[0] + "开始进行单词学习\r\n" + InfoText.Text;
                string wordstatu = MainFC.Study(str);
                InfoText.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 账户：" + arr[0] + "单词学习状态为：" + wordstatu + "\r\n" + InfoText.Text;
                InfoText.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 账户：" + arr[0] + "开始进行单元闯关\r\n" + InfoText.Text;
                string winstatu = MainFC.UnitTest(str, arr[0]);
                InfoText.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 账户：" + arr[0] + "单元闯关状态为：" + wordstatu + "\r\n" + InfoText.Text;
                //}

            }
            StartLineBtn.Enabled = true;
        }

        public void TestUnit()
        {
            foreach (string i in acclist)
            {
                string[] arr = i.Split(' ');
                if (arr.Length != 3)
                {
                    continue;
                }
                InfoText.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 账户：" + arr[0] + "即将开始登录\r\n" + InfoText.Text;
                string statu = MainFC.Login(arr[2], arr[0], arr[1]);
                InfoText.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 账户：" + arr[0] + "登录状态为：" + statu + "\r\n" + InfoText.Text;
                if (statu.IndexOf("成功") == -1)
                {
                    InfoText.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 账户：" + arr[0] + "登录不成功即将跳转下一个账户：\r\n" + InfoText.Text;
                    continue;
                }
                InfoText.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 账户：" + arr[0] + "开始进行教师测试\r\n" + InfoText.Text;
                Random irt = new Random();
                int timert = Int32.Parse(textBox1.Text) * 1000;
                int iScore = 100;
                if (radioButton12.Checked == true) { iScore = irt.Next(90, 100); }
                if (radioButton13.Checked == true) { iScore = 100; }
                acclist = AccListText.Lines.ToList();
                string teastatu = MainFC.TeachTest(iScore, timert, arr[0]);
                InfoText.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 账户：" + arr[0] + "教师测试状态为：" + teastatu + "\r\n" + InfoText.Text;
            }
            StartLineBtn.Enabled = true;
            button1.Enabled = true;
        }

        #endregion
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Environment.Exit(0);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //只做组卷 
            acclist = AccListText.Lines.ToList();
            Th = new Thread(new ThreadStart(TestUnit));
            Th.IsBackground = true;
            Th.Start();
            StartLineBtn.Enabled = false;
            button1.Enabled = false;
        }

        public void filllogs()
        {
            SQLiteConnection m_dbConnection = new SQLiteConnection("Data Source =cc.db; Version = 3; New = false; ");
            m_dbConnection.Open();
            SQLiteDataAdapter sqlda = new SQLiteDataAdapter("select * from logs", m_dbConnection);
            DataSet dt = new DataSet();
            sqlda.Fill(dt);
            dataGridView1.DataSource = dt.Tables[0];
            m_dbConnection.Close();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            filllogs();
        }

        private void 清除日志_Click(object sender, EventArgs e)
        {
            SQLiteConnection m_dbConnection = new SQLiteConnection("Data Source =cc.db; Version = 3; New = false; ");
            m_dbConnection.Open();
            string sql = "delete from logs where 1=1;";
            SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
            command.ExecuteNonQuery();
            m_dbConnection.Close();
        }





    }
}
