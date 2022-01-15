using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;//支持Chrome
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System.Data.SQLite;
using System.Windows.Forms;

namespace EnglishTool
{
    public class MainFC
    {
        public static IWebDriver selenium = null;

        public static void Init()
        {
            var option = new ChromeOptions();
            option.AddArgument("--incognito");
            option.AddArgument("--remote-debug-port=9222");
            option.AddArgument("--user-data-dir=C:\\tmp\\chromeprofiles\\profile\\" + DateTime.Now.ToString("yyyyMMddHHmmss"));
            selenium = new ChromeDriver(option);
        }

        public static string GetSource()
        {
            return selenium.PageSource;
        }

        public static List<string> GetUrl()
        {
            selenium.Navigate().GoToUrl("http://www.wxwords.com/student/space.html");
            Thread.Sleep(1000);
            string xpath = "/html/body/div[2]/div[1]/span[2]";
            if (!ExistElement(xpath, "XPath", 0))
            {
                xpath = "/html/body/div[4]/div[1]/span[2]";
            }
            Click(xpath, "XPath", 0);
            Thread.Sleep(1000);
            Regex reg = new Regex("<a href=\"/student/space.html\\?[^\"]+");
            MatchCollection collection = reg.Matches(selenium.PageSource);
            List<string> urllist = new List<string>();
            foreach (Match i in collection)
            {
                urllist.Add(i.Value.Replace("<a href=\"", ""));
            }

            return urllist;
        }

        public static string Login(string school, string acc, string psw)
        {
            selenium.Navigate().GoToUrl("http://www.wxwords.com");
            Thread.Sleep(500);
            selenium.Manage().Cookies.DeleteAllCookies();
            Thread.Sleep(1000);
            selenium.Navigate().Refresh();
            Thread.Sleep(1000);
            SendKey("allschool", "Id", 0, school);
            SendKey("userId", "Id", 0, acc);
            SendKey("password", "Id", 0, psw);
            Click("but", "ClassName", 0);
            try
            {
                if (ExistAlert())
                {
                    string tiptext = selenium.SwitchTo().Alert().Text;
                    selenium.SwitchTo().Alert().Accept();
                    return tiptext;
                }
                else if (selenium.PageSource.IndexOf("退出") != -1)
                {
                    return "登录成功";
                }
                else
                {
                    return "登录失败";
                }
            }
            catch
            {
                return "登录异常";
            }
        }

        /// <summary>
        /// 检查是否存在js的alert对话框
        /// </summary>
        /// <returns></returns>
        private static bool ExistAlert()
        {
            try
            {
                selenium.SwitchTo().Alert();
                return true;
            }
            catch (NoAlertPresentException ex)
            {
                return false;
            }
        }
        #region //传值学习
        //学前测试
        public static string PerStudy(string perurl,string account)
        {
            selenium.Navigate().GoToUrl("http://www.wxwords.com" + perurl);
            Thread.Sleep(5000);
            if (selenium.Title.IndexOf("智学") == -1)
            {
                return "无法跳转主页面";
            }
            if (selenium.PageSource.IndexOf("学前测试") == -1)
            {
                return "无学前测试";
            }

            selenium.Navigate().GoToUrl("http://www.wxwords.com" + perurl.Replace("space.html?", "quiz.html?") + "&quizType=3");
            Thread.Sleep(1500);
            IJavaScriptExecutor js = (IJavaScriptExecutor)selenium;
            js.ExecuteScript("$(\"#footer\").remove(\".foot\");");
            
            //banduan and xunzheti
            //Regex RightRegex = new Regex("<input id=\"question_\\d+_\\d\" type=\"radio\" value=\"true\".+?>");
            //Regex IdRegex = new Regex("question_\\d+_\\d");
            //MatchCollection collection = RightRegex.Matches(selenium.PageSource);
            //List<string> idlist = new List<string>();
            //foreach (Match matches in collection)
            //{
            //    idlist.Add(IdRegex.Match(matches.Value).Value);
            //}
            //idlist.Sort(compare);//排序
            //foreach (string id in idlist)//选择答案
            //{
            //    JumpToEle(id, "Id", 0);
            //    Thread.Sleep(100);
            //    Click(id, "Id", 0);
            //}
            //JumpToEle("/html/body/div[2]/div[8]/input[1]", "XPath", 0, true);//跳转到交卷按钮
            //Click("/html/body/div[2]/div[8]/input[1]", "XPath", 0);
            //JumpToEle("/html/body/div[2]/div[7]/div[2]/input[2]", "XPath", 0, true);//跳转到继续按钮
            //Thread.Sleep(100);
            //Click("/html/body/div[2]/div[7]/div[2]/input[2]", "XPath", 0);
            Regex IdRegex = new Regex("id=\"question_[\\d+]+\"");
            MatchCollection collection = IdRegex.Matches(selenium.PageSource);
            Random rd = new Random();
            int quesnum = 0;
            foreach (Match matchmes in collection)
            {
                quesnum++;
            }
            string quesid = "div#question_";
            string ids = "";
            string question_text = "";
            IWebElement quesdiv = null;
            IWebElement ansdiv = null;
            //分数随机。。。
            DateTime beforDT = System.DateTime.Now;  
            for (int ii = 1; ii <= quesnum; ii++)
            {
                Thread.Sleep(rd.Next(1000,3000));
                bool isChecked = false;
                ids = quesid + ii;
                quesdiv = selenium.FindElement(By.CssSelector(ids));
                string questionHtml = (string)js.ExecuteScript("return arguments[0].innerHTML;", quesdiv);
                ansdiv = selenium.FindElement(By.CssSelector(ids + " + div"));
                Regex qtRegex = new Regex("input type=\"text\" onfocus=");
                bool ques_type = qtRegex.IsMatch(questionHtml);
                if (!ques_type)
                {
                    SQLiteConnection m_dbConnection = new SQLiteConnection("Data Source =aa.db; Version = 3; New = false; ");
                    m_dbConnection.Open();
                    Regex Mp3Regex = new Regex("[a-zA-Z]+.MP3");
                    Regex EtoCRegex = new Regex("[a-zA-Z]+");
                    Regex aRegex = new Regex("<a href=");
                    Regex bRegex = new Regex("[\u4e00-\u9fa5]");
                    Regex AnswerIdRegex = new Regex("id=\"question_[\\d+]+_[\\d+]\"");
                    bool bQuestionType;
                    if (aRegex.IsMatch(questionHtml))
                    {
                        question_text = Mp3Regex.Match(questionHtml).Value.Replace(".MP3", "");
                        bQuestionType = true;
                    }
                    else if (!bRegex.IsMatch(questionHtml))
                    {
                        question_text = EtoCRegex.Match(quesdiv.GetAttribute("innerHTML")).Value;
                        bQuestionType = true;
                    }
                    else
                    {
                        string rep = ii + ".";
                        string temp = quesdiv.Text;
                        question_text = temp.Replace(rep, "").Replace(" ", "");
                        bQuestionType = false;
                    }

                    if (bQuestionType)
                    {
                        string eword = "";
                        string answerid = "question_";
                        string ans = "";
                        string innerHtml2 = (string)js.ExecuteScript("return arguments[0].innerHTML;", ansdiv);
                        MatchCollection collection2 = AnswerIdRegex.Matches(innerHtml2);
                        int ansids = 0;
                        foreach (Match m2 in collection2)
                        {
                            ansids++;
                        }
                        for (int j = 0; j < ansids; j++)
                        {
                            ans = answerid + ii + "_" + j;
                            string anstring = ansdiv.FindElement(By.Id(ans)).GetAttribute("value").Replace(" ", "");
                            string sql = "select * from words where cword = '" + anstring + "'";
                            SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
                            SQLiteDataReader reader = command.ExecuteReader();
                            if (reader.HasRows)
                            {

                                while (reader.Read())
                                {
                                    eword = reader["eword"].ToString();
                                    if (question_text == eword)
                                    {
                                        ansdiv.FindElement(By.Id(ans)).Click();
                                        isChecked = true;
                                    }
                                }
                            }
                            //else
                            //{
                            //    MessageBox.Show("未知单词" + question_text, "系统提示");
                            //}

                        }
                        m_dbConnection.Close();

                    }
                    else
                    {
                        string eword = "";
                        string answerid = "question_";
                        string ans = "";
                        string sql = "select * from words where cword = '" + question_text + "'";
                        SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
                        SQLiteDataReader reader = command.ExecuteReader();
                        if (reader.HasRows)
                        {

                            while (reader.Read())
                            {
                                eword = reader["eword"].ToString();
                            }
                        }
                        //else
                        //{
                        //    MessageBox.Show("未知单词" + question_text, "系统提示");
                        //}
                        string innerHtml2 = (string)js.ExecuteScript("return arguments[0].innerHTML;", ansdiv);

                        MatchCollection collection2 = AnswerIdRegex.Matches(innerHtml2);
                        int ansids = 0;
                        foreach (Match m2 in collection2)
                        {
                            ansids++;
                        }
                        for (int j = 0; j < ansids; j++)
                        {
                            ans = answerid + ii + "_" + j;
                            string anstring = ansdiv.FindElement(By.Id(ans)).GetAttribute("value");
                            if (anstring == eword)
                            {
                                ansdiv.FindElement(By.Id(ans)).Click();
                                isChecked = true;
                            }
                        }
                        m_dbConnection.Close();
                    }
                    if (!isChecked)
                    {
                        string ans = "question_" + ii + "_" + 1;
                        ansdiv.FindElement(By.Id(ans)).Click();
                    }

                }
                else
                {
                    tiankongti(ii);
                }

            }
            JumpToEle("/html/body/div[2]/div[8]/input[1]", "XPath", 0, true);//跳转到交卷按钮
            Click("/html/body/div[2]/div[8]/input[1]", "XPath", 0);
            JumpToEle("/html/body/div[2]/div[7]/div[2]/input[2]", "XPath", 0, true);//跳转到继续按钮
            Thread.Sleep(100);
            Click("/html/body/div[2]/div[7]/div[2]/input[2]", "XPath", 0);

            return "完成";
        }
        #endregion
        public static string Study(string perurl)
        {
            selenium.Navigate().GoToUrl("http://www.wxwords.com" + perurl);
            Thread.Sleep(5000);
            if (selenium.Title.IndexOf("智学") == -1)
            {
                return "无法跳转主页面";
            }
            Regex Reg = new Regex("<a href=\"javascript:to.+?</a>", RegexOptions.Singleline);
            Regex UnitReg = new Regex("Unit\\d+", RegexOptions.Singleline);
            Regex ProgramReg = new Regex("wordBrowse.html\\?programName=[^&]+", RegexOptions.Singleline);
            string proname = ProgramReg.Match(selenium.PageSource).Value.Replace("wordBrowse.html?programName=", "");
            if (proname == "")
            {
                return "获取内容失败";
            }
            MatchCollection match = Reg.Matches(selenium.PageSource);
            List<string> list = new List<string>();
            foreach (Match i in match)//获取未学习单元
            {
                if (i.Value.IndexOf("闯关") != -1)
                {
                    continue;
                }
                if (UnitReg.IsMatch(i.Value))
                {
                    list.Add(UnitReg.Match(i.Value).Value);
                }
            }
            //http://www.wxwords.com/student/memory.html?programName=CN-Level62&unitName=Unit3&isReview=false
            foreach (string i in list)
            {
                selenium.Navigate().GoToUrl("http://www.wxwords.com" + perurl.Replace("space.html?", "memory.html?") + "&unitName=" + i + "&isReview=false");
                Thread.Sleep(1500);
                if (IsExist())
                {
                    while (!ExistAlert())
                    {
                        Click("known", "Id", 0);
                        Thread.Sleep(100);
                        Click("right", "Id", 0);
                        Thread.Sleep(100);
                    }
                    selenium.SwitchTo().Alert().Accept();
                    Thread.Sleep(2000);
                }
                else
                {
                    selenium.Navigate().GoToUrl("http://www.wxwords.com" + perurl);
                }
            }

            return "完成";
        }

        public static bool IsExist()
        {
            IWebElement known = selenium.FindElement(By.Id("known"));
            IWebElement right = selenium.FindElement(By.Id("right"));
            if (known != null || right != null) { return true; }
            else
            {
                return false;
            }
        }
        //单元测试
        public static string UnitTest(string perurl,string account)
        {
            selenium.Navigate().GoToUrl("http://www.wxwords.com" + perurl);
            Thread.Sleep(5000);
            if (selenium.Title.IndexOf("智学") == -1)
            {
                return "无法跳转主页面";
            }
            Regex Reg = new Regex("<a href=\"javascript:to.+?</a>", RegexOptions.Singleline);
            Regex UnitReg = new Regex("Unit\\d+", RegexOptions.Singleline);
            Regex ProgramReg = new Regex("wordBrowse.html\\?programName=[^&]+", RegexOptions.Singleline);
            string proname = ProgramReg.Match(selenium.PageSource).Value.Replace("wordBrowse.html?programName=", "");
            if (proname == "")
            {
                return "获取内容失败";
            }
            MatchCollection match = Reg.Matches(selenium.PageSource);
            List<string> list = new List<string>();
            foreach (Match i in match)//获取未测试单元
            {
                if (i.Value.IndexOf("成功") != -1)
                {
                    continue;
                }
                if (i.Value.IndexOf("学习中") != -1)
                {
                    continue;
                }
                if (UnitReg.IsMatch(i.Value))
                {
                    list.Add(UnitReg.Match(i.Value).Value);
                }
            }
            foreach (string i in list)
            {
                selenium.Navigate().GoToUrl("http://www.wxwords.com" + perurl.Replace("space.html?", "quiz.html?") + "&unitName=" + i + "&quizType=8");
                Thread.Sleep(5000);
                IJavaScriptExecutor js = (IJavaScriptExecutor)selenium;
                js.ExecuteScript("$(\"#footer\").remove(\".foot\");");
                //Regex RightRegex = new Regex("<input id=\"question_\\d+_\\d\" type=\"radio\" value=\"true\".+?>");
                //Regex IdRegex = new Regex("question_\\d+_\\d");
                //MatchCollection collection = RightRegex.Matches(selenium.PageSource);
                //List<string> idlist = new List<string>();
                //foreach (Match matches in collection)
                //{
                //    idlist.Add(IdRegex.Match(matches.Value).Value);
                //}
                //idlist.Sort(compare);//排序
                //foreach (string id in idlist)//选择答案
                //{
                //    JumpToEle(id, "Id", 0);
                //    Thread.Sleep(100);
                //    Click(id, "Id", 0);
                //}
                Regex IdRegex = new Regex("id=\"question_[\\d+]+\"");
                MatchCollection collection = IdRegex.Matches(selenium.PageSource);
                Random rd = new Random();
                int quesnum = 0;
                foreach (Match matchmes in collection)
                {
                    quesnum++;
                }        
                string quesid = "div#question_";
                string ids = "";
                string question_text = "";
                IWebElement quesdiv = null;
                IWebElement ansdiv = null;
                //分数随机。。。
                DateTime beforDT = System.DateTime.Now;   
                for (int ii = 1; ii <= quesnum; ii++)
                {
                    Thread.Sleep(rd.Next(1000,3000));
                    bool isChecked = false;
                    ids = quesid + ii;
                    quesdiv = selenium.FindElement(By.CssSelector(ids));
                    string questionHtml = (string)js.ExecuteScript("return arguments[0].innerHTML;", quesdiv);
                    ansdiv = selenium.FindElement(By.CssSelector(ids + " + div"));                    
                    Regex qtRegex = new Regex("input type=\"text\" onfocus=");
                    bool ques_type = qtRegex.IsMatch(questionHtml);
                    if (!ques_type)
                    {
                        SQLiteConnection m_dbConnection = new SQLiteConnection("Data Source =aa.db; Version = 3; New = false; ");
                        m_dbConnection.Open();
                        Regex Mp3Regex = new Regex("[a-zA-Z]+.MP3");
                        Regex EtoCRegex = new Regex("[a-zA-Z]+");
                        Regex aRegex = new Regex("<a href=");
                        Regex bRegex = new Regex("[\u4e00-\u9fa5]");
                        Regex AnswerIdRegex = new Regex("id=\"question_[\\d+]+_[\\d+]\"");
                        bool bQuestionType;
                        if (aRegex.IsMatch(questionHtml))
                        {
                            question_text = Mp3Regex.Match(questionHtml).Value.Replace(".MP3", "");
                            bQuestionType = true;
                        }
                        else if (!bRegex.IsMatch(questionHtml))
                        {
                            question_text = EtoCRegex.Match(quesdiv.GetAttribute("innerHTML")).Value;
                            bQuestionType = true;
                        }
                        else
                        {
                            string rep = ii + ".";
                            string temp = quesdiv.Text;
                            question_text = temp.Replace(rep,"").Replace(" ","");
                            bQuestionType = false;
                        }

                        if (bQuestionType)
                        {
                            string eword = "";
                            string answerid = "question_";
                            string ans = "";                            
                            string innerHtml2 = (string)js.ExecuteScript("return arguments[0].innerHTML;", ansdiv);
                            MatchCollection collection2 = AnswerIdRegex.Matches(innerHtml2);
                            int ansids = 0;
                            foreach (Match m2 in collection2)
                            {
                                ansids++;
                            }
                            for (int j = 0; j < ansids; j++)
                            {
                                ans = answerid + ii + "_" + j;
                                string anstring = ansdiv.FindElement(By.Id(ans)).GetAttribute("value").Replace(" ","");
                                string sql = "select * from words where cword = '" + anstring + "'";
                                SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
                                SQLiteDataReader reader = command.ExecuteReader();
                                if (reader.HasRows)
                                {

                                    while (reader.Read())
                                    {
                                        eword = reader["eword"].ToString();
                                        if (question_text == eword)
                                        {
                                            ansdiv.FindElement(By.Id(ans)).Click();
                                            isChecked = true;
                                        }
                                    }
                                }
                                //else
                                //{
                                //    MessageBox.Show("未知单词" + question_text, "系统提示");
                                //}
                                
                            }
                            m_dbConnection.Close();

                        }
                        else
                        {
                            string eword = "";
                            string answerid = "question_";
                            string ans = "";
                            string sql = "select * from words where cword = '" + question_text + "'";
                            SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
                            SQLiteDataReader reader = command.ExecuteReader();
                            if (reader.HasRows)
                            {

                                while (reader.Read())
                                {
                                    eword = reader["eword"].ToString();
                                }
                            }
                            //else
                            //{
                            //    MessageBox.Show("未知单词" + question_text, "系统提示");
                            //}
                            string innerHtml2 = (string)js.ExecuteScript("return arguments[0].innerHTML;", ansdiv);

                            MatchCollection collection2 = AnswerIdRegex.Matches(innerHtml2);
                            int ansids = 0;
                            foreach (Match m2 in collection2)
                            {
                                ansids++;
                            }
                            for (int j = 0; j < ansids; j++)
                            {
                                ans = answerid + ii + "_" + j;
                                string anstring = ansdiv.FindElement(By.Id(ans)).GetAttribute("value");
                                if (anstring == eword)
                                {
                                    ansdiv.FindElement(By.Id(ans)).Click();
                                    isChecked = true;
                                }
                            }
                            m_dbConnection.Close();
                        }
                        if (!isChecked) {
                            string ans = "question_" + ii + "_" + 1;
                            ansdiv.FindElement(By.Id(ans)).Click();
                        }
                        
                    }
                    else
                    {
                        tiankongti(ii);
                    }

                }

                DateTime afterDT = System.DateTime.Now;
                JumpToEle("/html/body/div[2]/div[8]/input[1]", "XPath", 0, true);//跳转到交卷按钮
                Click("/html/body/div[2]/div[8]/input[1]", "XPath", 0);
                if (ExistAlert())
                {
                    selenium.SwitchTo().Alert().Accept();
                }
                TimeSpan ts = afterDT.Subtract(beforDT);
                string alltime = ts.TotalSeconds.ToString();
                Thread.Sleep(8000);
                string testName = selenium.FindElement(By.Id("testPaperName")).Text;
                string Score = selenium.FindElement(By.CssSelector("div#quizResultTip>span")).Text;

                SQLiteConnection m_dbConnection2 = new SQLiteConnection("Data Source =cc.db; Version = 3; New = false; ");
                m_dbConnection2.Open();
                string sql11 = "insert into logs values('" + account + "','" + testName + "','" + Score + "分','" + alltime + "秒');";
                SQLiteCommand command2 = new SQLiteCommand(sql11, m_dbConnection2);
                command2.ExecuteNonQuery();
                m_dbConnection2.Close();
                JumpToEle("/html/body/div[2]/div[7]/div[2]/input[2]", "XPath", 0, true);//跳转到继续按钮
                Thread.Sleep(100);
                Click("/html/body/div[2]/div[7]/div[2]/input[2]", "XPath", 0);
                Thread.Sleep(2500);
            }

            return "完成";
        }
        //组卷测试
        public static string TeachTest(int iScore,int timert,string account)
        {
                    
            selenium.Navigate().GoToUrl("http://www.wxwords.com/student/space.html");
            Thread.Sleep(5000);
            Regex Reg = new Regex("toTeacherQuiz\\(\\d+");
            if (!Reg.IsMatch(selenium.PageSource))
            {
                return "无组卷测试";
            }
            DateTime beforDT = System.DateTime.Now;    
            MatchCollection matches = Reg.Matches(selenium.PageSource);
            List<string> idlist = new List<string>();
            foreach (Match i in matches)//获取测试列表
            {
                idlist.Add(i.Value.Replace("toTeacherQuiz(", ""));
            }
            foreach (string i in idlist)
            {
                selenium.Navigate().GoToUrl("http://www.wxwords.com/student/quiz.html?testPaperId=" + i + "&quizType=14");
                Thread.Sleep(5000);
                IJavaScriptExecutor js = (IJavaScriptExecutor)selenium;
                js.ExecuteScript("$(\"#footer\").remove(\".foot\");");
                Regex IdRegex = new Regex("id=\"question_[\\d+]+\"");
                MatchCollection collection = IdRegex.Matches(selenium.PageSource);
                Random rd = new Random();
                int quesnum = 0;
                foreach (Match matchmes in collection)
                {
                    quesnum++;
                }        
                string quesid = "div#question_";
                string ids = "";
                string question_text = "";
                IWebElement quesdiv = null;
                IWebElement ansdiv = null;
                //分数随机。。。
                
                for (int ii = 1; ii <= quesnum; ii++)
                {
                    Thread.Sleep(rd.Next(timert,timert+1000));
                    bool isChecked = false;
                    ids = quesid + ii;
                    quesdiv = selenium.FindElement(By.CssSelector(ids));
                    string questionHtml = (string)js.ExecuteScript("return arguments[0].innerHTML;", quesdiv);
                    ansdiv = selenium.FindElement(By.CssSelector(ids + " + div"));                    
                    Regex qtRegex = new Regex("input type=\"text\" onfocus=");
                    bool ques_type = qtRegex.IsMatch(questionHtml);
                    if (!ques_type)
                    {
                        SQLiteConnection m_dbConnection = new SQLiteConnection("Data Source =aa.db; Version = 3; New = false; ");
                        m_dbConnection.Open();
                        Regex Mp3Regex = new Regex("[a-zA-Z]+.MP3");
                        Regex EtoCRegex = new Regex("[a-zA-Z]+");
                        Regex aRegex = new Regex("<a href=");
                        Regex bRegex = new Regex("[\u4e00-\u9fa5]");
                        Regex AnswerIdRegex = new Regex("id=\"question_[\\d+]+_[\\d+]\"");
                        bool bQuestionType;
                        if (aRegex.IsMatch(questionHtml))
                        {
                            question_text = Mp3Regex.Match(questionHtml).Value.Replace(".MP3", "");
                            bQuestionType = true;
                        }
                        else if (!bRegex.IsMatch(questionHtml))
                        {
                            question_text = EtoCRegex.Match(quesdiv.GetAttribute("innerHTML")).Value;
                            bQuestionType = true;
                        }
                        else
                        {
                            string rep = ii + ".";
                            string temp = quesdiv.Text;
                            question_text = temp.Replace(rep,"").Replace(" ","");
                            bQuestionType = false;
                        }

                        if (bQuestionType)
                        {
                            string eword = "";
                            string answerid = "question_";
                            string ans = "";                            
                            string innerHtml2 = (string)js.ExecuteScript("return arguments[0].innerHTML;", ansdiv);
                            MatchCollection collection2 = AnswerIdRegex.Matches(innerHtml2);
                            int ansids = 0;
                            foreach (Match m2 in collection2)
                            {
                                ansids++;
                            }
                            for (int j = 0; j < ansids; j++)
                            {
                                ans = answerid + ii + "_" + j;
                                string anstring = ansdiv.FindElement(By.Id(ans)).GetAttribute("value").Replace(" ","");
                                string sql = "select * from words where cword = '" + anstring + "'";
                                SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
                                SQLiteDataReader reader = command.ExecuteReader();
                                if (reader.HasRows)
                                {

                                    while (reader.Read())
                                    {
                                        eword = reader["eword"].ToString();
                                        if (question_text == eword)
                                        {
                                            ansdiv.FindElement(By.Id(ans)).Click();
                                            isChecked = true;
                                        }
                                    }
                                }
                                //else
                                //{
                                //    MessageBox.Show("未知单词" + question_text, "系统提示");
                                //}
                                
                            }
                            m_dbConnection.Close();

                        }
                        else
                        {
                            string eword = "";
                            string answerid = "question_";
                            string ans = "";
                            string sql = "select * from words where cword = '" + question_text + "'";
                            SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
                            SQLiteDataReader reader = command.ExecuteReader();
                            if (reader.HasRows)
                            {

                                while (reader.Read())
                                {
                                    eword = reader["eword"].ToString();
                                }
                            }
                            //else
                            //{
                            //    MessageBox.Show("未知单词" + question_text, "系统提示");
                            //}
                            string innerHtml2 = (string)js.ExecuteScript("return arguments[0].innerHTML;", ansdiv);

                            MatchCollection collection2 = AnswerIdRegex.Matches(innerHtml2);
                            int ansids = 0;
                            foreach (Match m2 in collection2)
                            {
                                ansids++;
                            }
                            for (int j = 0; j < ansids; j++)
                            {
                                ans = answerid + ii + "_" + j;
                                string anstring = ansdiv.FindElement(By.Id(ans)).GetAttribute("value");
                                if (anstring == eword)
                                {
                                    ansdiv.FindElement(By.Id(ans)).Click();
                                    isChecked = true;
                                }
                            }
                            m_dbConnection.Close();
                        }
                        if (!isChecked) {
                            string ans = "question_" + ii + "_" + 1;
                            ansdiv.FindElement(By.Id(ans)).Click();
                        }
                        
                    }
                    else
                    {
                        tiankongti(ii);
                    }

                }

                int s = 100 - iScore;
                for (int k = 0; k < s; k++) {
                    int kk = k + 2;
                    ids = "div#question_"+kk;
                    ansdiv = selenium.FindElement(By.CssSelector(ids + " + div"));
                    string ans = "question_" + kk + "_" + 1;
                    ansdiv.FindElement(By.Id(ans)).Click();                    
                }

                DateTime afterDT = System.DateTime.Now;
                JumpToEle("/html/body/div[2]/div[8]/input[1]", "XPath", 0, true);//跳转到交卷按钮
                Click("/html/body/div[2]/div[8]/input[1]", "XPath", 0);
                //<div class="fen" id="quizResultTip" style="display: block;">
                //<span style="font-size:24px;color:red;">100</span>分，赞！VERY GOOD！记得学而时习之哦！</div>
                //<div class="ceshi" id="testPaperName">组卷测试 - 六级核心词汇二6-10单元--测验四</div>
                //日志 账号   组卷测试名称 分数   测试用时 
                TimeSpan ts = afterDT.Subtract(beforDT);
                string alltime = ts.TotalSeconds.ToString();
                Thread.Sleep(8000);
                string testName = selenium.FindElement(By.Id("testPaperName")).Text;
                string Score = selenium.FindElement(By.CssSelector("div#quizResultTip>span")).Text;

                SQLiteConnection m_dbConnection2 = new SQLiteConnection("Data Source =cc.db; Version = 3; New = false; ");
                m_dbConnection2.Open();
                string sql11 = "insert into logs values('"+account+"','"+testName+"','"+Score+"分','"+alltime+"秒');";
                SQLiteCommand command2 = new SQLiteCommand(sql11, m_dbConnection2);
                command2.ExecuteNonQuery();
                m_dbConnection2.Close();
                //insert into logs values (account,testName,Score,alltime);                
                
                
                //   JumpToEle("/html/body/div[2]/div[7]/div[2]/input[2]", "XPath", 0, true);//跳转到继续按钮
                //   Click("/html/body/div[2]/div[7]/div[2]/input[2]", "XPath", 0);
            }
            return "完成";
        }

        public static bool ExistElement(string name, string type, int index)
        {
            try
            {
                Actions act = new Actions(selenium);
                IWebElement webElement = FindElement(name, type, index);
                act.MoveToElement(FindElement(name, type, index));
            }
            catch
            {
                return false;
            }
            return true;
        }

        public static void JumpToEle(string name, string type, int index, bool needspace = false)
        {
            Actions act = new Actions(selenium);
            IWebElement webElement = FindElement(name, type, index);
            if (webElement != null)
            {
                act.MoveToElement(FindElement(name, type, index));
                if (needspace)
                {
                    act.SendKeys(OpenQA.Selenium.Keys.Space);
                }
                act.Perform();
            }
        }



        public static int compare(string str1, string str2)
        {
            Regex idnum = new Regex("\\d+");
            int num1 = Convert.ToInt32(idnum.Match(str1).Value);
            int num2 = Convert.ToInt32(idnum.Match(str2).Value);

            if (num1 > num2)
            {
                return 1;
            }
            else if (num1 < num2)
            {
                return -1;
            }
            else
            {
                return 0;
            }
        }

        public static IWebElement FindElement(string idname, string type, int index)
        {
            try
            {
                if (type == "ClassName")
                {
                    return selenium.FindElements(By.ClassName(idname))[index];
                }
                else if (type == "CssSelector")
                {
                    return selenium.FindElements(By.CssSelector(idname))[index];
                }
                else if (type == "Id")
                {
                    return selenium.FindElements(By.Id(idname))[index];
                }
                else if (type == "LinkText")
                {
                    return selenium.FindElements(By.LinkText(idname))[index];
                }
                else if (type == "Name")
                {
                    return selenium.FindElements(By.Name(idname))[index];
                }
                else if (type == "PartialLinkText")
                {
                    return selenium.FindElements(By.PartialLinkText(idname))[index];
                }
                else if (type == "TagName")
                {
                    return selenium.FindElements(By.TagName(idname))[index];
                }
                else if (type == "XPath")
                {
                    return selenium.FindElements(By.XPath(idname))[index];
                }
                else if (type == "iframe")
                {
                    IList<IWebElement> frames = selenium.FindElements(By.TagName("iframe"));
                    IWebElement controlPanelFrame = null;
                    foreach (var frame in frames)
                    {
                        if (frame.GetAttribute("id") == idname)
                        {
                            controlPanelFrame = frame;
                            break;
                        }
                    }

                    return controlPanelFrame;
                }
            }
            catch (Exception e)
            {
                return null;
            }

            return null;
        }

        public static bool Click(string idname, string type, int index)
        {
            try
            {
                if (type == "ClassName")
                {
                    selenium.FindElements(By.ClassName(idname))[index].Click();
                }
                else if (type == "CssSelector")
                {
                    selenium.FindElements(By.CssSelector(idname))[index].Click();
                }
                else if (type == "Id")
                {
                    selenium.FindElements(By.Id(idname))[index].Click();
                }
                else if (type == "LinkText")
                {
                    selenium.FindElements(By.LinkText(idname))[index].Click();
                }
                else if (type == "Name")
                {
                    selenium.FindElements(By.Name(idname))[index].Click();
                }
                else if (type == "PartialLinkText")
                {
                    selenium.FindElements(By.PartialLinkText(idname))[index].Click();
                }
                else if (type == "TagName")
                {
                    selenium.FindElements(By.TagName(idname))[index].Click();
                }
                else if (type == "XPath")
                {
                    selenium.FindElements(By.XPath(idname))[index].Click();
                }
                else if (type == "iframe")
                {
                    IList<IWebElement> frames = selenium.FindElements(By.TagName("iframe"));
                    IWebElement controlPanelFrame = null;
                    foreach (var frame in frames)
                    {
                        if (frame.GetAttribute("id") == idname)
                        {
                            controlPanelFrame = frame;
                            break;
                        }
                    }

                    if (controlPanelFrame != null)
                    {
                        selenium.SwitchTo().Frame(controlPanelFrame);
                    }
                }
            }
            catch (Exception e)
            {
                return false;
            }

            return true;

        }

        public static void SendKey(string idname, string type, int index, string text)
        {
            try
            {
                if (type == "ClassName")
                {
                    selenium.FindElements(By.ClassName(idname))[index].SendKeys(text);
                }
                else if (type == "CssSelector")
                {
                    selenium.FindElements(By.CssSelector(idname))[index].SendKeys(text);
                }
                else if (type == "Id")
                {
                    selenium.FindElements(By.Id(idname))[index].SendKeys(text);
                }
                else if (type == "LinkText")
                {
                    selenium.FindElements(By.LinkText(idname))[index].SendKeys(text);
                }
                else if (type == "Name")
                {
                    selenium.FindElements(By.Name(idname))[index].SendKeys(text);
                }
                else if (type == "PartialLinkText")
                {
                    selenium.FindElements(By.PartialLinkText(idname))[index].SendKeys(text);
                }
                else if (type == "TagName")
                {
                    selenium.FindElements(By.TagName(idname))[index].SendKeys(text);
                }
                else if (type == "XPath")
                {
                    selenium.FindElements(By.XPath(idname))[index].SendKeys(text);
                }
                else if (type == "iframe")
                {
                    IList<IWebElement> frames = selenium.FindElements(By.TagName("iframe"));
                    IWebElement controlPanelFrame = null;
                    foreach (var frame in frames)
                    {
                        if (frame.GetAttribute("id") == idname)
                        {
                            controlPanelFrame = frame;
                            break;
                        }
                    }

                    if (controlPanelFrame != null)
                    {
                        selenium.SwitchTo().Frame(controlPanelFrame);
                    }
                }
            }
            catch (Exception e)
            {

            }
        }

        public static void tiankongti(int i)
        {
            string quesid = "div#question_";
            string ids = quesid + i;
            Regex Mp3Regex = new Regex("[a-zA-Z]+.MP3");
            IJavaScriptExecutor divjs = selenium as IJavaScriptExecutor;
            IWebElement quesdiv = selenium.FindElement(By.CssSelector(ids));
            IWebElement toinput = quesdiv.FindElement(By.TagName("input"));
            string innerHtml = (string)divjs.ExecuteScript("return arguments[0].innerHTML;", quesdiv);
            string question_sound = Mp3Regex.Match(innerHtml).Value.Replace(".MP3", "");
            toinput.SendKeys(question_sound);
        }

    }
}
