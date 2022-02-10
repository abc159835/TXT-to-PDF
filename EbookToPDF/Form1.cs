using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Web;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Configuration;
using System.Diagnostics;
using System.Drawing.Text;
using System.Threading;

namespace EbookToPDF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        Process process = new Process();
        string exePath = AppDomain.CurrentDomain.BaseDirectory.ToString();
        private void button1_Click(object sender, EventArgs e)
        {
            TextBox[] textBoxes = { textBox1, textBox2, textBox3, textBox4 };
            foreach(var a in textBoxes)
            {
                if (a.Text.Trim().Length <= 1)
                {
                    MessageBox.Show("请填写完整!");
                    return;
                }
            }
            Dictionary<string, List<string>> book = new Dictionary<string, List<string>>();
            StreamReader sr = new StreamReader(new FileStream(textBox1.Text, FileMode.Open), Encoding.UTF8);
            string line;
            string title = "";
            while ((line = sr.ReadLine()) != null)
            {
                int b = line.IndexOf("第");
                int c = line.IndexOf("章");
                if (line.Trim() == "")
                    continue;
                else if ((c - b) > 0 && (c - b) < 8) 
                {
                    if (!book.ContainsKey(line.Trim()))
                    {
                        book.Add(line.Trim(), new List<string>());
                        title = line.Trim();
                    }
                }
                else if (title != "")
                {
                    book[title].Add(line.Trim());
                }
            }
            sr.Close();
            txttoPDF(book);
        }
        private void txttoPDF(Dictionary<string, List<string>> book)
        {
            string head = File.ReadAllText("css.html").Replace("titlefont", comboBox1.Text)
                .Replace("textfont", comboBox2.Text).Replace("bookname", textBox2.Text);
            StringBuilder stringBuilder = new StringBuilder();
            string[] toP;
            if (textBox4.Text.Contains('+'))
                toP = textBox4.Text.Split('+');
            else
                toP = new string[] { textBox4.Text };
            for (int i = 0; i < 2; i++)
            {
                stringBuilder.AppendLine("<div class=\"center\">");
                for (int ii = 0; ii < 12; ii++)
                    stringBuilder.AppendLine("<br>");
                stringBuilder.AppendLine("<p style=\"font-size: 4em\">" + textBox2.Text + "</p>");
                stringBuilder.AppendLine("<p style=\"font-size: 1.5em\">" + textBox3.Text + " 著</p>");
                stringBuilder.AppendLine("<p style=\"font-size: 1em\">" + toP[i] + "</p>");
                stringBuilder.AppendLine("</div>");
            }
            foreach (var key in book)
            {
                say(key.Key + "--已获取");
                if(checkBox1.Checked)
                    stringBuilder.AppendLine("<div class=\"NO_next\">");
                else
                    stringBuilder.AppendLine("<div class=\"next\">");
                stringBuilder.AppendLine("<br><h1 id=\"title\" class=\"title\">" + key.Key + "</h1>");
                foreach (var line in key.Value)
                {
                    if(line.Contains("【")&& line.Contains("】"))
                        stringBuilder.AppendLine("<p class=\"p_title\">　　" + line + "</p>");
                    else
                        stringBuilder.AppendLine("<p class=\"p_text\">　　" + line + "</p>");
                }
                if (checkBox1.Checked)
                    for (int i = 0; i < 4; i++)
                        stringBuilder.AppendLine("<br>");
                stringBuilder.AppendLine("</div>");
            }
            stringBuilder.AppendLine("</body></html >");
            StreamWriter sw = new StreamWriter(new FileStream(textBox2.Text + ".html", FileMode.Create));
            sw.Write(head);
            sw.Write(stringBuilder.ToString());
            sw.Close();
            say("转换中,需要较长时间...");
            HtmlTextConvertToPdf(head + stringBuilder.ToString(), exePath + textBox2.Text + ".pdf");
            say("成功!");
            save();
        }
        /// <summary>
        /// 获取命令行参数
        /// </summary>
        /// <param name="htmlPath"></param>
        /// <param name="savePath"></param>
        /// <returns></returns>
        private string GetArguments(string htmlPath, string savePath)
        {
            if (string.IsNullOrEmpty(htmlPath))
            {
                throw new Exception("HTML local path or network address can not be empty.");
            }
            if (string.IsNullOrEmpty(savePath))
            {
                throw new Exception("The path saved by the PDF document can not be empty.");
            }
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.Append(" --page-height 210 ");        //页面高度100mm
            stringBuilder.Append(" --page-width 150 ");         //页面宽度100mm
            stringBuilder.Append(" --zoom 1 ");
            //stringBuilder.Append(" --header-center 我是页眉 ");  //设置居中显示页眉
            stringBuilder.Append(" --header-line ");         //页眉和内容之间显示一条直线
            stringBuilder.Append(" --footer-font-size 8 ");
            stringBuilder.Append(" --header-right [section] ");
            stringBuilder.Append(" --header-font-size 10 ");
            stringBuilder.Append(" --footer-center \"- [page] -\" ");
            //stringBuilder.Append(" --footer-line ");       //页脚和内容之间显示一条直线
            stringBuilder.Append(" " + htmlPath + " ");       //本地 HTML 的文件路径或网页 HTML 的URL地址
            stringBuilder.Append(" " + savePath + " ");       //生成的 PDF 文档的保存路径
            return stringBuilder.ToString();
        }

        /// <summary>
        /// HTML文本内容转换为PDF
        /// </summary>
        /// <param name="strHtml">HTML文本内容</param>
        /// <param name="savePath">PDF文件保存的路径</param>
        /// <returns></returns>
        public void HtmlTextConvertToPdf(string strHtml, string savePath)
        {
            string htmlPath = HtmlTextConvertFile(strHtml);
            HtmlConvertToPdf(htmlPath, savePath);
            File.Delete(htmlPath);
        }

        /// <summary>
        /// HTML转换为PDF
        /// </summary>
        /// <param name="htmlPath">可以是本地路径，也可以是网络地址</param>
        /// <param name="savePath">PDF文件保存的路径</param>
        /// <returns></returns>
        public void HtmlConvertToPdf(string htmlPath, string savePath)
        {
            CheckFilePath(savePath);
            ///这个路径为程序集的目录，因为我把应用程序 wkhtmltopdf.exe 放在了程序集同一个目录下
            string exePath = AppDomain.CurrentDomain.BaseDirectory.ToString() + "wkhtmltopdf.exe";
            if (!File.Exists(exePath))
            {
                throw new Exception("No application wkhtmltopdf.exe was found.");
            }
            string cmd = exePath + " " + GetArguments(htmlPath, savePath)+ "&exit";
            //启动Windows的cmd控制台
            process.StartInfo.FileName = "cmd.exe";
            //启动进程时不使用 shell
            process.StartInfo.UseShellExecute = false;
            //设置标准重定向输入
            process.StartInfo.RedirectStandardInput = true;
            //设置标准重定向输出
            process.StartInfo.RedirectStandardOutput = false;
            //设置标准重定向错误输出
            process.StartInfo.RedirectStandardError = false;
            //设置不显示cmd控制台窗体
            if(!checkBox2.Checked)
                process.StartInfo.CreateNoWindow = false;
            else
                process.StartInfo.CreateNoWindow = true;
            process.Start();
            process.StandardInput.WriteLine(cmd);
            process.WaitForExit();
        }
        /// <summary>
        /// 验证保存路径
        /// </summary>
        /// <param name="savePath"></param>
        private void CheckFilePath(string savePath)
        {
            string ext = string.Empty;
            string path = string.Empty;
            string fileName = string.Empty;

            ext = Path.GetExtension(savePath);
            if (string.IsNullOrEmpty(ext) || ext.ToLower() != ".pdf")
            {
                throw new Exception("Extension error:This method is used to generate PDF files.");
            }

            fileName = Path.GetFileName(savePath);
            if (string.IsNullOrEmpty(fileName))
            {
                throw new Exception("File name is empty.");
            }

            try
            {
                path = savePath.Substring(0, savePath.IndexOf(fileName));
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
            }
            catch
            {
                throw new Exception("The file path does not exist.");
            }
        }

        /// <summary>
        /// HTML文本内容转HTML文件
        /// </summary>
        /// <param name="strHtml">HTML文本内容</param>
        /// <returns>HTML文件的路径</returns>
        public string HtmlTextConvertFile(string strHtml)
        {
            if (string.IsNullOrEmpty(strHtml))
            {
                throw new Exception("HTML text content cannot be empty.");
            }

            try
            {
                string path = AppDomain.CurrentDomain.BaseDirectory.ToString() + @"html\";
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                string fileName = path + DateTime.Now.ToString("yyyyMMddHHmmssfff") + new Random().Next(1000, 10000) + ".html";
                FileStream fileStream = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                StreamWriter streamWriter = new StreamWriter(fileStream, Encoding.Default);
                streamWriter.Write(strHtml);
                streamWriter.Flush();

                streamWriter.Close();
                streamWriter.Dispose();
                fileStream.Close();
                fileStream.Dispose();
                return fileName;
            }
            catch
            {
                throw new Exception("HTML text content error.");
            }
        }
        private void say(string a)
        {
            richTextBox1.AppendText(a + "\n");
            richTextBox1.Focus();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            Control[] vs = { textBox1, textBox2, textBox3, textBox4, comboBox1, comboBox2 };
            var appSettings = ConfigurationManager.AppSettings;
            if (appSettings.Count == 0)
            {
                Console.WriteLine("AppSettings is empty.");
            }
            else
            {
                foreach (var key in appSettings.AllKeys)
                {
                    foreach (var a in vs)
                    {
                        if (key == a.Name)
                            a.Text = appSettings[key];
                    }
                }
            }
            InstalledFontCollection fonts = new InstalledFontCollection();
            foreach (FontFamily family in fonts.Families)
            {
                comboBox1.Items.Add(family.Name);
                comboBox2.Items.Add(family.Name);
            }
        }
        private void save()
        {
            Control[] vs = { textBox1, textBox2, textBox3, textBox4, comboBox1, comboBox2 };
            var configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            var settings = configFile.AppSettings.Settings;
            foreach (var a in vs)
            {
                if (settings[a.Name] == null)
                {
                    settings.Add(a.Name, a.Text);
                }
                else
                {
                    settings[a.Name].Value = a.Text;
                }
            }
            configFile.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name);
        }
    }
}
