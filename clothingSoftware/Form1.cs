using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;

using System.Data.SqlClient;
using System.Drawing.Printing;
using Printers;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace clothingSoftware
{
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public partial class Form1 : Form
    {
        //string OldActPrn;
        Externs Externs = new Externs();
        public BardCodeHooK BarCode = new BardCodeHooK();
        public delegate void ShowInfoDelegate(BardCodeHooK.BarCodes barCode);
        //为了让窗体获取焦点
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, ExactSpelling = true)]
        public static extern IntPtr GetForegroundWindow(); //获得本窗体的句柄
        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "SetForegroundWindow")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);//设置此窗体为活动窗体
        public IntPtr han;
        private void timer1_Tick_1(object sender, EventArgs e)
        {
            // if (han != GetForegroundWindow())
            //   {
            //      SetForegroundWindow(han);
            //  }
        }
        //判断是否是扫码枪输入，是的话执行事件
        void ShowInfo(BardCodeHooK.BarCodes barCode)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new ShowInfoDelegate(ShowInfo), new object[] { barCode });
            }
            else
            {
                //MessageBox.Show(barCode.IsValid.ToString());
                if (barCode.IsValid == true)
                {
                    object[] objects = new object[1];
                    objects[0] = barCode.BarCode;
                    webBrowser1.Document.InvokeScript("scalabel", objects);
                }

            }
        }
        public void BarCode_BarCodeEvent(BardCodeHooK.BarCodes barCode)
        {
            ShowInfo(barCode);
        }
        public Form1()
        {
            InitializeComponent();
            webBrowser1.Navigate("http://clath.dev.idaqi.com/idaqi_new_clothe/views/login.html");
            //webBrowser1.Navigate("http://cloth.dev.idaqi.com/idaqi_new_clothe/views/login.html");
            //webBrowser1.Navigate("http://cloth.test.idaqi.com/cloth/master/views/index.html");
            //webBrowser1.Navigate("http://localhost:8080/views/login.html");
            webBrowser1.ScriptErrorsSuppressed = true; //禁用错误脚本提示   
            webBrowser1.IsWebBrowserContextMenuEnabled = false;
            webBrowser1.AllowWebBrowserDrop = false;
            BarCode.BarCodeEvent += new BardCodeHooK.BardCodeDeletegate(BarCode_BarCodeEvent);
            RunCmd("RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8");

        }
        public void RunCmd(string cmd)
        {
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo.FileName = "cmd.exe";
            // 关闭Shell的使用
            p.StartInfo.UseShellExecute = false;
            // 重定向标准输入
            p.StartInfo.RedirectStandardInput = true;
            // 重定向标准输出
            p.StartInfo.RedirectStandardOutput = true;
            //重定向错误输出
            p.StartInfo.RedirectStandardError = true;
            p.StartInfo.CreateNoWindow = true;
            p.Start();
            p.StandardInput.WriteLine(cmd);
            p.StandardInput.WriteLine("exit");
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            webBrowser1.ObjectForScripting = this;
            han = this.Handle;
            //timer1.Enabled = true;
            BarCode.Start();
            string activeDir = @"D:";
            string newPath = System.IO.Path.Combine(activeDir, "图库");
            System.IO.Directory.CreateDirectory(newPath);
        }
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            BarCode.Stop();
        }
        //C工艺打印机时间
        public void cprintv(string paths,int size)
        {
            for(int i = 0; i < size; i++)
            {
                string paths1 = paths;
                System.Diagnostics.Process.Start(paths1);
                Thread.Sleep(1000);
                SendKeys.SendWait("{ENTER}");
                Thread.Sleep(1000);
                SendKeys.SendWait(@"%+({F4})");
            }
            
        }
        //打印标签函数
        public void printjs(int Printerstaus, string Printername)
        {
            //实例化打印对象
            PrintDocument printDocument1 = new PrintDocument();

            //设置打印用的纸张,可以自定义纸张的大小(单位：mm).   当打印高度不确定时也可以不设置
            printDocument1.DefaultPageSettings.PaperSize = new PaperSize("Custum", 200, 128);

            //注册PrintPage事件，打印每一页时会触发该事件
            printDocument1.PrintPage += new PrintPageEventHandler(this.printDocument1_PrintPage);
            if (Printerstaus == 0)
            {
                object[] objects = new object[1];
                objects[0] = "扫码枪处于空闲状态";
                webBrowser1.Document.InvokeScript("tipError1", objects);
            }
            else if (Printerstaus == 1)
            {
                Externs.SetDefaultPrinter(Printername);
                //if (Externs.SetDefaultPrinter(Printername))
                //{
                    //开始打印
                    printDocument1.Print();
                //}

            }
            else if (Printerstaus == 2)
            {
                object[] objects = new object[1];
                objects[0] = "扫码枪处于扫码分拣状态";
                webBrowser1.Document.InvokeScript("tipError1", objects);

            }
            else if (Printerstaus == 3)
            {
                Externs.SetDefaultPrinter(Printername);
               // if (Externs.SetDefaultPrinter(Printername))
                //{
                    //开始打印
                    printDocument1.Print();
               // }

            }
            //打印预览
            //PrintPreviewDialog ppd = new PrintPreviewDialog();
            //ppd.Document = printDocument1;
            //ppd.ShowDialog();
        }
        //导出打印
        public void printjs2(int Printerstaus, string Printername)
        {
            //实例化打印对象
            PrintDocument printDocument1 = new PrintDocument();

            //设置打印用的纸张,可以自定义纸张的大小(单位：mm).   当打印高度不确定时也可以不设置
            printDocument1.DefaultPageSettings.PaperSize = new PaperSize("Custum", 200, 128);

            //注册PrintPage事件，打印每一页时会触发该事件
            printDocument1.PrintPage += new PrintPageEventHandler(this.printDocument1_PrintPage2);
            Externs.SetDefaultPrinter(Printername);
            //if (Externs.SetDefaultPrinter(Printername))
            //{
                //开始打印
                printDocument1.Print();
           // }

            //打印预览
            //PrintPreviewDialog ppd = new PrintPreviewDialog();
            //ppd.Document = printDocument1;
            //ppd.ShowDialog();
        }
        string BarcodeString = "";//条码
        string BarcodeString2 = "";//条码
        string itemnum;
        string stname;
        string size;
        string num;
        string type;
        string colorname;
        string sizename;
        string stylename;

        string itemnum2;
        string stname2;
        string size2;
        int num2;
        string type2;
        string filepath2;
        double xx2;
        double yy2;
        double angle2;
        string colorname2;
        string sizename2;
        string stylename2;
        int ImgWidth = 200;
        int ImgHeight = 80;


        //为了让html页面可以传参
        public void OpenForm(string s)
        {
            BarcodeString = s;
        }
        public void OpenForm2(string s)
        {
            BarcodeString2 = s;
        }
        public void givesize(string a, string b, string c, string d, string e, string f, string g, string h)
        {
            itemnum = a;
            stname = b;
            size = c;
            num = d;
            type = e;
            colorname = f;
            sizename = g;
            stylename = h;
        }
        public void giveaisize(string a, string b, string c, int d, string e, string g, double x, double y, double angle, string n, string p, string q)
        {
            itemnum2 = a;
            stname2 = b;
            size2 = c;
            num2 = d;
            type2 = e;
            filepath2 = g;
            xx2 = x;
            yy2 = y;
            angle2 = angle;
            colorname2 = n;
            sizename2 = p;
            stylename2 = q;
        }

        //扫码贴标打印事件
        private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(itemnum + "\r\n");
            sb.Append(colorname + stylename + "\r\n");
            sb.Append("尺寸：" + sizename + " 数量：" + num + " \r\n");
            sb.Append("编号：" + BarcodeString);
            DrawPrint(e, sb.ToString(), BarcodeString, ImgWidth, ImgHeight);
        }
        //导出扫码贴标打印事件
        private void printDocument1_PrintPage2(object sender, PrintPageEventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(itemnum2  + "\r\n");
            sb.Append(colorname2 + stylename2 + "\r\n");
            sb.Append("尺寸：" + sizename2 + " 数量：" + num2 + " \r\n");
            sb.Append("编号：" + BarcodeString2);
            DrawPrint(e, sb.ToString(), BarcodeString2, ImgWidth, ImgHeight);
        }
        /// <summary>
        /// 绘制打印内容
        /// </summary>
        /// <param name="e">PrintPageEventArgs</param>
        /// <param name="PrintStr">需要打印的文本</param>
        /// <param name="BarcodeStr">条码</param>
        public void DrawPrint(PrintPageEventArgs e, string PrintStr, string BarcodeStr, int BarcodeWidth, int BarcodeHeight)
        {
            try
            {
                //绘制打印字符串
                e.Graphics.DrawString(PrintStr, new Font(new FontFamily("黑体"), 8), System.Drawing.Brushes.Black, 20, 10);

                if (!string.IsNullOrEmpty(BarcodeStr))
                {
                    int PrintWidth = 200;
                    int PrintHeight = 50;
                    //绘制打印图片
                    e.Graphics.DrawImage(CreateBarcodePicture(BarcodeStr, BarcodeWidth, BarcodeHeight), 0, 70, PrintWidth, PrintHeight);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        /// <summary>
        /// 根据字符串生成条码图片( 需添加引用：BarcodeLib.dll )
        /// </summary>
        /// <param name="BarcodeString">条码字符串</param>
        /// <param name="ImgWidth">图片宽带</param>
        /// <param name="ImgHeight">图片高度</param>
        /// <returns></returns>
        public System.Drawing.Image CreateBarcodePicture(string BarcodeString, int ImgWidth, int ImgHeight)
        {
            BarcodeLib.Barcode b = new BarcodeLib.Barcode();//实例化一个条码对象
            BarcodeLib.TYPE type = BarcodeLib.TYPE.CODE128;//编码类型

            //获取条码图片
            System.Drawing.Image BarcodePicture = b.Encode(type, BarcodeString, System.Drawing.Color.Black, System.Drawing.Color.White, ImgWidth, ImgHeight);

            BarcodePicture.Save(@"D:\Barcode.png");


            b.Dispose();

            return BarcodePicture;
        }
        private void CreateImage(string content)
        {
            //判断字符串不等于空和null
            if (content == null || content.Trim() == String.Empty)
                return;
            //创建一个位图对象
            Image img = System.Drawing.Image.FromFile(@"D:\Barcode.png");
            //Bitmap image = new Bitmap((int)Math.Ceiling((content.Length * 18.0)), 30);
            Bitmap image = new Bitmap(160, 110);
            //创建Graphics
            Graphics g = Graphics.FromImage(image);
            try
            {
                //清空图片背景颜色
                g.Clear(Color.White);
                Font font = new Font("Arial", 10f, (FontStyle.Bold));
                System.Drawing.Drawing2D.LinearGradientBrush brush = new System.Drawing.Drawing2D.LinearGradientBrush(new Rectangle(0, 0, image.Width, image.Height), Color.Black, Color.DarkRed, 1f, true);
                g.DrawString(content, font, brush, 2, 2);
                g.DrawImage(img, 0, 70, 160, 40);
                //画图片的边框线
                //g.DrawRectangle(new Pen(Color.Silver), 0, 0, image.Width - 1, image.Height - 1);
                image.Save(@"D:\str001.png");
            }
            finally
            {
                g.Dispose();
                image.Dispose();
            }
        }
        
        /// <summary>
        /// 打开路径并定位文件...对于@"h:\Bleacher Report - Hardaway with the safe call ??.mp4"这样的，explorer.exe /select,d:xxx不认，用API整它
        /// </summary>
        /// <param name="filePath">文件绝对路径</param>
        [DllImport("shell32.dll", ExactSpelling = true)]
        private static extern void ILFree(IntPtr pidlList);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        private static extern IntPtr ILCreateFromPathW(string pszPath);

        [DllImport("shell32.dll", ExactSpelling = true)]
        private static extern int SHOpenFolderAndSelectItems(IntPtr pidlList, uint cild, IntPtr children, uint dwFlags);

        public static void ExplorerFile(string filePath)//跳转到相应的文件夹并且选中
        {
            if (!File.Exists(filePath) && !Directory.Exists(filePath))
                return;

            if (Directory.Exists(filePath))
                Process.Start(@"explorer.exe", "/select,\"" + filePath + "\"");
            else
            {
                IntPtr pidlList = ILCreateFromPathW(filePath);
                if (pidlList != IntPtr.Zero)
                {
                    try
                    {
                        Marshal.ThrowExceptionForHR(SHOpenFolderAndSelectItems(pidlList, 0, IntPtr.Zero, 0));
                    }
                    finally
                    {
                        ILFree(pidlList);
                    }
                }
            }
        }
        /*
         * 获得指定路径下所有文件名
         * StreamWriter sw  文件写入流
         * string path      文件路径
         * int indent       输出时的缩进量
         */
        public void poFileName(string path)
        {
            string path1 = @path;
            DirectoryInfo root = new DirectoryInfo(path1);
            foreach (FileInfo f in root.GetFiles())
            {
                string myfile = f.Name;
                string mylastfile = f.Extension;
                string sourceFile = Path.Combine(path, myfile);
                string onlyname = myfile.Replace(mylastfile, "");
                string[] strarr = onlyname.Split('_');
                string[] sizearr = strarr[1].Split(',');
                for (int i = 0; i < sizearr.Length; i++)
                {
                    string sourceFile1 = sourceFile;
                    string destinationFile = strarr[0] + mylastfile;
                    string str1 = @"D:\图库";
                    string extension = Path.GetExtension(sourceFile1);
                    string str2 = "";
                    if (extension == ".jpg")
                    {
                        str2 = "A";
                    }
                    else if (extension == ".cdr")
                    {
                        str2 = "B";
                    }
                    else if (extension == ".arx4")
                    {
                        str2 = "C";
                    }
                    string str3 = Path.Combine(str1, str2,sizearr[i], destinationFile);
                    string newPath = System.IO.Path.Combine(str1, str2,sizearr[i]);
                    System.IO.Directory.CreateDirectory(newPath);
                    bool isrewrite = true; // true=覆盖已存在的同名文件,false则反之
                    if (sourceFile1 != str3)
                    {
                        System.IO.File.Copy(sourceFile1, str3, isrewrite);
                    }
                    object[] objects2 = new object[1];
                    objects2[0] = str3;
                    webBrowser1.Document.InvokeScript("givemeurl", objects2);
                    object[] objects = new object[5];
                    objects[0] = strarr[0];
                    objects[1] = sizearr[i];
                    objects[2] = strarr[2];
                    objects[3] = strarr[3];
                    objects[4] = str2;
                    webBrowser1.Document.InvokeScript("morepath", objects);
                }
            }
        }

        
        //拷贝图库到固定文件夹
        public void copyimg(string sourceFile, string mysize)
        {
            string sourceFile1 = @sourceFile;
            string destinationFile = System.IO.Path.GetFileName(sourceFile);
            string str1 = @"D:\图库";
            string str2 = "";
            string extension = Path.GetExtension(sourceFile1);
            if(extension == ".jpg")
            {
                str2 = "A";
            }else if(extension == ".cdr")
            {
                str2 = "B";
            }
            else if (extension == ".arx4")
            {
                str2 = "C";
            }
            string str3 = Path.Combine(str1,str2, mysize, destinationFile);
            string newPath = System.IO.Path.Combine(str1, str2,mysize);
            System.IO.Directory.CreateDirectory(newPath);
            bool isrewrite = true; // true=覆盖已存在的同名文件,false则反之
            if (sourceFile1 != str3)
            {
                System.IO.File.Copy(sourceFile1, str3, isrewrite);
            }
            object[] objects = new object[1];
            objects[0] = str3;
            webBrowser1.Document.InvokeScript("givemeurl", objects);
        }
        //跳转到相应的文件夹
        public void overstring()
        {
            string stri1 = @"D:\导出图库";
            string newtime = DateTime.Now.ToString("yyyy-MM-dd");
            string str4 = Path.Combine(stri1, newtime);
            ExplorerFile(str4);
        }
        //导出原图(D)
        public void cexptyunatu()
        {
            if (File.Exists(@filepath2))
            {

                CorelDRAW.Application core = new CorelDRAW.Application();
                core.Visible = false;
                core.OpenDocument(@filepath2, 0);
                string newtime = DateTime.Now.ToString("yyyy-MM-dd");
                string onePath = "D:\\导出图库\\" + newtime + "\\" + type2  + "\\" + size2;
                if (!Directory.Exists(onePath))
                {
                    Directory.CreateDirectory(onePath);//检查是否有同名的文件夹，没有就新建
                }
                for (int i = 1; i <= num2; i++)
                {
                    core.ActiveDocument.SaveAs(onePath + "\\" + stname2 + BarcodeString2 + "-" + i + ".cdr");
                }
                core.ActiveDocument.Close();
                core = null;
                webBrowser1.Document.InvokeScript("add");
                //core.Quit();
            }
            else
            {
                webBrowser1.Document.InvokeScript("add");
            }
        }
        //导出合成图(A,B)
        public void cexptai()
        {
            if (File.Exists(@filepath2))
            {
                CorelDRAW.Application core = new CorelDRAW.Application();
                core.Visible = false;
                if (type2 == "B")
                {
                    core.OpenDocument(@filepath2, 0);
                    VGCore.Layer pic1 = core.ActiveDocument.ActivePage.CreateLayer("pic1");
                    //置入图片  
                   // CreateBarcodePicture(BarcodeString2, 200, 40);
                    CreateImage(itemnum2  + "\r\n" +  colorname2 + stylename2 + "\r\n尺寸：" + sizename2 + "    数量：" + num2 + " \r\n编号：" + BarcodeString2);
                    core.ActiveDocument.ActivePage.ActiveLayer.Import("D://str001.png");
                    VGCore.Shape shape = core.ActiveSelection;
                    core.ActiveDocument.Unit = VGCore.cdrUnit.cdrMillimeter;
                    shape.Rotate(angle2);
                    shape.SetPosition(xx2, yy2);
                    string newtime = DateTime.Now.ToString("yyyy-MM-dd");
                    string onePath = "D:\\导出图库\\" + newtime + "\\" + type2  + "\\" + size2;
                    if (!Directory.Exists(onePath))
                    {
                        Directory.CreateDirectory(onePath);//检查是否有同名的文件夹，没有就新建
                    }
                    if (type2 == "B")
                    {
                        for (int i = 1; i <= num2; i++)
                        {
                            core.ActiveDocument.SaveAs(onePath + "\\" + stname2 + "-" + BarcodeString2 + "-" + i + ".cdr");
                        }
                    }
                    core.ActiveDocument.Close();
                    core = null;
                    webBrowser1.Document.InvokeScript("add");
                }
                if (type2 == "A")
                {
                    core.OpenDocument(@filepath2, 0);
                    VGCore.Layer pic1 = core.ActiveDocument.ActivePage.CreateLayer("pic1");
                    //置入图片  
                    //CreateBarcodePicture(BarcodeString2, 200, 40);
                    CreateImage(itemnum2  + "\r\n" + colorname2 + stylename2 + "\r\n尺寸：" + sizename2 + "    数量：" + num2 + " \r\n编号：" + BarcodeString2);
                    core.ActiveDocument.ActivePage.ActiveLayer.Import("D://str001.png");
                    VGCore.Shape shape = core.ActiveSelection;
                    core.ActiveDocument.Unit = VGCore.cdrUnit.cdrMillimeter;
                    shape.Rotate(angle2);
                    shape.SetPosition(xx2, yy2);
                    string newtime = DateTime.Now.ToString("yyyy-MM-dd");
                    string onePath = "D:\\导出图库\\" + newtime + "\\" + type2  + "\\" + size2;
                    if (!Directory.Exists(onePath))
                    {
                        Directory.CreateDirectory(onePath);//检查是否有同名的文件夹，没有就新建
                    }
                    if (type2 == "A")
                    {
                        for (int i = 1; i <= num2; i++)
                        {
                            //core.ActiveDocument.SaveAs(onePath + "//" + stname2 + BarcodeString2 + "-" + i + ".cdr");
                            core.ActiveDocument.Export(onePath + "\\" + stname2 + "-" + BarcodeString2 + "-" + i + ".jpg", VGCore.cdrFilter.cdrJPEG);
                        }
                    }
                    core.ActiveDocument.Close();
                    core = null;
                    webBrowser1.Document.InvokeScript("add");
                }
            }
            else
            {
                webBrowser1.Document.InvokeScript("add");//进度条增加
                object[] objects = new object[1];
                objects[0] = BarcodeString2;
                webBrowser1.Document.InvokeScript("changestatus", objects);
            }
        }
        //判断打印机是否存在，还有点问题
        public int islood(string printname)
        {
            int i;
            int ok = 0;
            string pkInstalledPrinters;
            using (PrintDocument pd = new PrintDocument())
            {
                for (i = 0; i < PrinterSettings.InstalledPrinters.Count; i++)  //开始遍历  
                {
                    pkInstalledPrinters = PrinterSettings.InstalledPrinters[i];  //取得名称 
                    if (pkInstalledPrinters == printname)
                    {
                        ok = 1;
                    }
                }
                if (ok == 0)
                {
                    object[] objects = new object[1];
                    objects[0] = "打印机未连接";
                    webBrowser1.Document.InvokeScript("tipError", objects);
                }
            }
            return ok;
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }
    }
}
