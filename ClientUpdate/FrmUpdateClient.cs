using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;
using System.Net;
using System.Web;
using System.Xml;
using System.Data.OleDb;
using System.Web.UI;
using System.Web.UI.Adapters;
using System.Web.UI.HtmlControls;
using System.Web.Util;
using System.Security.AccessControl;//160116
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

using ICSharpCode.SharpZipLib.Checksums;
using ICSharpCode.SharpZipLib.Zip.Compression;
using ICSharpCode.SharpZipLib.Zip.Compression.Streams;
using ICSharpCode.SharpZipLib.Zip;









namespace ClientUpdate
{
    public partial class FrmUpdateClient : Form
    {

        StringBuilder sqb = new StringBuilder();
        CFileInfo cfileinfo = new CFileInfo();
        string err;
        private string _ErrowInfo;
        public string ErrowInfo
        {

            set { _ErrowInfo = value; }
            get { return _ErrowInfo; }

        }
        basec bc = new basec();
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FrmUpdateClient());

        }
        public FrmUpdateClient()
        {
            InitializeComponent();
        }
        private void FrmUpdateClient_Load(object sender, EventArgs e)
        {

          
            try
            {          /*加载不显示FORM start 16/01/26*/

              
                bind();
            
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            //bind();
        }
        #region bind
        private void bind()
        {
            this.Hide();
            this.ShowInTaskbar = false;
            this.WindowState = FormWindowState.Minimized;
            /*加载不显示FORM end 16/01/26*/
            //MessageBox.Show(bc.RETURN_SERVER_IP_OR_DOMAIN());
            string url = "";
            if (bc.RETURN_SERVER_IP_OR_DOMAIN() == "192.168.1.9")
            {
                url = "http://" + bc.RETURN_SERVER_IP_OR_DOMAIN() + "/webserver_lan/s_connectionstring.aspx";
            }
            else
            {
                url = "http://" + bc.RETURN_SERVER_IP_OR_DOMAIN() + "/webserver/s_connectionstring.aspx";
            }

            JArray jar = bc.RETURN_JARRAY(url, "S_CONNECTIONSTRING=*");
            string M_str_sqlcon = "";
            if (jar.Count > 0)
            {
                M_str_sqlcon = jar[0].ToString();
             
            }
            else
            {
            
                ErrowInfo = "与服务器的通讯连接异常,檢查網絡是否正常，服務器的WEBCLIENT服務是否在啟動狀態";
            }
           
            string IP_OR_DOMAIN = bc.RETURN_APPOINT_UNTIL_CHAR(M_str_sqlcon, 13, ';');
           
            WebClient wclient = new WebClient();
            SaveFileDialog sfl = new SaveFileDialog();
       
            //MessageBox.Show(IP_OR_DOMAIN);
            string v1 = "http://" + bc.RETURN_SERVER_IP_OR_DOMAIN() + "/updateclient/Version.xml";
            string v4 = "http://" + bc.RETURN_SERVER_IP_OR_DOMAIN() + "/updateclient/UpdateClient.zip";
          
            string v5 = AppDomain.CurrentDomain.BaseDirectory.ToString();//获取应用程序之前安装的路径 16/01/10
            DirectoryInfo directoryinfo_1 = new DirectoryInfo(v5);//判断是第一次装软件还是系统已在存在之前的版本,如果第一次装就无续从服务器下载版本号
            if (directoryinfo_1.Exists == false)//默认第一次装的是最新版本
            {
               
            }
            else
            {
               
                DirectoryInfo directoryinfo = new DirectoryInfo(@"c:\temp");
                if (directoryinfo.Exists == false)//新建一个文件夹用户存放从服务器下载的含版本号的XML文件 16/01/11
                {
                    directoryinfo.Create();
                }
              
               
                   wclient.DownloadFile(v1, @"c:\\temp\Version.xml");//先下载服务器含版本号的XML文件用于获取服务器程序的版本号 16/01/11
                
             
                DateTime date1 = cfileinfo.GetTheLastUpdateTime(@"c:\\temp\Version.xml");//获取服务器的最近更新日期 16/01/11
                DateTime date2 = cfileinfo.GetTheLastUpdateTime(System.IO.Path.GetFullPath("Version.xml"));//获取客户端当前版本更新日期 16/01/11
                //DateTime date2 = Convert.ToDateTime("2015/01/01 0:00");
                if (date1 > date2)
                {
                    //Directory.Delete(v5, true);//先删除安装目录文件
                    //sqb.AppendFormat("当前应用程序所在路径：{0}, ", v5);
                    //sqb.AppendFormat("服务器更新时间：{0}, ", date1);
                    //sqb.AppendFormat("客户端当前版本时间：{0}, ", date2);
                    sqb.AppendFormat("检测到有新版本需要更新 最新版本更新时间为：");
                    sqb.AppendFormat("{0} 点确定按扭开始更新 更新完成后将打开登录窗口 请稍后...", date1.ToString("yyyy/MM/dd HH:mm").Replace("-", "/"));
                    MessageBox.Show(sqb.ToString(), "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    List<CFileInfo> list1 = new List<CFileInfo>();
                    CFileInfo cf = new CFileInfo(v5);
                    list1 = cf.FindFile();
                    for (int i = 0; i < list1.Count; i++)
                    {
                        CFileInfo cfile = list1[i];
                        if (cfile.FileName == "项目管理系统.exe" || cfile.FileName == "项目管理系统.exe.config"
                            || cfile.FileName == "Interop.Shell32.dll" || cfile.FileName == "Newtonsoft.Json.dll" || cfile.FileName == "ICSharpCode.SharpZipLib.dll")
                        {
                        }
                        else
                        {
                            File.Delete(cfile.FileNameAndPath);//在WIN10 系统下操作C盘文件删除提示出错，无权限 160116
                        }

                    }

                    if (Directory.Exists(v5 + "Image"))
                    {
                        Directory.Delete(v5 + "Image", true);//删除之前版本存放文件的文件夹及里面的所有图片
                    }
                    //下载服务器中含高版本的程序压缩包ZIP格式到安装目录 16/01/11
                    wclient.DownloadFile(v4, AppDomain.CurrentDomain.BaseDirectory.ToString() + "UpdateClient.zip");
                    //将下载好的文件压缩包解压到安装目录覆盖之前的程序文件 16/01/11*/
                    UnZipFile(v5 + "UpdateClient.zip");
                    
                        //MessageBox.Show("解压成功" + v5+err );

                        if (File.Exists(v5 + "UpdateClient.zip"))
                        {
                            File.Delete(v5 + "UpdateClient.zip");//删除压缩包ZIP文件使用卸载软件时能卸载干净
                        }
                    
                 


                    wclient.DownloadFile(v1, AppDomain.CurrentDomain.BaseDirectory.ToString() + "Version.xml");//将新的版本XML文件下载到安装目录
                }
                this.Close();
                //System.Diagnostics.Process process = System.Diagnostics.Process.GetCurrentProcess();
                //process.Close();
                this.Dispose();
                Application.Exit();
                System.Diagnostics.Process.Start(v5 + "项目管理系统客户端.exe");
              
            }
        }
        #endregion
       

        private static void UnZipFile(string zipFilePath)
        {
            try
            {
                if (!File.Exists(zipFilePath))
                {
                    Console.WriteLine("Cannot find file '{0}'", zipFilePath);
                    return;
                }

                using (ZipInputStream s = new ZipInputStream(File.OpenRead(zipFilePath)))
                {

                    ZipEntry theEntry;
                    while ((theEntry = s.GetNextEntry()) != null)
                    {

                        Console.WriteLine(theEntry.Name);

                        string directoryName = Path.GetDirectoryName(theEntry.Name);
                        string fileName = Path.GetFileName(theEntry.Name);

                        // create directory
                        if (directoryName.Length > 0)
                        {
                            Directory.CreateDirectory(directoryName);
                        }

                        if (fileName != String.Empty)
                        {
                            using (FileStream streamWriter = File.Create(theEntry.Name))
                            {

                                int size = 2048;
                                byte[] data = new byte[2048];
                                while (true)
                                {
                                    size = s.Read(data, 0, data.Length);
                                    if (size > 0)
                                    {
                                        streamWriter.Write(data, 0, size);
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        public static string GetExcelFirstTableName(string excelFileName)
        {
            string tableName = null;
            if (File.Exists(excelFileName))
            {
                using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.Jet." +
                  "OLEDB.4.0;Extended Properties=\"Excel 8.0\";Data Source=" + excelFileName))
                {
                    conn.Open();
                    DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    tableName = dt.Rows[0][2].ToString().Trim();

                }
            }
            return tableName;
        }
    }
}
