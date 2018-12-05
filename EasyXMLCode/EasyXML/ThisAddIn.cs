using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace EasyXML
{
    public partial class ThisAddIn
    {
        private static Excel.Application app;
        private static Dictionary<string, List<string>> dic_tempFiles;//打开的临时文件名称
        private static Dictionary<string, string> dic_tempFiles_times;//打开的临时文件所对应xml文件的修改时间
        private static string easyXmlPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\EasyXML\\";
        private static string configName = "EasyXml.config";
        public static bool isForbid_Multiple = true;
        public static bool isAuto_Close_After_Export = true;
        public static bool isAuto_Check_Out_Update = true;
        public static string ExportMsg = "XML数据导出成功~";
        public static string ExportOutUpdateMsg = "检测到文件已经从外部更改, 取消本次导出, 请妥善保存你的修改, 然后重新打开该XML文件进行数据导出操作~";
        public static string ExportCloseMsg = "您已设置导出后自动关闭选项, 即将为您关闭该XML~";
        public static string ForbidMultipleMsg = "您已设置禁止多开数据, 即将为你关闭该文件~";
        public static string NotContainExportDataMsg = "检测到有数据未包含在数据区域, 这些数据将无法导出,请处理后重试~";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            app = this.Application;
            dic_tempFiles = new Dictionary<string, List<string>>();
            dic_tempFiles_times = new Dictionary<string, string>();
            app.WorkbookOpen += App_WorkbookOpen;
            app.WorkbookBeforeXmlImport += App_WorkbookBeforeXmlImport;
            app.WorkbookBeforeClose += App_WorkbookBeforeClose;

            CheckConfig();
            if (app.ActiveWorkbook != null)
            {
                var Wb = app.ActiveWorkbook;
                CheckWb(Wb);
            }
        }

        private static void CheckConfig()
        {
            if (!Directory.Exists(easyXmlPath))
            {
                Directory.CreateDirectory(easyXmlPath);
            }

            if (!File.Exists(easyXmlPath + configName))
            {
                List<string> configs = new List<string>();
                configs.Add("#是否禁止多开");
                configs.Add("Forbid_Multiple:True");//禁止多开
                configs.Add("#是否检测外部更改");
                configs.Add("Auto_Check_Out_Update:True");//检测外部更改
                configs.Add("#是否导出自动关闭");
                configs.Add("Auto_Close_After_Export:True");//导出自动关闭
                configs.Add("ExportMsg:XML数据导出成功~");
                configs.Add("ExportOutUpdateMsg:检测到文件已经从外部更改, 取消本次导出, 请妥善保存你的修改, 然后重新打开该XML文件进行数据导出操作~");
                configs.Add("ExportCloseMsg:您已设置导出后自动关闭选项, 即将为您关闭该XML~");
                configs.Add("ForbidMultipleMsg:您已设置禁止多开数据, 即将为你关闭该文件~");
                configs.Add("NotContainExportDataMsg:检测到有数据未包含在数据区域, 这些数据将无法导出,请处理后重试~");
                isForbid_Multiple = true;
                isAuto_Close_After_Export = true;
                isAuto_Check_Out_Update = true;
                Write(easyXmlPath + configName, configs);
                return;
            }

            var list = Read(easyXmlPath + configName);
            for (int i = 0; i < list.Count; i++)
            {
                if (list[i].Contains("#"))
                {
                    continue;
                }
                list[i].Replace(" ", "");//去掉所有空格
                var keyValue = list[i].Split(':');
                if (keyValue[0].Equals("Forbid_Multiple"))
                {
                    bool.TryParse(keyValue[1], out isForbid_Multiple);
                }

                if (keyValue[0].Equals("Auto_Close_After_Export"))
                {
                    bool.TryParse(keyValue[1], out isAuto_Close_After_Export);
                }

                if (keyValue[0].Equals("Auto_Check_Out_Update"))
                {
                    bool.TryParse(keyValue[1], out isAuto_Check_Out_Update);
                }

                if (keyValue[0].Equals("ExportMsg"))
                {
                    ExportMsg = GB2312ToUTF8(keyValue[1]);
                }

                if (keyValue[0].Equals("ExportOutUpdateMsg"))
                {
                    ExportOutUpdateMsg = GB2312ToUTF8(keyValue[1]);
                }

                if (keyValue[0].Equals("ExportCloseMsg"))
                {
                    ExportCloseMsg = GB2312ToUTF8(keyValue[1]);
                }

                if (keyValue[0].Equals("ForbidMultipleMsg"))
                {
                    ForbidMultipleMsg = GB2312ToUTF8(keyValue[1]);
                }

                if (keyValue[0].Equals("NotContainExportDataMsg"))
                {
                    NotContainExportDataMsg = GB2312ToUTF8(keyValue[1]);
                }
            }
        }

        private static string GB2312ToUTF8(string str)
        {
            try
            {
                Encoding utf8 = Encoding.UTF8;
                Encoding gb2312 = Encoding.GetEncoding("GB2312");
                byte[] unicodeBytes = gb2312.GetBytes(str);
                byte[] asciiBytes = Encoding.Convert(gb2312, utf8, unicodeBytes);
                char[] asciiChars = new char[utf8.GetCharCount(asciiBytes, 0, asciiBytes.Length)];
                utf8.GetChars(asciiBytes, 0, asciiBytes.Length, asciiChars, 0);
                string result = new string(asciiChars);
                return result;
            }
            catch
            {
                return "";
            }
        }

        private void App_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            if (dic_tempFiles.ContainsKey(Wb.Title))
            {
                for (int i = 1; i <= app.Workbooks.Count; i++)
                {
                    if (app.Workbooks[i] != Wb && app.Workbooks[i].Title == Wb.Title)
                    {
                        return;
                    }
                }

                dic_tempFiles.Remove(Wb.Title);
            }
        }

        private void App_WorkbookBeforeXmlImport(Excel.Workbook Wb, Excel.XmlMap Map, string Url, bool IsRefresh, ref bool Cancel)
        {
            if (!Url.EndsWith(".xml"))
            {
                return;
            }

            Wb.Title = Url;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            System.IO.DirectoryInfo dirInfo = new System.IO.DirectoryInfo(easyXmlPath);
            var files = dirInfo.GetFiles();
            foreach (var file in files)
            {
                if (!file.FullName.EndsWith(".data"))
                {
                    continue;
                }

                if (File.Exists(file.FullName))
                {
                    try
                    {
                        File.Delete(file.FullName);
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            app = null;
        }

        private void App_WorkbookOpen(Excel.Workbook Wb)
        {
            CheckWb(Wb);
        }

        private static void CheckWb(Excel.Workbook Wb)
        {
            if (!Wb.Title.EndsWith(".xml"))
            {
                var map = Wb.XmlMaps[1];
                string url = map.DataBinding.SourceUrl;
                if (url.EndsWith(".xml"))
                {
                    Wb.Title = url;
                }
            }

            if (Wb.Title.EndsWith(".xml"))
            {
                int index = Wb.Title.LastIndexOf("\\");
                string temp = Wb.Title.Substring(index + 1);
                if (!Directory.Exists(easyXmlPath))
                {
                    Directory.CreateDirectory(easyXmlPath);
                }
                string fileName = "";
                if (!dic_tempFiles.ContainsKey(Wb.Title))
                {
                    fileName = easyXmlPath + temp.Substring(0, temp.Length - 4) + ".data";
                    dic_tempFiles[Wb.Title] = new List<string>() { fileName };
                    string time = "";
                    try
                    {
                        FileInfo fi = new FileInfo(Wb.Title);
                        time = fi.LastWriteTime.ToString();
                    }
                    catch (Exception)
                    {
                    }
                    dic_tempFiles_times[Wb.Title] = time;
                }
                else
                {

                    int count = dic_tempFiles[Wb.Title].Count;
                    if (isForbid_Multiple && count > 0)
                    {
                        try
                        {
                            //不允许多开文件，弹出提示并关闭
                            DialogResult dr = MessageBox.Show(ForbidMultipleMsg, "EasyXml提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            if (dr == DialogResult.OK)
                            {
                                //点击确定
                                Wb.Close(SaveOptions.None);
                            }
                            else
                            {
                                //关闭
                                Wb.Close(SaveOptions.None);
                            }
                        }
                        catch (Exception)
                        {

                        }
                        return;
                    }

                    if (count > 0)
                    {
                        fileName = easyXmlPath + temp.Substring(0, temp.Length - 4) + (count + 1) + ".data";
                    }
                    else
                    {
                        fileName = easyXmlPath + temp.Substring(0, temp.Length - 4) + ".data";
                    }
                    dic_tempFiles[Wb.Title].Add(fileName);
                    string time = "";
                    try
                    {
                        FileInfo fi = new FileInfo(Wb.Title);
                        time = fi.LastWriteTime.ToString();
                    }
                    catch (Exception)
                    {
                    }
                    dic_tempFiles_times[Wb.Title] = time;
                }
                if (File.Exists(fileName))
                {
                    try
                    {
                        File.Delete(fileName);
                    }
                    catch (Exception)
                    {
                    }
                }
                Wb.SaveAs(fileName);
            }
        }

        public static void OpenFile()
        {
            var dialog = app.GetOpenFilename();
            string xml = ((object)dialog).ToString();
            app.Workbooks.Add(xml);
        }

        public static void SaveFile()
        {
            if (app.Workbooks.Count <= 0)
            {
                MessageBox.Show("Office Workbooks异常", "EasyXml错误");
                return;
            }
            var wb = app.ActiveWorkbook;
            if (wb.XmlMaps.Count <= 0)
            {
                MessageBox.Show("没有可以保存的xml map数据", "EasyXml错误");
                return;
            }

            if (!wb.Title.Contains(".xml"))
            {
                SaveFileDialog ofd = new SaveFileDialog();
                ofd.Filter = "XML文件(*.xml)|*.xml|所有文件|*.*";
                ofd.ValidateNames = true;
                ofd.CheckPathExists = true;
                ofd.CheckFileExists = true;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    string strFileName = ofd.FileName;
                    wb.Title = strFileName;
                    var map = wb.XmlMaps[1];
                    CheckWb(wb);
                    wb.SaveAsXMLData(wb.Title, map);
                    var res = MessageBox.Show(ExportMsg, "导出成功", MessageBoxButtons.OK);
                }
            }
            else
            {
                //检测外部文件是否已经被修改，如果是，弹窗提示，取消导出操作
                FileInfo fi = new FileInfo(wb.Title);
                string time = fi.LastWriteTime.ToString();
                if (isAuto_Check_Out_Update && dic_tempFiles_times.ContainsKey(wb.Title))
                {
                    if (dic_tempFiles_times[wb.Title] != time.ToString())
                    {
                        //文件已经从外部修改
                        string info = ExportOutUpdateMsg;
                        MessageBox.Show(info, "外部变更提醒", MessageBoxButtons.OK);
                        return;
                    }
                }

                var map = wb.XmlMaps[1];
                Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[1];
                string checkStr;
                var exRes = map.ExportXml(out checkStr);
                if (ws != null && exRes == XlXmlExportResult.xlXmlExportSuccess)
                {
                    checkStr = checkStr.Replace("entry", "々");
                    int count = (checkStr.Split('々').Length - 1) / 2;
                    int dataRows = count + 1;
                    if (ws.UsedRange.Rows.Count > dataRows)
                    {
                        //数据域未更新，给出提醒
                        MessageBox.Show(NotContainExportDataMsg, "导出提醒", MessageBoxButtons.OK);
                        return;
                    }
                }

                wb.SaveAsXMLData(wb.Title, map);
                if (isAuto_Check_Out_Update && dic_tempFiles_times.ContainsKey(wb.Title))
                {
                    dic_tempFiles_times[wb.Title] = DateTime.Now.ToString();
                }

                string msg = ExportMsg;
                if (isAuto_Close_After_Export)
                {
                    msg += "\r\n" + ExportCloseMsg;
                }
                MessageBox.Show(msg, "导出成功", MessageBoxButtons.OK);

                if (isAuto_Close_After_Export)
                {
                    wb.Close(SaveOptions.None);
                }
            }



            if (app.Workbooks.Count <= 0)
            {
                app.Quit();
            }
        }

        public static void SetConfig(bool isForbid, bool isAutoCheck, bool isAutoClose)
        {
            List<string> configs = new List<string>();
            configs.Add("#是否禁止多开");
            configs.Add("Forbid_Multiple:" + isForbid);//禁止多开
            configs.Add("#是否检测外部更改");
            configs.Add("Auto_Check_Out_Update:" + isAutoCheck);//检测外部更改
            configs.Add("#是否导出自动关闭");
            configs.Add("Auto_Close_After_Export:" + isAutoClose);//导出自动关闭
            configs.Add("ExportMsg:" + ExportMsg);
            configs.Add("ExportOutUpdateMsg:" + ExportOutUpdateMsg);
            configs.Add("ExportCloseMsg:" + ExportCloseMsg);
            configs.Add("ForbidMultipleMsg:" + ForbidMultipleMsg);
            isForbid_Multiple = isForbid;
            isAuto_Close_After_Export = isAutoClose;
            isAuto_Check_Out_Update = isAutoCheck;
            Write(easyXmlPath + configName, configs);
        }

        public static void RedConfig(out bool isForbid, out bool isAutoCheck, out bool isAutoClose)
        {
            CheckConfig();
            isForbid = isForbid_Multiple;
            isAutoCheck = isAuto_Check_Out_Update;
            isAutoClose = isAuto_Close_After_Export;
        }

        private static List<string> Read(string path)
        {
            StreamReader sr = new StreamReader(path, Encoding.UTF8);
            string line;
            List<string> res = new List<string>();
            while ((line = sr.ReadLine()) != null)
            {
                res.Add(line);
            }
            sr.Close();
            return res;
        }

        private static void Write(string path, List<string> msgs)
        {
            FileStream fs = new FileStream(path, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);

            //开始写入
            foreach (var msg in msgs)
            {
                sw.WriteLine(msg);
            }

            //清空缓冲区
            sw.Flush();
            //关闭流
            sw.Close();
            fs.Close();
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
