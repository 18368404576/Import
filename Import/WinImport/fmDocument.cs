using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using TSES.Base;

namespace WinImport
{
    public partial class fmDocument : Form
    {
        #region API
        [DllImport("kernel32.dll")]
        public static extern IntPtr _lopen(string lpPathName, int iReadWrite);
        [DllImport("kernel32.dll")]
        public static extern bool CloseHandle(IntPtr hObject);
        public const int OF_READWRITE = 2;
        public const int OF_SHARE_DENY_NONE = 0x40;
        public readonly IntPtr HFILE_ERROR = new IntPtr(-1);
        #endregion

        #region 变量
        /// <summary>
        /// excel的设备数据
        /// </summary>
        DataTable _ECInfo_excel = null;

        ExcelManager _excelManager = new ExcelManager();
        /// <summary>
        /// 是否验证过
        /// </summary>
        bool isCheck = false;
        /// <summary>
        /// 需要手动执行的SQL语句
        /// </summary>
        string strSql = "";

        string FolderPath = AppDomain.CurrentDomain.BaseDirectory + "runsql";


        #endregion

        #region 构造函数
        public fmDocument()
        {
            InitializeComponent();
        }
        #endregion

        #region 事件
        private void fmDocument_Load(object sender, EventArgs e)
        {

        }

        private void btnCheck_Click(object sender, EventArgs e)
        {
            if (_ECInfo_excel == null)
            {
                WGMessage.ShowWarning("请选择[文件清单]文件!");
                return;
            }
            if (textBox1.Text.Trim() == "")
            {
                WGMessage.ShowWarning("请选择文档源目录!");
                return;
            }
            if (textBox2.Text.Trim() == null)
            {
                WGMessage.ShowWarning("请选择服务器文档存储目录!");
                return;
            }
            if (isCheck)
            {
                WGMessage.ShowAsterisk("已验证，不用重复验证！");
                return;
            }
            DirectoryInfo folder = new DirectoryInfo(textBox1.Text);
            FileInfo[] dirInfo = folder.GetFiles();
            if (dirInfo.Length != _ECInfo_excel.Rows.Count)
            {
                WGMessage.ShowAsterisk("文件清单和文档源目录文件数量不匹配！");
                return;
            }
            DataTable dataTablFile = new DataTable();
            dataTablFile.Columns.Add("FILENAME", typeof(string));
            dataTablFile.Columns.Add("FULLFILENAME", typeof(string));
            foreach (FileInfo item in dirInfo)
            {
                dataTablFile.Rows.Add(item.ToString().Trim(), item.FullName.ToString().Trim());
            }
            strSql = "";
            _ECInfo_excel.Columns.Add("FULLFILENAME", typeof(string));
            DataTable dt_error = _ECInfo_excel.Clone();
            foreach (DataRow item in _ECInfo_excel.Rows)
            {
                DataRow[] drs = dataTablFile.Select("FILENAME='" + ReturnString(item["文件名"].ToString()) + "'");
                if (drs.Length == 0)
                {
                    dt_error.Rows.Add(item.ItemArray);
                }
                FileInfo file = new FileInfo(drs[0]["FULLFILENAME"].ToString());
                string Fsize = (file.Length / 1024 == 0 ? 1 : file.Length / 1024) + "";
                if (strSql == "")
                {
                    strSql = "UPDATE DMVERSION SET FSIZE=" + Fsize + " WHERE FKEY='" + item["文件KEY"].ToString() + "'";
                }
                else
                {
                    strSql = strSql + "\r\n" + "UPDATE DMVERSION SET FSIZE=" + Fsize + " WHERE FKEY='" + item["文件KEY"].ToString() + "'";
                }
                item["FULLFILENAME"] = drs[0]["FULLFILENAME"].ToString();
            }
            if (dt_error.Rows.Count > 0)
            {
                isCheck = false;
                dgError.SetDataBinding(dt_error, "", false);
            }
            else
            {
                WGMessage.ShowAsterisk("数据检验通过！");
                isCheck = true;
            }
        }

        private void btnSelect1_Click(object sender, EventArgs e)
        {
            //得到excel数据源
            if (opfDialog.ShowDialog() == DialogResult.OK)
            {
                Button btn = sender as Button;
                if (btn.Name == "btnSelect1")
                {
                    txtFile1.Text = opfDialog.FileName;
                    _ECInfo_excel = _excelManager.GetExcelTableByOleDB(opfDialog.FileName, "DMLIBUARY");
                    MessageBox.Show("读取笔数:" + _ECInfo_excel.Rows.Count + "");
                }
                if (_ECInfo_excel == null || _ECInfo_excel.Rows.Count < 0)
                {
                    WGMessage.ShowWarning(@"无法读取当前Excel!");
                    return;
                }
            }
            else
            {
                return;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = dialog.SelectedPath;
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = dialog.SelectedPath;
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            if (!isCheck)
            {
                WGMessage.ShowAsterisk("还未验证，不能导入！");
                return;
            }
            foreach (DataRow item in _ECInfo_excel.Rows)
            {
                Create(item["文件名"].ToString(), item["文件KEY"].ToString());
                OpenWrite(item["FULLFILENAME"].ToString(), item["文件KEY"].ToString());
                string newcode = item["文档编号"].ToString();
                string dtype = item["受控类型"].ToString().Trim();
                string day = item["有效天数"].ToString().Trim() == "" ? "0" : item["有效天数"].ToString().Trim();
                if (dtype == "1")
                {
                    CreateWatermarkFile(item["文件名"].ToString(), item["文件KEY"].ToString(), dtype, newcode, DateTime.Now);
                }
                else
                {
                    CreateWatermarkFile(item["文件名"].ToString(), item["文件KEY"].ToString(), dtype, newcode, DateTime.Now.AddDays(int.Parse(day)));
                }
            }
            richTextBox1.Text = strSql;
            WriteLog(strSql, true);
            WGMessage.ShowAsterisk("导入成功！");
        }
        #endregion

        #region 私有方法
        /// <summary>
        /// 返回处理过单引号的和非法字符的字符串
        /// </summary>
        /// <param name="str">原string</param>
        /// <returns>处理后的string</returns>
        private string ReturnString(string str)
        {
            if (str != null)
            {
                if (str.Contains('\''))
                {
                    str = str.Replace("\'", "\''");
                }
            }
            return str;
        }

        private void WriteLog(string msg, bool writeTime)
        {
            DateTime TimeNow = DateTime.Now;
            string FolderPath2 = FolderPath;
            //判断文件夹是否存在
            if (Directory.Exists(FolderPath2) == false)
            {
                Directory.CreateDirectory(FolderPath2);
            }//不存在，重新新建

            if (File.Exists(FolderPath2 + "\\runsql.txt"))
            {
                File.Delete(FolderPath2 + "\\runsql.txt");
            }
            using (System.IO.StreamWriter sw = new System.IO.StreamWriter(FolderPath2 + "\\runsql.txt", true))
            {
                if (writeTime)
                {
                    sw.WriteLine(TimeNow.ToString("yyyy-MM-dd HH:mm:ss ") + msg);
                }
                else
                {
                    sw.WriteLine(msg);
                }
                sw.Close();
            }
        }

        public void Create(string filename, string key)
        {
            string Upload = textBox2.Text.Trim() + "\\Upload";
            if (!Directory.Exists(Upload))
                Directory.CreateDirectory(Upload);

            string tmp;
            string dat;
            string extension = Path.GetExtension(filename);

            while (true)
            {
                if (!Directory.Exists(Path.Combine(Upload, key.Substring(0, 2))))
                {
                    Directory.CreateDirectory(Path.Combine(Upload, key.Substring(0, 2)));
                }
                tmp = Path.Combine(Upload, key.Substring(0, 2), key);
                dat = Path.Combine(Upload, key.Substring(0, 2), key);

                if (!File.Exists(tmp) && !File.Exists(dat))
                    break;
            }
        }

        public void OpenWrite(string fileName, string filekey)
        {
            string Upload = textBox2.Text.Trim() + "\\Upload";
            string path = Path.Combine(Upload, filekey.Substring(0, 2), filekey);
            FileStream fs = File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
            BinaryReader reader = new BinaryReader(fs);
            byte[] data = reader.ReadBytes((int)fs.Length);
            Stream FStream = File.Open(path, FileMode.Create, FileAccess.Write, FileShare.None);
            FStream.Write(data, 0, data.Length);
            FStream.Flush();
            FStream.Close();
            FStream.Dispose();
            fs.Close();
            fs.Dispose();
            FStream = null;
        }


        /// <summary>
        /// 服务器生成水印文件
        /// </summary>
        /// <param name="fileKey">文件KEY</param>
        /// <param name="cType">1：受控水印；2：临时受控水印；3：作废水印</param>
        /// <param name="code">文档编号，非空时添加文档编号水印</param>
        /// <param name="aTime">评审日期 cType=1传入生效日期，cType=2传入有效期，cType=3传入作废日期</param>
        /// <returns></returns>
        public bool CreateWatermarkFile(string filename, string filekey, string cType, string code, DateTime aTime)
        {
            string tempPath = textBox2.Text.Trim() + @"\Temp";
            string UploadPath = textBox2.Text.Trim() + @"\Upload";
            if (!File.Exists(Path.Combine(UploadPath, filekey.Substring(0, 2), filekey)))
            {
                Log.Debug("CreateWatermarkFile时发现 {0} 文件不存在！", Path.Combine(UploadPath, filekey.Substring(0, 2), filekey));
                return false;
            }
            string path = Path.Combine(UploadPath, filekey.Substring(0, 2), filekey);
            if (!Directory.Exists(Path.Combine(tempPath, filekey.Substring(0, 2))))
            {
                Directory.CreateDirectory(Path.Combine(tempPath, filekey.Substring(0, 2)));
            }
            string targetFileName = string.Format("{0}-watermark", filekey);
            string targetPath = Path.Combine(tempPath, filekey.Substring(0, 2), string.Format("{0}temp", targetFileName));
            string extname = Path.GetExtension(filename);

            //添加水印效果
            if (extname == "")
            {
                Log.Debug("CreateWatermarkFile时发现BSFILEUPLOAD中找不到 {0} 的文件扩展名！", filekey);
                return false;
            }
            if (extname.ToLower() == ".doc" || extname.ToLower() == ".docx")
            {
                try
                {
                    using (Stream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        Aspose.Words.Document doc = new Aspose.Words.Document(path);
                        doc.Save(targetPath, Aspose.Words.SaveFormat.Pdf);
                    }
                }
                catch (Exception ex)
                {
                    Log.Exception(ex.Message, ex);
                    return false;
                }
            }
            else if (extname.ToLower() == ".xls" || extname.ToLower() == ".xlsx")
            {
                try
                {
                    Aspose.Cells.Workbook xls = new Aspose.Cells.Workbook(path);
                    Aspose.Cells.PdfSaveOptions xlsSaveOption = new Aspose.Cells.PdfSaveOptions();
                    xlsSaveOption.SecurityOptions = new Aspose.Cells.Rendering.PdfSecurity.PdfSecurityOptions();
                    xlsSaveOption.SecurityOptions.ExtractContentPermission = false;
                    xlsSaveOption.SecurityOptions.PrintPermission = false;
                    xlsSaveOption.AllColumnsInOnePagePerSheet = true;
                    xls.Save(targetPath, xlsSaveOption);
                }
                catch (Exception ex)
                {
                    Log.Exception(ex.Message, ex);
                    return false;
                }
            }
            else if (extname.ToLower() == ".pdf")
            {
                File.Copy(path, targetPath, true);
            }

            if (FileAddWaterMark(targetPath, Path.Combine(tempPath, filekey.Substring(0, 2), targetFileName), cType, code, aTime))
            {
                File.Delete(targetPath);
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// 博格华纳文件添加水印
        /// </summary>
        /// <param name="sourceFile">源文件</param>
        /// <param name="targetFile">目标文件</param>
        /// <param name="cType">1：受控水印；2：临时受控水印；3：作废水印</param>
        /// <param name="code">文档编号，非空时添加文档编号水印</param>
        /// <param name="aTime">评审日期</param>
        private bool FileAddWaterMark(string sourceFile, string targetFile, string cType, string code, DateTime aTime)
        {
            PdfReader pdfReader = null;
            PdfStamper pdfStamper = null;
            string tempPath = Environment.GetEnvironmentVariable("TEMP");
            try
            {
                IntPtr vHandle = _lopen(sourceFile, OF_READWRITE | OF_SHARE_DENY_NONE);
                CloseHandle(vHandle);
                string s = "";
                iTextSharp.text.BaseColor color = iTextSharp.text.BaseColor.DARK_GRAY;
                if (cType == "1")
                    s = "受    控";
                else if (cType == "2")
                    s = "临时受控";
                else if (cType == "3")
                {
                    s = "作    废";
                    color = iTextSharp.text.BaseColor.RED;
                }

                pdfReader = new PdfReader(sourceFile);
                pdfStamper = new PdfStamper(pdfReader, new FileStream(targetFile, FileMode.Create));
                int total = pdfReader.NumberOfPages + 1;
                PdfContentByte content;
                BaseFont font = BaseFont.CreateFont(@"C:\Windows\Fonts\msyh.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                PdfGState gs = new PdfGState();
                for (int i = 0; i < total; i++)
                {
                    content = pdfStamper.GetOverContent(i);//在内容上方加水印
                    if (content == null)
                        continue;
                    //透明度
                    gs.FillOpacity = 0.3f;
                    content.SetGState(gs);

                    content.SetLineWidth(0.05F);
                    content.SetColorStroke(color);
                    content.MoveTo(45F, 50F);
                    content.LineTo(120F, 50F);
                    content.LineTo(120F, 14F);
                    content.MoveTo(45F, 50F);
                    content.LineTo(45F, 14F);
                    content.LineTo(120F, 14F);
                    content.Stroke();

                    //开始写入文本
                    content.BeginText();
                    content.SetColorFill(color);
                    content.SetFontAndSize(font, 5);
                    content.SetTextMatrix(0, 0);
                    content.ShowTextAligned(iTextSharp.text.Element.ALIGN_LEFT, "博格华纳汽车零部件有限公司", 50, 42, 0);
                    content.SetFontAndSize(font, 8);
                    content.ShowTextAligned(iTextSharp.text.Element.ALIGN_CENTER, s, 82, 32, 0);
                    content.SetFontAndSize(font, 5);
                    if (code != "")
                        content.ShowTextAligned(iTextSharp.text.Element.ALIGN_LEFT, code, 50, 24, 0);
                    content.ShowTextAligned(iTextSharp.text.Element.ALIGN_LEFT, aTime.ToString("yyyy年MM月dd日"), 80, 16, 0);
                    content.EndText();
                }

                return true;
            }
            catch (Exception ex)
            {
                Log.Exception(ex.Message, ex);
                return false;
            }
            finally
            {
                if (pdfStamper != null)
                {
                    pdfStamper.Close();
                }
                if (pdfReader != null)
                {
                    pdfReader.Close();
                }
            }
        }
        #endregion
    }
}
