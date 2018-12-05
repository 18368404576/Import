using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TSES.Base;

namespace WinImport
{
    public partial class BSWCTYPE : Form
    {
        ExcelManager _excelManager = new ExcelManager();
        /// <summary>
        /// 工作中心类别DataTable
        /// </summary>
        DataTable _BSWCTYP_DB = null;
        /// <summary>
        /// excel的工作中心类别
        /// </summary>
        DataTable _BSWCTYPE_excel = null;
        /// <summary>
        /// 1、未使用 2、使用中 3、停用
        /// </summary>
        Dictionary<string, int> STs = new Dictionary<string, int>();

        /// <summary>
        /// 数值，文本
        /// </summary>
        Dictionary<string, string> CTYPEs = new Dictionary<string, string>();

        /// <summary>
        /// 通用，产品
        /// </summary>
        Dictionary<string, string> BTYPE = new Dictionary<string, string>();

        /// <summary>
        /// 是，否
        /// </summary>
        Dictionary<string, bool> ISINTERVALs = new Dictionary<string, bool>();

        /// <summary>
        /// 需要保存的sql
        /// </summary>
        List<string> sqlLs = new List<string>();

        /// <summary>
        ///  是否验证成功
        /// </summary>
        bool isCheck = false;

        public BSWCTYPE()
        {
            InitializeComponent();
            CTYPEs.Add("数值", "数值");
            CTYPEs.Add("文本", "文本");

            STs.Add("未使用", 1);
            STs.Add("使用中", 2);
            STs.Add("停用", 3);

            BTYPE.Add("通用", "通用");
            BTYPE.Add("产品", "产品");

            ISINTERVALs.Add("是", true);
            ISINTERVALs.Add("否", false);
        }

        private void btnSelect1_Click(object sender, EventArgs e)
        {
            //得到excel数据源
            if (opfDialog.ShowDialog() == DialogResult.OK)
            {
                Button btn = sender as Button;
                if (btn == null) return;
                switch (btn.Name)
                {
                    case "btnSelect1":
                        txtFile1.Text = opfDialog.FileName;
                        _BSWCTYPE_excel = _excelManager.GetExcelTableByOleDB(opfDialog.FileName, "BSWCTYPE");
                        _BSWCTYPE_excel.Columns.Add("GUID");
                        MessageBox.Show("读取笔数:" + _BSWCTYPE_excel.Rows.Count + "");
                        break;
                }

                ClearSql();
                if (_BSWCTYPE_excel == null || _BSWCTYPE_excel.Rows.Count <= 0)
                {
                    WGMessage.ShowWarning(@"无法读取当前Excel!");
                    return;
                }
            }
        }

        private DataTable CheckDBRepeat(DataTable exceldt)
        {
            DataTable dtRepet = exceldt.Clone();
            string sql = string.Format(" SELECT [BSWCTYPE].[CODE] 类别编号,[BSWCTYPE].[NAME] 类别名称 FROM [BSWCTYPE]  ");
            DataTable dt = FillDatatablde(sql, Main.CONN_Public);
            foreach (DataRow row1 in exceldt.Rows)
            {
                foreach (DataRow row2 in dt.Rows)
                {
                    if (row1["类别编号"].ToString() == row2["类别编号"].ToString()|| row1["类别名称"].ToString() == row2["类别名称"].ToString())
                    {
                        bool isHave = false;
                        foreach (DataRow row3 in dtRepet.Rows)
                        {
                            if (row3["GUID"].ToString() == row1["GUID"].ToString())
                                isHave = true;
                        }
                        if(!isHave)
                        dtRepet.Rows.Add(row1.ItemArray);
                    }
                }
            }
            dtRepet.Columns.Remove("GUID");
            return dtRepet;
        }

        private void btnCheck_Click(object sender, EventArgs e)
        {
            try
            {
                if (isCheck)
                {
                    WGMessage.ShowAsterisk("已验证，不用重复验证！");
                    return;
                }
                rbSql.Text = string.Empty;
                sqlLs = new List<string>();
                if (_BSWCTYPE_excel == null)
                {
                    WGMessage.ShowWarning("请选择[工作中心类别]文件!");
                    return;
                }
                string sql = string.Empty;


                //错误
                List<int[]> col_error1 = new List<int[]>();
                DataTable dt_error1 = _BSWCTYPE_excel.Clone();
                //重复数据
                
                DataTable dt_repet1 = _BSWCTYPE_excel.Clone();
                foreach (DataRow row in _BSWCTYPE_excel.Rows)
                {
                    bool isError1 = false;
                    bool isRepet1 = false;
                    if (string.IsNullOrWhiteSpace(row["类别编号"].ToString()))
                    {
                        //空
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 0 });
                        isError1 = true;
                    }
                    else
                    {
                        DataRow[] drss = _BSWCTYPE_excel.Select(string.Format("类别编号 = '{0}'", ReturnString(row["类别编号"].ToString())));
                        if (drss.Length > 1)
                        {
                            isRepet1 = true;
                        }
                    }

                    if (string.IsNullOrWhiteSpace(row["类别名称"].ToString()))
                    {
                        //空
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 0 });
                        isError1 = true;
                    }
                    else
                    {
                        DataRow[] drss = _BSWCTYPE_excel.Select(string.Format("类别名称 = '{0}'", ReturnString(row["类别名称"].ToString())));
                        if (drss.Length > 1)
                        {
                            isRepet1 = true;
                        }
                    }

                   
                    string BSWCTYP_GUID = Guid.NewGuid().ToString();
                    row["GUID"] = BSWCTYP_GUID;

                    if (isError1 || isRepet1)
                    {
                        if (isError1)
                        {
                            dt_error1.Rows.Add(row.ItemArray);
                        }
                        if (isRepet1)
                        {
                            dt_repet1.Rows.Add(row.ItemArray);
                        }
                        continue;
                    }

                    try
                    {
                        string temp = string.Format(@"  INSERT INTO [BSWCTYPE] ( [GUID], [CODE], [NAME], ST, [NOTE], [CC], [ND], [CD] )
                                                    VALUES  ( '{0}', -- GUID - uniqueidentifier
                                                              '{1}', -- CODE - nvarchar(20)
                                                              '{2}', -- NAME - nvarchar(40)
                                                               0,
                                                              '{3}', -- NOTE - nvarchar(200)
                                                              '1023', -- CC - nvarchar(20)
                                                              GETDATE(), -- ND - datetime
                                                              NULL  -- CD - datetime
                                                              )", row["GUID"], row["类别编号"], row["类别名称"], row["备注"]);

                        rbSql.Text += temp + Environment.NewLine;
                        sqlLs.Add(temp);
                    }
                    catch
                    { }
                }
                dt_error1.Columns.Remove("GUID");
                dt_repet1.Columns.Remove("GUID");

                dgError.DataSource = dt_error1; 
                dgRepet.DataSource = dt_repet1;
                DataTable dt_repetFromDB = CheckDBRepeat(_BSWCTYPE_excel.Copy());
                dgRepetFromDB.DataSource = dt_repetFromDB;

                if (dt_error1.Rows.Count > 0 || dt_repet1.Rows.Count > 0||dt_repetFromDB.Rows.Count>0)
                {
                    Main.SetErrorCell(dgError, col_error1);
                    rbSql.Text = "";
                    isCheck = false;
                    return;
                }
                isCheck = true;
            }
            catch (Exception ex)
            {
                WGMessage.ShowAsterisk("出现未知异常！请检查Excel文件正确性和顺序的正确性！" + ex.ToString());
                return;
            }
        }

        #region 方法
        /// <summary>
        /// 清理SQL
        /// </summary>
        public void ClearSql()
        {
            //重新上传后，清空原来的
            isCheck = false;
            sqlLs = new List<string>();
            rbSql.Text = string.Empty;
            dgError.DataSource = new DataTable();
            dgRepet.DataSource = new DataTable();
            dgRepetFromDB.DataSource = new DataTable();
        }
        /// <summary>
        /// 字符串单引号处理
        /// </summary>
        /// <param name="str">要处理单引号的字符串</param>
        /// <returns>处理单引号后的字符串</returns>
        public static string ReturnString(string str)
        {
            if (str != null && str.Contains('\''))
            {
                return str.Replace("\'", "\''");
            }
            return str;
        }
        private DataTable FillDatatablde(string sql, string connectionString)
        {

            DataTable dt = new DataTable();
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandTimeout = 120;
                cmd.CommandText = sql;
                SqlDataAdapter dadFill = new SqlDataAdapter(cmd);
                dadFill.Fill(dt);
                conn.Close();
                conn.Dispose();
                GC.Collect();
                return dt;
            }
        }

        public bool RunSql(List<string> SQLStringList, string connectionString)
        {
            //return true;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = connection;
                SqlTransaction tx = connection.BeginTransaction();
                cmd.Transaction = tx;
                try
                {
                    for (int n = 0; n < SQLStringList.Count; n++)
                    {
                        string strsql = SQLStringList[n].ToString();
                        if (strsql.Trim().Length > 1)
                        {
                            cmd.Parameters.Clear();
                            cmd.CommandText = strsql;
                            cmd.ExecuteNonQuery();
                        }
                    }
                    tx.Commit();
                    return true;
                }
                catch (Exception e)
                {
                    tx.Rollback();
                    throw e;
                }
                finally
                {
                    connection.Close();
                }
            }
        }
        #endregion

        private void btnImport_Click(object sender, EventArgs e)
        {
            if (!isCheck)
            {
                WGMessage.ShowAsterisk("还未验证，不能导入！");
                return;
            }
            if (RunSql(sqlLs, Main.CONN_Public))
            {
                WGMessage.ShowAsterisk("导入成功！");
                ClearSql();
            }
            else
            {
                WGMessage.ShowAsterisk("导入失败！");
            }
        }
    }
}
