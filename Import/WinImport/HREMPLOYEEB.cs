using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TSES.Base;
using System.Data.SqlClient;

namespace WinImport
{
    public partial class HREMPLOYEEB : Form
    {
        public HREMPLOYEEB()
        {
            InitializeComponent();
        }

        #region 变量
        ExcelManager _excelManager = new ExcelManager();
        DataTable _HREMPLOYEEB_excel = null;
        DataTable _HREMPLOYEE_DB = null;
        DataTable _HRPHYSICAL_DB = null;

        /// <summary>
        ///  是否验证成功
        /// </summary>
        bool isCheck = false;

        /// <summary>
        /// 需要保存的sql
        /// </summary>
        List<string> sqlLs = new List<string>();
        #endregion

        #region 方法
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

        private DataTable FillDatatable(string sql, string connectionString)
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

        public void ClearSql()
        {
            //重新上传后，清空原来的
            isCheck = false;
            sqlLs = new List<string>();
            rbSql.Text = "";

            dgError2.DataSource = new DataTable();
            dgRepet2.DataSource = new DataTable();
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

        private void button1_Click(object sender, EventArgs e)
        {
            //得到excel数据源
            if (opfDialog0.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    txtFileSon.Text = opfDialog0.FileName;
                    _HREMPLOYEEB_excel = _excelManager.GetExcelTableByOleDB(opfDialog0.FileName, "HREMPLOYEEB");
                    MessageBox.Show("读取笔数：" + _HREMPLOYEEB_excel.Rows.Count + "");
                }
                catch
                {

                }
                //ClearSql();
                if (_HREMPLOYEEB_excel == null || _HREMPLOYEEB_excel.Rows.Count <= 0)
                {
                    WGMessage.ShowWarning(@"无法读取当前Excel!");
                    return;
                }
                isCheck = false;
            }
            else
            {
                return;
            }
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
                rbSql.Text = "";
                sqlLs = new List<string>();

                string sql = "";
                List<int[]> col_error2 = new List<int[]>();

                if (_HREMPLOYEEB_excel == null)
                {
                    WGMessage.ShowWarning("请选择职工信息体检要求文件!");
                    return;
                }
                else
                {
                    sql = "SELECT * FROM HREMPLOYEE";
                    _HREMPLOYEE_DB = FillDatatable(sql, Main.CONN_Public);

                    sql = @"SELECT * FROM HRPHYSICAL WHERE ST=1";
                    _HRPHYSICAL_DB = FillDatatable(sql, Main.CONN_Public);

                    DataTable dt_error2 = _HREMPLOYEEB_excel.Clone();
                    //重复数据
                    DataTable dt_repet2 = _HREMPLOYEEB_excel.Clone();

                    string Mguid = "";
                    string Fguid = "";
                    #region 子
                    for (int i = 0; i < _HREMPLOYEEB_excel.Rows.Count; i++)
                    {
                        bool isError2 = false;
                        bool isRepet2 = false;

                        DataRow dr_excel = _HREMPLOYEEB_excel.Rows[i];

                        if (string.IsNullOrWhiteSpace(dr_excel["员工工号"].ToString()))
                        {
                            //空
                            col_error2.Add(new int[] { dt_error2.Rows.Count, 0 });
                            isError2 = true;
                        }
                        else
                        {
                            DataRow[] drs = _HREMPLOYEE_DB.Select("EMPCODE='" + ReturnString(dr_excel["员工工号"].ToString()) + "'");
                            if (drs.Length > 0)
                            {
                                Mguid = drs[0]["GUID"].ToString();
                            }
                            else
                            {
                                col_error2.Add(new int[] { dt_error2.Rows.Count, 0 });
                                isError2 = true;
                            }
                        }

                        int numa = 0;
                        if (string.IsNullOrWhiteSpace(dr_excel["序号"].ToString()) || !int.TryParse(dr_excel["序号"].ToString(), out numa))
                        {
                            col_error2.Add(new int[] { dt_error2.Rows.Count, 1 });
                            isError2 = true;
                        }

                        if (string.IsNullOrWhiteSpace(dr_excel["体检要求"].ToString()))
                        {
                            col_error2.Add(new int[] { dt_error2.Rows.Count, 2 });
                            isError2 = true;
                        }
                        else
                        {
                            DataRow[] drss = _HRPHYSICAL_DB.Select("NAME='" + ReturnString(dr_excel["体检要求"].ToString()) + "'");
                            if (drss.Length > 0)
                            {
                                Fguid = drss[0]["GUID"].ToString();
                            }
                            else
                            {
                                col_error2.Add(new int[] { dt_error2.Rows.Count, 2 });
                                isError2 = true;
                            }
                        }

                        DataRow[] drss1 = _HREMPLOYEEB_excel.Select(string.Format("员工工号='{0}' AND 体检要求='{1}'", ReturnString(dr_excel["员工工号"].ToString()), ReturnString(dr_excel["体检要求"].ToString())));
                        if (drss1.Length > 1)
                        {
                            isRepet2 = true;
                        }

                        if (isError2 || isRepet2)
                        {
                            if (isError2)
                            {
                                dt_error2.Rows.Add(dr_excel.ItemArray);
                            }
                            if (isRepet2)
                            {
                                dt_repet2.Rows.Add(dr_excel.ItemArray);
                            }
                            continue;
                        }

                        string temp = string.Format(@" INSERT INTO HREMPLOYEEB (GUID,PGUID,SNO,FGUID) 
                                                VALUES(NEWID(),'{0}',{1},'{2}')",
                                                    Mguid, numa, Fguid);

                        rbSql.Text += temp + Environment.NewLine;
                        sqlLs.Add(temp);
                    }
                    #endregion
                    dgError2.DataSource = dt_error2;
                    dgRepet2.DataSource = dt_repet2;
                    if (dt_error2.Rows.Count > 0 || dt_repet2.Rows.Count > 0)
                    {
                        Main.SetErrorCell(dgError2, col_error2);
                        rbSql.Text = "";
                        isCheck = false;
                        return;
                    }
                    isCheck = true;

                }

            }
            catch (Exception ex)
            {
                WGMessage.ShowAsterisk("出现未知异常！请检查Excel文件正确性！" + ex.ToString());
                return;
            }
        }

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
