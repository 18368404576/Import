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
    public partial class BSPROCESS : Form
    {
        public BSPROCESS()
        {
            InitializeComponent();
        }

        #region 变量
        ExcelManager _excelManager = new ExcelManager();
        DataTable _BSPROCESS_excel = null;
        DataTable _BSPROCESS_DB = null;

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

            dgError1.DataSource = new DataTable();
            dgRepet1.DataSource = new DataTable();
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
                    _BSPROCESS_excel = _excelManager.GetExcelTableByOleDB(opfDialog0.FileName, "BSPROCESS");
                    _BSPROCESS_excel.Columns.Add("GUID");
                    MessageBox.Show("读取笔数：" + _BSPROCESS_excel.Rows.Count + "");
                }
                catch
                { }
                if (_BSPROCESS_excel == null || _BSPROCESS_excel.Rows.Count <= 0)
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

                if (_BSPROCESS_excel == null)
                {
                    WGMessage.ShowWarning("请选择工艺信息文件!");
                    return;
                }

                Dictionary<string, bool> dic = new Dictionary<string, bool>();
                dic.Add("是", true);
                dic.Add("否", false);

                Dictionary<string, int> dicST = new Dictionary<string, int>();
                dicST.Add("未启用", 0);
                dicST.Add("已启用", 1);
                dicST.Add("已停用", 2);

                string sql = "";
                sql = "SELECT * FROM BSPROCESS";
                _BSPROCESS_DB = FillDatatable(sql, Main.CONN_Public);
             
                //错误
                List<int[]> col_error1 = new List<int[]>();
                DataTable dt_error1 = _BSPROCESS_excel.Clone();

                //重复数据
                DataTable dt_repet1 = _BSPROCESS_excel.Clone();
                DataTable dt_datarepet1 = _BSPROCESS_excel.Clone();


                #region 主

                for (int i = 0; i < _BSPROCESS_excel.Rows.Count; i++)
                {
                    bool isError1 = false;
                    bool isRepet1 = false;
                    bool isDataRepet1 = false;

                    DataRow dr_excel = _BSPROCESS_excel.Rows[i];

                    if (string.IsNullOrWhiteSpace(dr_excel["编码"].ToString()))
                    {
                        //空
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 0 });
                        isError1 = true;
                    }
                    else
                    {
                        DataRow[] drss = _BSPROCESS_excel.Select("编码 = '" + ReturnString(dr_excel["编码"].ToString()) + "'");
                        if (drss.Length > 1)
                        {
                            isRepet1 = true;
                        }

                        DataRow[] drss1 = _BSPROCESS_DB.Select(string.Format("CODE='{0}'", ReturnString(dr_excel["编码"].ToString())));
                        if (drss1.Length > 0)
                        {
                            isDataRepet1 = true;
                        }
                    }

                    if (string.IsNullOrWhiteSpace(dr_excel["名称"].ToString()))
                    {
                        //空
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 1 });
                        isError1 = true;
                    }

                    if (string.IsNullOrWhiteSpace(dr_excel["是否组装"].ToString()))
                    {
                        //空
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 2 });
                        isError1 = true;
                    }
                    else
                    {
                        if (!dic.ContainsKey(dr_excel["是否组装"].ToString()))
                        {
                            col_error1.Add(new int[] { dt_error1.Rows.Count, 2 });
                            isError1 = true;
                        }
                    }

                    if (string.IsNullOrWhiteSpace(dr_excel["是否包装"].ToString()))
                    {
                        //空
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 3 });
                        isError1 = true;
                    }
                    else
                    {
                        if (!dic.ContainsKey(dr_excel["是否包装"].ToString()))
                        {
                            col_error1.Add(new int[] { dt_error1.Rows.Count, 3 });
                            isError1 = true;
                        }
                    }


                    if (string.IsNullOrWhiteSpace(dr_excel["状态"].ToString()))
                    {
                        //空
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 4 });
                        isError1 = true;
                    }
                    else
                    {
                        if (!dicST.ContainsKey(dr_excel["状态"].ToString()))
                        {
                            col_error1.Add(new int[] { dt_error1.Rows.Count, 4 });
                            isError1 = true;
                        }
                    }


                    string BSPROCESS_GUID = Guid.NewGuid().ToString();

                    _BSPROCESS_excel.Rows[i]["GUID"] = BSPROCESS_GUID;


                    if (isError1 || isRepet1 || isDataRepet1)
                    {
                        if (isError1)
                        {
                            dt_error1.Rows.Add(dr_excel.ItemArray);
                        }
                        if (isRepet1)
                        {
                            dt_repet1.Rows.Add(dr_excel.ItemArray);
                        }
                        if (isDataRepet1)
                        {
                            dt_datarepet1.Rows.Add(dr_excel.ItemArray);
                        }
                        continue;
                    }

                    try
                    {
                        string temp = string.Format(@"INSERT INTO [BSPROCESS]
                                       ([GUID],[CODE],[NAME],[NOTE],[ST],[ISPACKAGE],[ISASSEMBLE],[ND],[CD])
                                        VALUES ('{0}','{1}','{2}','{3}',{4},'{5}','{6}','{7}','{7}')",
                                        BSPROCESS_GUID, ReturnString(dr_excel["编码"].ToString()), ReturnString(dr_excel["名称"].ToString()), ReturnString(dr_excel["备注"].ToString()),
                                        dicST[dr_excel["状态"].ToString()], dic[dr_excel["是否组装"].ToString()], dic[dr_excel["是否包装"].ToString()],
                                        DateTime.Now.ToString());

                        rbSql.Text += temp + Environment.NewLine;
                        sqlLs.Add(temp);
                    }
                    catch
                    { }
                }


                #endregion



                dt_error1.Columns.Remove("GUID");
                dt_repet1.Columns.Remove("GUID");
                dt_datarepet1.Columns.Remove("GUID");

                dgError1.DataSource = dt_error1;
                dgRepet1.DataSource = dt_repet1;
                dgDataRepet1.DataSource = dt_datarepet1;
                if (dt_error1.Rows.Count > 0 || dt_repet1.Rows.Count > 0 || dt_datarepet1.Rows.Count > 0)
                {
                    Main.SetErrorCell(dgError1, col_error1);
                    rbSql.Text = "";
                    isCheck = false;
                    return;
                }
                isCheck = true;
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
