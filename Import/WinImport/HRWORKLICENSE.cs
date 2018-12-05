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
    public partial class HRWORKLICENSE : Form
    {
        #region  变量
        ExcelManager _excelManager = new ExcelManager();

        DataTable _HRWORKLICENSE_excel = null;

        DataTable _HRWORKLICENSE_DB = null;

        DataTable _HRWORKLICENSES_excel = null;

        DataTable _HRWORKLICENSES_DB = null;

        DataTable _BSWORKSHOP_DB = null;//车间

        DataTable _ECTYPE_DB = null;//设备类别

        DataTable _ECINFO_DB = null;//设备信息
        /// <summary>
        ///  是否验证成功
        /// </summary>
        bool isCheck = false;

        /// <summary>
        /// 需要保存的sql
        /// </summary>
        List<string> sqlLs = new List<string>();
        #endregion

        #region 构造函数
        public HRWORKLICENSE()
        {
            InitializeComponent();
        }
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

        public void ClearSql()
        {
            //重新上传后，清空原来的
            isCheck = false;
            sqlLs = new List<string>();
            rbSql.Text = "";

            dgError1.DataSource = new DataTable();
            dgRepet1.DataSource = new DataTable();
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

        private void btnSelect_Click(object sender, EventArgs e)
        {
            //得到excel数据源
            if (opfDialog0.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    txtFileMain.Text = opfDialog0.FileName;
                    _HRWORKLICENSE_excel = _excelManager.GetExcelTableByOleDB(opfDialog0.FileName, "HRWORKLICENSE");
                    _HRWORKLICENSE_excel.Columns.Add("GUID");
                    MessageBox.Show("读取笔数：" + _HRWORKLICENSE_excel.Rows.Count + "");
                }
                catch
                { }
                if (_HRWORKLICENSE_excel == null || _HRWORKLICENSE_excel.Rows.Count <= 0)
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

        private void button1_Click(object sender, EventArgs e)
        {
            //得到excel数据源
            if (opfDialog0.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    txtFileSon.Text = opfDialog0.FileName;
                    _HRWORKLICENSES_excel = _excelManager.GetExcelTableByOleDB(opfDialog0.FileName, "HRWORKLICENSES");
                    MessageBox.Show("读取笔数：" + _HRWORKLICENSES_excel.Rows.Count + "");
                }
                catch
                {

                }
                //ClearSql();
                if (_HRWORKLICENSES_excel == null || _HRWORKLICENSES_excel.Rows.Count <= 0)
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

                if (_HRWORKLICENSE_excel == null)
                {
                    WGMessage.ShowWarning("请选择上岗证文件!");
                    return;
                }
                if (_HRWORKLICENSES_excel == null)
                {
                    WGMessage.ShowWarning("请选择上岗证技能分文件!");
                    return;
                }

                string sql = "";

                sql = "SELECT * FROM HRWORKLICENSE";
                _HRWORKLICENSE_DB = FillDatatablde(sql, Main.CONN_Public);

                sql = "SELECT * FROM HRWORKLICENSES";
                _HRWORKLICENSES_DB = FillDatatablde(sql, Main.CONN_Public);


                sql = @"SELECT * FROM BSWORKSHOP";
                _BSWORKSHOP_DB = FillDatatablde(sql, Main.CONN_Public);

                sql = @"SELECT * FROM ECTYPE";
                _ECTYPE_DB = FillDatatablde(sql, Main.CONN_Public);

                sql = @"SELECT * FROM ECINFO";
                _ECINFO_DB = FillDatatablde(sql, Main.CONN_Public);


                //错误
                List<int[]> col_error1 = new List<int[]>();
                List<int[]> col_error2 = new List<int[]>();
                List<int[]> col_error3 = new List<int[]>();
                DataTable dt_error1 = _HRWORKLICENSE_excel.Clone();

                //重复数据
                DataTable dt_repet1 = _HRWORKLICENSE_excel.Clone();
                DataTable dt_datarepet1 = _HRWORKLICENSE_excel.Clone();

                DataTable dt_lack1 = _HRWORKLICENSE_excel.Clone();

                DataTable dt_error2 = _HRWORKLICENSES_excel.Clone();

                //重复数据
                DataTable dt_repet2 = _HRWORKLICENSES_excel.Clone();

                DataTable dt_lack2 = _HRWORKLICENSES_excel.Clone();

                Dictionary<string, string> doclist = new Dictionary<string, string>();

                #region 主
                string Aguid = "";
                string Bguid = "";
                string Cguid = "";
                for (int i = 0; i < _HRWORKLICENSE_excel.Rows.Count; i++)
                {
                    bool isError1 = false;
                    bool isRepet1 = false;
                    bool islack1 = false;
                    bool isDataRepet1 = false;

                    DataRow dr_excel = _HRWORKLICENSE_excel.Rows[i];

                    if (string.IsNullOrWhiteSpace(dr_excel["上岗证名称"].ToString()))
                    {
                        //空
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 0 });
                        isError1 = true;
                    }
                    else
                    {
                        DataRow[] drss = _HRWORKLICENSE_excel.Select("上岗证名称 = '" + ReturnString(dr_excel["上岗证名称"].ToString()) + "'");
                        if (drss.Length > 1)
                        {
                            isRepet1 = true;
                        }

                        DataRow[] drss1 = _HRWORKLICENSE_DB.Select(string.Format("NAME='{0}'", ReturnString(dr_excel["上岗证名称"].ToString())));
                        if (drss1.Length > 0)
                        {
                            isDataRepet1 = true;
                        }
                    }

                    if (string.IsNullOrWhiteSpace(dr_excel["车间编号"].ToString()))
                    {
                        //空
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 1 });
                        isError1 = true;
                    }
                    else
                    {
                        DataRow[] drss = _BSWORKSHOP_DB.Select("CODE='" + ReturnString(dr_excel["车间编号"].ToString()) + "'");
                        if (drss.Count() == 0)
                        {
                            col_error1.Add(new int[] { dt_error1.Rows.Count, 1 });
                            isError1 = true;
                        }
                        else
                        {
                            Aguid = drss[0]["GUID"].ToString();
                        }
                    }

                    if (string.IsNullOrWhiteSpace(dr_excel["设备类别编号"].ToString()))
                    {
                        //空
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 2 });
                        isError1 = true;
                    }
                    else
                    {
                        DataRow[] drss = _ECTYPE_DB.Select("CODE='" + ReturnString(dr_excel["设备类别编号"].ToString()) + "'");
                        if (drss.Count() == 0)
                        {
                            col_error1.Add(new int[] { dt_error1.Rows.Count, 2 });
                            isError1 = true;
                        }
                        else
                        {
                            Bguid = drss[0]["GUID"].ToString();
                        }
                    }

                    if (string.IsNullOrWhiteSpace(dr_excel["设备编号"].ToString()))
                    {
                        DataRow[] drss = _HRWORKLICENSE_excel.Select("车间编号 = '" + ReturnString(dr_excel["车间编号"].ToString()) + "' AND 设备类别编号='" + ReturnString(dr_excel["设备类别编号"].ToString()) + "' AND 设备编号='' ");
                        if (drss.Length > 1)
                        {
                            isRepet1 = true;
                        }

                        drss = _HRWORKLICENSE_DB.Select("AGUID = '" + Aguid + "' AND BGUID='" + Bguid + "' AND CGUID IS NULL");
                        if (drss.Length > 0)
                        {
                            isDataRepet1 = true;
                        }
                    }
                    else
                    {
                        DataRow[] drss = _HRWORKLICENSE_excel.Select("车间编号 = '" + ReturnString(dr_excel["车间编号"].ToString()) + "' AND 设备类别编号='" + ReturnString(dr_excel["设备类别编号"].ToString()) + "' AND 设备编号='" + ReturnString(dr_excel["设备编号"].ToString()) + "'");
                        if (drss.Length > 1)
                        {
                            isRepet1 = true;
                        }
                        else
                        {
                            drss = _HRWORKLICENSE_DB.Select("AGUID = '" + Aguid + "' AND BGUID='" + Bguid + "' AND CGUID='" + Cguid + "'");
                            if (drss.Length > 0)
                            {
                                isDataRepet1 = true;
                            }
                            else
                            {
                                drss = _ECINFO_DB.Select("CODE='" + ReturnString(dr_excel["设备编号"].ToString()) + "' AND AGUID='" + Bguid + "'");
                                if (drss.Length > 0)
                                {
                                    Cguid = drss[0]["GUID"].ToString();
                                }
                                else
                                {
                                    col_error1.Add(new int[] { dt_error1.Rows.Count, 3 });
                                    isError1 = true;
                                }
                            }
                        }
                    }
                    DataRow[] drs1 = _HRWORKLICENSES_excel.Select(string.Format(@"车间编号='{0}' AND 设备类别编号='{1}'{2}", ReturnString(dr_excel["车间编号"].ToString()), ReturnString(dr_excel["设备类别编号"].ToString()), Cguid == "" ? " AND 设备编号 IS NULL" : " AND 设备编号='" + dr_excel["设备编号"] + "'"));

                    if (drs1.Length == 0)
                    {
                        islack1 = true;
                    }

                    string HRWORKLICENSE_GUID = Guid.NewGuid().ToString();

                    _HRWORKLICENSE_excel.Rows[i]["GUID"] = HRWORKLICENSE_GUID;


                    if (isError1 || isRepet1||islack1||isDataRepet1)
                    {
                        if (isError1)
                        {
                            dt_error1.Rows.Add(dr_excel.ItemArray);
                        }
                        if (isRepet1)
                        {
                            dt_repet1.Rows.Add(dr_excel.ItemArray);
                        }
                        if (islack1)
                        {
                            dt_lack1.Rows.Add(dr_excel.ItemArray);
                        }
                        if (isDataRepet1)
                        {
                            dt_datarepet1.Rows.Add(dr_excel.ItemArray);
                        }
                        continue;
                    }

                    try
                    {
                        string temp = string.Format(@"INSERT INTO [HRWORKLICENSE]
                                       ([GUID],[NAME],[AGUID]
                                       ,[BGUID],[CGUID],[NOTE]
                                       ,[ST],[CC],[ND],[CD])
                                        VALUES ('{0}','{1}','{2}','{3}',{4},'{5}',1,'1023','{6}','{6}')",
                                        HRWORKLICENSE_GUID, ReturnString(dr_excel["上岗证名称"].ToString()), Aguid,
                                        Bguid, Cguid == "" ? "null" : "'" + Cguid + "'", ReturnString(dr_excel["备注"].ToString()), DateTime.Now.ToString());

                        rbSql.Text += temp + Environment.NewLine;
                        sqlLs.Add(temp);
                    }
                    catch
                    { }
                }


                #endregion

                #region 子
                string Mguid = "";
                string EC = "";
                for (int i = 0; i < _HRWORKLICENSES_excel.Rows.Count; i++)
                {
                    bool isError2 = false;
                    bool isRepet2 = false;
                    bool isLack2 = false;

                    DataRow dr_excel = _HRWORKLICENSES_excel.Rows[i];

                    if (string.IsNullOrWhiteSpace(dr_excel["车间编号"].ToString()))
                    {
                        //空
                        col_error2.Add(new int[] { dt_error2.Rows.Count, 0 });
                        isError2 = true;
                    }

                    if (string.IsNullOrWhiteSpace(dr_excel["设备类别编号"].ToString()))
                    {
                        //空
                        col_error2.Add(new int[] { dt_error2.Rows.Count, 1 });
                        isError2 = true;
                    }

                    if (string.IsNullOrEmpty(dr_excel["设备编号"].ToString()))
                    {
                        EC = " AND 设备编号 is null";
                    }
                    else
                    {
                        EC = " AND 设备编号='" + dr_excel["设备编号"].ToString() + "'";
                    }

                    int numa;

                    if (!int.TryParse(dr_excel["星级"].ToString(), out numa))
                    {
                        //空
                        col_error2.Add(new int[] { dt_error2.Rows.Count, 3 });
                        isError2 = true;
                    }
                    else
                    {
                        if (numa <= 0 || numa >= 5)
                        {
                            col_error2.Add(new int[] { dt_error2.Rows.Count, 3 });
                            isError2 = true;
                        }
                        DataRow[] drss = _HRWORKLICENSES_excel.Select("车间编号='" + ReturnString(dr_excel["车间编号"].ToString()) + "' AND 设备类别编号='" + ReturnString(dr_excel["设备类别编号"].ToString()) + "' " + EC + " AND 星级=" + dr_excel["星级"].ToString());
                        if (drss.Length > 1)
                        {
                            isRepet2 = true;
                        }
                        else
                        {
                            drss = _HRWORKLICENSE_excel.Select("车间编号='" + ReturnString(dr_excel["车间编号"].ToString()) + "' AND 设备类别编号='" + ReturnString(dr_excel["设备类别编号"].ToString()) + "' " + EC + "");
                            if (drss.Length > 0)
                            {
                                Mguid = drss[0]["GUID"].ToString();
                            }
                            else
                            {
                                isError2 = true;
                                col_error2.Add(new int[] { dt_error2.Rows.Count, 3 });
                            }
                        }
                    }

                    if (!int.TryParse(dr_excel["技能分"].ToString(), out numa))
                    {
                        //空
                        col_error2.Add(new int[] { dt_error2.Rows.Count, 4 });
                        isError2 = true;
                    }

                    DataRow[] drs = _HRWORKLICENSES_excel.Select("车间编号='" + ReturnString(dr_excel["车间编号"].ToString()) + "' AND 设备类别编号='" + ReturnString(dr_excel["设备类别编号"].ToString()) + "' " + EC);
                    List<int> lst = new List<int>();
                    foreach (DataRow dr in drs)
                    {
                        if (int.TryParse(dr["星级"].ToString(), out numa) && !lst.Contains(numa))
                        {
                            lst.Add(numa);
                        }
                    }
                    if (lst.Count != 4)
                    {
                        isLack2 = true;
                    }

                    if (isError2 || isRepet2 || isLack2)
                    {
                        if (isError2)
                        {
                            dt_error2.Rows.Add(dr_excel.ItemArray);
                        }
                        if (isRepet2)
                        {
                            dt_repet2.Rows.Add(dr_excel.ItemArray);
                        }
                        if (isLack2)
                        {
                            dt_lack2.Rows.Add(dr_excel.ItemArray);
                        }
                        continue;
                    }

                    string temp = string.Format(@" INSERT INTO HRWORKLICENSES (GUID,PGUID,LVL,SPNUM) VALUES(NEWID(),'{0}',{1},{2})", Mguid, dr_excel["星级"], dr_excel["技能分"]);

                    rbSql.Text += temp + Environment.NewLine;
                    sqlLs.Add(temp);
                }
                #endregion


                dt_error1.Columns.Remove("GUID");
                dt_repet1.Columns.Remove("GUID");
                dt_lack1.Columns.Remove("GUID");
                dt_datarepet1.Columns.Remove("GUID");
                dgError1.DataSource = dt_error1; dgError2.DataSource = dt_error2; dgLack1.DataSource = dt_lack1;
                dgRepet1.DataSource = dt_repet1; dgRepet2.DataSource = dt_repet2; dgLack2.DataSource = dt_lack2;
                dgDataRepet1.DataSource = dt_datarepet1;
                if (dt_error1.Rows.Count > 0 || dt_error2.Rows.Count > 0 || dt_repet1.Rows.Count > 0 || dt_repet2.Rows.Count > 0 || dt_lack2.Rows.Count > 0 || dt_lack1.Rows.Count > 0 || dt_datarepet1.Rows.Count > 0)  
                {
                    Main.SetErrorCell(dgError1, col_error1);
                    Main.SetErrorCell(dgError2, col_error2);
                    rbSql.Text = "";
                    isCheck = false;
                    return;
                }
                isCheck = true;
            }
            catch (Exception ex)
            {
                WGMessage.ShowAsterisk("出现未知异常！请检查2个Excel文件正确性和顺序的正确性！" + ex.ToString());
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
