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
    public partial class HREMPLOYEES : Form
    {
        public HREMPLOYEES()
        {
            InitializeComponent();
        }

        #region  变量
        ExcelManager _excelManager = new ExcelManager();

        DataTable _HREMPLOYEE_DB = null;

        DataTable _HREMPLOYEES_excel = null;

        DataTable _HREMPLOYEES_DB = null;

        DataTable _BSDEPT_DB = null;//部门

        DataTable _BSDEPTPOS_DB = null;//岗位

        DataTable _BSWORKSHOP_DB = null;//车间

        DataTable _ECTYPE_DB = null;//设备类别

        DataTable _ECINFO_DB = null;//设备信息

        DataTable _HRWORKLICENSE_DB = null;//上岗证

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
                    _HREMPLOYEES_excel = _excelManager.GetExcelTableByOleDB(opfDialog0.FileName, "HREMPLOYEES");
                    MessageBox.Show("读取笔数：" + _HREMPLOYEES_excel.Rows.Count + "");
                }
                catch
                {

                }
                //ClearSql();
                if (_HREMPLOYEES_excel == null || _HREMPLOYEES_excel.Rows.Count <= 0)
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

                if (_HREMPLOYEES_excel == null)
                {
                    WGMessage.ShowWarning("请选择职工信息上岗证文件!");
                    return;
                }
                else
                {
                    sql = "SELECT * FROM HREMPLOYEE";
                    _HREMPLOYEE_DB = FillDatatable(sql, Main.CONN_Public);

                    sql = @"SELECT * FROM BSWORKSHOP";
                    _BSWORKSHOP_DB = FillDatatable(sql, Main.CONN_Public);

                    sql = @"SELECT * FROM ECTYPE";
                    _ECTYPE_DB = FillDatatable(sql, Main.CONN_Public);

                    sql = @"SELECT * FROM ECINFO";
                    _ECINFO_DB = FillDatatable(sql, Main.CONN_Public);

                    sql = @"SELECT * FROM HRWORKLICENSE";
                    _HRWORKLICENSE_DB = FillDatatable(sql, Main.CONN_Public);

                    DataTable dt_error2 = _HREMPLOYEES_excel.Clone();
                    //重复数据
                    DataTable dt_repet2 = _HREMPLOYEES_excel.Clone();
                    #region 子
                    string BWguid = "";
                    string ETGuid = "";
                    string ECguid = "";
                    string Mguid = "";
                    string Fguid = "";
                    string EC = "";
                    int LVL = 0;
                    Dictionary<string, int> dicST = new Dictionary<string, int>();
                    dicST.Add("试用", 0);
                    dicST.Add("待批", 1);
                    dicST.Add("正常", 2);
                    dicST.Add("冻结", 3);
                    dicST.Add("失效", 4);
                    List<int> lstLVL = new List<int>() { 1, 2, 3, 4 };
                    for (int i = 0; i < _HREMPLOYEES_excel.Rows.Count; i++)
                    {
                        bool isError2 = false;
                        bool isRepet2 = false;

                        DataRow dr_excel = _HREMPLOYEES_excel.Rows[i];

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

                        int numa;
                        if (string.IsNullOrWhiteSpace(dr_excel["序号"].ToString()) || !int.TryParse(dr_excel["序号"].ToString(), out numa))
                        {
                            col_error2.Add(new int[] { dt_error2.Rows.Count, 1 });
                            isError2 = true;
                        }

                        if (string.IsNullOrWhiteSpace(dr_excel["车间编号"].ToString()))
                        {
                            col_error2.Add(new int[] { dt_error2.Rows.Count, 2 });
                            isError2 = true;
                        }
                        else
                        {
                            DataRow[] drss = _BSWORKSHOP_DB.Select("CODE='" + ReturnString(dr_excel["车间编号"].ToString()) + "'");
                            if (drss.Length > 0)
                            {
                                BWguid = drss[0]["GUID"].ToString();
                            }
                            else
                            {
                                col_error2.Add(new int[] { dt_error2.Rows.Count, 2 });
                                isError2 = true;
                            }
                        }

                        if (string.IsNullOrWhiteSpace(dr_excel["设备类别编号"].ToString()))
                        {
                            //空
                            col_error2.Add(new int[] { dt_error2.Rows.Count, 3 });
                            isError2 = true;
                        }
                        else
                        {
                            DataRow[] drss = _ECTYPE_DB.Select("CODE='" + ReturnString(dr_excel["设备类别编号"].ToString()) + "'");
                            if (drss.Length > 0)
                            {
                                ETGuid = drss[0]["GUID"].ToString();
                            }
                            else
                            {
                                col_error2.Add(new int[] { dt_error2.Rows.Count, 3 });
                                isError2 = true;
                            }
                        }



                        if (string.IsNullOrEmpty(dr_excel["设备编号"].ToString()))
                        {
                            EC = " AND 设备编号 is null";
                            ECguid = " AND CGUID is null";
                        }
                        else
                        {
                            DataRow[] drss = _ECINFO_DB.Select(string.Format("CODE='{0}' AND AGUID='{1}'", ReturnString(dr_excel["设备编号"].ToString()), ETGuid == "" ? Guid.Empty.ToString() : ETGuid));
                            if (drss.Length > 0)
                            {
                                ECguid = "AND CGUID='" + drss[0]["GUID"].ToString() + "'";
                                EC = " AND 设备编号='" + drss[0]["CODE"].ToString() + "'";
                            }
                            else
                            {
                                col_error2.Add(new int[] { dt_error2.Rows.Count, 4 });
                                isError2 = true;
                            }
                        }


                        if (string.IsNullOrWhiteSpace(dr_excel["星级"].ToString()) || !int.TryParse(dr_excel["星级"].ToString(), out numa))
                        {
                            //空
                            col_error2.Add(new int[] { dt_error2.Rows.Count, 5 });
                            isError2 = true;
                        }
                        else
                        {
                            if (lstLVL.Contains(numa))
                            {
                                LVL = numa;
                            }
                            else
                            {
                                LVL = 0;
                                col_error2.Add(new int[] { dt_error2.Rows.Count, 5 });
                                isError2 = true;
                            }

                            DataRow[] drss = _HREMPLOYEES_excel.Select(string.Format("员工工号='{0}' AND 车间编号='{1}' AND 设备类别编号='{2}' {3}", ReturnString(dr_excel["员工工号"].ToString()), ReturnString(dr_excel["车间编号"].ToString()), ReturnString(dr_excel["设备类别编号"].ToString()), EC));
                            if (drss.Length > 1)
                            {
                                isRepet2 = true;
                            }
                            else
                            {
                                drss = _HRWORKLICENSE_DB.Select(string.Format("AGUID='{0}' AND BGUID='{1}' {2}", BWguid, ETGuid, ECguid));
                                if (drss.Length > 0)
                                {
                                    Fguid = drss[0]["GUID"].ToString();
                                }
                                else
                                {
                                    col_error2.Add(new int[] { dt_error2.Rows.Count, 4 });
                                    isError2 = true;
                                }
                            }
                        }

                        DateTime DATE = new DateTime();
                        if (string.IsNullOrWhiteSpace(dr_excel["上岗证颁发日期"].ToString()) || !DateTime.TryParse(dr_excel["上岗证颁发日期"].ToString(), out DATE))
                        {
                            col_error2.Add(new int[] { dt_error2.Rows.Count, 6 });
                            isError2 = true;
                        }

                        string EDATE = "";
                        if (string.IsNullOrWhiteSpace(dr_excel["最后一次上岗时间"].ToString()))
                        {
                            EDATE = "null";
                        }
                        else
                        {
                            if (DateTime.TryParse(dr_excel["最后一次上岗时间"].ToString(), out DATE))
                            {
                                EDATE = "'" + DATE.ToString() + "'";
                            }
                            else
                            {
                                col_error2.Add(new int[] { dt_error2.Rows.Count, 7 });
                                isError2 = true;
                            }
                        }

                        string ADATE = "";
                        if (string.IsNullOrWhiteSpace(dr_excel["冻结时间/失效时间"].ToString()))
                        {
                            ADATE = "null";
                        }
                        else
                        {
                            if (DateTime.TryParse(dr_excel["冻结时间/失效时间"].ToString(), out DATE))
                            {
                                ADATE = "'" + DATE.ToString() + "'";
                            }
                            else
                            {
                                col_error2.Add(new int[] { dt_error2.Rows.Count, 8 });
                                isError2 = true;
                            }
                        }

                        string CDATE = "";
                        if (string.IsNullOrWhiteSpace(dr_excel["试用期"].ToString()))
                        {
                            CDATE = "null";
                        }
                        else
                        {
                            if (DateTime.TryParse(dr_excel["试用期"].ToString(), out DATE))
                            {
                                CDATE = "'" + DATE.ToString() + "'";
                            }
                            else
                            {
                                col_error2.Add(new int[] { dt_error2.Rows.Count, 9 });
                                isError2 = true;
                            }
                        }

                        if (string.IsNullOrWhiteSpace(dr_excel["状态"].ToString()))
                        {
                            //空
                            col_error2.Add(new int[] { dt_error2.Rows.Count, 10 });
                            isError2 = true;
                        }
                        else
                        {
                            if (!dicST.ContainsKey(dr_excel["状态"].ToString()))
                            {
                                //空
                                col_error2.Add(new int[] { dt_error2.Rows.Count, 10 });
                                isError2 = true;
                            }
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

                        string temp = string.Format(@" INSERT INTO HREMPLOYEES (GUID,PGUID,FGUID,LVL,ADATE,CDATE,DDATE,EDATE,ISREMIND,ISTRIAL,ST) 
                                                VALUES(NEWID(),'{0}','{1}',{2},{3},{4},'{5}',{6},'0','0',{7})",
                                                    Mguid, Fguid, LVL, ADATE, CDATE, dr_excel["上岗证颁发日期"], EDATE, dicST[dr_excel["状态"].ToString()]);

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

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (isCheck)
                {
                    WGMessage.ShowAsterisk("已验证，不用重复验证！");
                    return;
                }

                if (_HREMPLOYEES_excel == null)
                {
                    WGMessage.ShowWarning("请选择职工上岗证文件!");
                    return;
                }
                bool isError = false;
                string message = "";
                foreach (DataRow dr in _HREMPLOYEE_DB.Rows)
                {
                    if (!UpdateSkillPoint(dr["GUID"].ToString()))
                    {
                        isError = true;
                        message += dr["员工编码"] + "技能分更新不成功" + Environment.NewLine;
                    }
                }
                if (isError)
                {
                    WGMessage.ShowAsterisk(message);
                }
                WGMessage.ShowAsterisk("技能分更新成功！");
            }
            catch (Exception ex)
            {

            }
        }

        /// <summary>
        /// 技能分更新
        /// </summary>
        /// <param name="guid"></param>
        /// <returns></returns>
        public bool UpdateSkillPoint(string guid)
        {
            bool isResult = false;
            int off = GetHRSkillPoint();
            DataTable dtLog = GetSikillPointLog(guid);
            decimal sum = 0;
            foreach (DataRow dr in dtLog.Rows)
            {
                decimal maxnum = Convert.ToDecimal(dr["MAXNUM"].ToString());
                decimal sumnum = Convert.ToDecimal(dr["SUMNUM"].ToString());
                if (dr["COUNTNUM"].ToString() == "1")
                {
                    sum += maxnum;
                }
                else
                {
                    sum += maxnum + (sumnum - maxnum) * off / 100;
                }
            }

            StringBuilder sb = new StringBuilder();
            sb.Append(string.Format(@"UPDATE HREMPLOYEE SET SKILLPOINTS={0} WHERE GUID='{1}'", sum, guid));
            using (SqlConnection conn = new SqlConnection(Main.CONN_Public))
            {
                SqlCommand cmd = new SqlCommand(sb.ToString());
                cmd.Connection = conn;
                conn.Open();
                if (cmd.ExecuteNonQuery() > 0)
                {
                    isResult = true;
                }
                else
                {
                    isResult = false;
                }
                conn.Close();
            }
            return isResult;
        }

        private int GetHRSkillPoint()
        {
            int result = 100;
            DataTable dt = GetSsparam("HRSkillPoint");
            if (dt != null && dt.Rows.Count > 0)
            {
                result = int.Parse(dt.Rows[0]["PVAL"].ToString());
            }
            return result;
        }

        private DataTable GetSikillPointLog(string guid)
        {
            string sql = string.Format(@"(SELECT COUNT(HREMPLOYEES.GUID) COUNTNUM,MAX(HRWORKLICENSES.SPNUM) MAXNUM,
SUM(HRWORKLICENSES.SPNUM) SUMNUM,HRLICENSETYPES.PGUID
FROM HREMPLOYEES 
LEFT JOIN HRWORKLICENSE ON HRWORKLICENSE.GUID=HREMPLOYEES.FGUID
LEFT JOIN HRWORKLICENSES ON  HRWORKLICENSES.PGUID=HRWORKLICENSE.GUID
LEFT JOIN HRLICENSETYPES ON HRLICENSETYPES.FGUID=HRWORKLICENSE.GUID
WHERE HRLICENSETYPES.PGUID IS NOT NULL AND HREMPLOYEES.PGUID='{0}' and HREMPLOYEES.LVL=HRWORKLICENSES.LVL and HREMPLOYEES.ST in(0,2)
GROUP BY HRLICENSETYPES.PGUID)
UNION ALL
(SELECT COUNT(HREMPLOYEES.GUID) COUNTNUM,MAX(HRWORKLICENSES.SPNUM) MAXNUM,
SUM(HRWORKLICENSES.SPNUM) SUMNUM,HRLICENSETYPES.PGUID
FROM HREMPLOYEES 
LEFT JOIN HRWORKLICENSE ON HRWORKLICENSE.GUID=HREMPLOYEES.FGUID
LEFT JOIN HRWORKLICENSES ON HRWORKLICENSES.PGUID=HRWORKLICENSE.GUID
LEFT JOIN HRLICENSETYPES ON HRLICENSETYPES.FGUID=HRWORKLICENSE.GUID
WHERE HRLICENSETYPES.PGUID IS NULL AND HREMPLOYEES.PGUID='{0}' and HREMPLOYEES.LVL=HRWORKLICENSES.LVL and HREMPLOYEES.ST in(0,2)
GROUP BY HRLICENSETYPES.PGUID,HREMPLOYEES.GUID)", guid);
            return FillDatatable(sql, Main.CONN_Public);
        }

        /// <summary>
        /// 得到系统参数
        /// </summary>
        /// <returns></returns>
        public DataTable GetSsparam(string PKEY)
        {
            string sql = string.Format("SELECT PVAL FROM SSPARAM WHERE PKEY='{0}'", PKEY);
            DataTable dt = FillDatatable(sql, Main.CONN_Public);
            return dt;
        }
    }
}
