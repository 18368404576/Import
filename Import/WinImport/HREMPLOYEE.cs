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
    public partial class HREMPLOYEE : Form
    {
        #region  变量
        ExcelManager _excelManager = new ExcelManager();

        DataTable _HREMPLOYEE_excel = null;

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

        #region 构造函数
        public HREMPLOYEE()
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

        private void btnSelect_Click(object sender, EventArgs e)
        {
            //得到excel数据源
            if (opfDialog0.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    txtFileMain.Text = opfDialog0.FileName;
                    _HREMPLOYEE_excel = _excelManager.GetExcelTableByOleDB(opfDialog0.FileName, "HREMPLOYEE");
                    _HREMPLOYEE_excel.Columns.Add("GUID");
                    MessageBox.Show("读取笔数：" + _HREMPLOYEE_excel.Rows.Count + "");
                }
                catch
                { }
                if (_HREMPLOYEE_excel == null || _HREMPLOYEE_excel.Rows.Count <= 0)
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

                if (_HREMPLOYEE_excel == null)
                {
                    WGMessage.ShowWarning("请选择职工信息文件!");
                    return;
                }
                

                Dictionary<string, int> dic = new Dictionary<string, int>();
                dic.Add("正式工", 1);
                dic.Add("派遣工", 2);
                dic.Add("其他", 3);

                Dictionary<string, bool> dicSex = new Dictionary<string, bool>();
                dicSex.Add("男", true);
                dicSex.Add("女", false);
                string sql = "";

                Dictionary<string, bool> dicISSTOP = new Dictionary<string, bool>();
                dicISSTOP.Add("是", true);
                dicISSTOP.Add("否", false);

                sql = "SELECT * FROM HREMPLOYEE";
                _HREMPLOYEE_DB = FillDatatable(sql, Main.CONN_Public);

                //sql = "SELECT * FROM HREMPLOYEES";
                //_HREMPLOYEES_DB = FillDatatable(sql, Main.CONN_Public);


                sql = @"SELECT * FROM BSDEPT";
                _BSDEPT_DB = FillDatatable(sql, Main.CONN_Public);

                sql = @"SELECT BSDEPTA.*,BSPOSITION.CODE FROM BSDEPTA LEFT JOIN BSPOSITION ON BSPOSITION.GUID=BSDEPTA.AGUID";
                _BSDEPTPOS_DB = FillDatatable(sql, Main.CONN_Public);

                sql = @"SELECT * FROM BSWORKSHOP";
                _BSWORKSHOP_DB = FillDatatable(sql, Main.CONN_Public);

                sql = @"SELECT * FROM ECTYPE";
                _ECTYPE_DB = FillDatatable(sql, Main.CONN_Public);

                sql = @"SELECT * FROM ECINFO";
                _ECINFO_DB = FillDatatable(sql, Main.CONN_Public);

                //sql = @"SELECT * FROM HRWORKLICENSE";
                //_HRWORKLICENSE_DB = FillDatatable(sql, Main.CONN_Public);

                //错误
                List<int[]> col_error1 = new List<int[]>();
                List<int[]> col_error2 = new List<int[]>();
                List<int[]> col_error3 = new List<int[]>();
                DataTable dt_error1 = _HREMPLOYEE_excel.Clone();

                //重复数据
                DataTable dt_repet1 = _HREMPLOYEE_excel.Clone();
                DataTable dt_datarepet1 = _HREMPLOYEE_excel.Clone();
                

                

                Dictionary<string, string> doclist = new Dictionary<string, string>();

                #region 主
                string Aguid = "";
                string Bguid = "";
                bool sex = false;
                string sBIRTHDAY = "";
                string sJOBSDATE = "";
                int POSST = 0;
                bool isSTOP = false;
                string pwd = "";

                sql = "SELECT * FROM SSPARAM WHERE PKEY='ITCPassword'";
                DataTable dtParam = FillDatatable(sql, Main.CONN_Public);
                if (dtParam != null && dtParam.Rows.Count > 0)
                {
                    pwd = dtParam.Rows[0]["PVAL"] == DBNull.Value || dtParam.Rows[0]["PVAL"].ToString() == "" ? c_Pwd.SEncrypt("0000") : dtParam.Rows[0]["PVAL"].ToString();
                }
                else
                {
                    pwd = c_Pwd.SEncrypt("0000");
                }

                for (int i = 0; i < _HREMPLOYEE_excel.Rows.Count; i++)
                {
                    bool isError1 = false;
                    bool isRepet1 = false;
                    bool isDataRepet1 = false;

                    DataRow dr_excel = _HREMPLOYEE_excel.Rows[i];

                    if (string.IsNullOrWhiteSpace(dr_excel["员工编码"].ToString()))
                    {
                        //空
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 0 });
                        isError1 = true;
                    }
                    else
                    {
                        DataRow[] drss = _HREMPLOYEE_excel.Select("员工编码 = '" + ReturnString(dr_excel["员工编码"].ToString()) + "'");
                        if (drss.Length > 1)
                        {
                            isRepet1 = true;
                        }

                        DataRow[] drss1 = _HREMPLOYEE_DB.Select(string.Format("EMPCODE='{0}'", ReturnString(dr_excel["员工编码"].ToString())));
                        if (drss1.Length > 0)
                        {
                            isDataRepet1 = true;
                        }
                    }

                    if (string.IsNullOrWhiteSpace(dr_excel["中文姓名"].ToString()))
                    {
                        //空
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 1 });
                        isError1 = true;
                    }

                    if (string.IsNullOrWhiteSpace(dr_excel["性别"].ToString()))
                    {
                        //空
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 2 });
                        isError1 = true;
                    }
                    else
                    {
                        if (dicSex.ContainsKey(dr_excel["性别"].ToString()))
                        {
                            sex = dicSex[dr_excel["性别"].ToString()];
                        }
                        else
                        {
                            col_error1.Add(new int[] { dt_error1.Rows.Count, 2 });
                            isError1 = true;
                        }
                    }

                    if (string.IsNullOrWhiteSpace(dr_excel["部门编号"].ToString()))
                    {
                        //空
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 3 });
                        isError1 = true;
                    }
                    else
                    {
                        DataRow[] drss = _BSDEPT_DB.Select("CODE='" + ReturnString(dr_excel["部门编号"].ToString()) + "'");
                        if (drss.Length == 0)
                        {
                            col_error1.Add(new int[] { dt_error1.Rows.Count, 3 });
                            isError1 = true;
                        }
                        else
                        {
                            Aguid = drss[0]["GUID"].ToString();
                        }
                    }

                    if (string.IsNullOrWhiteSpace(dr_excel["岗位编号"].ToString()))
                    {
                        //空
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 4 });
                        isError1 = true;
                    }
                    else
                    {
                        DataRow[] drss = _BSDEPTPOS_DB.Select("CODE='" + ReturnString(dr_excel["岗位编号"].ToString()) + "' AND PGUID='" + Aguid + "'");
                        if (drss.Count() == 0)
                        {
                            col_error1.Add(new int[] { dt_error1.Rows.Count, 4 });
                            isError1 = true;
                        }
                        else
                        {
                            Bguid = drss[0]["GUID"].ToString();
                        }
                    }

                    DateTime BIRTHDAY = new DateTime();
                    if (string.IsNullOrWhiteSpace(dr_excel["出生日期"].ToString()))
                    {
                        sBIRTHDAY = "null";
                        
                    }
                    else
                    {
                        if (!DateTime.TryParse(dr_excel["出生日期"].ToString(), out BIRTHDAY))
                        {
                            //空、类型不符
                            col_error1.Add(new int[] { dt_error1.Rows.Count, 5 });
                            isError1 = true;
                        }
                        else
                        {
                            sBIRTHDAY = "'" + BIRTHDAY.ToString("yyyy/MM/dd") + "'";
                        }
                    }

                    //if (string.IsNullOrWhiteSpace(dr_excel["员工类型"].ToString()))
                    //{
                    //    //空
                    //    col_error1.Add(new int[] { dt_error1.Rows.Count, 8 });
                    //    isError1 = true;
                    //}
                    //else
                    //{
                    //    if (dic.ContainsKey(dr_excel["员工类型"].ToString()))
                    //    {
                    //        POSST = dic[dr_excel["员工类型"].ToString()];
                    //    }
                    //    else
                    //    {
                    //        col_error1.Add(new int[] { dt_error1.Rows.Count, 8 });
                    //        isError1 = true;
                    //    }
                    //}


                    if (string.IsNullOrWhiteSpace(dr_excel["入职日期"].ToString()))
                    {
                        sJOBSDATE = "null";   
                    }
                    else
                    {
                        if (!DateTime.TryParse(dr_excel["入职日期"].ToString(), out BIRTHDAY))
                        {
                            //空、类型不符
                            col_error1.Add(new int[] { dt_error1.Rows.Count, 9 });
                            isError1 = true;
                        }
                        else
                        {
                            sJOBSDATE = "'" + BIRTHDAY.ToString("yyyy/MM/dd") + "'";
                        }
                    }

                    //if (string.IsNullOrWhiteSpace(dr_excel["是否离职"].ToString()))
                    //{
                    //    //空
                    //    col_error1.Add(new int[] { dt_error1.Rows.Count, 10 });
                    //    isError1 = true;
                    //}
                    //else
                    //{
                    //    if (dicISSTOP.ContainsKey(dr_excel["是否离职"].ToString()))
                    //    {
                    //        isSTOP = dicISSTOP[dr_excel["是否离职"].ToString()];
                    //    }
                    //    else
                    //    {
                    //        col_error1.Add(new int[] { dt_error1.Rows.Count, 10 });
                    //        isError1 = true;
                    //    }
                    //}

                    //if (string.IsNullOrWhiteSpace(dr_excel["英文名"].ToString()))
                    //{
                    //    //空
                    //    col_error1.Add(new int[] { dt_error1.Rows.Count, 13 });
                    //    isError1 = true;
                    //}

                    //if (string.IsNullOrWhiteSpace(dr_excel["紧急联系人"].ToString()))
                    //{
                    //    //空
                    //    col_error1.Add(new int[] { dt_error1.Rows.Count, 14 });
                    //    isError1 = true;
                    //}

                    //if (string.IsNullOrWhiteSpace(dr_excel["家庭住址"].ToString()))
                    //{
                    //    //空
                    //    col_error1.Add(new int[] { dt_error1.Rows.Count, 15 });
                    //    isError1 = true;
                    //}

                    string HREMPLOYEE_GUID = Guid.NewGuid().ToString();

                    _HREMPLOYEE_excel.Rows[i]["GUID"] = HREMPLOYEE_GUID;


                    if (isError1 || isRepet1|| isDataRepet1)
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
                        string temp = string.Format(@"INSERT INTO [HREMPLOYEE]
                                       ([GUID],[EMPCODE],[EMPNAME],[SEX],[AGUID]
                                       ,[BGUID],[BIRTHDAY],[PHONE],[MAIL],[POSST],[JOBSDATE],[ISSTOP],[LNOTE],[NOTE]
                                       ,[CC],[ND],[CD],[PWD],[ENAME],[SPECIALITY],[ADDRESS])
                                        VALUES ('{0}','{1}','{2}','{3}','{4}',
                                        '{5}',{6},'{7}','{8}',{9},{10},'{11}','{12}','{13}','1023','{14}','{14}','{15}','{16}','{17}','{18}')",
                                        HREMPLOYEE_GUID, ReturnString(dr_excel["员工编码"].ToString()), ReturnString(dr_excel["中文姓名"].ToString()), sex, Aguid,
                                        Bguid, sBIRTHDAY, ReturnString(dr_excel["联系方式"].ToString()), ReturnString(dr_excel["办公邮箱"].ToString()), POSST, sJOBSDATE, isSTOP, ReturnString(dr_excel["离职备注"].ToString()), dr_excel["备注"].ToString(), DateTime.Now.ToString(), pwd, ReturnString(dr_excel["英文名"].ToString()), ReturnString(dr_excel["紧急联系人"].ToString()), ReturnString(dr_excel["家庭住址"].ToString()));

                        rbSql.Text += temp + Environment.NewLine;
                        sqlLs.Add(temp);

                        string temp1 = string.Format("INSERT INTO HREMPLOYEEA(GUID,PGUID) VALUES(NEWID(),'{0}')", HREMPLOYEE_GUID);
                        rbSql.Text += temp1 + Environment.NewLine;
                        sqlLs.Add(temp1);
                        
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
