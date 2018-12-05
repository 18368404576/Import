using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using TSES.Base;

namespace WinImport
{
    public partial class BSWORKCENTERINFO : Form
    {
        public BSWORKCENTERINFO()
        {
            InitializeComponent();
        }

        #region  变量
        ExcelManager _excelManager = new ExcelManager();

        DataTable _BSWORKCENTER_excel = null;

        DataTable _BSWORKCENTER_DB = null;

        DataTable _BSWORKCENTERS_excel = null;

        DataTable _BSWORKCENTERS_DB = null;

        DataTable _BSWORKSHOP_DB = null;//车间

        DataTable _BSWCTYPE_DB = null;//工作中心类别

        DataTable _ECINFO_DB = null;//设备


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
                    _BSWORKCENTER_excel = _excelManager.GetExcelTableByOleDB(opfDialog0.FileName, "BSWORKCENTER");
                    _BSWORKCENTER_excel.Columns.Add("GUID");
                    MessageBox.Show("读取笔数：" + _BSWORKCENTER_excel.Rows.Count + "");
                }
                catch
                { }
                if (_BSWORKCENTER_excel == null || _BSWORKCENTER_excel.Rows.Count <= 0)
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
                    _BSWORKCENTERS_excel = _excelManager.GetExcelTableByOleDB(opfDialog0.FileName, "BSWORKCENTERA");
                    MessageBox.Show("读取笔数：" + _BSWORKCENTERS_excel.Rows.Count + "");
                }
                catch
                {

                }
                //ClearSql();
                if (_BSWORKCENTERS_excel == null || _BSWORKCENTERS_excel.Rows.Count <= 0)
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

                if (_BSWORKCENTER_excel == null)
                {
                    WGMessage.ShowWarning("请选择工作中心文件!");
                    return;
                }
                if (_BSWORKCENTERS_excel == null)
                {
                    WGMessage.ShowWarning("请选择工作中心工位文件!");
                    return;
                }

                string sql = "";

                sql = "SELECT * FROM BSWORKCENTER";
                _BSWORKCENTER_DB = FillDatatablde(sql, Main.CONN_Public);

                sql = "SELECT * FROM BSWORKCENTERA";
                _BSWORKCENTERS_DB = FillDatatablde(sql, Main.CONN_Public);

                sql = "SELECT * FROM BSWORKSHOP";
                _BSWORKSHOP_DB = FillDatatablde(sql, Main.CONN_Public);

                sql = "SELECT * FROM BSWCTYPE";
                _BSWCTYPE_DB = FillDatatablde(sql, Main.CONN_Public);

                sql = "SELECT * FROM ECINFO";
                _ECINFO_DB = FillDatatablde(sql, Main.CONN_Public);
                //错误
                List<int[]> col_error1 = new List<int[]>();
                List<int[]> col_error2 = new List<int[]>();

                DataTable dt_error1 = _BSWORKCENTER_excel.Clone();
                //重复数据
                DataTable dt_repet1 = _BSWORKCENTER_excel.Clone();
                DataTable dt_datarepet1 = _BSWORKCENTER_excel.Clone();

                DataTable dt_error2 = _BSWORKCENTERS_excel.Clone();
                //重复数据
                DataTable dt_repet2 = _BSWORKCENTERS_excel.Clone();
                DataTable dt_datarepet2 = _BSWORKCENTERS_excel.Clone();

                Dictionary<string, string> doclist = new Dictionary<string, string>();

                #region 主
                string Pguid = "";
                string Tguid = "";
                Dictionary<string, int> dic = new Dictionary<string, int>();
                dic.Add("正常", 0);
                dic.Add("停用", 1);
                for (int i = 0; i < _BSWORKCENTER_excel.Rows.Count; i++)
                {
                    bool isError1 = false;
                    bool isRepet1 = false;
                    bool isDataRepet1 = false;

                    DataRow dr_excel = _BSWORKCENTER_excel.Rows[i];

                    if (string.IsNullOrWhiteSpace(dr_excel["工作中心编号"].ToString()))
                    {
                        //空
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 0 });
                        isError1 = true;
                    }
                    else
                    {
                        DataRow[] drss = _BSWORKCENTER_excel.Select("工作中心编号 = '" + ReturnString(dr_excel["工作中心编号"].ToString()) + "'");
                        if (drss.Length > 1)
                        {
                            isRepet1 = true;
                        }

                        DataRow[] drss1 = _BSWORKCENTER_DB.Select(string.Format("CODE='{0}'", ReturnString(dr_excel["工作中心编号"].ToString())));
                        if (drss1.Length > 0)
                        {
                            isDataRepet1 = true;
                        }
                    }

                    if (string.IsNullOrWhiteSpace(dr_excel["工作中心名称"].ToString()))
                    {
                        //空
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 1 });
                        isError1 = true;
                    }

                    if (string.IsNullOrWhiteSpace(dr_excel["车间编号"].ToString()))
                    {
                        //空
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 2 });
                        isError1 = true;
                    }
                    else
                    {
                        DataRow[] drss = _BSWORKSHOP_DB.Select("CODE = '" + ReturnString(dr_excel["车间编号"].ToString()) + "'");
                        if (drss.Length > 0)
                        {
                            Pguid = drss[0]["GUID"].ToString();
                        }
                        else
                        {
                            col_error1.Add(new int[] { dt_error1.Rows.Count, 2 });
                            isError1 = true;
                        }
                    }

                    if (string.IsNullOrWhiteSpace(dr_excel["工作中心类别编号"].ToString()))
                    {
                        //空
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 2 });
                        isError1 = true;
                    }
                    else
                    {
                        DataRow[] drss = _BSWCTYPE_DB.Select("CODE = '" +ReturnString(dr_excel["工作中心类别编号"].ToString()) + "'");
                        if (drss.Length > 0)
                        {
                            Tguid = drss[0]["GUID"].ToString();
                        }
                        else
                        {
                            col_error1.Add(new int[] { dt_error1.Rows.Count, 2 });
                            isError1 = true;
                        }
                    }
                   
                    if (string.IsNullOrWhiteSpace(dr_excel["状态"].ToString()) || !dic.ContainsKey(dr_excel["状态"].ToString()))
                    {
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 6 });
                        isError1 = true;
                    }

                    string BSWORKCENTER_GUID = Guid.NewGuid().ToString();

                    _BSWORKCENTER_excel.Rows[i]["GUID"] = BSWORKCENTER_GUID;


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
                        string temp = string.Format(@"INSERT INTO BSWORKCENTER
                                                      (GUID,AGUID,ACODE,ANAME,CODE,NAME,BGUID,BCODE,BNAME,NOTE,ST,ND,CD)
                                                      VALUES  
                                                      ('{0}',{1},'{2}', '{3}', '{4}','{5}','{6}','{7}','{8}','{9}',0,GETDATE(),GETDATE())",
                                                      BSWORKCENTER_GUID, Tguid == "" ? "null" : "'" + Tguid + "'", ReturnString(dr_excel["工作中心类别编号"].ToString()), 
                                                      ReturnString(dr_excel["工作中心类别名称"].ToString()),ReturnString(dr_excel["工作中心编号"].ToString()), 
                                                      ReturnString(dr_excel["工作中心名称"].ToString()), Pguid, ReturnString(dr_excel["车间编号"].ToString()), ReturnString(dr_excel["车间编号"].ToString()),
                                                      ReturnString(dr_excel["备注"].ToString()), dic[dr_excel["状态"].ToString()],
                                                       DateTime.Now.ToString());

                        rbSql.Text += temp + Environment.NewLine;
                        sqlLs.Add(temp);
                    }
                    catch
                    { }
                }


                #endregion

                #region 子
                string Mguid = "";
                string FGUID = "null";
                Dictionary<string, bool> dicISTASKSTART = new Dictionary<string, bool>();
                dicISTASKSTART.Add("是", true);
                dicISTASKSTART.Add("否", false);

                Dictionary<string, bool> dicISREPORT = new Dictionary<string, bool>();
                dicISREPORT.Add("是", true);
                dicISREPORT.Add("否", false);

                for (int i = 0; i < _BSWORKCENTERS_excel.Rows.Count; i++)
                {
                    bool isError2 = false;
                    bool isRepet2 = false;
                    bool isdataRepet2 = false;

                    DataRow dr_excel = _BSWORKCENTERS_excel.Rows[i];

                    if (string.IsNullOrWhiteSpace(dr_excel["工作中心编号"].ToString()))
                    {
                        //空
                        col_error2.Add(new int[] { dt_error2.Rows.Count, 0 });
                        isError2 = true;
                    }
                    else
                    {
                        DataRow[] drss = _BSWORKCENTER_excel.Select("工作中心编号='" +ReturnString(dr_excel["工作中心编号"].ToString()) + "'");
                        if (drss.Length > 0)
                        {
                            Mguid = drss[0]["GUID"].ToString();
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
                        //空
                        col_error2.Add(new int[] { dt_error2.Rows.Count, 1 });
                        isError2 = true;
                    }
                    else
                    {
                        if (numa <= 0)
                        {
                            //空
                            col_error2.Add(new int[] { dt_error2.Rows.Count, 1 });
                            isError2 = true;
                        }
                    }

                    if (string.IsNullOrEmpty(dr_excel["工位编号"].ToString()))
                    {
                        //空
                        col_error2.Add(new int[] { dt_error2.Rows.Count, 2 });
                        isError2 = true;
                    }
                    else
                    {
                        DataRow[] drss = _BSWORKCENTERS_excel.Select("工作中心编号='" +ReturnString(dr_excel["工作中心编号"].ToString()) + "' AND 工位编号='" +ReturnString(dr_excel["工位编号"].ToString()) + "'");
                        if (drss.Length > 1)
                        {
                            
                            isRepet2 = true;
                        }

                        DataRow[] drs = _BSWORKCENTERS_DB.Select("CODE='" + ReturnString(dr_excel["工位编号"].ToString()) + "'");
                        if (drs.Length > 0)
                        {
                            isdataRepet2 = true;
                        }
                    }

                    if (string.IsNullOrWhiteSpace(dr_excel["工位名称"].ToString()))
                    {
                        col_error2.Add(new int[] { dt_error2.Rows.Count, 3 });
                        isError2 = true;
                    }

                    if (!string.IsNullOrWhiteSpace(dr_excel["设备编号"].ToString()))
                    {
                        DataRow[] drss = _ECINFO_DB.Select("CODE='" + dr_excel["设备编号"] + "'");
                        if (drss.Length > 0)
                        {
                            FGUID = "'" + drss[0]["GUID"].ToString() + "'";
                        }
                        else
                        {
                            col_error2.Add(new int[] { dt_error2.Rows.Count, 4 });
                            isError2 = true;
                        }
                    }
                    else
                    {
                        FGUID = "null";
                    }

                    //if (string.IsNullOrWhiteSpace(dr_excel["是否开启生产任务工位"].ToString()))
                    //{
                    //    col_error2.Add(new int[] { dt_error2.Rows.Count, 6 });
                    //    isError2 = true;
                    //}
                    //else
                    //{
                    //    if (!dicISTASKSTART.ContainsKey(dr_excel["是否开启生产任务工位"].ToString()))
                    //    {
                    //        col_error2.Add(new int[] { dt_error2.Rows.Count, 6 });
                    //        isError2 = true;
                    //    }
                    //}

                    //if (string.IsNullOrWhiteSpace(dr_excel["是否报工工位"].ToString()))
                    //{
                    //    col_error2.Add(new int[] { dt_error2.Rows.Count, 7 });
                    //    isError2 = true;
                    //}
                    //else
                    //{
                    //    if (!dicISREPORT.ContainsKey(dr_excel["是否报工工位"].ToString()))
                    //    {
                    //        col_error2.Add(new int[] { dt_error2.Rows.Count, 7 });
                    //        isError2 = true;
                    //    }
                    //}


                    if (isError2 || isRepet2||isdataRepet2)
                    {
                        if (isError2)
                        {
                            dt_error2.Rows.Add(dr_excel.ItemArray);
                        }
                        if (isRepet2)
                        {
                            dt_repet2.Rows.Add(dr_excel.ItemArray);
                        }
                        if (isdataRepet2)
                        {
                            dt_datarepet2.Rows.Add(dr_excel.ItemArray);
                        }
                        continue;
                    }

                    string temp = string.Format(@" INSERT INTO BSWORKCENTERA
                                                  (GUID,SNO,PGUID,CODE,NAME,AGUID,ACODE,ANAME,NOTE,ND,CD)
                                                  VALUES  
                                                  ( NEWID(),'{0}','{1}','{2}','{3}',{4},'{5}','{6}','{7}',GETDATE(), GETDATE())",
                                                   dr_excel["序号"], Mguid, dr_excel["工位编号"], dr_excel["工位名称"], FGUID, dr_excel["设备名称"], dr_excel["设备编号"],dr_excel["备注"]);

                    rbSql.Text += temp + Environment.NewLine;
                    sqlLs.Add(temp);
                }
                #endregion


                dt_error1.Columns.Remove("GUID");
                dt_repet1.Columns.Remove("GUID");
                dt_datarepet1.Columns.Remove("GUID");

                dgError1.DataSource = dt_error1; dgError2.DataSource = dt_error2;
                dgRepet1.DataSource = dt_repet1; dgRepet2.DataSource = dt_repet2;
                dgDataRepet1.DataSource = dt_datarepet1; dgDataRepet2.DataSource = dt_datarepet2;
                if (dt_error1.Rows.Count > 0 || dt_error2.Rows.Count > 0 || dt_repet1.Rows.Count > 0 || dt_repet2.Rows.Count > 0 || dt_datarepet1.Rows.Count > 0 || dt_datarepet2.Rows.Count > 0)
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
