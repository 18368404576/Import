using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NPOI.SS.Formula.Functions;
using TSES.Base;

namespace WinImport
{
    public partial class HRLIBRARY : Form
    {
        ExcelManager _excelManager = new ExcelManager();
        //考题类型
        DataTable _HRLIBRARYTYPE_DB = null;
        /// <summary>
        /// 人员
        /// </summary>
        DataTable _HREMPLOYEE_DB = null;

        /// <summary>
        /// 上岗证
        /// </summary>
        DataTable _HRWORKCENTER_DB = null;

        //题库
        private DataTable _HRLIBRARY_excel = new DataTable();

        //选项
        private DataTable _HRLIBRARYS_excel = new DataTable();

        //光联上岗证
     //   private DataTable _HRLIBRARYW_excel = null;

        //题库星级
        private DataTable _HRLIBRARYWS_excel = new DataTable();

        /// <summary>
        /// 是，否
        /// </summary>
        Dictionary<string, bool> ISINTERVALs = new Dictionary<string, bool>();

        /// <summary>
        /// 正确，错误
        /// </summary>
        Dictionary<string, bool> ISSURE = new Dictionary<string, bool>();

        /// <summary>
        /// 需要保存的sql
        /// </summary>
        List<string> sqlLs = new List<string>();

        /// <summary>
        ///  是否验证成功
        /// </summary>
        bool isCheck = false;

        public HRLIBRARY()
        {
            InitializeComponent();
            ISINTERVALs.Add("是", true);
            ISINTERVALs.Add("否", false);

            ISSURE.Add("正确",true);
            ISSURE.Add("错误", false);
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
            dgError3.DataSource = new DataTable();
            dgRepet3.DataSource = new DataTable();
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
                    _HRLIBRARY_excel = _excelManager.GetExcelTableByOleDB(opfDialog.FileName, "HRLIBRARY");
                    _HRLIBRARY_excel.Columns.Add("GUID");
                    MessageBox.Show("读取笔数:" + _HRLIBRARY_excel.Rows.Count + "");

                    if (_HRLIBRARY_excel == null || _HRLIBRARY_excel.Rows.Count <= 0)
                    {
                        WGMessage.ShowWarning(@"无法读取当前Excel!");
                        return;
                    }
                }
                else if (btn.Name == "btnSelect2")
                {
                    txtFile2.Text = opfDialog.FileName;
                    _HRLIBRARYS_excel = _excelManager.GetExcelTableByOleDB(opfDialog.FileName, "HRLIBRARYS");
                    _HRLIBRARYS_excel.Columns.Add("GUID");
                    MessageBox.Show("读取笔数:" + _HRLIBRARYS_excel.Rows.Count + "");

                    if (_HRLIBRARYS_excel == null || _HRLIBRARYS_excel.Rows.Count <= 0)
                    {
                        WGMessage.ShowWarning(@"无法读取当前Excel!");
                        return;
                    }
                }
                else if (btn.Name == "btnSelect3")
                {
                    txtFile3.Text = opfDialog.FileName;
                    _HRLIBRARYWS_excel = _excelManager.GetExcelTableByOleDB(opfDialog.FileName, "HRLIBRARYW");
                    MessageBox.Show("读取笔数:" + _HRLIBRARYWS_excel.Rows.Count + "");

                    if (_HRLIBRARYWS_excel == null || _HRLIBRARYWS_excel.Rows.Count <= 0)
                    {
                        WGMessage.ShowWarning(@"无法读取当前Excel!");
                        return;
                    }
                }
                //else if (btn.Name == "btnSelect4")
                //{
                //    txtFile4.Text = opfDialog.FileName;
                //    _BSPRODPLAN_excel = _excelManager.GetExcelTableByOleDB(opfDialog.FileName, "BSPRODPLAN");
                //    _BSPRODPLAN_excel.Columns.Add("GUID");
                //    MessageBox.Show("读取笔数:" + _BSPRODPLAN_excel.Rows.Count + "");

                //    if (_BSPRODPLAN_excel == null || _BSPRODPLAN_excel.Rows.Count <= 0)
                //    {
                //        WGMessage.ShowWarning(@"无法读取当前Excel!");
                //        return;
                //    }
                //}
                ClearSql();
            }
            else
            {
                return;
            }
        }

        private string GetISINTERVAL(string str)
        {
            if (str == "")
                return "null";
            else
            {
                return Main.SetDBValue(ISINTERVALs[str]);
            }
        }

        private string GetISSURE(string str)
        {
            if (str == "")
                return "null";
            else
            {
                return Main.SetDBValue(ISSURE[str]);
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

        private void btnCheck_Click(object sender, EventArgs e)
        {
            string sql = @"SELECT GUID,[NAME] FROM [dbo].[HRLIBRARYTYPE]";
            _HRLIBRARYTYPE_DB = FillDatatablde(sql, Main.CONN_Public);

            sql = @"		  SELECT GUID,[EMPCODE],[EMPNAME] FROM HREMPLOYEE";
            _HREMPLOYEE_DB = FillDatatablde(sql, Main.CONN_Public);

            sql = @"SELECT HRWORKLICENSE.GUID, HRWORKLICENSE.NAME 上岗证,[BSWORKSHOP].[CODE] 车间编号,[ECTYPE].[CODE] 设备类别编号,[ECINFO].[CODE] 设备编号  FROM HRWORKLICENSE
LEFT JOIN [dbo].[BSWORKSHOP] ON [BSWORKSHOP].[GUID] = [HRWORKLICENSE].[AGUID]
LEFT JOIN [dbo].[ECTYPE] ON [ECTYPE].[GUID] = HRWORKLICENSE.[BGUID]
LEFT JOIN [dbo].[ECINFO] ON [ECINFO].[GUID] = [HRWORKLICENSE].[CGUID]";
            _HRWORKCENTER_DB = FillDatatablde(sql, Main.CONN_Public);

            #region 题库

            DataTable HRLIBRARY_DB_ADD = new DataTable();
            HRLIBRARY_DB_ADD.Columns.Add("试题内容");
            HRLIBRARY_DB_ADD.Columns.Add("GUID");
            List<int[]> col_error1 = new List<int[]>();
            DataTable dt_repet1 = _HRLIBRARY_excel.Clone();
            DataTable dt_error1 = _HRLIBRARY_excel.Clone();

            for (int i = 0; i < _HRLIBRARY_excel.Rows.Count; i++)
            {
                bool isError = false;
                bool isRepet = false;

                DateTime dtTime = new DateTime();
                DataRow dr_excel = _HRLIBRARY_excel.Rows[i];

                if (!DateTime.TryParse(dr_excel["编制日期"].ToString(),out dtTime))
                {
                    col_error1.Add(new int[] { dt_error1.Rows.Count, 0 });
                    isError = true;
                }

                if (string.IsNullOrWhiteSpace(dr_excel["试题内容"].ToString()))
                {
                    //空
                    col_error1.Add(new int[] { dt_error1.Rows.Count, 1 });
                    isError = true;
                }

                DataRow[] drs_HRLIBRARY_name = _HRLIBRARY_excel.Select("试题内容 = '" + WGHelper.ReturnString(dr_excel["试题内容"].ToString()) + "'");
                if (drs_HRLIBRARY_name.Length > 1)
                {
                    isRepet = true;
                }

                DataRow[] drs_HRLIBRARY_HRLIBRARYTYPE = _HRLIBRARYTYPE_DB.Select("NAME = '" + WGHelper.ReturnString(dr_excel["考题类型"].ToString()) + "'");
                if (drs_HRLIBRARY_HRLIBRARYTYPE.Length == 0)
                {
                    col_error1.Add(new int[] { dt_error1.Rows.Count, 2 });
                    isError = true;
                }

                if (dr_excel["试题类型"].ToString() != "选择题" && dr_excel["试题类型"].ToString() != "判断题")
                {
                    col_error1.Add(new int[] { dt_error1.Rows.Count, 3 });
                    isError = true;
                }

                //DataRow[] drs_HRLIBRARY_HREMPLOYEE = _HREMPLOYEE_DB.Select("EMPCODE = '"+WGHelper.ReturnString(dr_excel["编制人员"].ToString())+"'");
                //if (drs_HRLIBRARY_HREMPLOYEE.Length == 0)
                //{
                //    col_error1.Add(new int[] { dt_error1.Rows.Count, 4 });
                //    isError = true;
                //}



                if (dr_excel["试题类型"].ToString() == "判断题")
                {
                    if (!ISSURE.ContainsKey(dr_excel["判断题正确答案"].ToString()))
                    {
                        col_error1.Add(new int[] {dt_error1.Rows.Count, 5});
                        isError = true;
                    }
                }

                if (isError || isRepet)
                {
                    if (isError)
                    {
                        dt_error1.Rows.Add(dr_excel.ItemArray);
                    }
                    if (isRepet)
                    {
                        dt_repet1.Rows.Add(dr_excel.ItemArray);
                    }
                    continue;
                }

                string MainGUID = Guid.NewGuid().ToString();
                DataRow dr_HRLIBRARY_DB_ADD_newor = HRLIBRARY_DB_ADD.NewRow();
                dr_HRLIBRARY_DB_ADD_newor["GUID"] = MainGUID;
                dr_HRLIBRARY_DB_ADD_newor["试题内容"] = dr_excel["试题内容"].ToString();

                string temp = string.Format(@"INSERT INTO [dbo].[HRLIBRARY]
        ( [GUID] ,
          [FGUID] ,
          [EDATE] ,
          [NAME] ,
          [CTYPE] ,
          [EUSER] ,
          [ANSWER] ,
          [ST] ,
          [CC] ,
          [ND] ,
          [CD]
        )
VALUES  ( '{0}' , -- GUID - uniqueidentifier
          '{1}' , -- FGUID - uniqueidentifier
          '{2}' , -- EDATE - datetime
          '{3}' , -- NAME - nvarchar(200)
          '{4}' , -- CTYPE - nvarchar(10)
          '{5}' , -- EUSER - nvarchar(20)
          {6} , -- ANSWER - bit
          1 , -- ST - bit
          NULL , -- CC - nvarchar(20)
          GETDATE() , -- ND - datetime
          GETDATE()  -- CD - datetime
        )", MainGUID, drs_HRLIBRARY_HRLIBRARYTYPE[0]["GUID"], dtTime, dr_excel["试题内容"], dr_excel["试题类型"], dr_excel["编制人员"], GetISSURE(dr_excel["判断题正确答案"].ToString()));
                rbSql.Text += temp + Environment.NewLine;
                sqlLs.Add(temp);
                HRLIBRARY_DB_ADD.Rows.Add(dr_HRLIBRARY_DB_ADD_newor);
            }

            dgError1.DataSource = dt_error1;
            dgRepet1.DataSource = dt_repet1;
            if (dt_error1.Rows.Count > 0 || dt_repet1.Rows.Count > 0)
            {
                Main.SetErrorCell(dgError1, col_error1);
                rbSql.Text = "";
                return;
            }
            #endregion

            #region 选项
            List<int[]> col_error2 = new List<int[]>();
            DataTable dt_repet2 = _HRLIBRARYS_excel.Clone();
            DataTable dt_error2 = _HRLIBRARYS_excel.Clone();

            for (int i = 0; i < _HRLIBRARYS_excel.Rows.Count; i++)
            {
                bool isError = false;
                bool isRepet = false;
                DataRow dr_excel = _HRLIBRARYS_excel.Rows[i];

                if (string.IsNullOrWhiteSpace(dr_excel["试题内容"].ToString()))
                {
                    //空
                    col_error2.Add(new int[] { dt_error2.Rows.Count, 0 });
                    isError = true;
                }

                DataRow[] drs_HRLIBRARY_name = HRLIBRARY_DB_ADD.Select("试题内容 = '" + WGHelper.ReturnString(dr_excel["试题内容"].ToString()) + "'");
                if (drs_HRLIBRARY_name.Length == 0)
                {
                    col_error2.Add(new int[] { dt_error2.Rows.Count, 0 });
                    isError = true;
                }

                int sno = 1;
                if (!int.TryParse(dr_excel["序号"].ToString(), out sno))
                {
                    col_error2.Add(new int[] { dt_error2.Rows.Count, 1 });
                    isError = true;
                }


                if (!ISINTERVALs.ContainsKey(dr_excel["是否正确答案"].ToString()))
                {
                    col_error2.Add(new int[] { dt_error2.Rows.Count, 3 });
                    isError = true;
                }

                DataRow[] _HRLIBRARYS_DB_RIGHTANWER_drs = _HRLIBRARYS_excel.Select("试题内容 = '"+WGHelper.ReturnString(dr_excel["试题内容"].ToString())+"' AND 是否正确答案 = '是'");
                if (_HRLIBRARYS_DB_RIGHTANWER_drs.Length == 0)
                {
                    col_error2.Add(new int[] { dt_error2.Rows.Count, 3 });
                    isError = true;
                }

                if (isError || isRepet)
                {
                    if (isError)
                    {
                        dt_error2.Rows.Add(dr_excel.ItemArray);
                    }
                    continue;
                }

                string temp = string.Format(@"
INSERT INTO [dbo].[HRLIBRARYS]
        ( [GUID] ,
          [PGUID] ,
          [SNO] ,
          [NAME] ,
          [ISRIGHT]
        )
VALUES  ( NEWID() , -- GUID - uniqueidentifier
          '{0}' , -- PGUID - uniqueidentifier
          '{1}' , -- SNO - int
          '{2}' , -- NAME - nvarchar(100)
          {3}  -- ISRIGHT - bit
        )", drs_HRLIBRARY_name[0]["GUID"], dr_excel["序号"], dr_excel["内容"], GetISINTERVAL(dr_excel["是否正确答案"].ToString()));
                rbSql.Text += temp + Environment.NewLine;
                sqlLs.Add(temp);
            }

            dgError2.DataSource = dt_error2;
            dgRepet2.DataSource = dt_repet2;
            if (dt_error2.Rows.Count > 0 || dt_repet2.Rows.Count > 0)
            {
                Main.SetErrorCell(dgError2, col_error2);
                rbSql.Text = "";
                return;
            }
            #endregion

            #region 关联上岗证
            List<int[]> col_error3 = new List<int[]>();
            DataTable dt_repet3 = _HRLIBRARYWS_excel.Clone();
            DataTable dt_error3 = _HRLIBRARYWS_excel.Clone();
            DataTable dt_HRLIBRARYW_DB = new DataTable();
            dt_HRLIBRARYW_DB.Columns.Add("GUID");
            dt_HRLIBRARYW_DB.Columns.Add("FGUID");
            dt_HRLIBRARYW_DB.Columns.Add("试题内容");

            for (int i = 0; i < _HRLIBRARYWS_excel.Rows.Count; i++)
            {
                bool isError = false;
                bool isRepet = false;
                DataRow dr_excel = _HRLIBRARYWS_excel.Rows[i];
                if (string.IsNullOrWhiteSpace(dr_excel["试题内容"].ToString()))
                {
                    //空
                    col_error3.Add(new int[] { dt_error3.Rows.Count, 0 });
                    isError = true;
                }

                DataRow[] drs_HRLIBRARY_name = HRLIBRARY_DB_ADD.Select("试题内容 = '" + WGHelper.ReturnString(dr_excel["试题内容"].ToString()) + "'");
                if (drs_HRLIBRARY_name.Length == 0)
                {
                    col_error3.Add(new int[] { dt_error3.Rows.Count, 0 });
                    isError = true;
                }

                int sno = 1;
                if (!int.TryParse(dr_excel["序号"].ToString(), out sno))
                {
                    col_error3.Add(new int[] { dt_error3.Rows.Count, 1 });
                    isError = true;
                }

                int star = 1;
                if (!int.TryParse(dr_excel["星级"].ToString(), out star))
                {
                    col_error3.Add(new int[] { dt_error3.Rows.Count, 1 });
                    isError = true;
                }

                string where = "";

                if (dr_excel["设备编号"].ToString() == "")
                {

                }
                else
                {
                    where = "AND 设备编号 = '"+WGHelper.ReturnString(dr_excel["设备编号"].ToString())+"'";
                }

                DataRow[] drs__HRWORKCENTER_DB = _HRWORKCENTER_DB.Select("上岗证 = '" + WGHelper.ReturnString(dr_excel["上岗证名称"].ToString()) + "' AND 车间编号 = '" + WGHelper.ReturnString(dr_excel["车间编号"].ToString()) + "' AND 设备类别编号 = '" + WGHelper.ReturnString(dr_excel["设备类别编号"].ToString()) + "' " + where);

                if(drs__HRWORKCENTER_DB.Length == 0)
                {
                    col_error3.Add(new int[] { dt_error3.Rows.Count, 2 });
                    col_error3.Add(new int[] { dt_error3.Rows.Count, 3 });
                    col_error3.Add(new int[] { dt_error3.Rows.Count, 4 });
                    col_error3.Add(new int[] { dt_error3.Rows.Count, 5 });
                    isError = true;
                }

                DataRow[] drs_repit = _HRLIBRARYWS_excel.Select("试题内容 = '" + WGHelper.ReturnString(dr_excel["试题内容"].ToString()) + "' AND 上岗证名称 = '" + WGHelper.ReturnString(dr_excel["上岗证名称"].ToString()) + "' AND 车间编号 = '" + WGHelper.ReturnString(dr_excel["车间编号"].ToString()) + "' AND 设备类别编号 = '" + WGHelper.ReturnString(dr_excel["设备类别编号"].ToString()) + "' AND 设备编号 = '" + WGHelper.ReturnString(dr_excel["设备编号"].ToString()) + "' AND 星级 = '" + dr_excel["星级"].ToString() + "'");

                if(drs_repit.Length > 1)
                {
                    isRepet = true;
                }

                if (isError || isRepet)
                {
                    if (isError)
                    {
                        dt_error3.Rows.Add(dr_excel.ItemArray);
                    }
                    if(isRepet)
                    {
                        dt_repet3.Rows.Add(dr_excel.ItemArray);
                    }
                    continue;
                }
                string KidGUID = Guid.NewGuid().ToString();
                DataRow DR_dt_HRLIBRARYW_DB = dt_HRLIBRARYW_DB.NewRow();
                DR_dt_HRLIBRARYW_DB["GUID"] = KidGUID;
                DR_dt_HRLIBRARYW_DB["FGUID"] = drs__HRWORKCENTER_DB[0]["GUID"];
                DR_dt_HRLIBRARYW_DB["试题内容"] = dr_excel["试题内容"];

                Boolean ISHAVE = false;
                foreach (DataRow dr in dt_HRLIBRARYW_DB.Rows)
                {
                    if (dr["FGUID"].ToString().ToUpper() == drs__HRWORKCENTER_DB[0]["GUID"].ToString().ToUpper() && dr["试题内容"].ToString() == dr_excel["试题内容"].ToString())
                    {
                        KidGUID = dr["GUID"].ToString();
                        ISHAVE = true;
                    }
                }

                if (!ISHAVE)
                {
                    string temp = string.Format(@"
INSERT INTO [dbo].[HRLIBRARYW]
        ( [GUID], [PGUID], [SNO], [FGUID] )
VALUES  ( '{0}', -- GUID - uniqueidentifier
          '{1}', -- PGUID - uniqueidentifier
          '{2}', -- SNO - int
          '{3}'  -- FGUID - uniqueidentifier
          )", KidGUID, drs_HRLIBRARY_name[0]["GUID"], sno, drs__HRWORKCENTER_DB[0]["GUID"]);
                    rbSql.Text += temp + Environment.NewLine;
                    sqlLs.Add(temp);
                    dt_HRLIBRARYW_DB.Rows.Add(DR_dt_HRLIBRARYW_DB);
                }
                else
                {

                }

                string temp1 = string.Format(@"
INSERT INTO [dbo].[HRLIBRARYWS]
        ( [GUID], [PGUID], [LVL] )
VALUES  ( NEWID(), -- GUID - uniqueidentifier
          '{0}', -- PGUID - uniqueidentifier
          '{1}'  -- LVL - int
          )", KidGUID,dr_excel["星级"]);
                rbSql.Text += temp1 + Environment.NewLine;
                sqlLs.Add(temp1);
            }

            dgError3.DataSource = dt_error3;
            dgRepet3.DataSource = dt_repet3;
            if (dt_error3.Rows.Count > 0 || dt_repet3.Rows.Count > 0)
            {
                Main.SetErrorCell(dgError3, col_error3);
                rbSql.Text = "";
                return;
            }

            #endregion
            isCheck = true;
        }
    }
}
