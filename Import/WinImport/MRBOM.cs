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
    public partial class MRBOM : Form
    {
        ExcelManager _excelManager = new ExcelManager();
        DataTable _MRBOM_excel = new DataTable();
        private DataTable _MRBOMSUB_excel = new DataTable();
        DataTable _HREMPLOYEE_DB = new DataTable();
        DataTable _MROPRODUCT_DB = new DataTable();

        DataTable _MRBOM_DB = new DataTable();
        DataTable _MRBOMSUB_DB = new DataTable();

        /// <summary>
        /// 需要保存的sql
        /// </summary>
        List<string> sqlLs = new List<string>();

        /// <summary>
        ///  是否验证成功
        /// </summary>
        bool isCheck = false;

        /// <summary>
        /// 0、未审核 1、已审核 
        /// </summary>
        Dictionary<string, int> STs = new Dictionary<string, int>();

        public MRBOM()
        {
            InitializeComponent();

            STs.Add("未审核", 0);
            STs.Add("已审核", 1);
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            //得到excel数据源
            if (opfDialog.ShowDialog() == DialogResult.OK)
            {
                Button btn = sender as Button;
                if (btn.Name == "btnSelect1")
                {
                    txtFile1.Text = opfDialog.FileName;
                    _MRBOM_excel = _excelManager.GetExcelTableByOleDB(opfDialog.FileName, "MRBOM");
                    MessageBox.Show("读取笔数:" + _MRBOM_excel.Rows.Count + "");
                    if (_MRBOM_excel == null || _MRBOM_excel.Rows.Count <= 0)
                    {
                        WGMessage.ShowWarning(@"无法读取当前Excel!");
                        return;
                    }
                }
                else if (btn.Name == "btnSelect3")
                {
                    //物料校验项
                    txtFile3.Text = opfDialog.FileName;
                    _MRBOMSUB_excel = _excelManager.GetExcelTableByOleDB(opfDialog.FileName, "MRBOMSUB");
                    MessageBox.Show("读取笔数:" + _MRBOMSUB_excel.Rows.Count + "");

                    if (_MRBOMSUB_excel == null || _MRBOMSUB_excel.Rows.Count <= 0)
                    {
                        WGMessage.ShowWarning(@"无法读取当前Excel!");
                        return;
                    }
                }
                ClearSql();
            }
            else
            {
                return;
            }
        }

        public void ClearSql()
        {
            //重新上传后，清空原来的
            isCheck = false;
            sqlLs = new List<string>();
            rbSql.Text = "";
            dgError.DataSource = new DataTable();
            dgRepet_excel.DataSource = new DataTable();
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

        /// <summary>
        /// 查询数据
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="Account"></param>
        /// <returns></returns>
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
            if (_MRBOM_excel == null)
            {
                WGMessage.ShowWarning("请选择[模具BOM]文件!");
                return;
            }
            if (_MRBOMSUB_excel == null)
            {
                WGMessage.ShowWarning("请选择[模具BOM明细]文件!");
                return;
            }
            if (isCheck)
            {
                WGMessage.ShowAsterisk("已验证，不用重复验证！");
                return;
            }
            //加载对应厂部的部门-岗位
            string sql = @"SELECT GUID,[EMPCODE],[EMPNAME] FROM [dbo].[HREMPLOYEE]";
            _HREMPLOYEE_DB = FillDatatablde(sql, Main.CONN_Public);
            //加载物料基本信息
            sql = @"select GUID,CODE,NAME from BSPRODUCT WHERE [CTYPE] IN (2)";
            _MROPRODUCT_DB = FillDatatablde(sql, Main.CONN_Public);

            sql = @"		SELECT [MRBOM].[CODE] 单号,[BSPRODUCT].[CODE] 物料 FROM [MRBOM]
		LEFT JOIN [dbo].[BSPRODUCT] ON [BSPRODUCT].[GUID] = [dbo].[MRBOM].[BGUID]";
            _MRBOM_DB = FillDatatablde(sql, Main.CONN_Public);

            sql = @"		SELECT [MRBOM].CODE 单号,[BSPRODUCT].[CODE] 物料 FROM MRBOMSUB
		LEFT JOIN [dbo].[MRBOM] ON [MRBOM].[GUID] = [MRBOMSUB].[PGUID]
		LEFT JOIN [dbo].[BSPRODUCT] ON [BSPRODUCT].[GUID] = [dbo].MRBOMSUB.WGUID";
            _MRBOMSUB_DB = FillDatatablde(sql, Main.CONN_Public);

            #region  模具BOM 验证

            DataTable dt_error = _MRBOM_excel.Clone();
            DataTable dt_repet_excel = _MRBOM_excel.Clone();
            //错误
            List<int[]> col_error = new List<int[]>();

            //重复数据
            DataTable dt_repet = _MRBOM_excel.Clone();

            Dictionary<string, Guid> newMRO = new Dictionary<string, Guid>();

            for (int i = 0; i < _MRBOM_excel.Rows.Count; i++)
            {
                bool isError = false;
                bool isRepet = false;
                bool isRepet_excel = false;

                DataRow dr_excel = _MRBOM_excel.Rows[i];

                //DataRow[] drs_ectype = _ECType_DB.Select("CODE='" + WGHelper.ReturnString(dr_excel["MRO类别"].ToString()) + "'");
                //if (string.IsNullOrWhiteSpace(dr_excel["MRO类别"].ToString()) || drs_ectype.Length == 0)
                //{
                //    //空、不存在
                //    col_error.Add(new int[] { dt_error.Rows.Count, 3 });
                //    isError = true;
                //}


                if (string.IsNullOrWhiteSpace(dr_excel["单号"].ToString()))
                {
                    //空、不存在
                    col_error.Add(new int[] { dt_error.Rows.Count, 0 });
                    isError = true;
                }

                DataRow[] drs_HREMPLOYEE1 = _HREMPLOYEE_DB.Select("EMPCODE='" + WGHelper.ReturnString(dr_excel["录入人工号"].ToString()) + "'");

                if (drs_HREMPLOYEE1.Length == 0)
                {
                    col_error.Add(new int[] { dt_error.Rows.Count, 1 });
                    isError = true;
                }

                DateTime dtTime1 = new DateTime();
                if (!DateTime.TryParse(dr_excel["录入日期"].ToString(), out dtTime1))
                {
                    col_error.Add(new int[] { dt_error.Rows.Count, 2 });
                    isError = true;
                }

                DataRow[] drs_mroproduct = _MROPRODUCT_DB.Select("CODE = '" + WGHelper.ReturnString(dr_excel["模具编号"].ToString()) + "'");
                if (drs_mroproduct.Length == 0)
                {
                    col_error.Add(new int[] { dt_error.Rows.Count, 3 });
                    isError = true;
                }

                if (!STs.ContainsKey(dr_excel["单据状态"].ToString()))
                {
                    col_error.Add(new int[] { dt_error.Rows.Count, 4 });
                    isError = true;
                }

                DataRow[] drs_HREMPLOYEE2 = _HREMPLOYEE_DB.Select("EMPCODE='" + WGHelper.ReturnString(dr_excel["审核人工号"].ToString()) + "'");

                if (drs_HREMPLOYEE2.Length == 0)
                {
                    col_error.Add(new int[] { dt_error.Rows.Count, 5 });
                    isError = true;
                }

                DateTime dtTime2 = new DateTime();
                if (!DateTime.TryParse(dr_excel["审核日期"].ToString(), out dtTime2))
                {
                    col_error.Add(new int[] { dt_error.Rows.Count, 6 });
                    isError = true;
                }

                DataRow[] drs_repet = _MRBOM_DB.Select("单号 = '"+WGHelper.ReturnString(dr_excel["单号"].ToString())+"'");
                if (drs_repet.Length > 0)
                    isRepet = true;

                drs_repet = _MRBOM_DB.Select("物料 = '" + WGHelper.ReturnString(dr_excel["模具编号"].ToString()) + "'");
                if (drs_repet.Length > 0)
                    isRepet = true;

                DataRow[] drs_repet_excel = _MRBOM_excel.Select("单号 = '" + WGHelper.ReturnString(dr_excel["单号"].ToString()) + "'");
                if (drs_repet_excel.Length > 1)
                {
                    isRepet_excel = true;
                }

                drs_repet_excel = _MRBOM_excel.Select("模具编号 = '" + WGHelper.ReturnString(dr_excel["模具编号"].ToString()) + "'");
                if (drs_repet_excel.Length > 1)
                {
                    isRepet_excel = true;
                }

                if (isError || isRepet || isRepet_excel)
                {
                    if (isError)
                    {
                        dt_error.Rows.Add(dr_excel.ItemArray);
                    }
                    if (isRepet)
                    {
                        dt_repet.Rows.Add(dr_excel.ItemArray);
                    }
                    if (isRepet_excel)
                    {
                        dt_repet_excel.Rows.Add(dr_excel.ItemArray);
                    }
                    continue;
                }

                Guid n = Guid.NewGuid();
                string temp = string.Format(@"INSERT INTO [dbo].[MRBOM]
        ( [GUID] ,
          [CODE] ,
          [AGUID] ,
          [INTIME] ,
          [BGUID] ,
          [ST] ,
          [AUDITGUID] ,
          [ADATE] ,
          [NOTE] ,
          [ND] ,
          [CD]
        )
VALUES  ( '{0}' , -- GUID - uniqueidentifier
          '{1}' , -- CODE - nvarchar(50)
          '{2}' , -- AGUID - uniqueidentifier
          '{3}' , -- INTIME - datetime
          '{4}' , -- BGUID - uniqueidentifier
          '{5}' , -- ST - bit
          '{6}' , -- AUDITGUID - uniqueidentifier
          '{7}' , -- ADATE - datetime
          '{8}' , -- NOTE - nvarchar(200)
          GETDATE() , -- ND - datetime
          GETDATE()  -- CD - datetime
        )"
                    ,n
                    ,WGHelper.ReturnString(dr_excel["单号"].ToString())
                    ,drs_HREMPLOYEE1[0]["GUID"].ToString()
                    ,dtTime1
                    , drs_mroproduct[0]["GUID"].ToString()
                    , STs[dr_excel["单据状态"].ToString()]
                    ,drs_HREMPLOYEE2[0]["GUID"].ToString()
                    ,dtTime2
                    ,WGHelper.ReturnString(dr_excel["备注"].ToString())
                    );

                //rbSql.Text += temp + Environment.NewLine;
                sqlLs.Add(temp);
                newMRO.Add(dr_excel["单号"].ToString(), n);
            }
            dgError.DataSource = dt_error;
            dgRepet.DataSource = dt_repet;
            dgRepet_excel.DataSource = dt_repet_excel;
            if (dt_error.Rows.Count > 0 || dt_repet.Rows.Count > 0 || dt_repet_excel.Rows.Count > 0)
            {
                Main.SetErrorCell(dgError, col_error);
                rbSql.Text = "";
                return;
            }
            #endregion

            #region  模具BOM明细 验证
            DataTable dt_error2 = _MRBOMSUB_excel.Clone();
            DataTable dt_repet2_excel = _MRBOMSUB_excel.Clone();
            //错误
            List<int[]> col_error2 = new List<int[]>();

            //重复数据
            DataTable dt_repet2 = _MRBOMSUB_excel.Clone();
            for (int i = 0; i < _MRBOMSUB_excel.Rows.Count; i++)
            {
                bool isError = false;
                bool isRepet = false;
                bool isRepet_excel = false;

                DataRow dr_excel = _MRBOMSUB_excel.Rows[i];

                if (!newMRO.ContainsKey(dr_excel["单号"].ToString()))
                {
                    col_error2.Add(new int[] { dt_error2.Rows.Count, 0 });
                    isError = true;
                }

                int SNO = 0;
                if (string.IsNullOrWhiteSpace(dr_excel["序号"].ToString())
                    || !int.TryParse(dr_excel["序号"].ToString(), out SNO))
                {
                    //空
                    col_error2.Add(new int[] { dt_error2.Rows.Count, 1 });
                    isError = true;
                }


                DataRow[] drs_mroproduct = _MROPRODUCT_DB.Select("CODE = '" + WGHelper.ReturnString(dr_excel["模具编号"].ToString()) + "'");
                if (drs_mroproduct.Length == 0)
                {
                    col_error2.Add(new int[] { dt_error2.Rows.Count, 2 });
                    isError = true;
                }

                int num = 0;

                if (!int.TryParse(dr_excel["数量"].ToString(),out num))
                {
                    col_error2.Add(new int[] { dt_error2.Rows.Count, 3 });
                    isError = true;
                }

                DataRow[] drs_repet = _MRBOMSUB_DB.Select("单号 = '" + WGHelper.ReturnString(dr_excel["单号"].ToString()) + "'");
                if (drs_repet.Length > 0)
                {
                    isRepet = true;
                }

                DataRow[] drs_repet_excel = _MRBOMSUB_excel.Select("单号 = '" + WGHelper.ReturnString(dr_excel["单号"].ToString()) + "' AND 模具编号 = '" + WGHelper.ReturnString(dr_excel["模具编号"].ToString()) + "'");

                if (drs_repet_excel.Length > 1)
                    isRepet_excel = true;

                if (isError || isRepet || isRepet_excel)
                {
                    if (isError)
                    {
                        dt_error2.Rows.Add(dr_excel.ItemArray);
                    }

                    if (isRepet)
                    {
                        dt_repet2.Rows.Add(dr_excel.ItemArray);
                    }

                    if (isRepet_excel)
                    {
                        dt_repet2_excel.Rows.Add(dr_excel.ItemArray);
                    }
                    continue;
                }

                string temp = string.Format(@"
                INSERT INTO [dbo].[MRBOMSUB]
                            ( [GUID] ,
                        [SNO] ,
                    [PGUID] ,
                    [WGUID] ,
                    [QTY]
                    )
                VALUES  ( NEWID() , -- GUID - uniqueidentifier
                '{0}' , -- SNO - int
                '{1}' , -- PGUID - uniqueidentifier
                '{2}' , -- WGUID - uniqueidentifier
                {3}  -- QTY - int
                    )"
                    ,dr_excel["序号"].ToString()
                    , newMRO[dr_excel["单号"].ToString()]
                    ,drs_mroproduct[0]["GUID"].ToString()
                    ,num);
                //rbSql.Text += temp + Environment.NewLine;
                sqlLs.Add(temp);
            }
            dgError2.DataSource = dt_error2;
            dgRepet2.DataSource = dt_repet2;
            dgRepet_excel2.DataSource = dt_repet2_excel;
            if (dt_error2.Rows.Count > 0 || dt_repet2.Rows.Count > 0 || dt_repet2_excel.Rows.Count > 0)
            {
                Main.SetErrorCell(dgError2, col_error2);
                rbSql.Text = "";
                return;
            }
            #endregion

            isCheck = true;

            StringBuilder last = new StringBuilder();
            foreach (string sql1 in sqlLs)
            {
                last.Append(sql1 + Environment.NewLine);
            }

            rbSql.Text = last.ToString();
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
