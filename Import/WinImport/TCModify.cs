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
    public partial class TCModify : Form
    {
        public TCModify()
        {
            InitializeComponent();
        }

        #region  变量
        ExcelManager _excelManager = new ExcelManager();

        DataTable _CT_excel = null;

        /// <summary>
        ///  是否验证成功
        /// </summary>
        bool isCheck = false;

        /// <summary>
        /// 需要保存的sql
        /// </summary>
        List<string> sqlLs = new List<string>();
        #endregion

        private void btnSelect_Click(object sender, EventArgs e)
        {
            //得到excel数据源
            if (opfDialog0.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    txtFileMain.Text = opfDialog0.FileName;
                    _CT_excel = _excelManager.GetExcelTableByOleDB(opfDialog0.FileName, "BSDEPT");
                    MessageBox.Show("读取笔数：" + _CT_excel.Rows.Count + "");
                }
                catch
                { }
                if (_CT_excel == null || _CT_excel.Rows.Count <= 0)
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
            if (_CT_excel == null)
            {
                WGMessage.ShowWarning("请选择[CT修改]文件!");
                return;
            }

            if (isCheck)
            {
                WGMessage.ShowAsterisk("已验证，不用重复验证！");
                return;
            }

            //错误
            List<int[]> col_error = new List<int[]>();

            //重复数据
            DataTable dt_error = _CT_excel.Clone();

            DataTable dt_repet_excel = _CT_excel.Clone();

            decimal CT = 0;
            decimal ACT = 0;
            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < _CT_excel.Rows.Count; i++)
            {
                bool isError = false;
                bool isRepet_excel = false;

                DataRow dr_excel = _CT_excel.Rows[i];

                if (string.IsNullOrWhiteSpace(dr_excel["零件号"].ToString()))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 0 });
                    isError = true;
                }
                if (string.IsNullOrWhiteSpace(dr_excel["版本"].ToString()))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 1 });
                    isError = true;
                }
                if (string.IsNullOrWhiteSpace(dr_excel["工序"].ToString()))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 2 });
                    isError = true;
                }

                if (string.IsNullOrWhiteSpace(dr_excel["工作中心编号"].ToString()))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 3 });
                    isError = true;
                }

                DataRow[] drs = _CT_excel.Select(string.Format("零件号='{0}' AND 版本='{1}' AND 工序='{2}' AND 工作中心编号='{3}'",
                    WGHelper.ReturnString(dr_excel["零件号"].ToString()),
                    WGHelper.ReturnString(dr_excel["版本"].ToString()), 
                    WGHelper.ReturnString(dr_excel["工序"].ToString()),
                    WGHelper.ReturnString(dr_excel["工作中心编号"].ToString())));
                if (drs.Length > 1)
                {
                    isRepet_excel = true;
                }


                if (string.IsNullOrWhiteSpace(dr_excel["标准工时"].ToString()))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 4 });
                    isError = true;
                }
                else
                {

                    if (!decimal.TryParse(dr_excel["标准工时"].ToString(), out CT))
                    {
                        col_error.Add(new int[] { dt_error.Rows.Count, 4 });
                        isError = true;
                    }
                }

                if (string.IsNullOrWhiteSpace(dr_excel["CT"].ToString()))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 5 });
                    isError = true;
                }
                else
                {
                    if (!decimal.TryParse(dr_excel["CT"].ToString(), out ACT))
                    {
                        col_error.Add(new int[] { dt_error.Rows.Count, 5 });
                        isError = true;
                    }
                }


                if (isError || isRepet_excel)
                {
                    if (isError)
                    {
                        dt_error.Rows.Add(dr_excel.ItemArray);
                    }

                    if (isRepet_excel)
                    {
                        dt_repet_excel.Rows.Add(dr_excel.ItemArray);
                    }

                    continue;
                }

                string temp = string.Format(@"UPDATE BSPRODSTDSS SET CT={4},ACT={5}
WHERE BSPRODSTDSS.GUID=
(SELECT BSPRODSTDSS.GUID FROM BSPRODSTDSS
JOIN BSPRODSTDS ON BSPRODSTDS.GUID=BSPRODSTDSS.PGUID
JOIN BSPRODSTD ON BSPRODSTD.GUID=BSPRODSTDS.PGUID
JOIN BSWORKCENTER ON BSWORKCENTER.GUID=BSPRODSTDSS.FGUID
JOIN BSPRODUCT ON BSPRODUCT.GUID=BSPRODSTD.PGUID
WHERE BSPRODUCT.CODE='{0}'
AND BSPRODSTD.VER='{1}'
AND BSPRODSTDS.CPCODE='{2}'
AND BSWORKCENTER.CODE='{3}')", WGHelper.ReturnString(dr_excel["零件号"].ToString()), WGHelper.ReturnString(dr_excel["版本"].ToString()),
                             WGHelper.ReturnString(dr_excel["工序"].ToString()), WGHelper.ReturnString(dr_excel["工作中心编号"].ToString()), CT, ACT);

                sb.AppendLine(temp);
                sb.AppendLine();
                sqlLs.Add(temp);
            }
            rbSql.Text = sb.ToString();

            dgError.DataSource = dt_error;
            dgRepet.DataSource = dt_repet_excel;
            if (dt_error.Rows.Count > 0 || dt_repet_excel.Rows.Count > 0)
            {
                Main.SetErrorCell(dgError, col_error);
                rbSql.Text = "";
                return;
            }
            isCheck = true;
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

        public void ClearSql()
        {
            //重新上传后，清空原来的
            isCheck = false;
            sqlLs = new List<string>();
            rbSql.Text = "";
            dgError.DataSource = new DataTable();
            dgRepet.DataSource = new DataTable();
        }
    }
    
}
