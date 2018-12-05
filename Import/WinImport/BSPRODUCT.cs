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
    public partial class BSPRODUCT : Form
    {
        ExcelManager _excelManager = new ExcelManager();

        /// <summary>
        /// 产品基本信息
        /// </summary>
        DataTable _BSPRODUCT_excel = null;

        /// <summary>
        /// 需要保存的sql
        /// </summary>
        List<string> sqlLs = new List<string>();

        /// <summary>
        /// 数据库的产品类别表
        /// </summary>
        DataTable _BSPRODTYPE_DB = null;
        /// <summary>
        /// 数据库的单位设定表
        /// </summary>
        DataTable _BSUNIT_DB = null;
        /// <summary>
        /// 数据库的产品基本信息表
        /// </summary>
        DataTable _BSPRODUCT_DB = null;

        DataTable _WORKSHOP_BD = null;

        DataTable _BSPRODSERIES_DB = null;
        /// <summary>
        /// 是，否
        /// </summary>
        Dictionary<string, bool> ISINTERVALs = new Dictionary<string, bool>();

        /// <summary>
        ///  是否验证成功
        /// </summary>
        bool isCheck = false;

        public BSPRODUCT()
        {
            InitializeComponent();
            ISINTERVALs.Add("是", true);
            ISINTERVALs.Add("否", false);
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
                    _BSPRODUCT_excel = _excelManager.GetExcelTableByOleDB(opfDialog.FileName, "BSPRODUCT");
                    MessageBox.Show("读取笔数:" + _BSPRODUCT_excel.Rows.Count + "");
                }
                ClearSql();
                if (_BSPRODUCT_excel == null || _BSPRODUCT_excel.Rows.Count <= 0)
                {
                    WGMessage.ShowWarning(@"无法读取当前Excel!");
                    return;
                }
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
            dgRepet.DataSource = new DataTable();
        }

        private void btnCheck_Click(object sender, EventArgs e)
        {
            if (_BSPRODUCT_excel == null)
            {
                WGMessage.ShowWarning("请选择[产品基本信息]文件!");
                return;
            }

            if (isCheck)
            {
                WGMessage.ShowAsterisk("已验证，不用重复验证！");
                return;
            }

            string sql = "SELECT GUID,CODE,NAME,CTYPE FROM BSPRODTYPE";
            _BSPRODTYPE_DB = FillDatatablde(sql,Main.CONN_Public);

            sql = "SELECT GUID,NAME FROM BSUNIT WHERE ST = 1";
            _BSUNIT_DB = FillDatatablde(sql, Main.CONN_Public);

            sql = "SELECT GUID,CODE,NAME FROM BSPRODUCT";
            _BSPRODUCT_DB = FillDatatablde(sql, Main.CONN_Public);


            //错误
            List<int[]> col_error = new List<int[]>();

            //重复数据
            DataTable dt_repet = _BSPRODUCT_excel.Clone();

            DataTable dt_error = _BSPRODUCT_excel.Clone();

            DataTable dt_repet_excel = _BSPRODUCT_excel.Clone();

            for (int i = 0; i < _BSPRODUCT_excel.Rows.Count; i++)
            {
                bool isError = false;
                bool isRepet = false;
                bool isRepet_excel = false;

                DataRow dr_excel = _BSPRODUCT_excel.Rows[i];

                if (string.IsNullOrWhiteSpace(dr_excel["产品编号"].ToString()))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 0 });
                    isError = true;
                }

                if (string.IsNullOrWhiteSpace(dr_excel["产品编号"].ToString())
                    || _BSPRODUCT_excel.Select("产品编号='" + WGHelper.ReturnString(dr_excel["产品编号"].ToString()) + "'").Length > 1)
                {
                    //空、重复
                    isRepet_excel = true;
                }

                if (_BSPRODUCT_DB.Select("CODE='" + WGHelper.ReturnString(dr_excel["产品编号"].ToString()) + "'").Length > 0)
                {
                    // 存在
                    isRepet = true;
                }

                if (string.IsNullOrWhiteSpace(dr_excel["产品名称"].ToString()))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 1 });
                    isError = true;
                }

                if (string.IsNullOrWhiteSpace(dr_excel["产品名称"].ToString())
                    || _BSPRODUCT_excel.Select("产品名称='" + WGHelper.ReturnString(dr_excel["产品名称"].ToString()) + "'").Length > 1)
                {
                    //空、重复
                    col_error.Add(new int[] { dt_error.Rows.Count, 1 });
                    isError = true;
                }
                else
                if (_BSPRODUCT_DB.Select("NAME='" + WGHelper.ReturnString(dr_excel["产品名称"].ToString()) + "'").Length > 0)
                {
                    // 存在
                    isRepet = true;
                }

                if (string.IsNullOrWhiteSpace(dr_excel["产品类别编号"].ToString()))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 3 });
                    isError = true;
                }

                DataRow[] drs_type = _BSPRODTYPE_DB.Select("CODE = '" + dr_excel["产品类别编号"] + "'");
                if (drs_type.Length == 0)
                {
                    col_error.Add(new int[] { dt_error.Rows.Count, 3 });
                    isError = true;
                }

                if (string.IsNullOrWhiteSpace(dr_excel["单位名称"].ToString()))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 4 });
                    isError = true;
                }

                DataRow[] drs_unit = _BSUNIT_DB.Select("NAME = '" + dr_excel["单位名称"] + "'");
                if (drs_unit.Length == 0)
                {
                    col_error.Add(new int[] { dt_error.Rows.Count, 4 });
                    isError = true;
                }
                int NEEDREADY = 0;
                int NEEDTRACE = 0;
                int ISSCHEDULING = 0;
                int ISDBBARCODE = 0;
                decimal SAFETYSTOCK = 0;
                if (string.IsNullOrWhiteSpace(dr_excel["是否需要备料"].ToString()) || !int.TryParse(dr_excel["是否需要备料"].ToString(), out NEEDREADY))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 3 });
                    isError = true;
                }
                if (string.IsNullOrWhiteSpace(dr_excel["安全库存量"].ToString()) || !decimal.TryParse(dr_excel["安全库存量"].ToString(), out SAFETYSTOCK))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 3 });
                    isError = true;
                }
                if (string.IsNullOrWhiteSpace(dr_excel["是否批次管理"].ToString()) || !int.TryParse(dr_excel["是否批次管理"].ToString(), out ISDBBARCODE))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 3 });
                    isError = true;
                }
                if (string.IsNullOrWhiteSpace(dr_excel["是否追溯"].ToString()) || !int.TryParse(dr_excel["是否追溯"].ToString(), out NEEDTRACE))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 3 });
                    isError = true;
                }
                if (string.IsNullOrWhiteSpace(dr_excel["是否排产"].ToString()) || !int.TryParse(dr_excel["是否排产"].ToString(), out ISSCHEDULING))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 3 });
                    isError = true;
                }
                decimal bzknum = 0;
                decimal dbznum = 0;
                decimal xbznum = 0;
                if (string.IsNullOrWhiteSpace(dr_excel["标准框数量"].ToString()) || !decimal.TryParse(dr_excel["标准框数量"].ToString(), out bzknum))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 3 });
                    isError = true;
                }

                if (string.IsNullOrWhiteSpace(dr_excel["小包装数量"].ToString()) || !decimal.TryParse(dr_excel["小包装数量"].ToString(), out dbznum))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 3 });
                    isError = true;
                }
                if (string.IsNullOrWhiteSpace(dr_excel["大包装数量"].ToString()) || !decimal.TryParse(dr_excel["大包装数量"].ToString(), out xbznum))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 3 });
                    isError = true;
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

                string NewGUID = Guid.NewGuid().ToString();

                string temp = string.Format(@"INSERT INTO BSPRODUCT
                                             (GUID,CODE,NAME,AGUID,ANAME,ACODE,CTYPE,SPEC,BGUID,BNAME,LPQTY,SPQTY,ANUM,NEEDREADY,NEEDTRACE,ISSCHEDULING,ISDBBARCODE,SAFETYSTOCK,NOTE,ST,ND,CD)
                                             VALUES 
                                             ('{0}',{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},GETDATE(),GETDATE())", 
                                             NewGUID,Main.SetDBValue(dr_excel["产品编号"].ToString()), Main.SetDBValue(dr_excel["产品名称"].ToString()), Main.SetDBValue(drs_type[0]["GUID"].ToString()), 
                                             Main.SetDBValue(dr_excel["产品类别编号"].ToString()), Main.SetDBValue(dr_excel["产品类别名称"].ToString()), Convert.ToInt32(drs_type[0]["CTYPE"].ToString()),
                                             Main.SetDBValue(dr_excel["产品描述"].ToString()), Main.SetDBValue(drs_unit[0]["GUID"].ToString()), Main.SetDBValue(dr_excel["单位名称"].ToString()),
                                             dbznum, xbznum, bzknum, NEEDREADY, NEEDTRACE, ISSCHEDULING, ISDBBARCODE, SAFETYSTOCK,"NULL",0);


                temp += string.Format(@"INSERT INTO BSPRODUCTVER
                                       (GUID,PGUID,PCODE,PNAME,VER,ST,STATE,VDATE)
                                       VALUES  
                                       (NEWID(),'{0}',{1},{2},'{3}',0, 1 ,GETDATE())", NewGUID, Main.SetDBValue(dr_excel["产品编号"].ToString()), 
                                       Main.SetDBValue(dr_excel["产品名称"].ToString()), "A.0", 0,0);

                rbSql.Text += temp + Environment.NewLine;
                sqlLs.Add(temp);
            }
            dgError.DataSource = dt_error;
            dgRepet.DataSource = dt_repet;
            dgRepet_Excel.DataSource = dt_repet_excel;
            if (dt_error.Rows.Count > 0 || dt_repet.Rows.Count > 0 || dt_repet_excel.Rows.Count >0)
            {
                Main.SetErrorCell(dgError, col_error);
                rbSql.Text = "";
                return; 
            }
            isCheck = true;
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
    }
}
