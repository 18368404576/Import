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
    public partial class fmImport : Form
    {
        public fmImport()
        {
            InitializeComponent();
        }

        private void btnCheck_Click(object sender, EventArgs e)
        {
            string sql = @"SELECT [DOMAIN].[DMNAME],[dbo].[DOMAIN].[CAPTION] NAME,[TBDMFLD].[TBDM],[dbo].[TBDMFLD].[CAPTION],[TBDMFLD].[RI] FROM [dbo].[DOMAIN]
LEFT JOIN [dbo].[TABLEDM] ON [TABLEDM].[DMRI] = [DOMAIN].[RI]
LEFT JOIN [dbo].[TBDMFLD] ON [TBDMFLD].[DMTBRI] = [TABLEDM].[RI]
WHERE [VIEWFLD] = 1 AND DOMAIN.DMNAME IN ("+textBox1.Text+")";

            DataTable dt = FillDatatablde(sql, Main.CONN_Model);
            DataTable dt_error = dt.Clone();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                bool isHave = false;
                int index = 0;
                for (int j = 0; j < _BSPRODUCT_excel.Rows.Count; j++)
                {
                    if (dt.Rows[i]["CAPTION"].ToString() == _BSPRODUCT_excel.Rows[j]["中文"].ToString())
                    {
                        isHave = true;
                        index = j;
                        break;
                    }
                }

                if (!isHave)
                {
                    dt_error.Rows.Add(dt.Rows[i].ItemArray);
                }
                else
                {
                    string str = _BSPRODUCT_excel.Rows[index]["中文"].ToString() + " " +
                                 _BSPRODUCT_excel.Rows[index]["Updated"].ToString();

                    int num = 20;
                    if (str.Length > 12)
                    {
                        num = 35;
                    }
                    if (str.Length > 20)
                    {
                        num = 50;
                    }

                    string temp = string.Format(@"UPDATE [TBDMFLD] SET CAPTION = '{0}',UILEN = '{2}' WHERE RI = '{1}'", str, dt.Rows[i]["RI"].ToString(),num) + Environment.NewLine;
                    rbSql.Text += temp + Environment.NewLine;
                    sqlLs.Add(temp);
                }
            }
            dgError.DataSource = dt_error;
        }
        bool isCheck = false;
        List<string> sqlLs = new List<string>();
        DataTable _BSPRODUCT_excel = null;
        ExcelManager _excelManager = new ExcelManager();

        public void ClearSql()
        {
            //重新上传后，清空原来的
            isCheck = false;
            sqlLs = new List<string>();
            rbSql.Text = "";
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

        private void btnImport_Click(object sender, EventArgs e)
        {
            if (RunSql(sqlLs, Main.CONN_Model))
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
