using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using TSES.Base;

namespace WinImport
{
    public partial class BSPROSTANDARD : Form
    {
        ExcelManager _excelManager = new ExcelManager();

        private bool istest = false;
        /// <summary>
        /// 产品工艺流程
        /// </summary>
        DataTable _BSPRODSTD_excel = new DataTable();

        /// <summary>
        /// 产品工艺流程工作中心设定
        /// </summary>
        DataTable _BSPRODSTDSS_excel = new DataTable();

        /// <summary>
        /// 制造BOM
        /// </summary>
        DataTable _BSPRODBOM_excel = new DataTable();

        private DataTable _BSPRODSTDS_DB = null;

        //MRO资源寿命
        private DataTable _BSPRODPLANSS_excel = new DataTable(); 

        /// <summary>
        /// MRO资源
        /// </summary>
        DataTable _BSPRODPLAN_excel = new DataTable();
        /// <summary>
        /// 检验标准
        /// </summary>
        DataTable _BSPRODTEST_excel = new DataTable();
        /// <summary>
        /// 工艺参数
        /// </summary>
        DataTable _BSPRODPARAM_excel = new DataTable();
        /// <summary>
        /// 数据库的部门-岗位表（级联部门和岗位）
        /// </summary>
        DataTable _hrDeptPos_DB = null;
        /// <summary>
        /// 数据库的设备类别表
        /// </summary>
        DataTable _ECType_DB = null;
        /// <summary>
        /// 数据库的供应商表
        /// </summary>
        DataTable _BSSupplier_DB = null;
        /// <summary>
        /// 数据库的保养类型表
        /// </summary>
        DataTable _ECUpkeep_DB = null;
        /// <summary>
        /// 部门
        /// </summary>
        DataTable _BSDEPT_DB = null;
        /// <summary>
        /// 产品基本信息
        /// </summary>
        DataTable _BSPRODUCT_DB = null;


        private DataTable _BSPRODSTD_DB = null;

        private DataTable _BSPRODBOM_DB = null;

        private DataTable _BSPRODPLAN_DB = null;

        private DataTable _BSPRODTEST_DB = null;

        private DataTable _BSPRODPARAM_DB = null;

        private bool isnew = true;

        //没有版本的产品
        private DataTable _BSPRODUCT_DB_A = null;

        //物料(2,3,4,6)
        private DataTable _BSPRODUCT_DB_B = null;

        //工作中心
        private DataTable _BSWORKCENTER_DB = null;

        //工艺信息
        private DataTable _BSPROCESS_DB = null;

        //设备类别
        private DataTable _ECTYPE_DB = null;

        /// <summary>
        /// 1、未使用 2、使用中 3、停用
        /// </summary>
        Dictionary<string, int> STs = new Dictionary<string, int>();

        /// <summary>
        /// 数值，文本
        /// </summary>
        Dictionary<string, string> CTYPEs = new Dictionary<string, string>();

        /// <summary>
        /// 通用，产品
        /// </summary>
        Dictionary<string, string> BTYPE = new Dictionary<string, string>();

        /// <summary>
        /// 是，否
        /// </summary>
        Dictionary<string, bool> ISINTERVALs = new Dictionary<string, bool>();

        /// <summary>
        /// 需要保存的sql
        /// </summary>
        List<string> sqlLs = new List<string>();

        /// <summary>
        ///  是否验证成功
        /// </summary>
        bool isCheck = false;

        public BSPROSTANDARD()
        {
            InitializeComponent();
            ISINTERVALs.Add("是", true);
            ISINTERVALs.Add("否", false);

            CTYPEs.Add("尺寸", "尺寸");
            CTYPEs.Add("外观", "外观");
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
                    _BSPRODSTD_excel = _excelManager.GetExcelTableByOleDB(opfDialog.FileName, "BSPRODSTDA");
                    _BSPRODSTD_excel.Columns.Add("GUID");
                    MessageBox.Show("读取笔数:" + _BSPRODSTD_excel.Rows.Count + "");

                    if (_BSPRODSTD_excel == null || _BSPRODSTD_excel.Rows.Count <= 0)
                    {
                        WGMessage.ShowWarning(@"无法读取当前Excel!");
                        return;
                    }
                }
                else if (btn.Name == "btnSelect2")
                {
                    txtFile2.Text = opfDialog.FileName;
                    _BSPRODSTDSS_excel = _excelManager.GetExcelTableByOleDB(opfDialog.FileName, "BSPRODSTDAA");
                    _BSPRODSTDSS_excel.Columns.Add("GUID");
                    MessageBox.Show("读取笔数:" + _BSPRODSTDSS_excel.Rows.Count + "");

                    if (_BSPRODSTDSS_excel == null || _BSPRODSTDSS_excel.Rows.Count <= 0)
                    {
                        WGMessage.ShowWarning(@"无法读取当前Excel!");
                        return;
                    }
                }
                else if (btn.Name == "btnSelect3")
                {
                    txtFile3.Text = opfDialog.FileName;
                    _BSPRODBOM_excel = _excelManager.GetExcelTableByOleDB(opfDialog.FileName, "BSPRODBOM");
                    MessageBox.Show("读取笔数:" + _BSPRODBOM_excel.Rows.Count + "");

                    if (_BSPRODBOM_excel == null || _BSPRODBOM_excel.Rows.Count <= 0)
                    {
                        WGMessage.ShowWarning(@"无法读取当前Excel!");
                        return;
                    }
                }
                #region 没用
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
                //else if (btn.Name == "btnSelect5")
                //{
                //    txtFile5.Text = opfDialog.FileName;
                //    _BSPRODPLANSS_excel = _excelManager.GetExcelTableByOleDB(opfDialog.FileName, "BSPRODPLANSS");
                //    MessageBox.Show("读取笔数:" + _BSPRODPLANSS_excel.Rows.Count + "");

                //    if (_BSPRODPLANSS_excel == null || _BSPRODPLANSS_excel.Rows.Count <= 0)
                //    {
                //        WGMessage.ShowWarning(@"无法读取当前Excel!");
                //        return;
                //    }
                //}
                //else if (btn.Name == "btnSelect6")
                //{
                //    txtFile6.Text = opfDialog.FileName;
                //    _BSPRODTEST_excel = _excelManager.GetExcelTableByOleDB(opfDialog.FileName, "BSPRODTEST");
                //    MessageBox.Show("读取笔数:" + _BSPRODTEST_excel.Rows.Count + "");

                //    if (_BSPRODTEST_excel == null || _BSPRODTEST_excel.Rows.Count <= 0)
                //    {
                //        WGMessage.ShowWarning(@"无法读取当前Excel!");
                //        return;
                //    }
                //}
                //else if (btn.Name == "btnSelect7")
                //{
                //    txtFile7.Text = opfDialog.FileName;
                //    _BSPRODPARAM_excel = _excelManager.GetExcelTableByOleDB(opfDialog.FileName, "BSPRODPARAM");
                //    MessageBox.Show("读取笔数:" + _BSPRODPARAM_excel.Rows.Count + "");

                //    if (_BSPRODPARAM_excel == null || _BSPRODPARAM_excel.Rows.Count <= 0)
                //    {
                //        WGMessage.ShowWarning(@"无法读取当前Excel!");
                //        return;
                //    }
                //}
                #endregion
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
            dgError1.DataSource = new DataTable();
            dgRepet1_excel.DataSource = new DataTable();
            dgError2.DataSource = new DataTable();
            dgRepet2_excel.DataSource = new DataTable();
            dgError3.DataSource = new DataTable();
            dgRepet3_excel.DataSource = new DataTable();
            dgError4.DataSource = new DataTable();
            dgRepet4_excel.DataSource = new DataTable();
            dgError5.DataSource = new DataTable();
            dgRepet5_excel.DataSource = new DataTable();
        }

        private void btnCheck_Click(object sender, EventArgs e)
        {
            if (_BSPRODSTD_excel == null)
            {
                WGMessage.ShowWarning("请选择[工艺路线]文件!");
                return;
            }
            if (_BSPRODSTDSS_excel == null)
            {
                WGMessage.ShowWarning("请选择[工艺路线工作中心]文件!");
                return;
            }
            if (_BSPRODBOM_excel == null)
            {
                WGMessage.ShowWarning("请选择[制造BOM]文件!");
                return;
            }
            //if (_BSPRODPLAN_excel == null)
            //{
            //    WGMessage.ShowWarning("请选择[MRO资源]文件!");
            //    return;
            //}
            //if (_BSPRODTEST_excel == null)
            //{
            //    WGMessage.ShowWarning("请选择[检验标准]文件!");
            //    return;
            //}
            //if (_BSPRODPARAM_excel == null)
            //{
            //    WGMessage.ShowWarning("请选择[工艺参数]文件!");
            //    return;
            //}

            //try
            {
                if (_BSPRODSTD_excel == null)
                {
                    _BSPRODSTD_excel = new DataTable();
                }
                if (_BSPRODSTDSS_excel == null)
                {
                    _BSPRODSTDSS_excel = new DataTable();
                }
                if (_BSPRODBOM_excel == null)
                {
                    _BSPRODBOM_excel = new DataTable();
                }
                if (_BSPRODPLAN_excel == null)
                {
                    _BSPRODPLAN_excel = new DataTable();
                }
                if (_BSPRODTEST_excel == null)
                {
                    _BSPRODTEST_excel = new DataTable();
                }
                if (_BSPRODPARAM_excel == null)
                {
                    _BSPRODPARAM_excel = new DataTable();
                }
                if (_BSPRODPLANSS_excel == null)
                {
                    _BSPRODPLANSS_excel = new DataTable();
                }

                if (isCheck)
                {
                    WGMessage.ShowAsterisk("已验证，不用重复验证！");
                    return;
                }

                string sql = @" SELECT BSPRODUCT.GUID,BSPRODUCT.CODE,BSPRODUCT.NAME,BSPRODUCTVER.[VER] FROM BSPRODUCT 
     LEFT JOIN BSPRODUCTVER ON BSPRODUCTVER.PGUID = BSPRODUCT.GUID";
                _BSPRODUCT_DB = FillDatatablde(sql, Main.CONN_Public);

                sql = @" SELECT GUID,CODE,NAME FROM BSPRODUCT";
                _BSPRODUCT_DB_A = FillDatatablde(sql, Main.CONN_Public);

                sql = @"SELECT GUID,CODE,NAME FROM BSPROCESS WHERE ST = 1";
                _BSPROCESS_DB = FillDatatablde(sql, Main.CONN_Public);

                sql = @"SELECT GUID,CODE,NAME FROM ECTYPE";
                _ECTYPE_DB = FillDatatablde(sql, Main.CONN_Public);

                sql = @"SELECT GUID,CODE,[NAME] FROM [dbo].[BSWORKCENTER] WHERE ST = 0";
                _BSWORKCENTER_DB = FillDatatablde(sql, Main.CONN_Public);

                sql = @"SELECT BSPRODUCT.[CODE] 产品编号,BSPRODSTD.VER 版本号,BSPRODSTDA.[CPCODE] 工序,BSPRODSTDA.[GUID] FROM BSPRODSTDA
                        LEFT JOIN BSPRODSTD ON [BSPRODSTD].[GUID] = [BSPRODSTDA].[PGUID]
                        LEFT JOIN BSPRODUCT ON BSPRODUCT.[GUID] = [dbo].[BSPRODSTD].[PGUID]
                        LEFT JOIN BSPRODUCTVER ON BSPRODUCTVER.[PGUID] = [dbo].[BSPRODUCT].[GUID]";
                _BSPRODSTDS_DB = FillDatatablde(sql, Main.CONN_Public);

                sql = @"SELECT BSPRODUCT.[CODE],[VER] FROM BSPRODSTD
                        LEFT JOIN BSPRODUCT ON [dbo].[BSPRODUCT].[GUID] = BSPRODSTD.[PGUID]";
                _BSPRODSTD_DB = FillDatatablde(sql, Main.CONN_Public);
                
                sql = @"SELECT  BSPRODUCT.[CODE],[VER] FROM BSPRODBOM
                        LEFT JOIN BSPRODUCT ON [dbo].[BSPRODUCT].[GUID] = BSPRODBOM.[PGUID]";
                _BSPRODBOM_DB = FillDatatablde(sql, Main.CONN_Public);
                
    //            sql = @"SELECT  BSPRODUCT.[CODE],[VER] FROM BSPRODPLAN
    //LEFT JOIN BSPRODUCT ON [dbo].[BSPRODUCT].[GUID] = BSPRODPLAN.[PGUID]";
    //            _BSPRODPLAN_DB = FillDatatablde(sql, Main.CONN_Public);
                
    //            sql = @"SELECT  BSPRODUCT.[CODE],[VER] FROM BSPRODTEST
    //LEFT JOIN BSPRODUCT ON [dbo].[BSPRODUCT].[GUID] = BSPRODTEST.[PGUID]";
    //            _BSPRODTEST_DB = FillDatatablde(sql, Main.CONN_Public);

    //            sql = @"SELECT  BSPRODUCT.[CODE],[VER] FROM BSPRODPARAM
    //LEFT JOIN BSPRODUCT ON [dbo].[BSPRODUCT].[GUID] = BSPRODPARAM.[PGUID]";
    //            _BSPRODPARAM_DB = FillDatatablde(sql, Main.CONN_Public);

                #region 工艺路线
                List<int[]> col_error1 = new List<int[]>();
                DataTable dt_repet1 = _BSPRODSTD_excel.Clone();
                DataTable dt_error1 = _BSPRODSTD_excel.Clone();
                DataTable dt_repet1_excel = _BSPRODSTD_excel.Clone();
                DataTable dt_BSPRODSTD_ADD = new DataTable();
                dt_BSPRODSTD_ADD.Columns.Add("产品编号");
                dt_BSPRODSTD_ADD.Columns.Add("版本号");
                dt_BSPRODSTD_ADD.Columns.Add("GUID");

                DataTable DT_BSPRODSTDS_ADD = new DataTable();
                DT_BSPRODSTDS_ADD.Columns.Add("产品编号");
                DT_BSPRODSTDS_ADD.Columns.Add("工序");
                DT_BSPRODSTDS_ADD.Columns.Add("GUID");

                for (int i = 0; i < _BSPRODSTD_excel.Rows.Count; i++)
                {
                    bool isError = false;
                    bool isRepet = false;
                    bool isRepet_excel = false;

                    DataRow dr_excel = _BSPRODSTD_excel.Rows[i];

                    DataRow[] drs_BSPRODSTD_PRODUCT =
                        _BSPRODUCT_DB.Select(string.Format("CODE = '{0}' AND VER = '{1}'", WGHelper.ReturnString(dr_excel["产品编号"].ToString()), dr_excel["版本号"].ToString()));

                    if (drs_BSPRODSTD_PRODUCT.Length == 0)
                    {
                        //空、不存在
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 0 });
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 1 });
                        isError = true;
                    }

                    DataRow[] drs_repet = _BSPRODSTDS_DB.Select(string.Format("产品编号 = '{0}' AND 版本号 = '{1}' AND 工序 = '{2}' ", WGHelper.ReturnString(dr_excel["产品编号"].ToString()), dr_excel["版本号"].ToString(),dr_excel["工序"].ToString()));                   
                    if (drs_repet.Length > 0)
                    {
                        isRepet = true;
                    }
                    DataRow[] drs_repet_excel = _BSPRODSTD_excel.Select(string.Format("产品编号 = '{0}' AND 版本号 = '{1}' AND 工序 = '{2}'", WGHelper.ReturnString(dr_excel["产品编号"].ToString()), dr_excel["版本号"].ToString(),dr_excel["工序"].ToString()));

                    if(drs_repet_excel.Length >1)
                    {
                        isRepet_excel = true;
                    }

                    int SNO = 0;
                    if (string.IsNullOrWhiteSpace(dr_excel["序号"].ToString())
                        || !int.TryParse(dr_excel["序号"].ToString(), out SNO))
                    {
                        //空
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 2 });
                        isError = true;
                    }

                    if (string.IsNullOrWhiteSpace(dr_excel["工序"].ToString()))
                    {
                        //空
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 3 });
                        isError = true;
                    }

                    DataRow[] drs_BSPROCESS = _BSPROCESS_DB.Select("NAME = '" + dr_excel["工艺信息名称"] + "'");

                    if (drs_BSPROCESS.Length == 0)
                    {
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 4 });
                        isError = true;
                    }

                    DataRow[] drs_ECTYPE = _ECTYPE_DB.Select("CODE = '" + WGHelper.ReturnString(dr_excel["设备类别编号"].ToString()) +"'");

                    if (drs_ECTYPE.Length == 0)
                    {
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 5 });
                        isError = true;
                    }
                    decimal dnum = 0;
                    if (!decimal.TryParse(dr_excel["允许报工超出量"].ToString(), out dnum))
                    {
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 6 });
                        isError = true;
                    }

                    if (isError || isRepet || isRepet_excel)
                    {
                        if (isError)
                        {
                            dt_error1.Rows.Add(dr_excel.ItemArray);
                        }
                        if (isRepet)
                        {
                            dt_repet1.Rows.Add(dr_excel.ItemArray);
                        }
                        if(isRepet_excel)
                        {
                            dt_repet1_excel.Rows.Add(dr_excel.ItemArray);
                        }
                        continue;
                    }

                    DataRow[] drs_BSPRODSTD_ADDNEW = dt_BSPRODSTD_ADD.Select(string.Format("产品编号 = '{0}' AND 版本号 = '{1}' ",dr_excel["产品编号"].ToString(),dr_excel["版本号"]));

                    string BSPRODSTD_PGUID;
                    if (drs_BSPRODSTD_ADDNEW.Length == 0)
                    {
                        Guid kidGUID = Guid.NewGuid();
                        BSPRODSTD_PGUID = kidGUID.ToString();
                        DataRow NEWROW = dt_BSPRODSTD_ADD.NewRow();
                        NEWROW["产品编号"] = dr_excel["产品编号"];
                        NEWROW["版本号"] = dr_excel["版本号"];
                        NEWROW["GUID"] = kidGUID;
                        dt_BSPRODSTD_ADD.Rows.Add(NEWROW);

                        string temp = string.Format(@" INSERT INTO [dbo].[BSPRODSTD]
                         ([GUID],[PGUID],[VER],[NOTE],[ST],[ND])
                          VALUES  
                         ('{0}','{1}', '{2}','',1,GETDATE() );", 
                         kidGUID, drs_BSPRODSTD_PRODUCT[0]["GUID"].ToString(), dr_excel["版本号"]);
                        sqlLs.Add(temp);
                    }
                    else
                    {
                        BSPRODSTD_PGUID = drs_BSPRODSTD_ADDNEW[0]["GUID"].ToString();
                    }
                    string SUNGUID = Guid.NewGuid().ToString();
                    dr_excel["GUID"] = SUNGUID;
                    string temp1 = string.Format(@" INSERT INTO [dbo].[BSPRODSTDA]
                                                    ([GUID],[PGUID],[SNO],[CPCODE],[AGUID],[ACODE],[ANAME],[BGUID],[BNAME],[OVERNUM])
                                                    VALUES  ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}',{9})",
                                                    SUNGUID,BSPRODSTD_PGUID, dr_excel["序号"], WGHelper.ReturnString(dr_excel["工序"].ToString()), 
                                                    drs_ECTYPE[0]["GUID"], dr_excel["设备类别编号"], dr_excel["设备类别名称"], drs_BSPROCESS[0]["GUID"], 
                                                    WGHelper.ReturnString(dr_excel["工艺信息名称"].ToString()), dnum);

                    DataRow drNew_BSPRODSTDS = DT_BSPRODSTDS_ADD.NewRow();
                    drNew_BSPRODSTDS["产品编号"] = dr_excel["产品编号"].ToString();
                    drNew_BSPRODSTDS["工序"] = dr_excel["工序"].ToString();
                    drNew_BSPRODSTDS["GUID"] = SUNGUID;
                    DT_BSPRODSTDS_ADD.Rows.Add(drNew_BSPRODSTDS);
                
                    sqlLs.Add(temp1);
                }
                //dt_error1.Columns.Remove("GUID");
                //dt_repet1.Columns.Remove("GUID");
                //dt_repet1_excel.Columns.Remove("GUID");
                dgError1.DataSource = dt_error1;
                dgRepet1.DataSource = dt_repet1;
                dgRepet1_excel.DataSource = dt_repet1_excel;
                if (dt_error1.Rows.Count > 0 || dt_repet1.Rows.Count > 0 || dt_repet1_excel.Rows.Count > 0)
                {
                    Main.SetErrorCell(dgError1, col_error1);
                    rbSql.Text = "";
                    return;
                }
            #endregion

                #region 工艺路线工作中心

                List<int[]> col_error2 = new List<int[]>();
                DataTable dt_repet2 = _BSPRODSTDSS_excel.Clone();
                DataTable dt_error2 = _BSPRODSTDSS_excel.Clone();
                DataTable dt_repet2_excel = _BSPRODSTDSS_excel.Clone();

                for (int i = 0; i < _BSPRODSTDSS_excel.Rows.Count; i++)
                {
                    bool isError = false;
                    bool isRepet = false;

                    DataRow dr_excel = _BSPRODSTDSS_excel.Rows[i];

                    DataRow[] drs_BSPRODSTDSS_PRODUCT =
                        _BSPRODSTD_excel.Select(string.Format("产品编号 = '{0}' AND 版本号 = '{1}' AND 工序 = '{2}' ",
                            WGHelper.ReturnString(dr_excel["产品编号"].ToString()), dr_excel["版本号"].ToString(), dr_excel["工序"].ToString()));

                    if (drs_BSPRODSTDSS_PRODUCT.Length == 0)
                    {
                        //空、不存在
                        col_error2.Add(new int[] {dt_error2.Rows.Count, 0});
                        col_error2.Add(new int[] {dt_error2.Rows.Count, 1});
                        col_error2.Add(new int[] { dt_error2.Rows.Count, 2 });
                        isError = true;
                    }

                    DataRow[] drs_repit = _BSPRODSTDSS_excel.Select("产品编号 = '"+WGHelper.ReturnString(dr_excel["产品编号"].ToString())+"' AND 版本号 = '"+WGHelper.ReturnString(dr_excel["版本号"].ToString())+"' AND 工序 = '"+dr_excel["工序"].ToString()+"' AND 工作中心编号 = '"+WGHelper.ReturnString(dr_excel["工作中心编号"].ToString())+"' ");
                    if(drs_repit.Length > 1)
                    {
                        isRepet = true;
                    }

                    //int SNO = 0;
                    //if (string.IsNullOrWhiteSpace(dr_excel["优先级"].ToString())
                    //    || !int.TryParse(dr_excel["优先级"].ToString(), out SNO))
                    //{
                    //    //空
                    //    col_error2.Add(new int[] { dt_error2.Rows.Count, 3 });
                    //    isError = true;
                    //}

                    DataRow[] drs_BSPRODSTDSS_BSWORKCENTER = _BSWORKCENTER_DB.Select("CODE = '" + dr_excel["工作中心编号"].ToString() + "'");

                    if (!istest)
                    {

                        if (drs_BSPRODSTDSS_BSWORKCENTER.Length == 0)
                        {
                            //空
                            col_error2.Add(new int[] {dt_error2.Rows.Count, 4});
                            isError = true;
                        }
                    }
                    else
                    {
                        drs_BSPRODSTDSS_BSWORKCENTER = drs_BSPRODSTDSS_PRODUCT;
                    }

                    double dnum = 0;
                    if (string.IsNullOrWhiteSpace(dr_excel["工价"].ToString())
                        || !double.TryParse(dr_excel["工价"].ToString(), out dnum))
                    {
                        //空
                        col_error2.Add(new int[] { dt_error2.Rows.Count, 5 });
                        isError = true;
                    }

                    //double dnum = 0;
                    //if (string.IsNullOrWhiteSpace(dr_excel["标准工时(秒)"].ToString())
                    //    || !double.TryParse(dr_excel["标准工时(秒)"].ToString(), out dnum))
                    //{
                    //    //空
                    //    col_error2.Add(new int[] { dt_error2.Rows.Count, 5 });
                    //    isError = true;
                    //}

                    //if (string.IsNullOrWhiteSpace(dr_excel["C/T(秒)"].ToString())
                    //   || !double.TryParse(dr_excel["C/T(秒)"].ToString(), out dnum))
                    //{
                    //    //空
                    //    col_error2.Add(new int[] { dt_error2.Rows.Count, 6 });
                    //    isError = true;
                    //}

                    //if (string.IsNullOrWhiteSpace(dr_excel["前置准备周期(分钟)"].ToString())
                    //   || !double.TryParse(dr_excel["前置准备周期(分钟)"].ToString(), out dnum))
                    //{
                    //    //空
                    //    col_error2.Add(new int[] { dt_error2.Rows.Count, 7 });
                    //    isError = true;
                    //}

                    //if (string.IsNullOrWhiteSpace(dr_excel["目标产能利用率(%)"].ToString())
                    //   || !double.TryParse(dr_excel["目标产能利用率(%)"].ToString(), out dnum))
                    //{
                    //    //空
                    //    col_error2.Add(new int[] { dt_error2.Rows.Count, 8 });
                    //    isError = true;
                    //}

                    if (isError || isRepet)
                    {
                        if (isError)
                        {
                            dt_error2.Rows.Add(dr_excel.ItemArray);
                        }
                        if (isRepet)
                        {
                            dt_repet2.Rows.Add(dr_excel.ItemArray);
                        }
                        continue;
                    }
                    string temp1 = string.Format(@" INSERT INTO [dbo].[BSPRODSTDAA]
                                                    ([GUID] ,[PGUID],[AGUID],[ACODE],[ANAME],[PRICE])
                                                    VALUES  
                                                    (NEWID(),'{0}','{1}','{2}','{3}',{4})", 
                                                    drs_BSPRODSTDSS_PRODUCT[0]["GUID"].ToString(), drs_BSPRODSTDSS_BSWORKCENTER[0]["GUID"], dr_excel["工作中心编号"], dr_excel["工作中心名称"], dnum);
                  //  rbSql.Text += temp1 + Environment.NewLine;
                    sqlLs.Add(temp1);
                }
                //dt_error2.Columns.Remove("GUID");
                //dt_error2.Columns.Remove("Columns6");
                //dt_error2.Columns.Remove("Columns7");
                //dt_error2.Columns.Remove("Columns8");
                //dt_repet2.Columns.Remove("GUID");
                //dt_repet2.Columns.Remove("Columns6");
                //dt_repet2.Columns.Remove("Columns7");
                //dt_repet2.Columns.Remove("Columns8");
                dgError2.DataSource = dt_error2;
                dgRepet2_excel.DataSource = dt_repet2;
                if (dt_error2.Rows.Count > 0 || dt_repet2.Rows.Count > 0 || dt_repet2_excel.Rows.Count >0)
                {
                    Main.SetErrorCell(dgError2, col_error2);
                    rbSql.Text = "";
                    return;
                }
            #endregion

                #region 制造BOM清单
                List<int[]> col_error3 = new List<int[]>();
                DataTable dt_repet3 = _BSPRODBOM_excel.Clone();
                DataTable dt_error3 = _BSPRODBOM_excel.Clone();
                DataTable dt_repet3_excel = _BSPRODBOM_excel.Clone();
                DataTable dt_BSPRODBOM_ADD = new DataTable();
                dt_BSPRODBOM_ADD.Columns.Add("产品编号");
                dt_BSPRODBOM_ADD.Columns.Add("版本号");
                dt_BSPRODBOM_ADD.Columns.Add("GUID");

                for (int i = 0; i < _BSPRODBOM_excel.Rows.Count; i++)
                {
                    bool isError = false;
                    bool isRepet = false;
                    bool isRepet_excel = false;

                    DataRow dr_excel = _BSPRODBOM_excel.Rows[i];

                    DataRow[] drs_BSPRODSTD_PRODUCT =
          _BSPRODUCT_DB.Select(string.Format("CODE = '{0}' AND VER = '{1}'", WGHelper.ReturnString(dr_excel["产品编号"].ToString()), dr_excel["版本号"].ToString()));

                    if (drs_BSPRODSTD_PRODUCT.Length == 0)
                    {
                        //空、不存在
                        col_error3.Add(new int[] { dt_error3.Rows.Count, 0 });
                        col_error3.Add(new int[] { dt_error3.Rows.Count, 1 });
                        isError = true;
                    }

                    DataRow[] drs_repet = _BSPRODBOM_DB.Select(string.Format("CODE = '{0}' AND VER = '{1}' ", WGHelper.ReturnString(dr_excel["产品编号"].ToString()), dr_excel["版本号"].ToString()));
                   
                    if(drs_repet.Length >0)
                    {
                        isRepet = true;
                    }

                    DataRow[] drs_repet_excel = _BSPRODBOM_excel.Select(string.Format("产品编号 = '{0}' AND 版本号 = '{1}' AND 工序 = '{2}' AND 子件编号 = '{3}' ", WGHelper.ReturnString(dr_excel["产品编号"].ToString()), dr_excel["版本号"].ToString(), WGHelper.ReturnString(dr_excel["工序"].ToString()), WGHelper.ReturnString(dr_excel["子件编号"].ToString())));

                    if (drs_repet_excel.Length > 1)
                    {
                        isRepet_excel = true;
                    }

                    int SNO = 0;
                    if (string.IsNullOrWhiteSpace(dr_excel["序号"].ToString())
                        || !int.TryParse(dr_excel["序号"].ToString(), out SNO))
                    {
                        //空
                        col_error3.Add(new int[] { dt_error3.Rows.Count, 2 });
                        isError = true;
                    }

                    if (string.IsNullOrWhiteSpace(dr_excel["工序"].ToString()))
                    {
                        //空
                        col_error3.Add(new int[] { dt_error3.Rows.Count, 3 });
                        isError = true;
                    }

                    DataRow[] drs_BSPRODSTD_PRODUCT_Tiny =
    _BSPRODUCT_DB.Select(string.Format("CODE = '{0}' ", WGHelper.ReturnString(dr_excel["产品编号"].ToString())));

                    if (drs_BSPRODSTD_PRODUCT.Length == 0)
                    {
                        //空、不存在
                        col_error3.Add(new int[] { dt_error3.Rows.Count, 4 });
                        isError = true;
                    }

                    decimal mynum = 0;
                    if (string.IsNullOrWhiteSpace(dr_excel["数量"].ToString())
                        || !decimal.TryParse(dr_excel["数量"].ToString(), out mynum))
                    {
                        //空
                        col_error3.Add(new int[] { dt_error3.Rows.Count, 5 });
                        isError = true;
                    }

                    DataRow[] drs_drNew_BSPRODSTDS;
                    if (isnew)
                    {
                        drs_drNew_BSPRODSTDS = _BSPRODSTDS_DB.Select("产品编号 = '" + dr_excel["产品编号"].ToString() + "' AND 工序 = '" + dr_excel["工序"].ToString() + "'");
                    }
                    else
                    {
                        drs_drNew_BSPRODSTDS = DT_BSPRODSTDS_ADD.Select("产品编号 = '" + dr_excel["产品编号"].ToString() + "' AND 工序 = '" + dr_excel["工序"].ToString() + "'");
                    }

                    if (drs_drNew_BSPRODSTDS.Length == 0)
                    {
                        col_error3.Add(new int[] { dt_error3.Rows.Count, 0 });
                        isError = true;
                        col_error3.Add(new int[] { dt_error3.Rows.Count, 3 });
                        isError = true;
                    }

                    DataRow[] drs_BSPRODUCT_DB_A = _BSPRODUCT_DB_A.Select("CODE = '" + dr_excel["子件编号"] + "'");

                    if (drs_BSPRODUCT_DB_A.Length == 0)
                    {
                        col_error3.Add(new int[] { dt_error3.Rows.Count, 4 });
                        isError = true;
                    }

                    if (isError || isRepet || isRepet_excel)
                    {
                        if (isError)
                        {
                            dt_error3.Rows.Add(dr_excel.ItemArray);
                        }
                        if(isRepet)
                        {
                            dt_repet3.Rows.Add(dr_excel.ItemArray);
                        }
                        if(isRepet_excel)
                        {
                            dt_repet3_excel.Rows.Add(dr_excel.ItemArray);
                        }
                        continue;
                    }

                    DataRow[] drs_BSPRODBOM_ADDNEW = dt_BSPRODBOM_ADD.Select(string.Format("产品编号 = '{0}' AND 版本号 = '{1}' ", dr_excel["产品编号"].ToString(), dr_excel["版本号"]));

                    string BSPRODBOM_PGUID;
                    if (drs_BSPRODBOM_ADDNEW.Length == 0)
                    {
                        Guid kidGUID = Guid.NewGuid();
                        BSPRODBOM_PGUID = kidGUID.ToString();
                        DataRow NEWROW = dt_BSPRODBOM_ADD.NewRow();
                        NEWROW["产品编号"] = dr_excel["产品编号"];
                        NEWROW["版本号"] = dr_excel["版本号"];
                        NEWROW["GUID"] = kidGUID;
                        dt_BSPRODBOM_ADD.Rows.Add(NEWROW);

                        string temp = string.Format(@" INSERT INTO [dbo].[BSPRODBOM]
                                                      ([GUID],[PGUID],[VER],[NOTE],[ST],[ND])
		                                               VALUES
                                                       ('{0}','{1}','{2}','',1 ,GETDATE());", 
                                                       kidGUID, drs_BSPRODSTD_PRODUCT[0]["GUID"].ToString(), dr_excel["版本号"]);
                        sqlLs.Add(temp);
                    }
                    else
                    {
                        BSPRODBOM_PGUID = drs_BSPRODBOM_ADDNEW[0]["GUID"].ToString();
                    }

                    string temp1 = string.Format(@" INSERT INTO [dbo].[BSPRODBOMA]
				                                    ([GUID],[PGUID],[SNO],[AGUID],[CPCODE],[BGUID],[BCODE],[BNAME],[PNUM])
				                                    VALUES 
                                                    ('{0}','{1}',{2},'{3}','{4}','{5}','{6}','{7}',{8})", 
                                                    Guid.NewGuid().ToString(), BSPRODBOM_PGUID, dr_excel["序号"], drs_drNew_BSPRODSTDS[0]["GUID"], dr_excel["工序"],
                                                    drs_BSPRODUCT_DB_A[0]["GUID"], dr_excel["产品编号"],dr_excel["产品名称"], dr_excel["数量"]);
                    //rbSql.Text += temp1 + Environment.NewLine;
                    sqlLs.Add(temp1);
                }
                dgError3.DataSource = dt_error3;
                dgRepet3.DataSource = dt_repet3;
                dgRepet3_excel.DataSource = dt_repet3_excel;
                if (dt_error3.Rows.Count > 0 || dt_repet3.Rows.Count > 0 || dt_repet3_excel.Rows.Count >0)
                {
                    Main.SetErrorCell(dgError3, col_error3);
                    rbSql.Text = "";
                    return;
                }
                #endregion
                #region 没用
                //            #region MRO资源
                //            List<int[]> col_error4 = new List<int[]>();
                //            DataTable dt_repet4 = _BSPRODPLAN_excel.Clone();
                //            DataTable dt_error4 = _BSPRODPLAN_excel.Clone();
                //            DataTable dt_repet4_excel = _BSPRODPLAN_excel.Clone();
                //            DataTable dt_BSPRODPLAN_ADD = new DataTable();
                //            dt_BSPRODPLAN_ADD.Columns.Add("产品编号");
                //            dt_BSPRODPLAN_ADD.Columns.Add("版本号");
                //            dt_BSPRODPLAN_ADD.Columns.Add("GUID");
                //            DataTable dt_BSPRODPLANS_ADD = new DataTable();
                //            dt_BSPRODPLANS_ADD.Columns.Add("产品编号");
                //            dt_BSPRODPLANS_ADD.Columns.Add("版本号");
                //            dt_BSPRODPLANS_ADD.Columns.Add("工序");
                //            dt_BSPRODPLANS_ADD.Columns.Add("物料母件编号");
                //            dt_BSPRODPLANS_ADD.Columns.Add("GUID");

                //            for (int i = 0; i < _BSPRODPLAN_excel.Rows.Count; i++)
                //            {
                //                bool isError = false;
                //                bool isRepet = false;
                //                bool isRepet_excel = false;

                //                DataRow dr_excel = _BSPRODPLAN_excel.Rows[i];

                //                DataRow[] drs_BSPRODPLAN_PRODUCT =
                //                    _BSPRODUCT_DB.Select(string.Format("CODE = '{0}' AND VER = '{1}'",
                //                        WGHelper.ReturnString(dr_excel["产品编号"].ToString()), dr_excel["版本号"].ToString()));

                //                DataRow[] drs_repet_excel = _BSPRODPLAN_excel.Select(string.Format("产品编号 = '{0}' AND 版本号 = '{1}' AND 工序 = '{2}' AND 物料编号 = '{3}'",
                //                        WGHelper.ReturnString(dr_excel["产品编号"].ToString()), dr_excel["版本号"].ToString(), WGHelper.ReturnString(dr_excel["工序"].ToString()), WGHelper.ReturnString(dr_excel["物料编号"].ToString())));

                //                if (drs_repet_excel.Length > 1)
                //                {
                //                    isRepet_excel = true;
                //                }

                //                DataRow[] drs_repet = _BSPRODPLAN_DB.Select(string.Format("CODE = '{0}' AND VER = '{1}'",
                //                        WGHelper.ReturnString(dr_excel["产品编号"].ToString()), dr_excel["版本号"].ToString()));

                //                if(drs_repet.Length > 0)
                //                {
                //                    isRepet = true;
                //                }


                //                if (drs_BSPRODPLAN_PRODUCT.Length == 0)
                //                {
                //                    //空、不存在
                //                    col_error4.Add(new int[] {dt_error4.Rows.Count, 0});
                //                    col_error4.Add(new int[] {dt_error4.Rows.Count, 1});
                //                    isError = true;
                //                }

                //                int SNO = 0;
                //                if (string.IsNullOrWhiteSpace(dr_excel["序号"].ToString())
                //                    || !int.TryParse(dr_excel["序号"].ToString(), out SNO))
                //                {
                //                    //空
                //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 2 });
                //                    isError = true;
                //                }

                //                if (string.IsNullOrWhiteSpace(dr_excel["工序"].ToString()))
                //                {
                //                    //空
                //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 3 });
                //                    isError = true;
                //                }

                //                DataRow[] drs_BSPRODSTD_PRODUCT_B_Tiny =
                //_BSPRODUCT_DB_B.Select(string.Format("CODE = '{0}' ", WGHelper.ReturnString(dr_excel["物料编号"].ToString())));

                //                if (drs_BSPRODSTD_PRODUCT_B_Tiny.Length == 0)
                //                {
                //                    //空、不存在
                //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 4 });
                //                    isError = true;
                //                }

                //                SNO = 0;
                //                if (string.IsNullOrWhiteSpace(dr_excel["数量"].ToString())
                //                    || !int.TryParse(dr_excel["数量"].ToString(), out SNO))
                //                {
                //                    //空
                //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 5 });
                //                    isError = true;
                //                }

                //                SNO = 0;
                //                if (!string.IsNullOrWhiteSpace(dr_excel["寿命(pcs)"].ToString()) )
                //                {
                //                    if (!int.TryParse(dr_excel["寿命(pcs)"].ToString(), out SNO))
                //                    {
                //                        //空
                //                        col_error4.Add(new int[] {dt_error4.Rows.Count, 6});
                //                        isError = true;
                //                    }
                //                }
                //                else
                //                {
                //                    dr_excel["寿命(pcs)"] = "null";
                //                }

                //                DataRow[] drs_BSPRODPLANSS_PRODUCT;
                //                if (isnew)
                //                {
                //                    drs_BSPRODPLANSS_PRODUCT =
                //                        _BSPRODSTDS_DB.Select(string.Format("产品编号 = '{0}' AND 版本号 = '{1}' AND 工序 = '{2}' ",
                //                            WGHelper.ReturnString(dr_excel["产品编号"].ToString()), dr_excel["版本号"].ToString(), dr_excel["工序"].ToString()));
                //                }
                //                else
                //                {
                //                    drs_BSPRODPLANSS_PRODUCT =
                //                        _BSPRODSTD_excel.Select(string.Format("产品编号 = '{0}' AND 版本号 = '{1}' AND 工序 = '{2}' ",
                //                            WGHelper.ReturnString(dr_excel["产品编号"].ToString()), dr_excel["版本号"].ToString(), dr_excel["工序"].ToString()));
                //                }


                //                if (drs_BSPRODPLANSS_PRODUCT.Length == 0)
                //                {
                //                    //空、不存在
                //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 0 });
                //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 1 });
                //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 2 });
                //                    isError = true;
                //                }

                //                DataRow[] drs_BSPRODPLAN_ADDNEW = dt_BSPRODPLAN_ADD.Select(string.Format("产品编号 = '{0}' AND 版本号 = '{1}' ", dr_excel["产品编号"].ToString(), dr_excel["版本号"]));

                //                if (isError || isRepet || isRepet_excel)
                //                {
                //                    if (isError)
                //                    {
                //                        dt_error4.Rows.Add(dr_excel.ItemArray);
                //                    }
                //                    if (isRepet)
                //                    {
                //                        dt_repet4.Rows.Add(dr_excel.ItemArray);
                //                    }
                //                    if(isRepet_excel)
                //                    {
                //                        dt_repet4_excel.Rows.Add(dr_excel.ItemArray);
                //                    }
                //                    continue;
                //                }

                //                string BSPRODPLAN_PGUID;
                //                if (drs_BSPRODPLAN_ADDNEW.Length == 0)
                //                {
                //                    Guid kidGUID = Guid.NewGuid();
                //                    BSPRODPLAN_PGUID = kidGUID.ToString();
                //                    DataRow NEWROW = dt_BSPRODPLAN_ADD.NewRow();
                //                    NEWROW["产品编号"] = dr_excel["产品编号"];
                //                    NEWROW["版本号"] = dr_excel["版本号"];
                //                    NEWROW["GUID"] = kidGUID;
                //                    dt_BSPRODPLAN_ADD.Rows.Add(NEWROW);

                //                    string temp = string.Format(@" 
                //            INSERT INTO [dbo].[BSPRODPLAN]
                //                    ( [GUID] ,
                //                      [PGUID] ,
                //                      [VER] ,
                //                      [NOTE] ,
                //                      [ST] ,
                //                      [ND]
                //                    )
                //            VALUES  ( '{0}' , -- GUID - uniqueidentifier
                //                      '{1}' , -- PGUID - uniqueidentifier
                //                      '{2}' , -- VER - nvarchar(10)
                //                      N'' , -- NOTE - nvarchar(200)
                //                      1 , -- ST - int
                //                      GETDATE() 
                //                    );", kidGUID, drs_BSPRODPLAN_PRODUCT[0]["GUID"].ToString(), dr_excel["版本号"]);
                //                   // rbSql.Text += temp + Environment.NewLine;
                //                    sqlLs.Add(temp);
                //                }
                //                else
                //                {
                //                    BSPRODPLAN_PGUID = drs_BSPRODPLAN_ADDNEW[0]["GUID"].ToString();
                //                }

                //                string newguid = Guid.NewGuid().ToString();
                //                dr_excel["GUID"] = newguid;
                //                DataRow drNew_BSPRODPLANS_ADD = dt_BSPRODPLANS_ADD.NewRow();
                //                drNew_BSPRODPLANS_ADD["产品编号"] = dr_excel["产品编号"].ToString();
                //                drNew_BSPRODPLANS_ADD["版本号"] = dr_excel["版本号"].ToString();
                //                drNew_BSPRODPLANS_ADD["工序"] = dr_excel["工序"].ToString();
                //                drNew_BSPRODPLANS_ADD["物料母件编号"] = dr_excel["物料编号"].ToString();
                //                drNew_BSPRODPLANS_ADD["GUID"] = newguid;
                //                dt_BSPRODPLANS_ADD.Rows.Add(drNew_BSPRODPLANS_ADD);

                //                string temp1 = string.Format(@" 
                //                INSERT INTO [dbo].[BSPRODPLANS]
                //                    ( [GUID] ,
                //                      [PGUID] ,
                //                      [SNO] ,
                //                      [FGUID] ,
                //                      [MGUID] ,
                //                      [NQTY] ,
                //                      [LIFE] ,
                //                      [NOTE] ,
                //                      [ISBOM]
                //                    )
                //            VALUES  ( '{0}' , -- GUID - uniqueidentifier
                //                      '{1}' , -- PGUID - uniqueidentifier
                //                      {2} , -- SNO - int
                //                      '{3}' , -- FGUID - uniqueidentifier
                //                      '{4}' , -- MGUID - uniqueidentifier
                //                      {5} , -- NQTY - decimal
                //                      {6} , -- LIFE - int
                //                      N'' , -- NOTE - nvarchar(200)
                //                      {7}  -- ISBOM - bit
                //                    )
                //", newguid, BSPRODPLAN_PGUID, dr_excel["序号"], drs_BSPRODPLANSS_PRODUCT[0]["GUID"], drs_BSPRODSTD_PRODUCT_B_Tiny[0]["GUID"], dr_excel["数量"], dr_excel["寿命(pcs)"], GetISINTERVAL(dr_excel["是否组合件"].ToString()));
                //                //rbSql.Text += temp1 + Environment.NewLine;
                //                sqlLs.Add(temp1);
                //            }

                //            dgError4.DataSource = dt_error4;
                //            dgRepet4.DataSource = dt_repet4;
                //            dgRepet4_excel.DataSource = dt_repet4_excel;
                //            if (dt_error4.Rows.Count > 0 || dt_repet4.Rows.Count > 0 || dt_repet4_excel.Rows.Count > 0)
                //            {
                //                Main.SetErrorCell(dgError4, col_error4);
                //                rbSql.Text = "";
                //                return;
                //            }
                //        #endregion

                //            #region MRO资源寿命
                //            List<int[]> col_error5 = new List<int[]>();
                //            DataTable dt_repet5 = _BSPRODPLANSS_excel.Clone();
                //            DataTable dt_error5 = _BSPRODPLANSS_excel.Clone();

                //            for (int i = 0; i < _BSPRODPLANSS_excel.Rows.Count; i++)
                //            {
                //                bool isError = false;
                //                bool isRepet = false;

                //                DataRow dr_excel = _BSPRODPLANSS_excel.Rows[i];

                //                DataRow[] drs_BSPRODPLANSS_PRODUCT =
                //    _BSPRODPLAN_excel.Select(string.Format("产品编号 = '{0}' AND 版本号 = '{1}' AND 物料编号 = '{2}' AND 工序 = '{3}' ",
                //        WGHelper.ReturnString(dr_excel["产品编号"].ToString()), dr_excel["版本号"].ToString(), dr_excel["物料母件编号"].ToString(), dr_excel["工序"].ToString()));


                //                if (drs_BSPRODPLANSS_PRODUCT.Length == 0)
                //                {
                //                    col_error5.Add(new int[] { dt_error5.Rows.Count, 0 });
                //                    col_error5.Add(new int[] { dt_error5.Rows.Count, 1 });
                //                    col_error5.Add(new int[] { dt_error5.Rows.Count, 2 });
                //                    col_error5.Add(new int[] { dt_error5.Rows.Count, 3 });
                //                    isError = true;
                //                }

                //                DataRow[] drs_BSPRODUCT_DB_B_BSPRODPLANSS = _BSPRODUCT_DB_B.Select("CODE = '" + dr_excel["物料子件编号"] + "'");

                //                if (drs_BSPRODUCT_DB_B_BSPRODPLANSS.Length == 0)
                //                {
                //                    col_error5.Add(new int[] { dt_error5.Rows.Count, 5 });
                //                    isError = true;
                //                }

                //                int SNO = 0;
                //                if (string.IsNullOrWhiteSpace(dr_excel["序号"].ToString())
                //                    || !int.TryParse(dr_excel["序号"].ToString(), out SNO))
                //                {
                //                    //空
                //                    col_error5.Add(new int[] { dt_error5.Rows.Count, 4 });
                //                    isError = true;
                //                }

                //                DataRow[] drs_BSPRODPLANSS_BSPRODUCT = _BSPRODUCT_DB_B.Select(" CODE = '" + WGHelper.ReturnString(dr_excel["物料子件编号"].ToString()) + "' ");

                //                if (drs_BSPRODPLANSS_BSPRODUCT.Length == 0)
                //                {
                //                    //空
                //                    col_error5.Add(new int[] { dt_error5.Rows.Count, 5 });
                //                    isError = true;
                //                }

                //                SNO = 0;
                //                if (!string.IsNullOrWhiteSpace(dr_excel["寿命(pcs)"].ToString()))
                //                {
                //                    if (!int.TryParse(dr_excel["寿命(pcs)"].ToString(), out SNO))
                //                    {
                //                        //空
                //                        col_error5.Add(new int[] {dt_error5.Rows.Count, 6});
                //                        isError = true;
                //                    }
                //                }
                //                else
                //                {
                //                    dr_excel["寿命(pcs)"] = "null";
                //                }

                //                if (isError || isRepet)
                //                {
                //                    if (isError)
                //                    {
                //                        dt_error5.Rows.Add(dr_excel.ItemArray);
                //                    }
                //                    continue;
                //                }

                //                string temp = string.Format(@" 
                //            INSERT INTO [dbo].[BSPRODPLANSS]
                //        ( [GUID] ,
                //          [PGUID] ,
                //          [SNO] ,
                //          [FGUID] ,
                //          [LIFE] ,
                //          [NOTE]
                //        )
                //VALUES  ( NEWID() , -- GUID - uniqueidentifier
                //          '{0}' , -- PGUID - uniqueidentifier
                //          {1} , -- SNO - int
                //          '{2}' , -- FGUID - uniqueidentifier
                //          {3} , -- LIFE - int
                //          N''  -- NOTE - nvarchar(200)
                //        );", drs_BSPRODPLANSS_PRODUCT[0]["GUID"], dr_excel["序号"], drs_BSPRODPLANSS_BSPRODUCT[0]["GUID"], dr_excel["寿命(pcs)"]);
                //               // rbSql.Text += temp + Environment.NewLine;
                //                sqlLs.Add(temp);
                //            }
                //            dgError5.DataSource = dt_error5;
                //            dgRepet5_excel.DataSource = dt_repet5;
                //            if (dt_error5.Rows.Count > 0 || dt_repet5.Rows.Count > 0 )
                //            {
                //                Main.SetErrorCell(dgError5, col_error5);
                //                rbSql.Text = "";
                //                return;
                //            }
                //        #endregion

                //            #region 检验标准
                //            List<int[]> col_error6 = new List<int[]>();
                //            DataTable dt_repet6 = _BSPRODTEST_excel.Clone();
                //            DataTable dt_error6 = _BSPRODTEST_excel.Clone();
                //            DataTable dt_repet6_excel = _BSPRODTEST_excel.Clone();
                //            DataTable dt_BSPRODTEST_ADD = new DataTable();
                //            dt_BSPRODTEST_ADD.Columns.Add("产品编号");
                //            dt_BSPRODTEST_ADD.Columns.Add("版本号");
                //            dt_BSPRODTEST_ADD.Columns.Add("GUID");

                //            for (int i = 0; i < _BSPRODTEST_excel.Rows.Count; i++)
                //            {
                //                bool isError = false;
                //                bool isRepet = false;
                //                bool isRepet_excel = false;

                //                DataRow dr_excel = _BSPRODTEST_excel.Rows[i];

                //                DataRow[] drs_BSPRODUCT_BSPRODTEST =
                //                    _BSPRODUCT_DB.Select("CODE = '" + dr_excel["产品编号"] + "' AND VER = '" + dr_excel["版本号"].ToString() +
                //                                         "' ");

                //                if (drs_BSPRODUCT_BSPRODTEST.Length == 0)
                //                {
                //                    col_error6.Add(new int[] {dt_error6.Rows.Count, 0});
                //                    col_error6.Add(new int[] {dt_error6.Rows.Count, 1});
                //                    isError = true;
                //                }

                //                DataRow[] drs_repet = _BSPRODTEST_DB.Select("CODE = '" + dr_excel["产品编号"] + "' AND VER = '" + dr_excel["版本号"].ToString() +"' ");
                //                if (drs_repet.Length > 0)
                //                {
                //                    isRepet = true;
                //                }

                //                //DataRow[] drs_repet_excel = _BSPRODTEST_excel.Select("产品编号 = '" + dr_excel["产品编号"] + "' AND 版本号 = '" + dr_excel["版本号"].ToString() + "' AND 工序 = '"+dr_excel["工序"].ToString()+"' ");
                //                //if (drs_repet_excel.Length > 1)
                //                //{
                //                //    isRepet_excel = true;
                //                //}


                //                DataRow[] drs_BSPRODTESTS_PRODUCT;
                //                if (isnew)
                //                {
                //                    drs_BSPRODTESTS_PRODUCT = _BSPRODSTDS_DB.Select("产品编号 = '" + dr_excel["产品编号"].ToString() + "' AND 工序 = '" + dr_excel["工序"].ToString() + "'");
                //                }
                //                else
                //                {
                //                    drs_BSPRODTESTS_PRODUCT =
                //                                            _BSPRODSTD_excel.Select(string.Format("产品编号 = '{0}' AND 版本号 = '{1}' AND 工序 = '{2}' ",
                //                                                WGHelper.ReturnString(dr_excel["产品编号"].ToString()), dr_excel["版本号"].ToString(),
                //                                                dr_excel["工序"].ToString()));
                //                }

                //                if (drs_BSPRODTESTS_PRODUCT.Length == 0)
                //                {
                //                    col_error6.Add(new int[] {dt_error6.Rows.Count, 0});
                //                    col_error6.Add(new int[] {dt_error6.Rows.Count, 1});
                //                    col_error6.Add(new int[] {dt_error6.Rows.Count, 3});
                //                    isError = true;
                //                }

                //                int SNO = 0;
                //                if (string.IsNullOrWhiteSpace(dr_excel["序号"].ToString())
                //                    || !int.TryParse(dr_excel["序号"].ToString(), out SNO))
                //                {
                //                    //空
                //                    col_error6.Add(new int[] {dt_error6.Rows.Count, 2});
                //                    isError = true;
                //                }

                //                if (!CTYPEs.ContainsKey(dr_excel["测量类型"].ToString()))
                //                {
                //                    col_error6.Add(new int[] {dt_error6.Rows.Count, 5});
                //                    isError = true;
                //                }

                //                if (string.IsNullOrWhiteSpace(dr_excel["工序"].ToString()))
                //                {
                //                    //空
                //                    col_error6.Add(new int[] {dt_error6.Rows.Count, 3});
                //                    isError = true;
                //                }
                //                if (string.IsNullOrWhiteSpace(dr_excel["检验项"].ToString()))
                //                {
                //                    //空
                //                    col_error6.Add(new int[] {dt_error6.Rows.Count, 4});
                //                    isError = true;
                //                }
                //                if (string.IsNullOrWhiteSpace(dr_excel["测量类型"].ToString()))
                //                {
                //                    //空
                //                    col_error6.Add(new int[] {dt_error6.Rows.Count, 5});
                //                    isError = true;
                //                }
                //                if (string.IsNullOrWhiteSpace(dr_excel["测量方法"].ToString()))
                //                {
                //                    //空
                //                    col_error6.Add(new int[] {dt_error6.Rows.Count, 6});
                //                    isError = true;
                //                }
                //                if (string.IsNullOrWhiteSpace(dr_excel["标准"].ToString()) && dr_excel["测量类型"].ToString() == "外观")
                //                {
                //                    //空
                //                    col_error6.Add(new int[] {dt_error6.Rows.Count, 7});
                //                    isError = true;
                //                }

                //                decimal nums = 0;
                //                decimal numsgc = 0;
                //                decimal numxgc = 0;
                //                decimal numsx = 0;
                //                decimal numxx = 0;

                //                if (decimal.TryParse(dr_excel["标准值"].ToString(), out nums) && decimal.TryParse(dr_excel["上公差"].ToString(), out numsgc))
                //                {
                //                    numsx = nums + numsgc;
                //                    dr_excel["标准上限"] = numsx;
                //                }

                //                if (decimal.TryParse(dr_excel["标准值"].ToString(), out nums) && decimal.TryParse(dr_excel["下公差"].ToString(), out numxgc))
                //                {
                //                    numxx = nums - numxgc;
                //                    dr_excel["标准下限"] = numxx;
                //                }

                //                if (dr_excel["测量类型"].ToString() == "尺寸" )
                //                {
                //                    if (string.IsNullOrWhiteSpace(dr_excel["标准值"].ToString()))
                //                    {
                //                        col_error6.Add(new int[] {dt_error6.Rows.Count, 8});
                //                        isError = true;
                //                    }

                //                    if (!ISINTERVALs.ContainsKey(dr_excel["正无穷大"].ToString()))
                //                    {
                //                        col_error6.Add(new int[] { dt_error6.Rows.Count, 9 });
                //                        isError = true;
                //                    }

                //                    if (!ISINTERVALs.ContainsKey(dr_excel["负无穷大"].ToString()))
                //                    {
                //                        col_error6.Add(new int[] { dt_error6.Rows.Count, 11 });
                //                        isError = true;
                //                    }

                //                    if (!ISINTERVALs.ContainsKey(dr_excel["是否首检"].ToString()))
                //                    {
                //                        col_error6.Add(new int[] { dt_error6.Rows.Count, 15 });
                //                        isError = true;
                //                    }

                //                    if (!ISINTERVALs.ContainsKey(dr_excel["是否末检"].ToString()))
                //                    {
                //                        col_error6.Add(new int[] { dt_error6.Rows.Count, 17 });
                //                        isError = true;
                //                    }

                //                    if (!ISINTERVALs.ContainsKey(dr_excel["是否巡检"].ToString()))
                //                    {
                //                        col_error6.Add(new int[] { dt_error6.Rows.Count, 19 });
                //                        isError = true;
                //                    }

                //                    if (!ISINTERVALs.ContainsKey(dr_excel["是否送检"].ToString()))
                //                    {
                //                        col_error6.Add(new int[] { dt_error6.Rows.Count, 22 });
                //                        isError = true;
                //                    }

                //                    if (!ISINTERVALs.ContainsKey(dr_excel["是否自检"].ToString()))
                //                    {
                //                        col_error6.Add(new int[] { dt_error6.Rows.Count, 23 });
                //                        isError = true;
                //                    }
                //                }
                //                DataRow[] drs_BSPRODTEST_ADDNEW = dt_BSPRODTEST_ADD.Select(string.Format("产品编号 = '{0}' AND 版本号 = '{1}' ", dr_excel["产品编号"].ToString(), dr_excel["版本号"]));

                //                if (isError || isRepet || isRepet_excel)
                //                {
                //                    if (isError)
                //                    {
                //                        dt_error6.Rows.Add(dr_excel.ItemArray);
                //                    }
                //                    if(isRepet)
                //                    {
                //                        dt_repet6.Rows.Add(dr_excel.ItemArray);
                //                    }
                //                    if (isRepet_excel)
                //                    {
                //                        dt_repet6_excel.Rows.Add(dr_excel.ItemArray);
                //                    }
                //                    continue;
                //                }

                //                string BSPRODTEST_PGUID;
                //                if (drs_BSPRODTEST_ADDNEW.Length == 0)
                //                {
                //                    Guid kidGUID = Guid.NewGuid();
                //                    BSPRODTEST_PGUID = kidGUID.ToString();
                //                    DataRow NEWROW = dt_BSPRODTEST_ADD.NewRow();
                //                    NEWROW["产品编号"] = dr_excel["产品编号"];
                //                    NEWROW["版本号"] = dr_excel["版本号"];
                //                    NEWROW["GUID"] = kidGUID;
                //                    dt_BSPRODTEST_ADD.Rows.Add(NEWROW);

                //                    string temp = string.Format(@" 
                //        INSERT INTO [dbo].[BSPRODTEST]
                //          ( [GUID] ,
                //            [PGUID] ,
                //            [VER] ,
                //            [NOTE] ,
                //            [ST] ,
                //            [ND] 
                //          )
                //  VALUES  ( '{0}' , -- GUID - uniqueidentifier
                //            '{1}' , -- PGUID - uniqueidentifier
                //            '{2}' , -- VER - nvarchar(10)
                //            N'' , -- NOTE - nvarchar(200)
                //            1 , -- ST - int
                //            GETDATE() 
                //          );", kidGUID, drs_BSPRODUCT_BSPRODTEST[0]["GUID"].ToString(), dr_excel["版本号"]);
                //                   // rbSql.Text += temp + Environment.NewLine;
                //                    sqlLs.Add(temp);
                //                }
                //                else
                //                {
                //                    BSPRODTEST_PGUID = drs_BSPRODTEST_ADDNEW[0]["GUID"].ToString();
                //                }
                //                string HK_STR = GetHKstr(dr_excel["是否首检"].ToString(), dr_excel["首检数量"].ToString(), dr_excel["是否末检"].ToString(), dr_excel["末检数量"].ToString(), dr_excel["是否巡检"].ToString(), dr_excel["巡检数量"].ToString(), dr_excel["巡检周期"].ToString(), dr_excel["是否送检"].ToString(),dr_excel["是否自检"].ToString());
                //                string temp1 = string.Format(@"
                //                INSERT INTO [dbo].[BSPRODTESTS]
                //                        ( [GUID] ,
                //                          [PGUID] ,
                //                          [SNO] ,
                //                          [FGUID] ,
                //                          [CTYPE] ,
                //                          [CITEM] ,
                //                          [CWAY] ,
                //                          [SVALUE] ,
                //                          [STDVALUE] ,
                //                          [ISMAX] ,
                //                          [MAXOFFSET] ,
                //                          [ISMIN] ,
                //                          [MINOFFSET] ,
                //                          [MAXDV] ,
                //                          [MINDV] ,
                //                          [HZ] ,
                //                          [NOTE] ,
                //                          [ISFC] ,
                //                          [FNUM] ,
                //                          [ISLC] ,
                //                          [LNUM] ,
                //                          [ISPC] ,
                //                          [PCYCLE] ,
                //                          [PNUM] ,
                //                          [ISSC] ,
                //                          [ISOC]
                //                        )
                //                VALUES  ( NEWID() , -- GUID - uniqueidentifier
                //                          '{0}' , -- PGUID - uniqueidentifier
                //                          {1} , -- SNO - int
                //                          '{2}' , -- FGUID - uniqueidentifier
                //                          '{3}' , -- CTYPE - int
                //                          '{4}' , -- CITEM - nvarchar(50)
                //                          '{5}' , -- CWAY - nvarchar(50)
                //                          '{6}' , -- SVALUE - nvarchar(50)
                //                          {7} , -- STDVALUE - decimal
                //                          {8} , -- ISMAX - bit
                //                          {9} , -- MAXOFFSET - decimal
                //                          {10} , -- ISMIN - bit
                //                          {11} , -- MINOFFSET - decimal
                //                          {12} , -- MAXDV - decimal
                //                          {13} , -- MINDV - decimal
                //                          '{14}' , -- HZ - nvarchar(100)
                //                          N'' , -- NOTE - nvarchar(200)
                //                          {15} , -- ISFC - bit
                //                          {16} , -- FNUM - int
                //                          {17} , -- ISLC - bit
                //                          {18} , -- LNUM - int
                //                          {19} , -- ISPC - bit
                //                          {20} , -- PCYCLE - int
                //                          {21} , -- PNUM - int
                //                          {22} , -- ISSC - bit
                //                          {23}
                //                        )", BSPRODTEST_PGUID, dr_excel["序号"], drs_BSPRODTESTS_PRODUCT[0]["GUID"], GetCTYPE(dr_excel["测量类型"].ToString()), dr_excel["检验项"], dr_excel["测量方法"], dr_excel["标准"], GetSTR(dr_excel["标准值"].ToString()), GetISINTERVAL(dr_excel["正无穷大"].ToString()), GetSTR(dr_excel["上公差"].ToString()), GetISINTERVAL(dr_excel["负无穷大"].ToString()), GetSTR(dr_excel["下公差"].ToString()), GetSTR(dr_excel["标准上限"].ToString()), GetSTR(dr_excel["标准下限"].ToString()), HK_STR, GetISINTERVAL(dr_excel["是否首检"].ToString()), GetSTR(dr_excel["首检数量"].ToString()), GetISINTERVAL(dr_excel["是否末检"].ToString()), GetSTR(dr_excel["末检数量"].ToString()), GetISINTERVAL(dr_excel["是否巡检"].ToString()), GetSTR(dr_excel["巡检周期"].ToString()), GetSTR(dr_excel["巡检数量"].ToString()), GetISINTERVAL(dr_excel["是否送检"].ToString()), GetISINTERVAL(dr_excel["是否自检"].ToString()));
                //                rbSql.Text += temp1 + Environment.NewLine;
                //                sqlLs.Add(temp1);
                //            }

                //            dgError6.DataSource = dt_error6;
                //            dgRepet6.DataSource = dt_repet6;
                //            dgRepet6_excel.DataSource = dt_repet6_excel;
                //            if (dt_error6.Rows.Count > 0 || dt_repet6.Rows.Count > 0 || dt_repet6_excel.Rows.Count > 0)
                //            {
                //                Main.SetErrorCell(dgError6, col_error6);
                //                rbSql.Text = "";
                //                return;
                //            }

                //        #endregion

                //            #region 工艺参数

                //            List<int[]> col_error7 = new List<int[]>();
                //            DataTable dt_repet7 = _BSPRODPARAM_excel.Clone();
                //            DataTable dt_error7 = _BSPRODPARAM_excel.Clone();
                //            DataTable dt_repet7_excel = _BSPRODPARAM_excel.Clone();
                //            DataTable dt_BSPRODPARAM_ADD = new DataTable();
                //            dt_BSPRODPARAM_ADD.Columns.Add("产品编号");
                //            dt_BSPRODPARAM_ADD.Columns.Add("版本号");
                //            dt_BSPRODPARAM_ADD.Columns.Add("GUID");

                //            for (int i = 0; i < _BSPRODPARAM_excel.Rows.Count; i++)
                //            {
                //                bool isError = false;
                //                bool isRepet = false;
                //                bool isRepet_excel = false;

                //                DataRow dr_excel = _BSPRODPARAM_excel.Rows[i];

                //                DataRow[] drs_BSPRODUCT_BSPRODTEST =
                //                    _BSPRODUCT_DB.Select("CODE = '" + dr_excel["产品编号"] + "' AND VER = '" + dr_excel["版本号"].ToString() +
                //                                         "' ");
                //                DataRow[] drs_repet = _BSPRODPARAM_DB.Select("CODE = '" + dr_excel["产品编号"] + "' AND VER = '" + dr_excel["版本号"].ToString() +
                //                                         "' ");
                //                if (drs_repet.Length > 0)
                //                {
                //                    isRepet = true;
                //                }

                //                //DataRow[] drs_repet_excel = _BSPRODPARAM_excel.Select("产品编号 = '" + dr_excel["产品编号"] + "' AND 版本号 = '" + dr_excel["版本号"].ToString() +
                //                //                       "' ");
                //                //if (drs_repet_excel.Length > 1)
                //                //{
                //                //    isRepet_excel = true;
                //                //}

                //                if (drs_BSPRODUCT_BSPRODTEST.Length == 0)
                //                {
                //                    col_error7.Add(new int[] {dt_error7.Rows.Count, 0});
                //                    col_error7.Add(new int[] {dt_error7.Rows.Count, 1});
                //                    isError = true;
                //                }

                //                int SNO = 0;
                //                if (string.IsNullOrWhiteSpace(dr_excel["序号"].ToString())
                //                    || !int.TryParse(dr_excel["序号"].ToString(), out SNO))
                //                {
                //                    //空
                //                    col_error7.Add(new int[] {dt_error7.Rows.Count, 2});
                //                    isError = true;
                //                }

                //                if (string.IsNullOrWhiteSpace(dr_excel["工序"].ToString()))
                //                {
                //                    //空
                //                    col_error7.Add(new int[] {dt_error7.Rows.Count, 2});
                //                    isError = true;
                //                }


                //                DataRow[] drs_BSPRODTESTS_PRODUCT;
                //                if (isnew)
                //                {
                //                    drs_BSPRODTESTS_PRODUCT = _BSPRODSTDS_DB.Select("产品编号 = '" + dr_excel["产品编号"].ToString() + "' AND 工序 = '" + dr_excel["工序"].ToString() + "'");
                //                }
                //                else
                //                {
                //                                            drs_BSPRODTESTS_PRODUCT =
                //                    _BSPRODSTD_excel.Select(string.Format("产品编号 = '{0}' AND 版本号 = '{1}' AND 工序 = '{2}' ",
                //                        WGHelper.ReturnString(dr_excel["产品编号"].ToString()), dr_excel["版本号"].ToString(),
                //                        dr_excel["工序"].ToString()));
                //                }

                //                if (drs_BSPRODTESTS_PRODUCT.Length == 0)
                //                {
                //                    col_error7.Add(new int[] {dt_error7.Rows.Count, 0});
                //                    col_error7.Add(new int[] {dt_error7.Rows.Count, 1});
                //                    col_error7.Add(new int[] {dt_error7.Rows.Count, 3});
                //                    isError = true;
                //                }

                //                if (string.IsNullOrWhiteSpace(dr_excel["标准"].ToString()))
                //                {
                //                    //空
                //                    col_error7.Add(new int[] {dt_error7.Rows.Count, 6});
                //                    isError = true;
                //                }

                //                if (dr_excel["类型"].ToString() == "数值")
                //                {
                //                    if (!ISINTERVALs.ContainsKey(dr_excel["是否区间"].ToString()))
                //                    {
                //                        col_error7.Add(new int[] {dt_error7.Rows.Count, 7});
                //                        isError = true;
                //                    }

                //                    //string stand_str = dr_excel["符号1"].ToString() + dr_excel["最小值"].ToString() +
                //                    //                   dr_excel["单位"].ToString() + " " + dr_excel["符号2"].ToString() +
                //                    //                   dr_excel["最大值"].ToString() + dr_excel["单位"].ToString();

                //                    //if (!dr_excel["标准"].ToString().Contains(stand_str))
                //                    //{
                //                    //    col_error7.Add(new int[] {dt_error7.Rows.Count, 6});
                //                    //    col_error7.Add(new int[] {dt_error7.Rows.Count, 8});
                //                    //    col_error7.Add(new int[] {dt_error7.Rows.Count, 9});
                //                    //    col_error7.Add(new int[] {dt_error7.Rows.Count, 10});
                //                    //    col_error7.Add(new int[] {dt_error7.Rows.Count, 11});
                //                    //    col_error7.Add(new int[] {dt_error7.Rows.Count, 12});
                //                    //    isError = true;
                //                    //}    
                //                }

                //                if (isError || isRepet || isRepet_excel)
                //                {
                //                    if (isError)
                //                    {
                //                        dt_error7.Rows.Add(dr_excel.ItemArray);
                //                    }
                //                    if (isRepet)
                //                    {
                //                        dt_repet7.Rows.Add(dr_excel.ItemArray);
                //                    }
                //                    if (isRepet_excel)
                //                    {
                //                        dt_repet7_excel.Rows.Add(dr_excel.ItemArray);
                //                    }
                //                    continue;
                //                }




                //                DataRow[] drs_BSPRODPARAM_ADDNEW =
                //                      dt_BSPRODPARAM_ADD.Select(string.Format("产品编号 = '{0}' AND 版本号 = '{1}' ",
                //                          dr_excel["产品编号"].ToString(), dr_excel["版本号"]));

                //                string BSPRODPARAM_PGUID;
                //                if (drs_BSPRODPARAM_ADDNEW.Length == 0)
                //                {
                //                    Guid kidGUID = Guid.NewGuid();
                //                    BSPRODPARAM_PGUID = kidGUID.ToString();
                //                    DataRow NEWROW = dt_BSPRODBOM_ADD.NewRow();
                //                    NEWROW["产品编号"] = dr_excel["产品编号"];
                //                    NEWROW["版本号"] = dr_excel["版本号"];
                //                    NEWROW["GUID"] = kidGUID;
                //                    dt_BSPRODPARAM_ADD.Rows.Add(NEWROW.ItemArray);

                //                    string temp = string.Format(@" 
                //        INSERT INTO [dbo].[BSPRODPARAM]
                //        ( [GUID] ,
                //          [PGUID] ,
                //          [VER] ,
                //          [NOTE] ,
                //          [ST] ,
                //          [ND] 
                //        )
                //VALUES  ( '{0}' , -- GUID - uniqueidentifier
                //          '{1}' , -- PGUID - uniqueidentifier
                //          '{2}' , -- VER - nvarchar(10)
                //          N'' , -- NOTE - nvarchar(200)
                //          1 , -- ST - int
                //          GETDATE() 
                //        ) ;", kidGUID, drs_BSPRODUCT_BSPRODTEST[0]["GUID"].ToString(), dr_excel["版本号"]);
                //                   // rbSql.Text += temp + Environment.NewLine;
                //                    sqlLs.Add(temp);
                //                }
                //                else
                //                {
                //                    BSPRODPARAM_PGUID = drs_BSPRODPARAM_ADDNEW[0]["GUID"].ToString();
                //                }
                //                string temp1 = string.Format(@"
                //INSERT INTO [dbo].[BSPRODPARAMS]
                //        ( [GUID] ,
                //          [PGUID] ,
                //          [SNO] ,
                //          [FGUID] ,
                //          [CTYPE] ,
                //          [CITEM] ,
                //          [SVALUE] ,
                //          [ISINTERVAL] ,
                //          [MAXDV] ,
                //          [MINDV] ,
                //          [SIGN1] ,
                //          [SIGN2] ,
                //          [UNIT] ,
                //          [NOTE]
                //        )
                //VALUES  ( NEWID() , -- GUID - uniqueidentifier
                //          '{0}' , -- PGUID - uniqueidentifier
                //          {1} , -- SNO - int
                //          '{2}' , -- FGUID - uniqueidentifier
                //          '{3}' , -- CTYPE - nvarchar(10)
                //          '{4}' , -- CITEM - nvarchar(50)
                //          '{5}' , -- SVALUE - nvarchar(50)
                //          {6} , -- ISINTERVAL - bit
                //          {7} , -- MAXDV - decimal
                //          {8} , -- MINDV - decimal
                //          {9} , -- SIGN1 - nvarchar(10)
                //          {10} , -- SIGN2 - nvarchar(10)
                //          '{11}' , -- UNIT - nvarchar(10)
                //          N''  -- NOTE - nvarchar(200)
                //        )", BSPRODPARAM_PGUID, dr_excel["序号"].ToString(), drs_BSPRODTESTS_PRODUCT[0]["GUID"].ToString(),
                //                    dr_excel["类型"].ToString(), dr_excel["参数"].ToString(), dr_excel["标准"].ToString(),
                //                    GetISINTERVAL(dr_excel["是否区间"].ToString()), GetSTR(dr_excel["最大值"].ToString()),
                //                    GetSTR(dr_excel["最小值"].ToString()), GetSTR_D(dr_excel["符号1"].ToString()),
                //                    GetSTR_D(dr_excel["符号2"].ToString()), dr_excel["单位"].ToString());
                //                rbSql.Text += temp1 + Environment.NewLine;
                //                sqlLs.Add(temp1);
                //            }
                //            dgError7.DataSource = dt_error7;
                //            dgRepet7.DataSource = dt_repet7;
                //            dgRepet7_excel.DataSource = dt_repet7_excel;
                //            if (dt_error7.Rows.Count > 0 || dt_repet7.Rows.Count > 0 || dt_repet7_excel.Rows.Count > 0)
                //            {
                //                Main.SetErrorCell(dgError7, col_error7);
                //                rbSql.Text = "";
                //                return;
                //            }

                //        #endregion
#endregion
                isCheck = true;

                StringBuilder last = new StringBuilder();
                foreach (string sql1 in sqlLs)
                {
                    last.Append(sql1 + Environment.NewLine);
                }

                rbSql.Text = last.ToString();
            }
            //catch (Exception eee)
            //{
            //    throw eee;
            //}
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

        private string GetCTYPE(string str)
        {
            if (str == "尺寸")
            {
                return "1";
            }
            else if (str == "外观")
            {
                return "2";
            }
            return null;
        }

        private string GetSTR(string str)
        {
            if (str == "")
                return "null";
            else
            {
                return str;
            }
        }

        private string GetSTR_D(string str)
        {
            if (str == "")
                return "null";
            else
            {
                return "'"+ str+"'";
            }
        }

        //private string GetHKstr(string 是否首检, string 首检数量, string 是否末检, string 末检数量, string 是否巡检, string 巡检数量, string 巡检周期, string 是否送检,string 是否自检)
        //{
        //    string return_string = "";
        //    if (是否首检 == "是")
        //        return_string += "首件" + 首检数量 + "PCS,";
        //    if (是否巡检 == "是")
        //        return_string += "过程每" + 巡检周期 + "PCS检验" + 巡检数量 + "PCS,";
        //    if (是否末检 == "是")
        //        return_string += "末件" + 末检数量 + "PCS,";
        //    if (是否送检 == "是")
        //        return_string += "送检,";
        //    if (是否自检 == "是")
        //        return_string += "自检,";
        //    if (return_string != "")
        //        return_string = return_string.Substring(0, return_string.Length - 1);

        //    return return_string;
        //}

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

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            if (tabControl1.SelectedIndex == 0)
            {
                dt = dgError1.DataSource as DataTable;
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                dt = dgRepet1_excel.DataSource as DataTable;
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                dt = dgRepet1.DataSource as DataTable;
            }
            else if (tabControl1.SelectedIndex == 3)
            {
                dt = dgError2.DataSource as DataTable;
            }
            else if (tabControl1.SelectedIndex == 4)
            {
                dt = dgRepet2_excel.DataSource as DataTable;
            }
            else if (tabControl1.SelectedIndex == 5)
            {
                dt = dgRepet2.DataSource as DataTable;
            }
            else if (tabControl1.SelectedIndex == 6)
            {
                dt = dgError3.DataSource as DataTable;
            }
            else if (tabControl1.SelectedIndex == 7)
            {
                dt = dgRepet3_excel.DataSource as DataTable;
            }
            else if (tabControl1.SelectedIndex == 8)
            {
                dt = dgRepet3.DataSource as DataTable;
            }
            else if (tabControl1.SelectedIndex == 9)
            {
                dt = dgError4.DataSource as DataTable;
            }
            else if (tabControl1.SelectedIndex == 10)
            {
                dt = dgRepet4_excel.DataSource as DataTable;
            }
            else if (tabControl1.SelectedIndex == 11)
            {
                dt = dgRepet4.DataSource as DataTable;
            }
            else if (tabControl1.SelectedIndex == 12)
            {
                dt = dgError5.DataSource as DataTable;
            }
            else if (tabControl1.SelectedIndex == 13)
            {
                dt = dgRepet5_excel.DataSource as DataTable;
            }
            else if (tabControl1.SelectedIndex == 14)
            {
                dt = dgRepet5.DataSource as DataTable;
            }
            else if (tabControl1.SelectedIndex == 15)
            {
                dt = dgError6.DataSource as DataTable;
            }
            else if (tabControl1.SelectedIndex == 16)
            {
                dt = dgRepet6_excel.DataSource as DataTable;
            }
            else if (tabControl1.SelectedIndex == 17)
            {
                dt = dgRepet6.DataSource as DataTable;
            }

            Dictionary<string, string> _dicDataFieldCaption = new Dictionary<string, string>();
            ExportToExcel("错误数据.xlsx", "错误数据", _dicDataFieldCaption, dt, null);
        }

        public static void ExportToExcel(string fileName, string title, Dictionary<string, string> ColDic,
       DataTable dtSource, IList<Image> jpgImagesList = null, Dictionary<bool, string> boolKeyValue = null)
        {
            DataSet ds = new DataSet();
            DataTable dt = dtSource.Copy();
            dt.TableName = title;
            DataTable dtTemp = FormatDataTable(dt, ColDic);

            ds.Tables.Add(dtTemp);
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.DefaultExt = "xlsx";
            dlg.Title = "导出Excel";
            dlg.FileName = fileName;
            dlg.Filter = "Excel|*.xlsx;*.xls";
            if (dlg.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(dlg.FileName) && !string.IsNullOrWhiteSpace(Path.GetExtension(dlg.FileName)))
            {
                string fileExt = Path.GetExtension(dlg.FileName).ToLower();
                if (!(fileExt == ".xls" || fileExt == ".xlsx"))
                {
                    MessageBox.Show("指定文件格式错误!导出格式Excel格式应为xlsx或xls!");
                    return;
                }
                MemoryStream stream = Export(ds, jpgImagesList, boolKeyValue, fileExt);
                try
                {
                    byte[] bytes = stream.ToArray();
                    FileStream fs = new FileStream(dlg.FileName, FileMode.Create);
                    BinaryWriter bw = new BinaryWriter(fs);
                    bw.Write(bytes);
                    bw.Close();
                    fs.Close();
                }
                catch (Exception ex)
                {
                    WGMessage.ShowAsterisk("导出文件时出错,文件可能正被打开！\n" + ex.Message);
                    return;
                }
                MessageBox.Show("请您注意，导出完成！ ", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public static MemoryStream Export(DataSet dataSet, IList<Image> jpgImagesList, Dictionary<bool, string> boolKeyValue, string fileExt)
        {
            DataSet ds = dataSet.Copy();
            IWorkbook workbook;
            //XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
            switch (fileExt)
            {
                case ".xlsx":
                    workbook = new XSSFWorkbook();
                    break;
                case ".xls":
                    workbook = new HSSFWorkbook();
                    break;
                default:
                    workbook = null;
                    break;
            }
            if (workbook == null) { return null; }

            ISheet sheet = workbook.CreateSheet();
            ICellStyle dateStyle = workbook.CreateCellStyle();
            IDataFormat format = workbook.CreateDataFormat();
            dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");
            int rowIndex = 0;
            int tableIndex = -1;

            for (int i = 0; i < 32; i++)
            {
                sheet.SetColumnWidth(i, 20 * 256);
            }

            switch (fileExt)
            {
                case ".xlsx":
                    {
                        if (jpgImagesList != null)
                        {
                            XSSFDrawing patriarch = sheet.CreateDrawingPatriarch() as XSSFDrawing;
                            foreach (Image image in jpgImagesList)
                            {
                                byte[] bytes = new byte[0];
                                int pictureIdx = workbook.AddPicture(bytes, PictureType.JPEG);
                                XSSFClientAnchor anchor = new XSSFClientAnchor { Row1 = rowIndex };
                                XSSFPicture pict = patriarch.CreatePicture(anchor, pictureIdx) as XSSFPicture;
                                if (pict != null) pict.Resize();
                                rowIndex = anchor.Row2 + 2;
                            }
                        }
                    }
                    break;
                case ".xls":
                    {
                        if (jpgImagesList != null)
                        {
                            HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;
                            foreach (Image image in jpgImagesList)
                            {
                                byte[] bytes = new byte[0];
                                int pictureIdx = workbook.AddPicture(bytes, PictureType.JPEG);
                                HSSFClientAnchor anchor = new HSSFClientAnchor { Row1 = rowIndex };
                                HSSFPicture pict = patriarch.CreatePicture(anchor, pictureIdx) as HSSFPicture;
                                if (pict != null) pict.Resize();
                                rowIndex = anchor.Row2 + 2;
                            }
                        }
                    }
                    break;
            }

            foreach (DataTable dtSource in ds.Tables)
            {
                foreach (DataRow row in dtSource.Rows)
                {
                    #region 新建表，填充表头，填充列头，样式

                    if (rowIndex >= 60000)
                    {
                        sheet = workbook.CreateSheet();
                        rowIndex = 0;
                    }

                    if (rowIndex == 0 || tableIndex != ds.Tables.IndexOf(dtSource))
                    {
                        tableIndex = ds.Tables.IndexOf(dtSource);

                        #region 表头及样式
                        {
                            if (dtSource.TableName.IndexOf("Table", StringComparison.Ordinal) != 0)
                            {
                                IRow headerRow = sheet.CreateRow(rowIndex);
                                headerRow.HeightInPoints = dtSource.TableName.Length != new Regex(" ").Replace(dtSource.TableName, "\n", 1).Length ? 25 : 50;
                                headerRow.CreateCell(0).SetCellValue(new Regex(" ").Replace(dtSource.TableName, "\n", 1));

                                ICellStyle headStyle = workbook.CreateCellStyle();
                                headStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CENTER;
                                IFont font = workbook.CreateFont();
                                font.FontHeightInPoints = 20;
                                font.Boldweight = 700;
                                //font.Color = HSSFColor.WHITE.index;
                                headStyle.SetFont(font);
                                headStyle.WrapText = true;
                                //headStyle.FillForegroundColor = GetXLColour(workbook, AppConfig.ZhongTaiLightRed);
                                //headStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;
                                headerRow.GetCell(0).CellStyle = headStyle;
                                sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(rowIndex, rowIndex, 0, dtSource.Columns.Count - 1 < 0 ? 0 : dtSource.Columns.Count - 1));
                            }
                        }
                        #endregion

                        #region 列头及样式
                        {
                            IRow headerRow = sheet.CreateRow(rowIndex + 1);
                            ICellStyle headStyle = workbook.CreateCellStyle();
                            headStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CENTER;
                            IFont font = workbook.CreateFont();
                            font.FontHeightInPoints = 10;
                            font.Boldweight = 700;
                            //font.Color = HSSFColor.WHITE.index;
                            headStyle.SetFont(font);
                            headStyle.WrapText = true;
                            //headStyle.FillForegroundColor = GetXLColour(workbook, AppConfig.ZhongTaiLightRed);
                            //headStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;
                            foreach (DataColumn column in dtSource.Columns)
                            {
                                headerRow.CreateCell(column.Ordinal).SetCellValue(new Regex(" ").Replace(column.ColumnName, "\n", 1));
                                headerRow.GetCell(column.Ordinal).CellStyle = headStyle;
                            }
                        }
                        #endregion

                        rowIndex += 2;
                    }
                    #endregion

                    #region 填充内容
                    IRow dataRow = sheet.CreateRow(rowIndex);
                    foreach (DataColumn column in dtSource.Columns)
                    {
                        ICell newCell = dataRow.CreateCell(column.Ordinal);

                        string drValue = row[column].ToString();

                        switch (column.DataType.ToString())
                        {
                            case "System.String"://字符串类型
                                newCell.SetCellValue(drValue);
                                break;
                            case "System.DateTime"://日期类型
                                System.DateTime dateV;
                                System.DateTime.TryParse(drValue, out dateV);
                                newCell.SetCellValue(dateV);

                                newCell.CellStyle = dateStyle;//格式化显示
                                break;
                            case "System.Boolean"://布尔型
                                bool boolV = false;
                                bool.TryParse(drValue, out boolV);
                                if (boolKeyValue != null)
                                {
                                    string value;
                                    boolKeyValue.TryGetValue(boolV, out value);
                                    newCell.SetCellValue(value);
                                }
                                else
                                {
                                    newCell.SetCellValue(boolV ? "√" : "×");
                                }
                                break;
                            case "System.Int16"://整型
                            case "System.Int32":
                            case "System.Int64":
                            case "System.Byte":
                                int intV = 0;
                                int.TryParse(drValue, out intV);
                                newCell.SetCellValue(intV);
                                break;
                            case "System.Decimal"://浮点型
                            case "System.Double":
                                double doubV = 0;
                                double.TryParse(drValue, out doubV);
                                newCell.SetCellValue(doubV);
                                break;
                            case "System.DBNull"://空值处理
                                newCell.SetCellValue("");
                                break;
                            default:
                                newCell.SetCellValue("");
                                break;
                        }
                    }
                    #endregion

                    rowIndex++;
                }
                rowIndex += 2;
            }
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Write(ms);
                ms.Flush();
                //ms.Position = 0;
                //sheet.Dispose();
                return ms;
            }
        }


        public static DataTable FormatDataTable(DataTable dataTable, Dictionary<string, string> ColDic = null, Dictionary<bool, string> boolKeyValue = null)
        {
            DataTable dt = dataTable.Copy();
            DataTable dtTemp = dt.Clone();
            List<int> removeIndexList = new List<int>();
            foreach (DataColumn col in dt.Columns)
            {
                if (ColDic != null && ColDic.Count > 0)
                {
                    bool isHave = false;
                    foreach (KeyValuePair<string, string> colDic in ColDic.Where(colDic => col.ColumnName == colDic.Key))
                    {
                        dtTemp.Columns[col.Ordinal].ColumnName = colDic.Value;
                        isHave = true;
                    }
                    if (!isHave)
                    {
                        removeIndexList.Add(col.Ordinal);
                    }
                }

                if (col.ColumnName.Contains("日期") || col.ColumnName.Contains("时间"))
                {
                    dtTemp.Columns[col.Ordinal].DataType = typeof(string);
                }
                if (col.ColumnName.ToUpper().EndsWith("GUID") || col.ColumnName.EndsWith("RI"))
                {
                    removeIndexList.Add(col.Ordinal);
                }
            }

            foreach (DataRow row in dt.Rows)
            {
                DataRow newDtRow = dtTemp.NewRow();
                foreach (DataColumn col in dt.Columns)
                {
                    string value = row[col.Ordinal].ToString();
                    newDtRow[col.Ordinal] = row[col.Ordinal];
                    if (col.ColumnName.Contains("日期"))
                    {
                        try
                        {
                            newDtRow[col.Ordinal] = Convert.ToDateTime(row[col.Ordinal]).ToString("yyyy/MM/dd");
                        }
                        catch (Exception)
                        {
                            newDtRow[col.Ordinal] = value;
                        }
                    }
                    if (col.ColumnName.Contains("时间"))
                    {
                        try
                        {
                            if (row[col.Ordinal].ToString().Length <= 12)
                            {
                                newDtRow[col.Ordinal] = value;
                            }
                            else
                            {
                                newDtRow[col.Ordinal] = Convert.ToDateTime(row[col.Ordinal]).ToString("yyyy/MM/dd HH:mm");
                            }
                        }
                        catch (Exception)
                        {
                            newDtRow[col.Ordinal] = value;
                        }
                    }
                }
                dtTemp.Rows.Add(newDtRow);
            }

            removeIndexList = removeIndexList.Where((x, i) => removeIndexList.FindIndex(z => z == x) == i).OrderByDescending(x => x).ToList();
            foreach (int removeIndex in removeIndexList)
            {
                dtTemp.Columns.RemoveAt(removeIndex);
            }

            return dtTemp;
        }

    }
}
