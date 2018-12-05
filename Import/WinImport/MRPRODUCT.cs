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
    public partial class MRPRODUCT : Form
    {
        ExcelManager _excelManager = new ExcelManager();
        /// <summary>
        /// excel的设备数据
        /// </summary>
        DataTable _ECInfo_excel = null;
        /// <summary>
        /// 数据库的设备数据
        /// </summary>
        DataTable _ECInfo_DB = null;
        /// <summary>
        /// excel的物料校验项数据
        /// </summary>
        DataTable _ECInfoD_excel = null;
        /// <summary>
        /// excel的物料保养项数据
        /// </summary>
        DataTable _ECInfoE_excel = null;

        /// <summary>
        /// 数据库的部门-岗位表（级联部门和岗位）
        /// </summary>
        DataTable _hrDeptPos_DB = null;

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

        /// <summary>
        /// 数据库的设备类别表
        /// </summary>
        DataTable _ECType_DB = null;

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
        public MRPRODUCT()
        {
            InitializeComponent();

            CTYPEs.Add("数值", "数值");
            CTYPEs.Add("文本", "文本");

            ISINTERVALs.Add("是", true);
            ISINTERVALs.Add("否", false);
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
                    _ECInfo_excel = _excelManager.GetExcelTableByOleDB(opfDialog.FileName, "MRPRODUCT");
                    MessageBox.Show("读取笔数:" + _ECInfo_excel.Rows.Count + "");
                }
                else if (btn.Name == "btnSelect3")
                {
                    //物料校验项
                    txtFile3.Text = opfDialog.FileName;
                    _ECInfoD_excel = _excelManager.GetExcelTableByOleDB(opfDialog.FileName, "MRCHECKCT");
                    MessageBox.Show("读取笔数:" + _ECInfoD_excel.Rows.Count + "");
                }
                else if (btn.Name == "btnSelect4")
                {
                    //物料保养项
                    txtFile4.Text = opfDialog.FileName;
                    _ECInfoE_excel = _excelManager.GetExcelTableByOleDB(opfDialog.FileName, "MRMAINT");
                    MessageBox.Show("读取笔数:" + _ECInfoE_excel.Rows.Count + "");
                }
                ClearSql();
                if (_ECInfo_excel == null || _ECInfo_excel.Rows.Count <= 0)
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
            dgRepet_excel.DataSource = new DataTable();
        }

        private void btnCheck_Click(object sender, EventArgs e)
        {
            if (_ECInfo_excel == null)
            {
                WGMessage.ShowWarning("请选择[设备信息]文件!");
                return;
            }
            if (_ECInfoD_excel == null)
            {
                WGMessage.ShowWarning("请选择[物料校验项]文件!");
                return;
            }
            if (_ECInfoE_excel == null)
            {
                WGMessage.ShowWarning("请选择[物料保养项]文件!");
                return;
            }
            if (isCheck)
            {
                WGMessage.ShowAsterisk("已验证，不用重复验证！");
                return;
            }
            //加载对应厂部的部门-岗位
            string sql = @"SELECT t1.[GUID] BSDEPTPOS_GUID,t2.GUID BSDEPT_GUID,t2.CODE BSDEPT_CODE,t2.NAME BSDEPT_NAME,t3.CODE BSPOSITION_CODE,t3.NAME BSPOSITION_NAME
                              FROM [BSDEPTPOS] t1
                              left join BSDEPT t2 on t1.PGUID=t2.GUID
                              left join BSPOSITION t3 on t1.FGUID=t3.GUID";
            _hrDeptPos_DB = FillDatatablde(sql, Main.CONN_Public);
            //加载物料基本信息
            sql = @"select GUID,CODE,NAME from BSPRODUCT WHERE [CTYPE] IN (2,3,4,5,6)";
            _ECInfo_DB = FillDatatablde(sql, Main.CONN_Public);

            //加载物料类别
            sql = @"select GUID,CODE,NAME,CTYPE from [BSPRODTYPE] WHERE CTYPE IN (2,3,4,5,6)";
            _ECType_DB = FillDatatablde(sql, Main.CONN_Public);

            sql = @"select GUID,CODE,NAME FROM BSPRODUCT WHERE CTYPE IN (2,3,4,5,6)";
            _BSPRODUCT_DB = FillDatatablde(sql, Main.CONN_Public);

            //错误
            List<int[]> col_error = new List<int[]>();

            //重复数据
            DataTable dt_repet = _ECInfo_excel.Clone();

            #region  物料信息 验证
            DataTable dt_error = _ECInfo_excel.Clone();
            DataTable dt_repet_excel = _ECInfo_excel.Clone();

            DataTable dt_repet_excel1 = _ECInfoD_excel.Clone();
            DataTable dt_repet_excel2 = _ECInfoE_excel.Clone();

            //存放要保存的设备厂内编号对应的guid
            Dictionary<string, Guid> newEC = new Dictionary<string, Guid>();
            Dictionary<string, Guid> newDJName = new Dictionary<string, Guid>();

            for (int i = 0; i < _ECInfo_excel.Rows.Count; i++)
            {
                bool isError = false;
                bool isRepet = false;
                bool isRepet_excel = false;

                DataRow dr_excel = _ECInfo_excel.Rows[i];

                DataRow[] drs_ectype = _ECType_DB.Select("CODE='" + WGHelper.ReturnString(dr_excel["MRO类别"].ToString()) + "'");
                if (string.IsNullOrWhiteSpace(dr_excel["MRO类别"].ToString()) || drs_ectype.Length == 0)
                {
                    //空、不存在
                    col_error.Add(new int[] { dt_error.Rows.Count, 3 });
                    isError = true;
                }

               
                if (string.IsNullOrWhiteSpace(dr_excel["MRO编号"].ToString()) )
                {
                    //空、不存在
                    col_error.Add(new int[] { dt_error.Rows.Count, 0 });
                    isError = true;
                }

                DataRow[] drs_mrcode = _ECInfo_DB.Select("CODE='" + WGHelper.ReturnString(dr_excel["MRO编号"].ToString()) + "'");

                if (drs_mrcode.Length > 0)
                {
                    isRepet = true;
                }

                DataRow[] drs_repet_excel = _ECInfo_excel.Select("MRO编号 = '" + WGHelper.ReturnString(dr_excel["MRO编号"].ToString()) + "'");

                if(drs_repet_excel.Length > 1)
                {
                    isRepet_excel = true;
                }

                if (string.IsNullOrWhiteSpace(dr_excel["MRO名称"].ToString()))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 1 });
                    isError = true;
                }
                if (!string.IsNullOrWhiteSpace(dr_excel["安全库存管理"].ToString()) &&
                    dr_excel["安全库存管理"].ToString().ToLower() == "true" && string.IsNullOrWhiteSpace(dr_excel["安全库存"].ToString()))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 2 });
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
                    if(isRepet_excel)
                    {
                        dt_repet_excel.Rows.Add(dr_excel.ItemArray);
                    }
                    continue;
                }

                Guid n = Guid.NewGuid();
                string temp = @"INSERT INTO [BSPRODUCT]
                                       ([GUID],[CODE],[NAME],[PGUID],[CTYPE]
                                       ,[ISMAINTMGR],[ISCHECKMGR],[ISSTORAGEMGR],[STORAGENUM],[ISSNMGR],[ISSEQUESTMGR],[ND])
                                 VALUES
                                       (" + Main.SetDBValue(n) + "," + Main.SetDBValue(dr_excel["MRO编号"]) + "," + Main.SetDBValue(dr_excel["MRO名称"]) + "," + Main.SetDBValue(drs_ectype[0]["GUID"]) + @"
                                        ," + Main.SetDBValue(drs_ectype[0]["CTYPE"]) + "," + Main.SetDBValue(dr_excel["保养管理"]) + "," + Main.SetDBValue(dr_excel["校验管理"]) + "," + Main.SetDBValue(dr_excel["安全库存管理"]) + "," + Main.SetDBValue(dr_excel["安全库存"]) + "," +
                                        "" + Main.SetDBValue(dr_excel["标识管理"]) + " ," + Main.SetDBValue(dr_excel["封存管理"]) + "," + Main.SetDBValue(DateTime.Now) + ")";
                rbSql.Text += temp + Environment.NewLine;
                sqlLs.Add(temp);
                newEC.Add(dr_excel["MRO编号"].ToString(), n);
            }
            dgError.DataSource = dt_error;
            dgRepet.DataSource = dt_repet;
            dgRepet_excel.DataSource = dt_repet_excel;
            if (dt_error.Rows.Count > 0 || dt_repet.Rows.Count > 0||dt_repet_excel.Rows.Count >0)
            {
                Main.SetErrorCell(dgError, col_error);
                rbSql.Text = "";
                return;
            }
            #endregion

            #region 物料校验项 验证
            dt_error = _ECInfoD_excel.Clone();
            for (int i = 0; i < _ECInfoD_excel.Rows.Count; i++)
            {
                bool isError = false;
                bool isRepet = false;
                bool isRepet_excel = false;

                DataRow dr_excel = _ECInfoD_excel.Rows[i];

                int SNO = 0;
                if (string.IsNullOrWhiteSpace(dr_excel["序号"].ToString())
                    || !int.TryParse(dr_excel["序号"].ToString(), out SNO))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 0 });
                    isError = true;
                }


                if (string.IsNullOrWhiteSpace(dr_excel["MRO编号"].ToString()))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 1 });
                    isError = true;
                }
                else if (!newEC.ContainsKey(dr_excel["MRO编号"].ToString()))
                {
                    // 不存在
                    col_error.Add(new int[] { dt_error.Rows.Count, 1 });
                    isError = true;
                }

                if (string.IsNullOrWhiteSpace(dr_excel["校验项"].ToString()))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 2 });
                    isError = true;
                }

                if (string.IsNullOrWhiteSpace(dr_excel["类型"].ToString())
                   || !CTYPEs.ContainsKey(dr_excel["类型"].ToString()))
                {
                    //空、不包含
                    col_error.Add(new int[] { dt_error.Rows.Count, 5 });
                    isError = true;
                }

                DataRow[] drs_repet = _ECInfoD_excel.Select("校验项 = '" + dr_excel["校验项"].ToString() + "' AND MRO编号 = '" + dr_excel["MRO编号"] + "' ");
                if (drs_repet.Length > 1)
                {
                    isRepet_excel = true;
                }

                DataRow[] drs_dept_pos_a = _hrDeptPos_DB.Select("BSDEPT_CODE='"
                                 + WGHelper.ReturnString(dr_excel["处理部门"].ToString())
                                 + "' and BSPOSITION_CODE='" + WGHelper.ReturnString(dr_excel["处理岗位"].ToString()) + "'");
                if (string.IsNullOrWhiteSpace(dr_excel["处理部门"].ToString())
                    || string.IsNullOrWhiteSpace(dr_excel["处理岗位"].ToString())
                    || drs_dept_pos_a.Length == 0)
                {
                    //空、不存在
                    col_error.Add(new int[] { dt_error.Rows.Count, 8 });
                    col_error.Add(new int[] { dt_error.Rows.Count, 9 });
                    isError = true;
                }

                DataRow[] drs_dept_pos_b = _hrDeptPos_DB.Select("BSDEPT_CODE='"
                 + WGHelper.ReturnString(dr_excel["响应部门"].ToString())
                 + "' and BSPOSITION_CODE='" + WGHelper.ReturnString(dr_excel["响应岗位"].ToString()) + "'");
                if (string.IsNullOrWhiteSpace(dr_excel["响应部门"].ToString())
                    || string.IsNullOrWhiteSpace(dr_excel["响应岗位"].ToString())
                    || drs_dept_pos_b.Length == 0)
                {
                    //空、不存在
                    col_error.Add(new int[] { dt_error.Rows.Count, 10 });
                    col_error.Add(new int[] { dt_error.Rows.Count, 11 });
                    isError = true;
                }

                if (dr_excel["类型"].ToString() == "数值")
                {
                    if (string.IsNullOrWhiteSpace(dr_excel["是否区间"].ToString())
                        || !ISINTERVALs.ContainsKey(dr_excel["是否区间"].ToString()))
                    {
                        //空、不包含
                        col_error.Add(new int[] { dt_error.Rows.Count, 12 });
                        isError = true;
                    }
                }


                int DAYS = 0;
                //if (string.IsNullOrWhiteSpace(dr_excel["周期(天)"].ToString())
                //    && !int.TryParse(dr_excel["周期(天)"].ToString(), out DAYS))
                //{
                //    //空
                //    col_error.Add(new int[] { dt_error.Rows.Count, 17 });
                //    isError = true;
                //}

                string Stand_Str = dr_excel["符号1"].ToString() + dr_excel["最小值"].ToString() + " " + dr_excel["符号2"].ToString() + dr_excel["最大值"].ToString();

                if (dr_excel["类型"].ToString() == "文本" && string.IsNullOrWhiteSpace(dr_excel["标准"].ToString()))
                {
                    col_error.Add(new int[] { dt_error.Rows.Count, 7 });
                    col_error.Add(new int[] { dt_error.Rows.Count, 13 });
                    col_error.Add(new int[] { dt_error.Rows.Count, 14 });
                    col_error.Add(new int[] { dt_error.Rows.Count, 15 });
                    col_error.Add(new int[] { dt_error.Rows.Count, 16 });
                    isError = true;
                }

                if (dr_excel["类型"].ToString() == "数值")
                {
                    dr_excel["标准"] = Stand_Str;
                }
                if ((dr_excel["是否区间"].ToString() == "是" && dr_excel["符号1"].ToString() == "=") || ((dr_excel["是否区间"].ToString() == "否" && dr_excel["符号1"].ToString() != "=")))
                {
                    col_error.Add(new int[] { dt_error.Rows.Count, 12 });
                    col_error.Add(new int[] { dt_error.Rows.Count, 15 });
                    isError = true;
                }

                if (isError || isRepet||isRepet_excel)
                {
                    if (isError)
                    {
                        dt_error.Rows.Add(dr_excel.ItemArray);
                    }
                    if (isRepet_excel)
                    {
                        dt_repet_excel1.Rows.Add(dr_excel.ItemArray);
                    }
                    continue;
                }

                string temp = @"INSERT INTO [dbo].[MRCHECKCT]
                                       ([GUID],[PGUID],[SNO],[NAME],[METHOD]
                                       ,[REQUEST],[CTYPE],[SVALUE],[ISINTERVAL],[MINVALUE]
                                       ,[MAXVALUE],[SG1],[SG2],[CYCLE],[AGUID]
                                       ,[BGUID])
                                 VALUES
                                       (" + Main.SetDBValue(Guid.NewGuid()) + "," + Main.SetDBValue(newEC[dr_excel["MRO编号"].ToString()]) + "," + Main.SetDBValue(dr_excel["序号"]) + "," + Main.SetDBValue(dr_excel["校验项"]) + "," + Main.SetDBValue(dr_excel["方法"]) + @"
                                        ," + Main.SetDBValue(dr_excel["要求"]) + "," + Main.SetDBValue(CTYPEs[dr_excel["类型"].ToString()]) + "," + Main.SetDBValue(dr_excel["标准"]) + "," + (GetISINTERVAL(dr_excel["是否区间"].ToString())) + "," + Main.SetDBValue(dr_excel["最小值"]) + @"
                                        ," + Main.SetDBValue(dr_excel["最大值"]) + "," + Main.SetDBValue(dr_excel["符号1"]) + "," + Main.SetDBValue(dr_excel["符号2"]) + "," + Main.SetDBValue(dr_excel["周期(天)"]) + "," + Main.SetDBValue(drs_dept_pos_a[0]["BSDEPTPOS_GUID"]) + @"
                                        ," + Main.SetDBValue(drs_dept_pos_b[0]["BSDEPTPOS_GUID"]) + ")";
                rbSql.Text += temp + Environment.NewLine;
                sqlLs.Add(temp);
            }
            dgError.DataSource = dt_error;
            dgRepet_excel1.DataSource = dt_repet_excel1;
            if (dt_error.Rows.Count > 0 || dt_repet_excel1.Rows.Count >0)
            {
                Main.SetErrorCell(dgError, col_error);
                rbSql.Text = "";
                return;
            }
            #endregion

            #region 物料保养项 验证
            dt_error = _ECInfoE_excel.Clone();

            for (int i = 0; i < _ECInfoE_excel.Rows.Count; i++)
            {
                bool isError = false;
                bool isRepet = false;
                bool isRepet_excel = false;

                DataRow dr_excel = _ECInfoE_excel.Rows[i];


                int SNO = 0;
                if (string.IsNullOrWhiteSpace(dr_excel["序号"].ToString())
                    || !int.TryParse(dr_excel["序号"].ToString(), out SNO))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 0 });
                    isError = true;
                }

                if (string.IsNullOrWhiteSpace(dr_excel["MRO编号"].ToString()))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 1 });
                    isError = true;
                }
                else if (!newEC.ContainsKey(dr_excel["MRO编号"].ToString()))
                {
                    // 不存在
                    col_error.Add(new int[] { dt_error.Rows.Count, 1 });
                    isError = true;
                }

                if (string.IsNullOrWhiteSpace(dr_excel["保养项"].ToString()))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 2 });
                    isError = true;
                }

                if (string.IsNullOrWhiteSpace(dr_excel["类型"].ToString())
                   || !CTYPEs.ContainsKey(dr_excel["类型"].ToString()))
                {
                    //空、不包含
                    col_error.Add(new int[] { dt_error.Rows.Count, 5 });
                    isError = true;
                }

                DataRow[] drs_repet = _ECInfoE_excel.Select("保养项 = '" + dr_excel["保养项"].ToString() + "' AND MRO编号 = '" + dr_excel["MRO编号"] + "' ");
                if (drs_repet.Length > 1)
                {
                    isRepet_excel = true;
                }

                DataRow[] drs_dept_pos_a = _hrDeptPos_DB.Select("BSDEPT_CODE='"
                                                 + WGHelper.ReturnString(dr_excel["处理部门"].ToString())
                                                 + "' and BSPOSITION_CODE='" + WGHelper.ReturnString(dr_excel["处理岗位"].ToString()) + "'");
                if (string.IsNullOrWhiteSpace(dr_excel["处理部门"].ToString())
                    || string.IsNullOrWhiteSpace(dr_excel["处理岗位"].ToString())
                    || drs_dept_pos_a.Length == 0)
                {
                    //空、不存在
                    col_error.Add(new int[] { dt_error.Rows.Count, 8 });
                    col_error.Add(new int[] { dt_error.Rows.Count, 9 });
                    isError = true;
                }

                DataRow[] drs_dept_pos_b = _hrDeptPos_DB.Select("BSDEPT_CODE='"
                                                  + WGHelper.ReturnString(dr_excel["响应部门"].ToString())
                                                  + "' and BSPOSITION_CODE='" + WGHelper.ReturnString(dr_excel["响应岗位"].ToString()) + "'");
                if (string.IsNullOrWhiteSpace(dr_excel["响应部门"].ToString())
                    || string.IsNullOrWhiteSpace(dr_excel["响应岗位"].ToString())
                    || drs_dept_pos_b.Length == 0)
                {
                    //空、不存在
                    col_error.Add(new int[] { dt_error.Rows.Count, 10 });
                    col_error.Add(new int[] { dt_error.Rows.Count, 11 });
                    isError = true;
                }

                int DAYS = 0;
                if (string.IsNullOrWhiteSpace(dr_excel["周期(天)"].ToString())
                    && !int.TryParse(dr_excel["周期(天)"].ToString(), out DAYS))
                {
                    //空
                    col_error.Add(new int[] { dt_error.Rows.Count, 17 });
                    isError = true;
                }

                string Stand_Str = dr_excel["符号1"].ToString() + dr_excel["最小值"].ToString() + " " + dr_excel["符号2"].ToString() + dr_excel["最大值"].ToString();

                if (dr_excel["类型"].ToString() == "文本" && string.IsNullOrWhiteSpace(dr_excel["标准"].ToString()))
                {
                    col_error.Add(new int[] { dt_error.Rows.Count, 7 });
                    col_error.Add(new int[] { dt_error.Rows.Count, 13 });
                    col_error.Add(new int[] { dt_error.Rows.Count, 14 });
                    col_error.Add(new int[] { dt_error.Rows.Count, 15 });
                    col_error.Add(new int[] { dt_error.Rows.Count, 16 });
                    isError = true;
                }

                if (dr_excel["类型"].ToString() == "数值")
                {
                    dr_excel["标准"] = Stand_Str;
                }
                if ((dr_excel["是否区间"].ToString() == "是" && dr_excel["符号1"].ToString() == "=") || ((dr_excel["是否区间"].ToString() == "否" && dr_excel["符号1"].ToString() != "=")))
                {
                    col_error.Add(new int[] { dt_error.Rows.Count, 12 });
                    col_error.Add(new int[] { dt_error.Rows.Count, 15 });
                    isError = true;
                }

                if (isError || isRepet || isRepet_excel)
                {
                    if (isError)
                    {
                        dt_error.Rows.Add(dr_excel.ItemArray);
                    }
                    if (isRepet_excel)
                    {
                        dt_repet_excel2.Rows.Add(dr_excel.ItemArray);
                    }
                    continue;
                }
                string temp = @"INSERT INTO [dbo].[MRMAINT]
                                       ([GUID],[PGUID],[SNO],[NAME],[METHOD]
                                       ,[REQUEST],[CTYPE],[SVALUE],[ISINTERVAL],[MINVALUE]
                                       ,[MAXVALUE],[SG1],[SG2],[CYCLE],[AGUID]
                                       ,[BGUID])
                                 VALUES
                                       (" + Main.SetDBValue(Guid.NewGuid()) + "," + Main.SetDBValue(newEC[dr_excel["MRO编号"].ToString()]) + "," + Main.SetDBValue(dr_excel["序号"]) + "," + Main.SetDBValue(dr_excel["保养项"]) + "," + Main.SetDBValue(dr_excel["方法"]) + @"
                                        ," + Main.SetDBValue(dr_excel["要求"]) + "," + Main.SetDBValue(CTYPEs[dr_excel["类型"].ToString()]) + "," + Main.SetDBValue(dr_excel["标准"]) + "," + (GetISINTERVAL(dr_excel["是否区间"].ToString())) + "," + Main.SetDBValue(dr_excel["最小值"]) + @"
                                        ," + Main.SetDBValue(dr_excel["最大值"]) + "," + Main.SetDBValue(dr_excel["符号1"]) + "," + Main.SetDBValue(dr_excel["符号2"]) + "," + Main.SetDBValue(dr_excel["周期(天)"]) + "," + Main.SetDBValue(drs_dept_pos_a[0]["BSDEPTPOS_GUID"]) + @"
                                        ," + Main.SetDBValue(drs_dept_pos_b[0]["BSDEPTPOS_GUID"]) + ")";
                rbSql.Text += temp + Environment.NewLine;
                sqlLs.Add(temp);
            }
            dgError.DataSource = dt_error;
            dgRepet_excel2.DataSource = dt_repet_excel2;
            if (dt_error.Rows.Count > 0 || dt_repet_excel2.Rows.Count > 0)
            {
                Main.SetErrorCell(dgError, col_error);
                rbSql.Text = "";
                return;
            }
            #endregion
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

    }
}
