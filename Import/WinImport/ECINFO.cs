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
    public partial class ECINFO : Form
    {
        ExcelManager _excelManager = new ExcelManager();
        /// <summary>
        /// excel的设备数据
        /// </summary>
        DataTable _ECInfo_excel = null;
        /// <summary>
        /// 数据库的设备数据
        /// </summary>
        DataTable _ECInfo_DB = new DataTable();
        /// <summary>
        /// excel的AM点检数据
        /// </summary>
        DataTable _ECInfoC_excel = new DataTable();

        /// <summary>
        /// excel的AM点检产品
        /// </summary>
        DataTable _ECInfoCS_excel = new DataTable();

        /// <summary>
        /// excel的PM巡检数据
        /// </summary>
        DataTable _ECInfoD_excel = new DataTable();
        /// <summary>
        /// excel的PM保养数据
        /// </summary>
        DataTable _ECInfoE_excel = new DataTable();
        /// <summary>
        /// excel的技术规格数据
        /// </summary>
        DataTable _ECInfoA_excel = new DataTable();

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

        //重要等级
        DataTable _ECCLASS_DB = null;

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
        public ECINFO()
        {
            InitializeComponent();

            //CTYPEs.Add("数值", "数值");
            //CTYPEs.Add("文本", "文本");

            STs.Add("未使用", 1);
            STs.Add("使用中", 2);
            STs.Add("停用", 3);

            //BTYPE.Add("通用", "通用");
            //BTYPE.Add("产品", "产品");

            //ISINTERVALs.Add("是", true);
            //ISINTERVALs.Add("否", false);
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
                    _ECInfo_excel = _excelManager.GetExcelTableByOleDB(opfDialog.FileName, "ECINFO");
                    MessageBox.Show("读取笔数:"+_ECInfo_excel.Rows.Count + "");
                }
                //else if (btn.Name == "btnSelect2")
                //{
                //    //AM点检
                //    txtFile2.Text = opfDialog.FileName;
                //    _ECInfoC_excel = _excelManager.GetExcelTableByOleDB(opfDialog.FileName, "ECINFOC");
                //    MessageBox.Show("读取笔数:" + _ECInfoC_excel.Rows.Count + "");
                //}
                //else if (btn.Name == "btnSelect21")
                //{
                //    //AM点检产品
                //    txtFile21.Text = opfDialog.FileName;
                //    _ECInfoCS_excel = _excelManager.GetExcelTableByOleDB(opfDialog.FileName, "ECINFOCS");
                //    MessageBox.Show("读取笔数:" + _ECInfoCS_excel.Rows.Count + "");
                //}
                //else if (btn.Name == "btnSelect3")
                //{
                //    //PM巡检
                //    txtFile3.Text = opfDialog.FileName;
                //    _ECInfoD_excel = _excelManager.GetExcelTableByOleDB(opfDialog.FileName, "ECINFOD");
                //    MessageBox.Show("读取笔数:" + _ECInfoD_excel.Rows.Count + "");
                //}
                //else if (btn.Name == "btnSelect4")
                //{
                //    //PM保养
                //    txtFile4.Text = opfDialog.FileName;
                //    _ECInfoE_excel = _excelManager.GetExcelTableByOleDB(opfDialog.FileName, "ECINFOE");
                //    MessageBox.Show("读取笔数:" + _ECInfoE_excel.Rows.Count + "");
                //}
                //else if (btn.Name == "btnSelect5")
                //{
                //    //技术规格
                //    txtFile5.Text = opfDialog.FileName;
                //    _ECInfoA_excel = _excelManager.GetExcelTableByOleDB(opfDialog.FileName, "ECINFOA");
                //    MessageBox.Show("读取笔数:" + _ECInfoA_excel.Rows.Count + "");
                //}
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
            dgError1.DataSource = new DataTable();
            dgRepet1.DataSource = new DataTable();
            dgError2.DataSource = new DataTable();
            dgRepet2.DataSource = new DataTable();
            dgError3.DataSource = new DataTable();
            dgRepet3.DataSource = new DataTable();
            dgError4.DataSource = new DataTable();
            dgRepet4.DataSource = new DataTable();
            dgError5.DataSource = new DataTable();
            dgRepet5.DataSource = new DataTable();
        }

        private void btnCheck_Click(object sender, EventArgs e)
        {
            if (_ECInfo_excel == null)
            {
                WGMessage.ShowWarning("请选择[设备信息]文件!");
                return;
            }
            //if (_ECInfoC_excel == null)
            //{
            //    WGMessage.ShowWarning("请选择[AM点检]文件!");
            //    return;
            //}
            //if (_ECInfoCS_excel == null)
            //{
            //    WGMessage.ShowWarning("请选择[AM点检产品]文件!");
            //    return;
            //}
            //if (_ECInfoD_excel == null)
            //{
            //    WGMessage.ShowWarning("请选择[PM巡检]文件!");
            //    return;
            //}
            //if (_ECInfoE_excel == null)
            //{
            //    WGMessage.ShowWarning("请选择[PM保养]文件!");
            //    return;
            //}
            //if (_ECInfoA_excel == null)
            //{
            //    WGMessage.ShowWarning("请选择[技术规格]文件!");
            //    return;
            //}

            if (_ECInfoC_excel == null)
            {
                _ECInfoC_excel = new DataTable();
            }
            if (_ECInfoCS_excel == null)
            {
                _ECInfoCS_excel = new DataTable();
            }
            if (_ECInfoD_excel == null)
            {
                _ECInfoD_excel = new DataTable();
            }
            if (_ECInfoE_excel == null)
            {
                _ECInfoE_excel = new DataTable();
            }
            if (_ECInfoA_excel == null)
            {
                _ECInfoA_excel = new DataTable();
            }

            if (isCheck)
            {
                WGMessage.ShowAsterisk("已验证，不用重复验证！");
                return;
            }
            ////加载对应厂部的部门-岗位
            //string sql = @"SELECT t1.[GUID] BSDEPTPOS_GUID,t2.GUID BSDEPT_GUID,t2.CODE BSDEPT_CODE,t2.NAME BSDEPT_NAME,t3.CODE BSPOSITION_CODE,t3.NAME BSPOSITION_NAME
            //                  FROM [BSDEPTPOS] t1
            //                  left join BSDEPT t2 on t1.PGUID=t2.GUID
            //                  left join BSPOSITION t3 on t1.FGUID=t3.GUID";
            //_hrDeptPos_DB = FillDatatablde(sql, Main.CONN_Public);
            //加载设备信息
            string sql = @"select GUID,CODE,NAME from ECINFO";
            _ECInfo_DB = FillDatatablde(sql, Main.CONN_Public);
            //加载设备类别
            sql = @"select GUID,CODE,NAME from ECTYPE";
            _ECType_DB = FillDatatablde(sql, Main.CONN_Public);
            //加载供应商
            sql = @"select GUID,CODE,NAME from BSSUPPLIER";
            _BSSupplier_DB = FillDatatablde(sql, Main.CONN_Public);

            sql = @"select GUID,CODE,NAME FROM BSDEPT";
            _BSDEPT_DB = FillDatatablde(sql, Main.CONN_Public);

            //sql = @"select GUID,CODE,NAME FROM BSPRODUCT WHERE CTYPE = 1";
            //_BSPRODUCT_DB = FillDatatablde(sql, Main.CONN_Public);

            //sql = @"SELECT GUID,CODE,NAME FROM ECCLASS";
            //_ECCLASS_DB = FillDatatablde(sql, Main.CONN_Public);

            #region  设备信息 验证
            //重复数据
            //错误
            List<int[]> col_error1 = new List<int[]>();
            DataTable dt_repet1 = _ECInfo_excel.Clone();
            DataTable dt_error1 = _ECInfo_excel.Clone();
            DataTable dt_repet1_excel = _ECInfo_excel.Clone();

            //存放要保存的设备厂内编号对应的guid
            Dictionary<string, Guid> newEC = new Dictionary<string, Guid>();
            Dictionary<string, Guid> newDJName = new Dictionary<string, Guid>();

            for (int i = 0; i < _ECInfo_excel.Rows.Count; i++)
            {
                bool isError = false;
                bool isRepet = false;
                bool isRepet_excle = false;

                DataRow dr_excel = _ECInfo_excel.Rows[i];

                DataRow[] drs_ectype = _ECType_DB.Select("CODE='" + WGHelper.ReturnString(dr_excel["设备类别编号"].ToString()) + "'");
                if (string.IsNullOrWhiteSpace(dr_excel["设备类别编号"].ToString())
                    || drs_ectype.Length == 0)
                {
                    //空、不存在
                    col_error1.Add(new int[] { dt_error1.Rows.Count, 0 });
                    isError = true;
                }

                if (string.IsNullOrWhiteSpace(dr_excel["设备编号"].ToString()) || _ECInfo_excel.Select("设备编号='" + WGHelper.ReturnString(dr_excel["设备编号"].ToString()) + "'").Length > 1)
                {
                    // 存在
                    isRepet_excle = true;
                }
                
                if (_ECInfo_DB.Select("CODE='" + WGHelper.ReturnString(dr_excel["设备编号"].ToString()) + "'").Length > 0)
                {
                    isRepet = true;
                }

                if (string.IsNullOrWhiteSpace(dr_excel["设备名称"].ToString()))
                {
                    //空
                    col_error1.Add(new int[] { dt_error1.Rows.Count, 2 });
                    isError = true;
                }

                DataRow[] drs_bssupplier = _BSSupplier_DB.Select("CODE='" + WGHelper.ReturnString(dr_excel["供应商编号"].ToString()) + "'");
                if ((!string.IsNullOrWhiteSpace(dr_excel["供应商编号"].ToString()))
                    && drs_bssupplier.Length == 0)
                {
                    // 非空、不存在
                    col_error1.Add(new int[] { dt_error1.Rows.Count, 4 });
                    isError = true;
                }

                DateTime IDATE = new DateTime();
                if ((!string.IsNullOrWhiteSpace(dr_excel["启用日期"].ToString()))
                    && !DateTime.TryParse(dr_excel["启用日期"].ToString(), out IDATE))
                {
                    //非空、类型不符
                    col_error1.Add(new int[] { dt_error1.Rows.Count, 5 });
                    isError = true;
                }

                DateTime ADATE = new DateTime();
                if ((!string.IsNullOrWhiteSpace(dr_excel["投产日期"].ToString()))
                    && !DateTime.TryParse(dr_excel["投产日期"].ToString(), out ADATE))
                {
                    //非空、类型不符
                    col_error1.Add(new int[] { dt_error1.Rows.Count, 6 });
                    isError = true;
                }

                DateTime BDATE = new DateTime();
                if ((!string.IsNullOrWhiteSpace(dr_excel["购买日期"].ToString()))
                    && !DateTime.TryParse(dr_excel["购买日期"].ToString(), out BDATE))
                {
                    //非空、类型不符
                    col_error1.Add(new int[] { dt_error1.Rows.Count, 7 });
                    isError = true;
                }

                DataRow[] drs_dept_pos = _BSDEPT_DB.Select("CODE='"
                                                  + WGHelper.ReturnString(dr_excel["所属部门编号"].ToString()) + "'");
                if ((!string.IsNullOrWhiteSpace(dr_excel["所属部门编号"].ToString()))
                    && drs_dept_pos.Length == 0)
                {
                    //非空、不存在
                    col_error1.Add(new int[] { dt_error1.Rows.Count, 8 });
                    isError = true;
                }

                //DataRow[] drs_ecclass = _ECCLASS_DB.Select("CODE = '" + WGHelper.ReturnString(dr_excel["重要等级编号"].ToString()) + "'");

                //if ((!string.IsNullOrWhiteSpace(dr_excel["重要等级编号"].ToString()))
                //&& drs_ecclass.Length == 0)
                //{
                //    //非空、不存在
                //    col_error1.Add(new int[] { dt_error1.Rows.Count, 11 });
                //    isError = true;
                //}

                if (string.IsNullOrWhiteSpace(dr_excel["状态"].ToString())
                   || !STs.ContainsKey(dr_excel["状态"].ToString()))
                {
                    //空、不包含
                    col_error1.Add(new int[] { dt_error1.Rows.Count, 9 });
                    isError = true;
                }
                if (isError || isRepet || isRepet_excle)
                {
                    if (isError)
                    {
                        dt_error1.Rows.Add(dr_excel.ItemArray);
                    }
                    if (isRepet)
                    {
                        dt_repet1.Rows.Add(dr_excel.ItemArray);
                    }
                    if(isRepet_excle)
                    {
                        dt_repet1_excel.Rows.Add(dr_excel.ItemArray);
                    }
                    continue;
                }

                Guid n = Guid.NewGuid();
                string temp = string.Format(@"INSERT INTO ECINFO
                                             (GUID,AGUID,CODE,NAME,SPEC,BGUID,BCODE,BNAME,IDATE,ADATE,BDATE,DGUID,DCODE,DNAME,ST,NOTE,ND,CD)
                                             VALUES  
                                             (NEWID(),{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},GETDATE(),GETDATE())",
                                             Main.SetDBValue(drs_ectype[0]["GUID"]), Main.SetDBValue(dr_excel["设备编号"]), Main.SetDBValue(dr_excel["设备名称"]), Main.SetDBValue(dr_excel["设备型号"]),
                                             Main.SetDBValue(drs_bssupplier.Length == 0 ? "" : drs_bssupplier[0]["GUID"]), Main.SetDBValue(dr_excel["供应商编号"]), Main.SetDBValue(dr_excel["供应商名称"]), 
                                             Main.SetDBValue(dr_excel["启用日期"]), Main.SetDBValue(dr_excel["投产日期"]),Main.SetDBValue(dr_excel["购买日期"]), Main.SetDBValue(drs_dept_pos.Length == 0 ? "" : drs_dept_pos[0]["GUID"]),
                                             Main.SetDBValue(dr_excel["所属部门编号"]), Main.SetDBValue(dr_excel["所属部门名称"]), Main.SetDBValue(STs[dr_excel["状态"].ToString()]), Main.SetDBValue(dr_excel["备注"])); 
                sqlLs.Add(temp);
            }
            dgError1.DataSource = dt_error1;
            dgRepet1.DataSource = dt_repet1;
            dgRepet1_Excel.DataSource = dt_repet1_excel;
            if (dt_error1.Rows.Count > 0 || dt_repet1.Rows.Count > 0 || dt_repet1_excel.Rows.Count > 0)
            {
                Main.SetErrorCell(dgError1, col_error1);
                rbSql.Text = "";
                return;
            }
            #endregion

            #region 不需要

            //            #region AM点检 验证
            //            List<int[]> col_error2 = new List<int[]>();
            //            //重复数据
            //            DataTable dt_repet2 = _ECInfoC_excel.Clone();
            //            DataTable dt_error2 = _ECInfoC_excel.Clone();

            //            dt_error2 = _ECInfoC_excel.Clone();
            //            for (int i = 0; i < _ECInfoC_excel.Rows.Count; i++)
            //            {
            //                bool isError = false;
            //                bool isRepet = false;

            //                DataRow dr_excel = _ECInfoC_excel.Rows[i];

            //                int SNO = 0;
            //                if (string.IsNullOrWhiteSpace(dr_excel["序号"].ToString())
            //                    || !int.TryParse(dr_excel["序号"].ToString(), out SNO))
            //                {
            //                    //空
            //                    col_error2.Add(new int[] { dt_error2.Rows.Count, 0 });
            //                    isError = true;
            //                }

            //                if (string.IsNullOrWhiteSpace(dr_excel["设备编号"].ToString()))
            //                {
            //                    //空
            //                    col_error2.Add(new int[] { dt_error2.Rows.Count, 1 });
            //                    isError = true;
            //                }
            //                else if (!newEC.ContainsKey(dr_excel["设备编号"].ToString()))
            //                {
            //                    // 不存在
            //                    col_error2.Add(new int[] { dt_error2.Rows.Count, 1 });
            //                    isError = true;
            //                }

            //                if (string.IsNullOrWhiteSpace(dr_excel["点检项"].ToString()))
            //                {
            //                    //空
            //                    col_error2.Add(new int[] { dt_error2.Rows.Count, 2 });
            //                    isError = true;
            //                }

            //                if (string.IsNullOrWhiteSpace(dr_excel["类型"].ToString())
            //                   || !CTYPEs.ContainsKey(dr_excel["类型"].ToString()))
            //                {
            //                    //空、不包含
            //                    col_error2.Add(new int[] { dt_error2.Rows.Count, 5 });
            //                    isError = true;
            //                }

            //                if (string.IsNullOrWhiteSpace(dr_excel["标准"].ToString()) && dr_excel["类型"].ToString() == "文本")
            //                {
            //                    //空
            //                    col_error2.Add(new int[] { dt_error2.Rows.Count, 6 });
            //                    isError = true;
            //                }

            //                DataRow[] drs_dept_pos = _hrDeptPos_DB.Select("BSDEPT_CODE='"
            //                                  + WGHelper.ReturnString(dr_excel["响应部门编号"].ToString())
            //                                  + "' and BSPOSITION_CODE='" + WGHelper.ReturnString(dr_excel["响应岗位编号"].ToString()) + "'");
            //                if (string.IsNullOrWhiteSpace(dr_excel["响应部门编号"].ToString())
            //                    || string.IsNullOrWhiteSpace(dr_excel["响应岗位编号"].ToString())
            //                    || drs_dept_pos.Length == 0)
            //                {
            //                    //空、不存在
            //                    col_error2.Add(new int[] { dt_error2.Rows.Count, 7 });
            //                    col_error2.Add(new int[] { dt_error2.Rows.Count, 8 });
            //                    isError = true;
            //                }

            //                if (dr_excel["类型"].ToString() == "数值")
            //                {
            //                    if (string.IsNullOrWhiteSpace(dr_excel["是否区间"].ToString())
            //                        || !ISINTERVALs.ContainsKey(dr_excel["是否区间"].ToString()))
            //                    {
            //                        //空、不包含
            //                        col_error2.Add(new int[] {dt_error2.Rows.Count, 9});
            //                        isError = true;
            //                    }
            //                }

            //                if (string.IsNullOrWhiteSpace(dr_excel["适用范围"].ToString())
            //  || !BTYPE.ContainsKey(dr_excel["适用范围"].ToString()))
            //                {
            //                    //空、不包含
            //                    col_error2.Add(new int[] { dt_error2.Rows.Count, 14 });
            //                    isError = true;
            //                }


            //                if(!string.IsNullOrWhiteSpace(dr_excel["周期(天)"].ToString()) && !int.TryParse(dr_excel["周期(天)"].ToString(), out SNO))
            //                {
            //                    //空
            //                    col_error2.Add(new int[] { dt_error2.Rows.Count, 16 });
            //                    isError = true;
            //                }

            //                string Stand_Str = dr_excel["符号1"].ToString() + dr_excel["最小值"].ToString() + " " + dr_excel["符号2"].ToString() + dr_excel["最大值"].ToString();

            //                if (dr_excel["类型"].ToString() == "数值")
            //                {
            //                    dr_excel["标准"] = Stand_Str;
            //                }

            //                //if (newDJName.ContainsKey(dr_excel["点检项"].ToString()))
            //                //{
            //                //    //空
            //                //    col_error2.Add(new int[] { dt_error2.Rows.Count, 2 });
            //                //    isError = true;
            //                //}

            //                //if (Stand_Str.Contains(dr_excel["标准"].ToString()) || dr_excel["类型"].ToString() == "文本")
            //                //{

            //                //}
            //                //else
            //                //{
            //                //    col_error2.Add(new int[] { dt_error2.Rows.Count, 6 });
            //                //    col_error2.Add(new int[] { dt_error2.Rows.Count, 10 });
            //                //    col_error2.Add(new int[] { dt_error2.Rows.Count, 11 });
            //                //    col_error2.Add(new int[] { dt_error2.Rows.Count, 12 });
            //                //    col_error2.Add(new int[] { dt_error2.Rows.Count, 13 });
            //                //    isError = true;
            //                //}

            //                if ((dr_excel["是否区间"].ToString() == "是" && dr_excel["符号1"].ToString() == "=") || ((dr_excel["是否区间"].ToString() == "否" && dr_excel["符号1"].ToString() != "=")))
            //                {
            //                    col_error2.Add(new int[] { dt_error2.Rows.Count, 9 });
            //                    col_error2.Add(new int[] { dt_error2.Rows.Count, 12 });
            //                    isError = true;
            //                }

            //                if (dr_excel["适用范围"].ToString() == "通用" && dr_excel["点检频率"].ToString() == "")
            //                {
            //                    col_error2.Add(new int[] { dt_error2.Rows.Count, 15 });
            //                    isError = true;
            //                }

            //                if (dr_excel["适用范围"].ToString() == "产品" && dr_excel["点检频率"].ToString() != "")
            //                {
            //                    col_error2.Add(new int[] { dt_error2.Rows.Count, 15 });
            //                    isError = true;
            //                }

            //                if (dr_excel["点检频率"].ToString() == "按天" && dr_excel["周期(天)"].ToString() == "")
            //                {
            //                    col_error2.Add(new int[] { dt_error2.Rows.Count, 16 });
            //                    isError = true;
            //                }

            //                if (dr_excel["点检频率"].ToString() == "换班" && dr_excel["周期(天)"].ToString() != "")
            //                {
            //                    col_error2.Add(new int[] { dt_error2.Rows.Count, 16 });
            //                    isError = true;
            //                }

            //                DataRow[] drs_repit = _ECInfoC_excel.Select("设备编号 = '"+WGHelper.ReturnString(dr_excel["设备编号"].ToString())+"' AND 点检项 = '"+WGHelper.ReturnString(dr_excel["点检项"].ToString())+"'");

            //                if (drs_repit.Length > 1)
            //                {
            //                    isRepet = true;
            //                }

            //                if (isError || isRepet)
            //                {
            //                    if (isError)
            //                    {
            //                        dt_error2.Rows.Add(dr_excel.ItemArray);
            //                    }
            //                    if (isRepet)
            //                    {
            //                        dt_repet2.Rows.Add(dr_excel.ItemArray);
            //                    }
            //                    continue;
            //                }
            //                Guid kidGUID = Guid.NewGuid();
            //                string temp = @"INSERT INTO [dbo].[ECINFOC]
            //                                       ([GUID],[PGUID],[SNO],[NAME],[METHOD]
            //                                       ,[REQUEST],[CTYPE],[SVALUE],[ISINTERVAL],[MINVALUE]
            //                                       ,[MAXVALUE],[SG1],[SG2]
            //                                       ,[ATYPE],[BTYPE],[CYCLE],[AGUID])
            //                                 VALUES
            //                                       (" + Main.SetDBValue(kidGUID) + "," + Main.SetDBValue(newEC[dr_excel["设备编号"].ToString()]) + "," + Main.SetDBValue(dr_excel["序号"]) + "," + Main.SetDBValue(dr_excel["点检项"]) + "," + Main.SetDBValue(dr_excel["方法"]) + @"
            //                                        ," + Main.SetDBValue(dr_excel["要求"]) + "," + Main.SetDBValue(CTYPEs[dr_excel["类型"].ToString()]) + "," + Main.SetDBValue(dr_excel["标准"]) + "," + GetISINTERVAL(dr_excel["是否区间"].ToString()) + "," + Main.SetDBValue(dr_excel["最小值"]) + @"
            //                                        ," + Main.SetDBValue(dr_excel["最大值"]) + "," + Main.SetDBValue(dr_excel["符号1"]) + "," + Main.SetDBValue(dr_excel["符号2"]) +
            //                                        "," + Main.SetDBValue(dr_excel["适用范围"]) + "," + Main.SetDBValue(dr_excel["点检频率"]) + "," +  Main.SetDBValue(dr_excel["周期（天）"].ToString()) + "," + Main.SetDBValue(drs_dept_pos[0]["BSDEPTPOS_GUID"]) + ")";
            //                try
            //                {
            //                    newDJName.Add(dr_excel["点检项"].ToString() + dr_excel["设备编号"], kidGUID);
            //                }
            //                catch
            //                {
            //                    MessageBox.Show(dr_excel["设备编号"].ToString() + dr_excel["点检项"].ToString() +"重复！");
            //                    rbSql.Text = "";
            //                    return;
            //                }

            //                //rbSql.Text += temp + Environment.NewLine;
            //                sqlLs.Add(temp);
            //            }
            //            dgError2.DataSource = dt_error2;
            //            dgRepet2.DataSource = dt_repet2;
            //            if (dt_error2.Rows.Count > 0 || dt_repet2.Rows.Count > 0)
            //            {
            //                Main.SetErrorCell(dgError2, col_error2);
            //                rbSql.Text = "";
            //                return;
            //            }
            //            #endregion

            //            #region AM点检产品
            //            List<int[]> col_error3 = new List<int[]>();
            //            DataTable dt_repet3 = _ECInfoCS_excel.Clone();
            //            DataTable dt_error3 = _ECInfoCS_excel.Clone(); 

            //            for (int i = 0; i < _ECInfoCS_excel.Rows.Count; i++)
            //            {
            //                bool isError = false;
            //                bool isRepet = false;

            //                DataRow dr_excel = _ECInfoCS_excel.Rows[i];

            //                int SNO = 0;
            //                if (string.IsNullOrWhiteSpace(dr_excel["序号"].ToString())
            //                    || !int.TryParse(dr_excel["序号"].ToString(), out SNO))
            //                {
            //                    //空
            //                    col_error3.Add(new int[] {dt_error3.Rows.Count, 0});
            //                    isError = true;
            //                }

            //                if (string.IsNullOrWhiteSpace(dr_excel["设备编号"].ToString()))
            //                {
            //                    //空
            //                    col_error3.Add(new int[] { dt_error3.Rows.Count, 1 });
            //                    isError = true;
            //                }
            //                else if (!newEC.ContainsKey(dr_excel["设备编号"].ToString()))
            //                {
            //                    // 不存在
            //                    col_error3.Add(new int[] { dt_error3.Rows.Count, 1 });
            //                    isError = true;
            //                }

            //                if (string.IsNullOrWhiteSpace(dr_excel["点检项"].ToString()))
            //                {
            //                    //空
            //                    col_error3.Add(new int[] { dt_error3.Rows.Count, 2 });
            //                    isError = true;
            //                }
            //                else if (!newDJName.ContainsKey(dr_excel["点检项"].ToString() + dr_excel["设备编号"]))
            //                {
            //                    // 不存在
            //                    col_error3.Add(new int[] { dt_error3.Rows.Count, 2 });
            //                    isError = true;
            //                }

            //                DataRow[] drs_bsproduct = _BSPRODUCT_DB.Select("CODE = '" + dr_excel["产品编号"].ToString() + "'");

            //                if (string.IsNullOrWhiteSpace(dr_excel["产品编号"].ToString()))
            //                {
            //                    //空
            //                    col_error3.Add(new int[] { dt_error3.Rows.Count, 3 });
            //                    isError = true;
            //                }

            //                if (drs_bsproduct.Length == 0)
            //                {
            //                    col_error3.Add(new int[] { dt_error3.Rows.Count, 3 });
            //                    isError = true;
            //                }

            //                if (isError || isRepet)
            //                {
            //                    if (isError)
            //                    {
            //                        dt_error3.Rows.Add(dr_excel.ItemArray);
            //                    }
            //                    continue;
            //                }
            //                string temp = string.Format(@"INSERT INTO [dbo].[ECINFOCS]
            //        ( [GUID], [PGUID], [SNO], [FGUID] )
            //VALUES  ( NEWID(), -- GUID - uniqueidentifier
            //          {0}, -- PGUID - uniqueidentifier
            //          {1}, -- SNO - int
            //          {2}  -- FGUID - uniqueidentifier
            //          )", Main.SetDBValue(newDJName[dr_excel["点检项"].ToString()+dr_excel["设备编号"].ToString()]  ), dr_excel["序号"].ToString(), Main.SetDBValue(drs_bsproduct[0]["GUID"].ToString()));
            //                //rbSql.Text += temp + Environment.NewLine;
            //                sqlLs.Add(temp);
            //            }
            //            dgError3.DataSource = dt_error3;
            //            dgRepet3.DataSource = dt_repet3;
            //            if (dt_error3.Rows.Count > 0 || dt_repet3.Rows.Count > 0)
            //            {
            //                Main.SetErrorCell(dgError3, col_error3);
            //                rbSql.Text = "";
            //                return;
            //            }
            //            #endregion

            //            #region PM巡检 验证
            //            List<int[]> col_error4 = new List<int[]>();
            //            DataTable dt_repet4 = _ECInfoD_excel.Clone();
            //            DataTable dt_error4 = _ECInfoD_excel.Clone(); 
            //            for (int i = 0; i < _ECInfoD_excel.Rows.Count; i++)
            //            {
            //                bool isError = false;
            //                bool isRepet = false;

            //                DataRow dr_excel = _ECInfoD_excel.Rows[i];

            //                int SNO = 0;
            //                if (string.IsNullOrWhiteSpace(dr_excel["序号"].ToString())
            //                    || !int.TryParse(dr_excel["序号"].ToString(), out SNO))
            //                {
            //                    //空
            //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 0 });
            //                    isError = true;
            //                }

            //                if (string.IsNullOrWhiteSpace(dr_excel["设备编号"].ToString()))
            //                {
            //                    //空
            //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 1 });
            //                    isError = true;
            //                }
            //                else if (!newEC.ContainsKey(dr_excel["设备编号"].ToString()))
            //                {
            //                    // 不存在
            //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 1 });
            //                    isError = true;
            //                }

            //                if (string.IsNullOrWhiteSpace(dr_excel["巡检项"].ToString()))
            //                {
            //                    //空
            //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 2 });
            //                    isError = true;
            //                }

            //                if (string.IsNullOrWhiteSpace(dr_excel["类型"].ToString())
            //                   || !CTYPEs.ContainsKey(dr_excel["类型"].ToString()))
            //                {
            //                    //空、不包含
            //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 5 });
            //                    isError = true;
            //                }

            //                if (string.IsNullOrWhiteSpace(dr_excel["标准"].ToString()) && dr_excel["类型"].ToString() == "文本")
            //                {
            //                    //空
            //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 6 });
            //                    isError = true;
            //                }

            //                DataRow[] drs_dept_pos_a = _hrDeptPos_DB.Select("BSDEPT_CODE='"
            //                                 + WGHelper.ReturnString(dr_excel["处理部门编号"].ToString())
            //                                 + "' and BSPOSITION_CODE='" + WGHelper.ReturnString(dr_excel["处理岗位编号"].ToString()) + "'");
            //                if (string.IsNullOrWhiteSpace(dr_excel["处理部门编号"].ToString())
            //                    || string.IsNullOrWhiteSpace(dr_excel["处理岗位编号"].ToString())
            //                    || drs_dept_pos_a.Length == 0)
            //                {
            //                    //空、不存在
            //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 7 });
            //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 8 });
            //                    isError = true;
            //                }

            //                DataRow[] drs_dept_pos_b = _hrDeptPos_DB.Select("BSDEPT_CODE='"
            //                 + WGHelper.ReturnString(dr_excel["响应部门编号"].ToString())
            //                 + "' and BSPOSITION_CODE='" + WGHelper.ReturnString(dr_excel["响应岗位编号"].ToString()) + "'");
            //                if (string.IsNullOrWhiteSpace(dr_excel["响应部门编号"].ToString())
            //                    || string.IsNullOrWhiteSpace(dr_excel["响应岗位编号"].ToString())
            //                    || drs_dept_pos_b.Length == 0)
            //                {
            //                    //空、不存在
            //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 9 });
            //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 10 });
            //                    isError = true;
            //                }

            //                if (dr_excel["类型"].ToString() == "数值")
            //                {
            //                    if (string.IsNullOrWhiteSpace(dr_excel["是否区间"].ToString())
            //                        || !ISINTERVALs.ContainsKey(dr_excel["是否区间"].ToString()))
            //                    {
            //                        //空、不包含
            //                        col_error4.Add(new int[] { dt_error4.Rows.Count, 12 });
            //                        isError = true;
            //                    }
            //                }


            //                int DAYS = 0;
            //                if (string.IsNullOrWhiteSpace(dr_excel["周期(天)"].ToString())
            //                    && !int.TryParse(dr_excel["周期(天)"].ToString(), out DAYS))
            //                {
            //                    //空
            //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 17 });
            //                    isError = true;
            //                }

            //                string Stand_Str = dr_excel["符号1"].ToString() + dr_excel["最小值"].ToString() + " " + dr_excel["符号2"].ToString() + dr_excel["最大值"].ToString();

            //                if (dr_excel["类型"].ToString() == "数值")
            //                {
            //                    dr_excel["标准"] = Stand_Str;
            //                }
            //                if (Stand_Str.Contains(dr_excel["标准"].ToString()) || dr_excel["类型"].ToString() == "文本")
            //                {

            //                }
            //                else
            //                {
            //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 7 });
            //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 13 });
            //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 14 });
            //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 15 });
            //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 16 });
            //                    isError = true;
            //                }

            //                if ((dr_excel["是否区间"].ToString() == "是" && dr_excel["符号1"].ToString() == "=") || ((dr_excel["是否区间"].ToString() == "否" && dr_excel["符号1"].ToString() != "=")))
            //                {
            //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 12 });
            //                    col_error4.Add(new int[] { dt_error4.Rows.Count, 15 });
            //                    isError = true;
            //                }


            //                DataRow[] drs_repit = _ECInfoD_excel.Select("设备编号 = '" + WGHelper.ReturnString(dr_excel["设备编号"].ToString()) + "' AND 巡检项 = '" + WGHelper.ReturnString(dr_excel["巡检项"].ToString()) + "'");

            //                if (drs_repit.Length > 1)
            //                {
            //                    isRepet = true;
            //                }

            //                if (isError || isRepet)
            //                {
            //                    if (isError)
            //                    {
            //                        dt_error4.Rows.Add(dr_excel.ItemArray);
            //                    }
            //                    if (isRepet)
            //                    {
            //                        dt_repet4.Rows.Add(dr_excel.ItemArray);
            //                    }
            //                    continue;
            //                }

            //                string temp = @"INSERT INTO [dbo].[ECINFOD]
            //                                       ([GUID],[PGUID],[SNO],[NAME],[METHOD]
            //                                       ,[REQUEST],[CTYPE],[SVALUE],[ISINTERVAL],[MINVALUE]
            //                                       ,[MAXVALUE],[SG1],[SG2],[DAYS],[AGUID]
            //                                       ,[BGUID])
            //                                 VALUES
            //                                       (" + Main.SetDBValue(Guid.NewGuid()) + "," + Main.SetDBValue(newEC[dr_excel["设备编号"].ToString()]) + "," + Main.SetDBValue(dr_excel["序号"]) + "," + Main.SetDBValue(dr_excel["巡检项"]) + "," + Main.SetDBValue(dr_excel["方法"]) + @"
            //                                        ," + Main.SetDBValue(dr_excel["要求"]) + "," + Main.SetDBValue(CTYPEs[dr_excel["类型"].ToString()]) + "," + Main.SetDBValue(dr_excel["标准"]) + "," + (GetISINTERVAL(dr_excel["是否区间"].ToString())) + "," + Main.SetDBValue(dr_excel["最小值"]) + @"
            //                                        ," + Main.SetDBValue(dr_excel["最大值"]) + "," + Main.SetDBValue(dr_excel["符号1"]) + "," + Main.SetDBValue(dr_excel["符号2"]) + "," + Main.SetDBValue(dr_excel["周期(天)"]) + "," + Main.SetDBValue(drs_dept_pos_a[0]["BSDEPTPOS_GUID"]) + @"
            //                                        ," + Main.SetDBValue(drs_dept_pos_b[0]["BSDEPTPOS_GUID"]) + ")";
            //                //rbSql.Text += temp + Environment.NewLine;
            //                sqlLs.Add(temp);
            //            }
            //            dgError4.DataSource = dt_error4;
            //            dgRepet4.DataSource = dt_repet4;
            //            if (dt_error4.Rows.Count > 0 || dt_repet4.Rows.Count > 0)
            //            {
            //                Main.SetErrorCell(dgError4, col_error4);
            //                rbSql.Text = "";
            //                return;
            //            }
            //            #endregion

            //            #region PM保养 验证
            //            List<int[]> col_error5 = new List<int[]>();
            //            DataTable dt_repet5 = _ECInfoE_excel.Clone();
            //            DataTable dt_error5 = _ECInfoE_excel.Clone(); 
            //            for (int i = 0; i < _ECInfoE_excel.Rows.Count; i++)
            //            {
            //                bool isError = false;
            //                bool isRepet = false;

            //                DataRow dr_excel = _ECInfoE_excel.Rows[i];


            //                int SNO = 0;
            //                if (string.IsNullOrWhiteSpace(dr_excel["序号"].ToString())
            //                    || !int.TryParse(dr_excel["序号"].ToString(), out SNO))
            //                {
            //                    //空
            //                    col_error5.Add(new int[] { dt_error5.Rows.Count, 0 });
            //                    isError = true;
            //                }

            //                if (string.IsNullOrWhiteSpace(dr_excel["设备编号"].ToString()))
            //                {
            //                    //空
            //                    col_error5.Add(new int[] { dt_error5.Rows.Count, 1 });
            //                    isError = true;
            //                }
            //                else if (!newEC.ContainsKey(dr_excel["设备编号"].ToString()))
            //                {
            //                    // 不存在
            //                    col_error5.Add(new int[] { dt_error5.Rows.Count, 1 });
            //                    isError = true;
            //                }

            //                if (string.IsNullOrWhiteSpace(dr_excel["保养项"].ToString()))
            //                {
            //                    //空
            //                    col_error5.Add(new int[] { dt_error5.Rows.Count, 2 });
            //                    isError = true;
            //                }

            //                if (string.IsNullOrWhiteSpace(dr_excel["类型"].ToString())
            //                   || !CTYPEs.ContainsKey(dr_excel["类型"].ToString()))
            //                {
            //                    //空、不包含
            //                    col_error5.Add(new int[] { dt_error5.Rows.Count, 5 });
            //                    isError = true;
            //                }

            //                if (string.IsNullOrWhiteSpace(dr_excel["标准"].ToString()) && dr_excel["类型"].ToString() == "文本")
            //                {
            //                    //空
            //                    col_error5.Add(new int[] { dt_error5.Rows.Count, 6 });
            //                    isError = true;
            //                }

            //                DataRow[] drs_dept_pos_a = _hrDeptPos_DB.Select("BSDEPT_CODE='"
            //                                                 + WGHelper.ReturnString(dr_excel["处理部门编号"].ToString())
            //                                                 + "' and BSPOSITION_CODE='" + WGHelper.ReturnString(dr_excel["处理岗位编号"].ToString()) + "'");
            //                if (string.IsNullOrWhiteSpace(dr_excel["处理部门编号"].ToString())
            //                    || string.IsNullOrWhiteSpace(dr_excel["处理岗位编号"].ToString())
            //                    || drs_dept_pos_a.Length == 0)
            //                {
            //                    //空、不存在
            //                    col_error5.Add(new int[] { dt_error5.Rows.Count, 8 });
            //                    col_error5.Add(new int[] { dt_error5.Rows.Count, 9 });
            //                    isError = true;
            //                }

            //                DataRow[] drs_dept_pos_b = _hrDeptPos_DB.Select("BSDEPT_CODE='"
            //                                                  + WGHelper.ReturnString(dr_excel["响应部门编号"].ToString())
            //                                                  + "' and BSPOSITION_CODE='" + WGHelper.ReturnString(dr_excel["响应岗位编号"].ToString()) + "'");
            //                if (string.IsNullOrWhiteSpace(dr_excel["响应部门编号"].ToString())
            //                    || string.IsNullOrWhiteSpace(dr_excel["响应岗位编号"].ToString())
            //                    || drs_dept_pos_b.Length == 0)
            //                {
            //                    //空、不存在
            //                    col_error5.Add(new int[] { dt_error5.Rows.Count, 10 });
            //                    col_error5.Add(new int[] { dt_error5.Rows.Count, 11 });
            //                    isError = true;
            //                }

            //                int DAYS = 0;
            //                if (string.IsNullOrWhiteSpace(dr_excel["周期(天)"].ToString())
            //                    && !int.TryParse(dr_excel["周期(天)"].ToString(), out DAYS))
            //                {
            //                    //空
            //                    col_error5.Add(new int[] { dt_error5.Rows.Count, 16 });
            //                    isError = true;
            //                }

            //                string Stand_Str = dr_excel["符号1"].ToString() + dr_excel["最小值"].ToString() + " " + dr_excel["符号2"].ToString() + dr_excel["最大值"].ToString();
            //                if (dr_excel["类型"].ToString() == "数值")
            //                {
            //                    dr_excel["标准"] = Stand_Str;
            //                }

            //                //if (Stand_Str.Contains(dr_excel["标准"].ToString()) || dr_excel["类型"].ToString() == "文本")
            //                //{

            //                //}
            //                //else
            //                //{
            //                //    col_error5.Add(new int[] { dt_error5.Rows.Count, 7 });
            //                //    col_error5.Add(new int[] { dt_error5.Rows.Count, 13 });
            //                //    col_error5.Add(new int[] { dt_error5.Rows.Count, 14 });
            //                //    col_error5.Add(new int[] { dt_error5.Rows.Count, 15 });
            //                //    col_error5.Add(new int[] { dt_error5.Rows.Count, 16 });
            //                //    isError = true;
            //                //}

            //                if ((dr_excel["是否区间"].ToString() == "是" && dr_excel["符号1"].ToString() == "=") || ((dr_excel["是否区间"].ToString() == "否" && dr_excel["符号1"].ToString() != "=")))
            //                {
            //                    col_error5.Add(new int[] { dt_error5.Rows.Count, 12 });
            //                    col_error5.Add(new int[] { dt_error5.Rows.Count, 15 });
            //                    isError = true;
            //                }

            //                DataRow[] drs_repit = _ECInfoE_excel.Select("设备编号 = '" + WGHelper.ReturnString(dr_excel["设备编号"].ToString()) + "' AND 保养项 = '" + WGHelper.ReturnString(dr_excel["保养项"].ToString()) + "'");

            //                if (drs_repit.Length > 1)
            //                {
            //                    isRepet = true;
            //                }

            //                if (isError || isRepet)
            //                {
            //                    if (isError)
            //                    {
            //                        dt_error5.Rows.Add(dr_excel.ItemArray);
            //                    }
            //                    if (isRepet)
            //                    {
            //                        dt_repet5.Rows.Add(dr_excel.ItemArray);
            //                    }

            //                    continue;
            //                }

            //                string temp = @"INSERT INTO [dbo].[ECINFOE]
            //                                       ([GUID],[PGUID],[SNO],[NAME],[METHOD]
            //                                       ,[REQUEST],[CTYPE],[SVALUE],[ISINTERVAL],[MINVALUE]
            //                                       ,[MAXVALUE],[SG1],[SG2],[DAYS],[AGUID]
            //                                       ,[BGUID])
            //                                 VALUES
            //                                       (" + Main.SetDBValue(Guid.NewGuid()) + "," + Main.SetDBValue(newEC[dr_excel["设备编号"].ToString()]) + "," + Main.SetDBValue(dr_excel["序号"]) + "," + Main.SetDBValue(dr_excel["保养项"]) + "," + Main.SetDBValue(dr_excel["方法"]) + @"
            //                                        ," + Main.SetDBValue(dr_excel["要求"]) + "," + Main.SetDBValue(CTYPEs[dr_excel["类型"].ToString()]) + "," + Main.SetDBValue(dr_excel["标准"]) + "," + (GetISINTERVAL(dr_excel["是否区间"].ToString())) + "," + Main.SetDBValue(dr_excel["最小值"]) + @"
            //                                        ," + Main.SetDBValue(dr_excel["最大值"]) + "," + Main.SetDBValue(dr_excel["符号1"]) + "," + Main.SetDBValue(dr_excel["符号2"]) + "," + Main.SetDBValue(dr_excel["周期(天)"]) + "," + Main.SetDBValue(drs_dept_pos_a[0]["BSDEPTPOS_GUID"]) + @"
            //                                        ," + Main.SetDBValue(drs_dept_pos_b[0]["BSDEPTPOS_GUID"]) + ")";
            //                //rbSql.Text += temp + Environment.NewLine;
            //                sqlLs.Add(temp);
            //            }
            //            dgError5.DataSource = dt_error5;
            //            dgRepet5.DataSource = dt_repet5;
            //            if (dt_error5.Rows.Count > 0 || dt_repet5.Rows .Count > 0)
            //            {
            //                Main.SetErrorCell(dgError5, col_error5);
            //                rbSql.Text = "";
            //                return;
            //            }
            //            #endregion

            //            #region 技术规格 验证
            ////            dt_error = _ECInfoA_excel.Clone();
            ////            for (int i = 0; i < _ECInfoA_excel.Rows.Count; i++)
            ////            {
            ////                bool isError = false;
            ////                bool isRepet = false;

            ////                DataRow dr_excel = _ECInfoA_excel.Rows[i];
            ////                if (string.IsNullOrWhiteSpace(dr_excel["设备厂内编号"].ToString()))
            ////                {
            ////                    //空
            ////                    col_error.Add(new int[] { dt_error.Rows.Count, 0 });
            ////                    isError = true;
            ////                }
            ////                else if (!newEC.ContainsKey(dr_excel["设备厂内编号"].ToString()))
            ////                {
            ////                    // 存在
            ////                    isRepet = true;
            ////                }

            ////                int SNO = 0;
            ////                if (string.IsNullOrWhiteSpace(dr_excel["序号"].ToString())
            ////                    || !int.TryParse(dr_excel["序号"].ToString(), out SNO))
            ////                {
            ////                    //空
            ////                    col_error.Add(new int[] { dt_error.Rows.Count, 1 });
            ////                    isError = true;
            ////                }

            ////                if (string.IsNullOrWhiteSpace(dr_excel["属性"].ToString()))
            ////                {
            ////                    //空
            ////                    col_error.Add(new int[] { dt_error.Rows.Count, 2 });
            ////                    isError = true;
            ////                }

            ////                if (string.IsNullOrWhiteSpace(dr_excel["值"].ToString()))
            ////                {
            ////                    //空
            ////                    col_error.Add(new int[] { dt_error.Rows.Count, 3 });
            ////                    isError = true;
            ////                }

            ////                if (isError || isRepet)
            ////                {
            ////                    if (isError)
            ////                    {
            ////                        dt_error.Rows.Add(dr_excel.ItemArray);
            ////                    }
            ////                    continue;
            ////                }
            ////                string temp = @"INSERT INTO [dbo].[ECINFOA]
            ////                                       ([GUID],[PGUID],[SNO],[PROPERTY],[VALUE]
            ////                                       ,[DES])
            ////                                 VALUES
            ////                                       (" + Main.SetDBValue(Guid.NewGuid()) + "," + Main.SetDBValue(newEC[dr_excel["设备厂内编号"].ToString()]) + "," + Main.SetDBValue(dr_excel["序号"]) + "," + Main.SetDBValue(dr_excel["属性"]) + "," + Main.SetDBValue(dr_excel["值"]) + @"
            ////                                        ," + Main.SetDBValue(dr_excel["描述"]) + ")";
            ////                rbSql.Text += temp + Environment.NewLine;
            ////                sqlLs.Add(temp);
            ////            }
            ////            dgError.DataSource = dt_error;
            ////            if (dt_error.Rows.Count > 0)
            ////            {
            ////                Main.SetErrorCell(dgError, col_error);
            ////                rbSql.Text = "";
            ////                return;
            ////            }
            //            #endregion
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
