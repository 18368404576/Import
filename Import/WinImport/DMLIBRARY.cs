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
using System.IO;
using System.Collections;
using System.Text.RegularExpressions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace WinImport
{
    public partial class DMLIBRARY : Form
    {
        ExcelManager _excelManager = new ExcelManager();

        DataTable _DMLIBRARY_excel = new DataTable();

        DataTable _DMLIBRARY_DB = null;

        public DataTable _DMLINK_excel = new DataTable();

        DataTable _WORKSHOP_DB = null;

        DataTable _WORKCENTER_DB = null;

        DataTable _ECINFO_DB = null;

        DataTable _DMTYPE_DB = null;

        DataTable _HREMPLOYEE_DB = null;

        DataTable _DEPT_DB = null;

        DataTable _BSPRODUCT_DB = null;

        private DataTable dtout = null;
        string[] DOC;

        /// <summary>
        /// 需要保存的sql
        /// </summary>
        List<string> sqlLs = new List<string>();

        /// <summary>
        ///  是否验证成功
        /// </summary>
        bool isCheck = false;

        bool isError = false;
        bool isRepet = false;

        StringBuilder sb = new StringBuilder();

        public DMLIBRARY()
        {
            InitializeComponent();
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
            dgRepet2.DataSource = new DataTable();
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            //得到excel数据源
            if (opfDialog0.ShowDialog() == DialogResult.OK)
            {
                txtFileMain.Text = opfDialog0.FileName;
                _DMLIBRARY_excel = _excelManager.GetExcelTableByOleDB(opfDialog0.FileName, "DMLIBRARY");

                _DMLIBRARY_excel.Columns.Add("MGUID");
                _DMLIBRARY_excel.Columns.Add("LGUID");
                _DMLIBRARY_excel.Columns.Add("FKEY");
                _DMLIBRARY_excel.Columns.Add("FSIZE");
                isCheck = false;

                MessageBox.Show("读取笔数：" + _DMLIBRARY_excel.Rows.Count + "");
                //ClearSql();
                if (_DMLIBRARY_excel == null || _DMLIBRARY_excel.Rows.Count <= 0)
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

        private void button3_Click(object sender, EventArgs e)
        {
            //得到excel数据源
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtFileSon.Text = openFileDialog1.FileName;
                _DMLINK_excel = _excelManager.GetExcelTableByOleDB(openFileDialog1.FileName, "DMLINK");
                //ClearSql();

                MessageBox.Show("读取笔数：" + _DMLINK_excel.Rows.Count + "");
                isCheck = false;
                if (_DMLINK_excel == null || _DMLINK_excel.Rows.Count <= 0)
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


        private void btnCheck_Click(object sender, EventArgs e)
        {
            sb.Clear();
            if (isCheck)
            {
                WGMessage.ShowAsterisk("已验证，不用重复验证！");
                return;
            }
            rbSql.Text = "";
            sqlLs = new List<string>();

            if (_DMLIBRARY_excel == null)
            {
                WGMessage.ShowWarning("请选择主文件!");
                return;
            }
            if (_DMLINK_excel == null)
            {
                WGMessage.ShowWarning("请选择子文件!");
                return;
            }

            //if (DOC == null)
            //{
            //    WGMessage.ShowWarning("请选择含PDF的文件夹!");
            //    return;
            //}

            //if (DOC.Length == 0)
            //{
            //    WGMessage.ShowWarning("请选择含PDF的文件夹!");
            //    return;
            //}

            //if (txtFileAim.Text == "")
            //{
            //    WGMessage.ShowWarning("请选择目的地文件夹!");
            //    return;
            //}

            string sql = @"SELECT * FROM BSWORKSHOP";
            _WORKSHOP_DB = FillDatatablde(sql, Main.CONN_Public);

            sql = @"SELECT * FROM BSWORKCENTERS";
            _ECINFO_DB = FillDatatablde(sql, Main.CONN_Public);

            sql = @"SELECT * FROM BSWORKCENTER";
            _WORKCENTER_DB = FillDatatablde(sql, Main.CONN_Public);

            sql = @"SELECT GUID,EMPCODE FROM HREMPLOYEE";
            _HREMPLOYEE_DB = FillDatatablde(sql, Main.CONN_Public);

            sql = @"SELECT GUID,CODE FROM BSDEPT";
            _DEPT_DB = FillDatatablde(sql, Main.CONN_Public);

            sql = @"SELECT * FROM DMTYPE";
            _DMTYPE_DB = FillDatatablde(sql, Main.CONN_Public);

            sql = @"SELECT * FROM DMLIBRARY
LEFT JOIN DMVERSION ON DMVERSION.PGUID = DMLIBRARY.GUID";
            _DMLIBRARY_DB = FillDatatablde(sql, Main.CONN_Public);

            sql = @"SELECT BSPRODUCT.GUID,[BSPRODUCT].[CODE],[dbo].[BSPRODUCTVER].[VER] FROM [dbo].[BSPRODUCT]
LEFT JOIN BSPRODUCTVER ON BSPRODUCTVER.[PGUID] = [BSPRODUCT].[GUID]";
            _BSPRODUCT_DB = FillDatatablde(sql,Main.CONN_Public);

            //错误
            List<int[]> col_error1 = new List<int[]>();
            List<int[]> col_error2 = new List<int[]>();
            DataTable dt_error1 = _DMLIBRARY_excel.Clone();

            //重复数据
            DataTable dt_repet1 = _DMLIBRARY_excel.Clone();
            DataTable dt_repet1_excel = _DMLIBRARY_excel.Clone();

            DataTable dt_error2 = _DMLINK_excel.Clone();

            //重复数据
            DataTable dt_repet2 = _DMLINK_excel.Clone();


            //需要保持的数据
            DataTable dt = _DMLIBRARY_excel.Clone();
            dt.Columns.Add();

            Dictionary<string, string> doclist = new Dictionary<string, string>();

            dtout = new DataTable();
            dtout.Columns.Add("文件名");
            dtout.Columns.Add("文件KEY");
            dtout.Columns.Add("受控类型");
            dtout.Columns.Add("有效天数");
            dtout.Columns.Add("文档编号");

            #region 主子

            for (int i = 0; i < _DMLIBRARY_excel.Rows.Count; i++)
            {
                bool isError = false;
                bool isRepet = false;
                bool isRepet_excel = false;

                DataRow dr_excel = _DMLIBRARY_excel.Rows[i];

                doclist.Add(Guid.NewGuid().ToString(), dr_excel["文档名"].ToString());

                if (string.IsNullOrWhiteSpace(dr_excel["文档编号"].ToString())
                    || _DMLIBRARY_excel.Select("文档编号='" + WGHelper.ReturnString(dr_excel["文档编号"].ToString()) + "'").Length > 1)
                {
                    //空、重复
                    isRepet_excel = true;
                }

                if (_DMLIBRARY_DB.Select("CODE='" + WGHelper.ReturnString(dr_excel["文档编号"].ToString()) + "'").Length > 0)
                {
                    // 存在
                    isRepet = true;
                }

                if (string.IsNullOrWhiteSpace(dr_excel["文档名"].ToString()))
                {
                    //空
                    col_error1.Add(new int[] { dt_error1.Rows.Count, 2 });
                    isError = true;
                }


                DataRow[] drs_DMTYPE_pos = _DMTYPE_DB.Select("CODE='"
                                                  + WGHelper.ReturnString(dr_excel["文档类别编号"].ToString()) + "'");
                if (string.IsNullOrWhiteSpace(dr_excel["文档类别编号"].ToString())
                    || drs_DMTYPE_pos.Length == 0)
                {
                    //空、不存在
                    col_error1.Add(new int[] { dt_error1.Rows.Count, 0 });
                    isError = true;
                }

                //DataRow[] drs_DMCATALOG_pos = _DMCATALOG_DB.Select("NAME='"
                //                                  + WGHelper.ReturnString(dr_excel["文档目录名称"].ToString()) + "'");
                //if (string.IsNullOrWhiteSpace(dr_excel["文档目录名称"].ToString())
                //    || drs_DMCATALOG_pos.Length == 0)
                //{
                //    //空、不存在
                //    col_error1.Add(new int[] { dt_error1.Rows.Count, 3 });
                //    isError = true;
                //}

                DataRow[] drs_HREMPLOYEE_pos = _HREMPLOYEE_DB.Select("EMPCODE ='"
                                  + WGHelper.ReturnString(dr_excel["上传人工号"].ToString()) + "'");
                if (string.IsNullOrWhiteSpace(dr_excel["上传人工号"].ToString())
                    || drs_HREMPLOYEE_pos.Length == 0)
                {
                    //空、不存在
                    col_error1.Add(new int[] { dt_error1.Rows.Count, 4 });
                    isError = true;
                }


                DataRow[] drs_DEPT_pos = _DEPT_DB.Select("CODE ='"
                  + WGHelper.ReturnString(dr_excel["上传人所在部门编号"].ToString()) + "'");
                if (string.IsNullOrWhiteSpace(dr_excel["上传人所在部门编号"].ToString() + "'")
                    || drs_DEPT_pos.Length == 0)
                {
                    //空、不存在
                    col_error1.Add(new int[] { dt_error1.Rows.Count, 5 });
                    isError = true;
                }

                DateTime ADATE = new DateTime();
                if (string.IsNullOrWhiteSpace(dr_excel["上传时间"].ToString())
                    || !DateTime.TryParse(dr_excel["上传时间"].ToString(), out ADATE))
                {
                    //空、类型不符
                    col_error1.Add(new int[] { dt_error1.Rows.Count, 6 });
                    isError = true;
                }

                if (dr_excel["受控类型"].ToString() != "1" && dr_excel["受控类型"].ToString() != "2")
                {
                    col_error1.Add(new int[] { dt_error1.Rows.Count, 7 });
                    isError = true;
                }

                DateTime VDATE = new DateTime();


                int day = 0;

                if (dr_excel["受控类型"].ToString() == "2")
                {
                    if (string.IsNullOrWhiteSpace(dr_excel["有效期"].ToString())
                        || !DateTime.TryParse(dr_excel["有效期"].ToString(), out VDATE))
                    {
                        //空、类型不符
                        col_error1.Add(new int[] {dt_error1.Rows.Count, 9});
                        isError = true;
                    }

                    if (!int.TryParse(dr_excel["有效天数"].ToString(),out day))
                    {
                        col_error1.Add(new int[] { dt_error1.Rows.Count, 8 });
                        isError = true;
                    }
                }
                else if (dr_excel["受控类型"].ToString() == "1")
                {
                    dr_excel["有效天数"] = 0;
                    dr_excel["有效期"] = "";
                }

                if (string.IsNullOrWhiteSpace(dr_excel["正式版本号"].ToString()))
                {
                    //空
                    col_error1.Add(new int[] { dt_error1.Rows.Count, 10 });
                    isError = true;
                }

                if (dr_excel["文档来源"].ToString() != "1" && dr_excel["文档来源"].ToString() != "3")
                {
                    col_error1.Add(new int[] { dt_error1.Rows.Count, 11 });
                    isError = true;
                }
                DateTime ndTime = DateTime.Now;
                DateTime cdTime = ndTime;

                DateTime.TryParse(dr_excel["创建时间"].ToString(),out ndTime);
                DateTime.TryParse(dr_excel["修改时间"].ToString(), out cdTime);

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
                //bool ishavedoc = false;

                //FileInfo dr_info = new FileInfo (DOC[0]);

                //for (int j = 0; j < DOC.Length; j++)
                //{
                //    FileInfo info = new FileInfo(DOC[j]);

                //    if (info.Name.Contains(".pdf") && info.Name == dr_excel["文档名"].ToString())
                //    {
                //        ishavedoc = true;
                //        dr_info = info;
                //        break;
                //    }
                //}

                //if (ishavedoc == false)
                //{
                //    WGMessage.ShowAsterisk("选择文件夹中不包含"+dr_excel["文档名"]+"!");
                //    return;
                //}

                string DMLIBRARY_GUID = Guid.NewGuid().ToString();
                string DMVERSION_GUID = Guid.NewGuid().ToString();

                _DMLIBRARY_excel.Rows[i]["MGUID"] = DMLIBRARY_GUID;
                _DMLIBRARY_excel.Rows[i]["LGUID"] = DMVERSION_GUID;

                try
                {
                    string vdate = "";

                    if (dr_excel["受控类型"].ToString() == "1")
                    {
                        vdate = "null";
                    }
                    else
                    {
                        vdate = "'" + dr_excel["有效期"] + "'";
                    }
                    string temp = @"INSERT INTO [DMLIBRARY]
                                       ([GUID],[FGUID],[CODE]
                                       ,[NAME],[KEYWORD],[AGUID]
                                       ,[DEPTRI],[ADATE],[VDATE],[VER],[ST]
                                       ,[NOTE],[CC],[ND],[CD],[CTYPE]
                                       ,[VDAYS],[FTYPE])
                                 VALUES
                                       (" + Main.SetDBValue(DMLIBRARY_GUID) + "," + Main.SetDBValue(drs_DMTYPE_pos[0]["GUID"]) + "," + Main.SetDBValue(dr_excel["文档编号"].ToString()) + @"
                                        ," + Main.SetDBValue(dr_excel["文档名"].ToString()) + "," + Main.SetDBValue(dr_excel["关键字"]) + "," + Main.SetDBValue(drs_HREMPLOYEE_pos[0]["GUID"]) + @"
                                        ," + Main.SetDBValue(Guid.NewGuid().ToString()/*drs_DEPT_pos[0]["GUID"]*/) + "," + Main.SetDBValue(dr_excel["上传时间"]) + "," + (vdate) + "," + Main.SetDBValue(dr_excel["正式版本号"]) + ",2" + @"
                                        ," + Main.SetDBValue(dr_excel["备注"].ToString()) + ",''," + Main.SetDBValue(ndTime) + "," + Main.SetDBValue(cdTime) + ",'" + dr_excel["受控类型"].ToString() + "'" + "" +
                                        "," + day + ",'" + dr_excel["文档来源"].ToString() + "');";


                    string FKEY = Guid.NewGuid().ToString();

                    dr_excel["FKEY"] = FKEY;
                    temp += @" INSERT INTO DMVERSION (GUID,PGUID,VER,FKEY,AGUID,ADATE,ST) VALUES
                        ('" + DMVERSION_GUID + "','" + DMLIBRARY_GUID + "'," + Main.SetDBValue(dr_excel["正式版本号"]) + ",'" + FKEY + "'," + Main.SetDBValue(drs_HREMPLOYEE_pos[0]["GUID"]) + "," + Main.SetDBValue(dr_excel["上传时间"]) + ",2);";

                    sb.Append(temp + Environment.NewLine);
                    sqlLs.Add(temp);

                    DataRow drnew = dtout.NewRow();
                    drnew["文件名"] = dr_excel["文档名"].ToString();
                    drnew["文件KEY"] = FKEY;
                    drnew["受控类型"] = dr_excel["受控类型"].ToString();
                    drnew["有效天数"] = day;
                    drnew["文档编号"] = dr_excel["文档编号"].ToString();
                    dtout.Rows.Add(drnew);
                }
                catch
                { }
            }


            #endregion

            #region 孙孙


            for (int i = 0; i < _DMLINK_excel.Rows.Count; i++)
            {
                isError = false;
                isRepet = false;

                DataRow dr_excel = _DMLINK_excel.Rows[i];

                if (string.IsNullOrWhiteSpace(dr_excel["文档编号"].ToString()))
                {

                    col_error2.Add(new int[] { dt_error2.Rows.Count, 0 });
                    isError = true;
                }

                if (_DMLINK_excel.Select("文档编号='" + WGHelper.ReturnString(dr_excel["文档编号"].ToString()) + "' AND 正式版本号 = '" + WGHelper.ReturnString(dr_excel["正式版本号"].ToString()) + "' and 车间编号 = '"+WGHelper.ReturnString(dr_excel["车间编号"].ToString())+"' AND 工作中心编号 = '"+WGHelper.ReturnString(dr_excel["工作中心编号"].ToString())+"' AND 工位编号 = '"+WGHelper.ReturnString(dr_excel["工位编号"].ToString())+"'").Length > 1)
                {
                    // 存在
                    isRepet = true;
                }


                if (_DMLINK_excel.Select("文档编号='" + WGHelper.ReturnString(dr_excel["文档编号"].ToString()) + "' AND 正式版本号 = '" + WGHelper.ReturnString(dr_excel["正式版本号"].ToString()) + "' AND 产品编号 = '"+WGHelper.ReturnString(dr_excel["产品编号"].ToString())+"' AND 产品版本 = '"+WGHelper.ReturnString(dr_excel["产品版本"].ToString())+"' AND 工序码 = '"+WGHelper.ReturnString(dr_excel["工序码"].ToString())+"'" ).Length > 1)
                {
                     isRepet = true;
                }

               /* DataRow[] drss = _DMLINK_excel.Select("文档编号 = '"+dr_excel["文档编号"].ToString()+ "' AND 正式版本号 = '"+ dr_excel["正式版本号"].ToString() + "'");

                if (drss.Length > 1)
                {
                    col_error2.Add(new int[] { dt_error2.Rows.Count, 0 });
                    col_error2.Add(new int[] { dt_error2.Rows.Count, 1 });
                    isError = true;
                }*/

                if (string.IsNullOrWhiteSpace(dr_excel["正式版本号"].ToString()))
                {
                    //空
                    col_error2.Add(new int[] { dt_error2.Rows.Count, 1 });
                    isError = true;
                }
                int numa;

                if (dr_excel["类型"].ToString() != "0" && dr_excel["类型"].ToString() != "1")
                {
                    col_error2.Add(new int[] { dt_error2.Rows.Count, 2 });
                    isError = true;
                }

                DataRow[] drs_WORKCENTER_pos = new DataRow[0];
                DataRow[] drs_WORKSHOP_pos = new DataRow[0];
                if (dr_excel["类型"].ToString() == "0")
                {
                    if (dr_excel["序号"].ToString() == "" && dr_excel["车间编号"].ToString() != "")
                    {
                        col_error2.Add(new int[] {dt_error2.Rows.Count, 3});
                        isError = true;
                    }
                    else if (dr_excel["序号"].ToString() == "" && dr_excel["车间编号"].ToString() == "")
                    {

                    }
                    else if (!int.TryParse(dr_excel["序号"].ToString(), out numa))
                    {
                        //空
                        col_error2.Add(new int[] {dt_error2.Rows.Count, 3});
                        isError = true;
                    }

                    drs_WORKSHOP_pos = _WORKSHOP_DB.Select("CODE='"
                                                                     +
                                                                     WGHelper.ReturnString(dr_excel["车间编号"].ToString()) +
                                                                     "'");

                    if (dr_excel["车间编号"].ToString() == "" && dr_excel["工作中心编号"].ToString() != "")
                    {
                        col_error2.Add(new int[] {dt_error2.Rows.Count, 4});
                        isError = true;
                    }
                    else if (dr_excel["车间编号"].ToString() == "" && dr_excel["工作中心编号"].ToString() == "")
                    {

                    }
                    else if (drs_WORKSHOP_pos.Length == 0)
                    {
                        //空、不存在
                        col_error2.Add(new int[] {dt_error2.Rows.Count, 4});
                        isError = true;
                    }

                    drs_WORKCENTER_pos = _WORKCENTER_DB.Select("CODE ='"
                                                                         +
                                                                         WGHelper.ReturnString(
                                                                             dr_excel["工作中心编号"].ToString()) + "'");

                    if (dr_excel["工作中心编号"].ToString() == "" && dr_excel["工位编号"].ToString() != "")
                    {
                        col_error2.Add(new int[] {dt_error2.Rows.Count, 5});
                        isError = true;
                    }
                    else if (dr_excel["工作中心编号"].ToString() == "" && dr_excel["工位编号"].ToString() == "")
                    {

                    }
                    else if (drs_WORKCENTER_pos.Length == 0)
                    {
                        col_error2.Add(new int[] {dt_error2.Rows.Count, 5});
                        isError = true;
                    }

                    if (dr_excel["工位编号"].ToString() != "" && drs_WORKCENTER_pos.Length == 0)
                    {
                        //空、不存在
                        col_error2.Add(new int[] {dt_error2.Rows.Count, 6});
                        isError = true;
                    }

                }

                string[] ECINFO = dr_excel["工位编号"].ToString().Split(',');
                string[] ECINFOGUID = new string[ECINFO.Length];
                string[] ECINFONAME = new string[ECINFO.Length];
                int index = 0;

                if (dr_excel["工位编号"].ToString() == "")
                {
                }
                else
                {
                    foreach (var ei in ECINFO)
                    {
                        if (drs_WORKCENTER_pos.Length == 0)
                        {
                            col_error2.Add(new int[] { dt_error2.Rows.Count, 6 });
                            isError = true;
                            break;
                        }

                        DataRow[] drs_ECINFO_pos = _ECINFO_DB.Select("CODE ='"
                                     + WGHelper.ReturnString(ECINFO[index].ToString()) + "' AND PGUID = '" + drs_WORKCENTER_pos [0]["GUID"].ToString()+ "'");
                        if (drs_ECINFO_pos.Length == 0)
                        {
                            //空、不存在
                            col_error2.Add(new int[] { dt_error2.Rows.Count, 6 });
                            isError = true;
                            break;
                        }
                        ECINFOGUID[index] = drs_ECINFO_pos[0]["GUID"].ToString();
                        ECINFONAME[index] = drs_ECINFO_pos[0]["NAME"].ToString();
                        index++;
                    }
                }

                DataRow[] drs_DMLINK =  new DataRow[0];
                if (dr_excel["类型"].ToString() == "1")
                {
                    if (dr_excel["产品编号"].ToString() == "")
                    {
                        col_error2.Add(new int[] { dt_error2.Rows.Count, 7 });
                        isError = true;
                    }

                    if (dr_excel["产品版本"].ToString() == "")
                    {
                        col_error2.Add(new int[] { dt_error2.Rows.Count, 8 });
                        isError = true;
                    }

                    if (dr_excel["工序码"].ToString() == "")
                    {
                        col_error2.Add(new int[] { dt_error2.Rows.Count, 9 });
                        isError = true;
                    }

                    drs_DMLINK = _BSPRODUCT_DB.Select("CODE = '" + WGHelper.ReturnString(dr_excel["产品编号"].ToString()) + "' AND VER = '"+dr_excel["产品版本"].ToString()+"' ");

                    if (drs_DMLINK.Length == 0)
                    {
                        col_error2.Add(new int[] { dt_error2.Rows.Count, 7 });
                        col_error2.Add(new int[] { dt_error2.Rows.Count, 8 });
                        isError = true;
                    }
                }

                DataRow[] drs = _DMLIBRARY_excel.Select("文档编号 = '" + dr_excel["文档编号"].ToString() + "' AND 正式版本号 = '" + dr_excel["正式版本号"].ToString() + "' ");

                if (drs.Length == 0)
                {
                    col_error2.Add(new int[] { dt_error2.Rows.Count, 0 });
                    col_error2.Add(new int[] { dt_error2.Rows.Count, 1 });
                    isError = true;
                }


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

                string DMLINK_GUID = Guid.NewGuid().ToString();

                string AGUID = Guid.NewGuid().ToString();
                string BGUID = Guid.NewGuid().ToString();

                try
                {
                    AGUID = drs_WORKSHOP_pos[0]["GUID"].ToString();
                }
                catch
                {

                }

                try
                {
                    BGUID = drs_WORKCENTER_pos[0]["GUID"].ToString();
                }
                catch
                {

                }

                if (dr_excel["序号"].ToString() != "")
                {
                    string temp = "";

                    if (dr_excel["类型"].ToString() == "0")
                    {
                        temp = @" INSERT INTO DMLINK (GUID,PGUID,SNO,AGUID,BGUID,CTYPE) VALUES
('" + DMLINK_GUID + "','" + drs[0]["LGUID"].ToString() + "','" + dr_excel["序号"].ToString() + "','" + AGUID +
                               "','" + BGUID + "','" + dr_excel["类型"] + "');";
                    }
                    else if (dr_excel["类型"].ToString() == "1")
                    {
                        temp = @" INSERT INTO DMLINK (GUID,PGUID,SNO,CTYPE,CGUID,PVER,CODE) VALUES
('" + DMLINK_GUID + "','" + drs[0]["LGUID"].ToString() + "','" + dr_excel["序号"].ToString() + "','" + dr_excel["类型"] + "','" + drs_DMLINK[0]["GUID"] + "','" +
                               dr_excel["产品版本"] + "','" + WGHelper.ReturnString(dr_excel["工序码"].ToString()) + "');";
                    }

                    for (int j = 0; j < ECINFO.Length; j++)
                    {
                        if (ECINFOGUID[j] == null)
                        {
                            continue;
                        }

                        temp += @" INSERT INTO DMLINKS (GUID,PGUID,SNO,FGUID) VALUES
('" + Guid.NewGuid().ToString() + "','" + DMLINK_GUID + "','" + (j + 1) + "','" + ECINFOGUID[j] + "')";
                    }
                    sb.Append(temp + Environment.NewLine);
                    //rbSql.Text += temp + Environment.NewLine;
                    sqlLs.Add(temp);
                }
            }

            #endregion

            #region （BSFILEUPLOAD）
            for (int i = 0; i < _DMLIBRARY_excel.Rows.Count; i++)
            {
                string temp = @"INSERT INTO BSFILEUPLOAD (GUID,DOMAIN,TBDM,PGUID,PKEY,FNAME,SAVENAME,EXTNAME,ST)
            VALUES ('" + Guid.NewGuid().ToString() + "','DMLIBRARY','DMVERSION','" + _DMLIBRARY_excel.Rows[i]["LGUID"].ToString() + "','" + _DMLIBRARY_excel.Rows[i]["FKEY"].ToString() + "','" + _DMLIBRARY_excel.Rows[i]["文档名"] + "','" + _DMLIBRARY_excel.Rows[i]["FKEY"].ToString() + ".pdf','.pdf',1);";

                sb.Append(temp + Environment.NewLine);
                //rbSql.Text += temp + Environment.NewLine;
                sqlLs.Add(temp);
            }

            #endregion
            dt_error1.Columns.Remove("MGUID"); 
            dt_error1.Columns.Remove("LGUID");
            dt_error1.Columns.Remove("FKEY");
            dt_error1.Columns.Remove("FSIZE");
            dt_repet1.Columns.Remove("MGUID");
            dt_repet1.Columns.Remove("LGUID");
            dt_repet1.Columns.Remove("FKEY");
            dt_repet1.Columns.Remove("FSIZE");

            //dt_error2.Columns.Remove("F7");
            //dt_error2.Columns.Remove("F8");
            //dt_error2.Columns.Remove("F9");
            //dt_error2.Columns.Remove("F10");
            //dt_repet2.Columns.Remove("F7");
            //dt_repet2.Columns.Remove("F8");
            //dt_repet2.Columns.Remove("F9");
            //dt_repet2.Columns.Remove("F10");

            dgError1.DataSource = dt_error1; dgError2.DataSource = dt_error2;
            dgRepet1_excel.DataSource = dt_repet1_excel; dgRepet2.DataSource = dt_repet2;
            dgRepet1.DataSource = dt_repet1;
            if (dt_error1.Rows.Count > 0 || dt_error2.Rows.Count >0 || dt_repet1_excel.Rows.Count >0  || dt_repet1.Rows.Count > 0 || dt_repet2.Rows.Count > 0)
            {
                Main.SetErrorCell(dgError1, col_error1);
                Main.SetErrorCell(dgError2, col_error2);
                rbSql.Text = "";
                isCheck = false;
                return;
            }
            this.rbSql.Text = sb.ToString();
            isCheck = true;
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

        private void btnImport_Click(object sender, EventArgs e)
        {
            if (!isCheck)
            {
                WGMessage.ShowAsterisk("还未验证，不能导入！");
                return;
            }
            if (isError)
            {
                WGMessage.ShowAsterisk("有误不得导入！");
                return;
            }
            if (isRepet)
            {
                WGMessage.ShowAsterisk("有重复不得导入！");
                return;
            }

            string result = "";
            if (RunSql(sqlLs, Main.CONN_Public))
            {
                result += "导入成功！";
                ClearSql();
            }
            else
            {
                result += "导入失败！";
            }

            //if (UploadDoc())
            //{
            //    result += "上传成功！";
            //}
            //else
            //{
            //    result += "上传失败！";
            //}

            WGMessage.ShowAsterisk(result);
            Dictionary<string, string> _dicDataFieldCaption = new Dictionary<string, string>();
            ExportToExcel("文档清单.xlsx", "文档清单", _dicDataFieldCaption, dtout);
            ClearSql();
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
                                byte[] bytes = ConvertImage(image);
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
                                byte[] bytes = ConvertImage(image);
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

                        //#region 表头及样式
                        //{
                        //    if (dtSource.TableName.IndexOf("Table", StringComparison.Ordinal) != 0)
                        //    {
                        //        IRow headerRow = sheet.CreateRow(rowIndex);
                        //        headerRow.HeightInPoints = dtSource.TableName.Length != new Regex(" ").Replace(dtSource.TableName, "\n", 1).Length ? 25 : 50;
                        //        headerRow.CreateCell(0).SetCellValue(new Regex(" ").Replace(dtSource.TableName, "\n", 1));

                        //        ICellStyle headStyle = workbook.CreateCellStyle();
                        //        headStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CENTER;
                        //        IFont font = workbook.CreateFont();
                        //        font.FontHeightInPoints = 20;
                        //        font.Boldweight = 700;
                        //        //font.Color = HSSFColor.WHITE.index;
                        //        headStyle.SetFont(font);
                        //        headStyle.WrapText = true;
                        //        //headStyle.FillForegroundColor = GetXLColour(workbook, AppConfig.ZhongTaiLightRed);
                        //        //headStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;
                        //        headerRow.GetCell(0).CellStyle = headStyle;
                        //        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(rowIndex, rowIndex, 0, dtSource.Columns.Count - 1 < 0 ? 0 : dtSource.Columns.Count - 1));
                        //    }
                        //}
                        //#endregion

                        rowIndex = -1;
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

        public static byte[] ConvertImage(Image image)
        {
            byte[] bt = null;
            if (image.Equals(null)) return null;
            using (MemoryStream mostream = new MemoryStream())
            {
                Bitmap bmp = new Bitmap(image);
                bmp.Save(mostream, System.Drawing.Imaging.ImageFormat.Jpeg);//将图像以指定的格式存入缓存内存流
                bt = new byte[mostream.Length];
                mostream.Position = 0;//设置留的初始位置
                mostream.Read(bt, 0, Convert.ToInt32(bt.Length));
            }
            return bt;
        }

        //public bool UploadDoc()
        //{
        //    for (int i = 0; i < _DMLIBRARY_excel.Rows.Count; i++)
        //    {
        //        for (int j = 0; j < DOC.Length; j++)
        //        {
        //            if (Path.GetFileName(DOC[j]) == _DMLIBRARY_excel.Rows[i]["文档名"].ToString())
        //           {                      
        //                FileInfo info = new FileInfo(DOC[j]);

        //                string path_name = txtFileAim.Text + "\\" + _DMLIBRARY_excel.Rows[i]["FKEY"].ToString() + @".pdf";
        //                File.Copy(DOC[j], path_name);
        //                File.SetLastWriteTime(path_name, DateTime.Now);
        //            }
        //        }
        //    }

        //    return true;
        //}

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

        private void DMLIBRARY_Load(object sender, EventArgs e)
        {

        }
    }
}
