using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace WinImport
{
    public class ExcelManager
    {

        /// <summary>
        /// Excel导入成Datable
        /// </summary>
        /// <param name="strExcelPath">导入路径(包含文件名与扩展名)</param>
        /// <param name="tableName">表名</param>
        /// <returns></returns>
        public DataTable GetExcelTableByOleDB(string strExcelPath, string tableName)
        {
            return ExcelToTable(strExcelPath, tableName);
        }

        /// <summary>
        /// Excel导入成Datable
        /// </summary>
        /// <param name="file">导入路径(包含文件名与扩展名)</param>
        /// <param name="tableName">表名</param>
        /// <returns></returns>
        public static DataTable ExcelToTable(string file, string tableName)
        {
            using (DataTable dt = new DataTable(tableName))
            {
                string extension = Path.GetExtension(file);
                if (extension == null) return dt;
                string fileExt = extension.ToLower();
                using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
                {
                    //XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
                    IWorkbook workbook;
                    switch (fileExt)
                    {
                        case ".xlsx":
                            workbook = new XSSFWorkbook(fs);
                            break;
                        case ".xls":
                            workbook = new HSSFWorkbook(fs);
                            break;
                        default:
                            workbook = null;
                            break;
                    }
                    if (workbook == null) { return null; }
                    ISheet sheet = workbook.GetSheetAt(0);

                    //表头  
                    IRow header = sheet.GetRow(sheet.FirstRowNum);
                    List<int> columns = new List<int>();
                    for (int i = 0; i < header.LastCellNum; i++)
                    {
                        object obj = GetValueType(header.GetCell(i));
                        if (obj == null || obj.ToString() == string.Empty)
                        {
                            dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                        }
                        else
                            dt.Columns.Add(new DataColumn(obj.ToString()));
                        columns.Add(i);
                    }
                    //数据  
                    for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                    {
                        DataRow dr = dt.NewRow();
                        bool hasValue = false;
                        foreach (int j in columns)
                        {
                            dr[j] = GetValueType(sheet.GetRow(i).GetCell(j));
                            if (dr[j] != null && dr[j].ToString() != string.Empty)
                            {
                                hasValue = true;
                            }
                        }
                        if (hasValue)
                        {
                            dt.Rows.Add(dr);
                        }
                    }
                }
                return dt;
            }
        }

        /// <summary>
        /// 获取单元格类型
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static object GetValueType(ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.BLANK: //BLANK:  
                    return null;
                case CellType.BOOLEAN: //BOOLEAN:  
                    return cell.BooleanCellValue;
                case CellType.NUMERIC: //NUMERIC:  
                    //Cell为非NUMERIC时，调用IsCellDateFormatted方法会报错，所以先要进行类型判断
                    return DateUtil.IsCellDateFormatted(cell)? (object) cell.DateCellValue.ToString("yyyy-MM-dd HH:mm:ss") : cell.NumericCellValue;
                case CellType.STRING: //STRING:  
                    return cell.StringCellValue;
                case CellType.ERROR: //ERROR:  
                    return cell.ErrorCellValue;
                case CellType.FORMULA: //FORMULA:  
                default:
                    return "=" + cell.CellFormula;
            }
        }

        #region Old Code
        //public DataTable GetExcelTableByOleDB(string strExcelPath, string tableName)
        //{
        //    try
        //    {
        //        DataTable dtExcel = new DataTable();
        //        //数据表
        //        DataSet ds = new DataSet();
        //        //获取文件扩展名
        //        string strExtension = System.IO.Path.GetExtension(strExcelPath);
        //        string strFileName = System.IO.Path.GetFileName(strExcelPath);
        //        //Excel的连接
        //        OleDbConnection objConn = null;
        //        switch (strExtension)
        //        {
        //            case ".xls":
        //                objConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strExcelPath + ";" + "Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1;\"");
        //                break;
        //            case ".xlsx":
        //                objConn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strExcelPath + ";" + "Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1;\"");
        //                break;
        //            default:
        //                objConn = null;
        //                break;
        //        }
        //        if (objConn == null)
        //        {
        //            return null;
        //        }
        //        objConn.Open();
        //        //获取Excel中所有Sheet表的信息
        //        //System.Data.DataTable schemaTable = objConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);
        //        //objConn为读取Excel的链接，下面通过过滤来获取有效的Sheet页名称集合
        //        System.Data.DataTable schemaTable = objConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);

        //        List<string> lstSheetNames = new List<string>();
        //        for (int i = 0; i < schemaTable.Rows.Count; i++)
        //        {
        //            string strSheetName = (string)schemaTable.Rows[i]["TABLE_NAME"];
        //            if (strSheetName.Contains("$") && !strSheetName.Replace("'", "").EndsWith("$"))
        //            {
        //                //过滤无效SheetName完毕....
        //                continue;
        //            }
        //            if (lstSheetNames != null && !lstSheetNames.Contains(strSheetName))
        //                lstSheetNames.Add(strSheetName);
        //        }
        //        if (lstSheetNames.Count <= 0) return null;
        //        string sheetName = lstSheetNames[0];
        //        //获取Excel的第一个Sheet表名
        //        //string tableName = schemaTable.Rows[0][2].ToString().Trim();
        //        string strSql = "select * from [" + sheetName + "]";
        //        //获取Excel指定Sheet表中的信息
        //        OleDbCommand objCmd = new OleDbCommand(strSql, objConn);
        //        OleDbDataAdapter myData = new OleDbDataAdapter(strSql, objConn);
        //        myData.Fill(ds, tableName);//填充数据
        //        objConn.Close();
        //        //dtExcel即为excel文件中指定表中存储的信息
        //        dtExcel = ds.Tables[tableName];
        //        return dtExcel;
        //    }
        //    catch (Exception e)
        //    {
        //        MessageBox.Show(e.Message);
        //        return null;
        //    }
        //}
        #endregion
    }
}
