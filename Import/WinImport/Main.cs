using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using TSES.Base;

namespace WinImport
{
    public partial class Main : Form
    {
        /// <summary>
        /// 公共数据库连接
        /// </summary>
        public static string CONN_Public = ConfigurationManager.AppSettings["CONN_ACCOUNT"];
        public static string CONN_Model = ConfigurationManager.AppSettings["CONN_MODEL"];

        public Main()
        {
            InitializeComponent();
        }

        private void btnHRWORKLICENSE_Click(object sender, EventArgs e)
        {
            HRWORKLICENSE fm = new HRWORKLICENSE();
            fm.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ECINFO fm = new ECINFO();
            fm.ShowDialog();
        } 
        
        /// <summary>
        /// 设置错误单元格
        /// </summary>
        /// <param name="c1FlexGrid1"></param>
        /// <param name="errorCell"></param>
        public static void SetErrorCell(C1.Win.C1FlexGrid.C1FlexGrid c1FlexGrid1, List<int[]> errorCell)
        {
            if (c1FlexGrid1.Styles["ErrorCell"] == null)
            {
                C1.Win.C1FlexGrid.CellStyle errorStyle = c1FlexGrid1.Styles.Add("ErrorCell");
                errorStyle.BackColor = Color.Red;
            }
            foreach (int[] item in errorCell)
            {
                c1FlexGrid1.SetCellStyle(item[0] + 1, item[1] + 1, "ErrorCell");
            }
        }

        /// <summary>
        /// 用于数据库赋值(包括单引号转义)
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string SetDBValue(object value)
        {
            return (value == null || value.ToString() == "") ? "null" : "'" + WGHelper.ReturnString(value.ToString()) + "'";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            BSPRODUCT fm = new BSPRODUCT();
            fm.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            BSPROSTANDARD fm = new BSPROSTANDARD();
            fm.ShowDialog();
        }

        private void btnHREMPLOYEE_Click(object sender, EventArgs e)
        {
            HREMPLOYEE fm = new HREMPLOYEE();
            fm.ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            BSPOSITION fm = new BSPOSITION();
            fm.ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            BSDEPT fm = new BSDEPT();
            fm.ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            BSWORKCENTERINFO fm = new BSWORKCENTERINFO();
            fm.ShowDialog();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            BSWCTYPE fm = new BSWCTYPE();
            fm.ShowDialog();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            MRPRODUCT fm = new MRPRODUCT();
            fm.ShowDialog();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            fmImport fm = new fmImport();
            fm.ShowDialog();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            HREMPLOYEES fm = new HREMPLOYEES();
            fm.ShowDialog();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            HREMPLOYEEB fm = new HREMPLOYEEB();
            fm.ShowDialog();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            HRWORKLICENSEA fm = new HRWORKLICENSEA();
            fm.ShowDialog();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            BSPROCESS fm = new BSPROCESS();
            fm.ShowDialog();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            HRLIBRARY fm = new HRLIBRARY();
            fm.ShowDialog();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            DMLIBRARY fm = new DMLIBRARY();
            fm.ShowDialog();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            fmDocument fm = new fmDocument();
            fm.ShowDialog();
        }

        private void button17_Click(object sender, EventArgs e)
        {
            MRBOM fm = new MRBOM();
            fm.ShowDialog();
        }

        private void button18_Click(object sender, EventArgs e)
        {
            ECINFOS fm = new ECINFOS();
            fm.ShowDialog();
        }

        private void button19_Click(object sender, EventArgs e)
        {
            DiffenctDocument doc = new DiffenctDocument();
            doc.ShowDialog();
        }

        private void button20_Click(object sender, EventArgs e)
        {
            TCModify fm = new TCModify();
            fm.ShowDialog();
        }

        /// <summary>
        /// 供应商
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button21_Click(object sender, EventArgs e)
        {
            BSSUPPLIER fm = new BSSUPPLIER();
            fm.ShowDialog();
        }

        /// <summary>
        /// 客户
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button22_Click(object sender, EventArgs e)
        {
            BSCLIENT fm = new BSCLIENT();
            fm.ShowDialog();
        }
    }
}
