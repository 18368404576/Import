using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace WinImport
{
    public partial class DiffenctDocument : Form
    {
        List<string> files = new List<string>();

        List<string> files1 = new List<string>();

        ExcelManager _excelManager = new ExcelManager();

        public DiffenctDocument()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                this.txtPath.Text = this.folderBrowserDialog1.SelectedPath;
                List<string> list = Directory.GetFiles(this.txtPath.Text).ToList();
                foreach (string item in list)
                {
                    FileInfo info = new FileInfo(item);
                    files.Add(info.Name);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.txtExcel.Text = this.openFileDialog1.FileName;
                DataTable tbl = _excelManager.GetExcelTableByOleDB(this.openFileDialog1.FileName, "Sheet0");
                files1 = new List<string>();
                StringBuilder sb1 = new StringBuilder();
                foreach (DataRow item in tbl.Rows)
                {
                    if (!files1.Contains(item["文件名"].ToString()))
                    {
                        files1.Add(item["文件名"].ToString());
                    }
                    else
                    {
                        sb1.Append(item["文件名"].ToString() + Environment.NewLine);
                    }
                }
                this.textBox2.Text = sb1.ToString();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            List<string> result = files1.Except(files).ToList();
            StringBuilder sb2 = new StringBuilder();
            foreach (string item in result)
            {
                sb2.Append(item + Environment.NewLine);
            }

            result = files.Except(files1).ToList();
            foreach (string item in result)
            {
                sb2.Append(item + Environment.NewLine);
            }

            this.textBox1.Text = sb2.ToString();
        }
    }
}
