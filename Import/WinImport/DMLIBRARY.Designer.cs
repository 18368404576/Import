namespace WinImport
{
    partial class DMLIBRARY
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.opfDialog0 = new System.Windows.Forms.OpenFileDialog();
            this.label4 = new System.Windows.Forms.Label();
            this.panel3 = new System.Windows.Forms.Panel();
            this.dgRepet1_excel = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.dgError1 = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage5 = new System.Windows.Forms.TabPage();
            this.dgRepet1 = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.dgError2 = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.dgRepet2 = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.txtFileSon = new System.Windows.Forms.TextBox();
            this.button3 = new System.Windows.Forms.Button();
            this.txtFileMain = new System.Windows.Forms.TextBox();
            this.btnImport = new System.Windows.Forms.Button();
            this.btnSelect = new System.Windows.Forms.Button();
            this.btnCheck = new System.Windows.Forms.Button();
            this.rbSql = new System.Windows.Forms.RichTextBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.folderBrowserDialog2 = new System.Windows.Forms.FolderBrowserDialog();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet1_excel)).BeginInit();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgError1)).BeginInit();
            this.tabPage2.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet1)).BeginInit();
            this.tabPage3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgError2)).BeginInit();
            this.tabPage4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet2)).BeginInit();
            this.panel2.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // opfDialog0
            // 
            this.opfDialog0.Filter = "Excel|*.xls;*.xlsx";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(21, 14);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(59, 12);
            this.label4.TabIndex = 11;
            this.label4.Text = "Sql语句：";
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.label4);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel3.Location = new System.Drawing.Point(0, 576);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(1257, 37);
            this.panel3.TabIndex = 14;
            // 
            // dgRepet1_excel
            // 
            this.dgRepet1_excel.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgRepet1_excel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgRepet1_excel.Location = new System.Drawing.Point(3, 3);
            this.dgRepet1_excel.Name = "dgRepet1_excel";
            this.dgRepet1_excel.Rows.DefaultSize = 20;
            this.dgRepet1_excel.Size = new System.Drawing.Size(1243, 495);
            this.dgRepet1_excel.TabIndex = 8;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.dgError1);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1249, 501);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "错误数据1";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // dgError1
            // 
            this.dgError1.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgError1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgError1.Location = new System.Drawing.Point(3, 3);
            this.dgError1.Name = "dgError1";
            this.dgError1.Rows.DefaultSize = 20;
            this.dgError1.Size = new System.Drawing.Size(1243, 495);
            this.dgError1.TabIndex = 7;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.dgRepet1_excel);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1249, 501);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Excel重复1";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage5);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1257, 527);
            this.tabControl1.TabIndex = 8;
            // 
            // tabPage5
            // 
            this.tabPage5.Controls.Add(this.dgRepet1);
            this.tabPage5.Location = new System.Drawing.Point(4, 22);
            this.tabPage5.Name = "tabPage5";
            this.tabPage5.Size = new System.Drawing.Size(1249, 501);
            this.tabPage5.TabIndex = 4;
            this.tabPage5.Text = "数据库重复";
            this.tabPage5.UseVisualStyleBackColor = true;
            // 
            // dgRepet1
            // 
            this.dgRepet1.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgRepet1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgRepet1.Location = new System.Drawing.Point(0, 0);
            this.dgRepet1.Name = "dgRepet1";
            this.dgRepet1.Rows.DefaultSize = 20;
            this.dgRepet1.Size = new System.Drawing.Size(1249, 501);
            this.dgRepet1.TabIndex = 9;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.dgError2);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(1249, 501);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "错误数据2";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // dgError2
            // 
            this.dgError2.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgError2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgError2.Location = new System.Drawing.Point(0, 0);
            this.dgError2.Name = "dgError2";
            this.dgError2.Rows.DefaultSize = 20;
            this.dgError2.Size = new System.Drawing.Size(1249, 501);
            this.dgError2.TabIndex = 8;
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.dgRepet2);
            this.tabPage4.Location = new System.Drawing.Point(4, 22);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Size = new System.Drawing.Size(1249, 501);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "excel重复2";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // dgRepet2
            // 
            this.dgRepet2.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgRepet2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgRepet2.Location = new System.Drawing.Point(0, 0);
            this.dgRepet2.Name = "dgRepet2";
            this.dgRepet2.Rows.DefaultSize = 20;
            this.dgRepet2.Size = new System.Drawing.Size(1249, 501);
            this.dgRepet2.TabIndex = 9;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.tabControl1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 86);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1257, 527);
            this.panel2.TabIndex = 13;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.label8);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.txtFileSon);
            this.panel1.Controls.Add(this.button3);
            this.panel1.Controls.Add(this.txtFileMain);
            this.panel1.Controls.Add(this.btnImport);
            this.panel1.Controls.Add(this.btnSelect);
            this.panel1.Controls.Add(this.btnCheck);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1257, 86);
            this.panel1.TabIndex = 12;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.Red;
            this.label1.Location = new System.Drawing.Point(846, 56);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(251, 12);
            this.label1.TabIndex = 19;
            this.label1.Text = "文档导入成功后要修改对应文档类别表的CNO值";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(21, 51);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(89, 12);
            this.label8.TabIndex = 18;
            this.label8.Text = "文档库关联项：";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(21, 23);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(53, 12);
            this.label7.TabIndex = 17;
            this.label7.Text = "文档库：";
            // 
            // txtFileSon
            // 
            this.txtFileSon.Location = new System.Drawing.Point(121, 47);
            this.txtFileSon.Name = "txtFileSon";
            this.txtFileSon.Size = new System.Drawing.Size(468, 21);
            this.txtFileSon.TabIndex = 15;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(591, 47);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 16;
            this.button3.Text = "选择文件";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // txtFileMain
            // 
            this.txtFileMain.Location = new System.Drawing.Point(121, 20);
            this.txtFileMain.Name = "txtFileMain";
            this.txtFileMain.Size = new System.Drawing.Size(468, 21);
            this.txtFileMain.TabIndex = 1;
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(753, 47);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(75, 23);
            this.btnImport.TabIndex = 6;
            this.btnImport.Text = "开始导入";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // btnSelect
            // 
            this.btnSelect.Location = new System.Drawing.Point(591, 18);
            this.btnSelect.Name = "btnSelect";
            this.btnSelect.Size = new System.Drawing.Size(75, 23);
            this.btnSelect.TabIndex = 2;
            this.btnSelect.Text = "选择文件";
            this.btnSelect.UseVisualStyleBackColor = true;
            this.btnSelect.Click += new System.EventHandler(this.btnSelect_Click);
            // 
            // btnCheck
            // 
            this.btnCheck.Location = new System.Drawing.Point(672, 47);
            this.btnCheck.Name = "btnCheck";
            this.btnCheck.Size = new System.Drawing.Size(75, 23);
            this.btnCheck.TabIndex = 5;
            this.btnCheck.Text = "数据检验";
            this.btnCheck.UseVisualStyleBackColor = true;
            this.btnCheck.Click += new System.EventHandler(this.btnCheck_Click);
            // 
            // rbSql
            // 
            this.rbSql.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.rbSql.Location = new System.Drawing.Point(0, 613);
            this.rbSql.Name = "rbSql";
            this.rbSql.Size = new System.Drawing.Size(1257, 110);
            this.rbSql.TabIndex = 15;
            this.rbSql.Text = "";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.Filter = "Excel|*.xls;*.xlsx";
            // 
            // DMLIBRARY
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1257, 723);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.rbSql);
            this.Name = "DMLIBRARY";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "文档库";
            this.Load += new System.EventHandler(this.DMLIBRARY_Load);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet1_excel)).EndInit();
            this.tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgError1)).EndInit();
            this.tabPage2.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage5.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet1)).EndInit();
            this.tabPage3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgError2)).EndInit();
            this.tabPage4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet2)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog opfDialog0;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Panel panel3;
        private C1.Win.C1FlexGrid.C1FlexGrid dgRepet1_excel;
        private System.Windows.Forms.TabPage tabPage1;
        private C1.Win.C1FlexGrid.C1FlexGrid dgError1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox txtFileMain;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.Button btnSelect;
        private System.Windows.Forms.Button btnCheck;
        private System.Windows.Forms.RichTextBox rbSql;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtFileSon;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog2;
        private System.Windows.Forms.TabPage tabPage3;
        private C1.Win.C1FlexGrid.C1FlexGrid dgError2;
        private System.Windows.Forms.TabPage tabPage4;
        private C1.Win.C1FlexGrid.C1FlexGrid dgRepet2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TabPage tabPage5;
        private C1.Win.C1FlexGrid.C1FlexGrid dgRepet1;
    }
}