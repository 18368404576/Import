namespace WinImport
{
    partial class BSPOSITION
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
            this.dgError1 = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.dgRepet1 = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.folderBrowserDialog2 = new System.Windows.Forms.FolderBrowserDialog();
            this.panel2 = new System.Windows.Forms.Panel();
            this.rbSql = new System.Windows.Forms.RichTextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtFileMain = new System.Windows.Forms.TextBox();
            this.btnSelect = new System.Windows.Forms.Button();
            this.opfDialog0 = new System.Windows.Forms.OpenFileDialog();
            this.panel3 = new System.Windows.Forms.Panel();
            this.btnImport = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnCheck = new System.Windows.Forms.Button();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.dgDataRepet1 = new C1.Win.C1FlexGrid.C1FlexGrid();
            ((System.ComponentModel.ISupportInitialize)(this.dgError1)).BeginInit();
            this.tabPage1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet1)).BeginInit();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel1.SuspendLayout();
            this.tabPage3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgDataRepet1)).BeginInit();
            this.SuspendLayout();
            // 
            // dgError1
            // 
            this.dgError1.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgError1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgError1.Location = new System.Drawing.Point(3, 3);
            this.dgError1.Name = "dgError1";
            this.dgError1.Rows.DefaultSize = 20;
            this.dgError1.Size = new System.Drawing.Size(1143, 494);
            this.dgError1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.dgError1);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1149, 500);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "错误数据";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1157, 526);
            this.tabControl1.TabIndex = 8;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.dgRepet1);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1149, 500);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "EXCEL重复数据";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // dgRepet1
            // 
            this.dgRepet1.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgRepet1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgRepet1.Location = new System.Drawing.Point(3, 3);
            this.dgRepet1.Name = "dgRepet1";
            this.dgRepet1.Rows.DefaultSize = 20;
            this.dgRepet1.Size = new System.Drawing.Size(1143, 494);
            this.dgRepet1.TabIndex = 0;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.Filter = "Excel|*.xls;*.xlsx";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.tabControl1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 50);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1157, 526);
            this.panel2.TabIndex = 26;
            // 
            // rbSql
            // 
            this.rbSql.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.rbSql.Location = new System.Drawing.Point(0, 576);
            this.rbSql.Name = "rbSql";
            this.rbSql.Size = new System.Drawing.Size(1157, 110);
            this.rbSql.TabIndex = 25;
            this.rbSql.Text = "";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(21, 19);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(65, 12);
            this.label7.TabIndex = 17;
            this.label7.Text = "职位信息：";
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
            // txtFileMain
            // 
            this.txtFileMain.Location = new System.Drawing.Point(116, 14);
            this.txtFileMain.Name = "txtFileMain";
            this.txtFileMain.Size = new System.Drawing.Size(468, 21);
            this.txtFileMain.TabIndex = 15;
            // 
            // btnSelect
            // 
            this.btnSelect.Location = new System.Drawing.Point(586, 14);
            this.btnSelect.Name = "btnSelect";
            this.btnSelect.Size = new System.Drawing.Size(75, 23);
            this.btnSelect.TabIndex = 16;
            this.btnSelect.Text = "选择文件";
            this.btnSelect.UseVisualStyleBackColor = true;
            this.btnSelect.Click += new System.EventHandler(this.btnSelect_Click);
            // 
            // opfDialog0
            // 
            this.opfDialog0.Filter = "Excel|*.xls;*.xlsx";
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.label4);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel3.Location = new System.Drawing.Point(0, 686);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(1157, 37);
            this.panel3.TabIndex = 24;
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(753, 14);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(75, 23);
            this.btnImport.TabIndex = 6;
            this.btnImport.Text = "开始导入";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click_1);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.txtFileMain);
            this.panel1.Controls.Add(this.btnSelect);
            this.panel1.Controls.Add(this.btnImport);
            this.panel1.Controls.Add(this.btnCheck);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1157, 50);
            this.panel1.TabIndex = 23;
            // 
            // btnCheck
            // 
            this.btnCheck.Location = new System.Drawing.Point(672, 14);
            this.btnCheck.Name = "btnCheck";
            this.btnCheck.Size = new System.Drawing.Size(75, 23);
            this.btnCheck.TabIndex = 5;
            this.btnCheck.Text = "数据检验";
            this.btnCheck.UseVisualStyleBackColor = true;
            this.btnCheck.Click += new System.EventHandler(this.btnCheck_Click);
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.dgDataRepet1);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(1149, 500);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "数据库重复数据";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // dgDataRepet1
            // 
            this.dgDataRepet1.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgDataRepet1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgDataRepet1.Location = new System.Drawing.Point(0, 0);
            this.dgDataRepet1.Name = "dgDataRepet1";
            this.dgDataRepet1.Rows.DefaultSize = 20;
            this.dgDataRepet1.Size = new System.Drawing.Size(1149, 500);
            this.dgDataRepet1.TabIndex = 1;
            // 
            // BSPOSITION
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1157, 723);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.rbSql);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel1);
            this.Name = "BSPOSITION";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "职位设定";
            ((System.ComponentModel.ISupportInitialize)(this.dgError1)).EndInit();
            this.tabPage1.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet1)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgDataRepet1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private C1.Win.C1FlexGrid.C1FlexGrid dgError1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage2;
        private C1.Win.C1FlexGrid.C1FlexGrid dgRepet1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog2;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.RichTextBox rbSql;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtFileMain;
        private System.Windows.Forms.Button btnSelect;
        private System.Windows.Forms.OpenFileDialog opfDialog0;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnCheck;
        private System.Windows.Forms.TabPage tabPage3;
        private C1.Win.C1FlexGrid.C1FlexGrid dgDataRepet1;
    }
}