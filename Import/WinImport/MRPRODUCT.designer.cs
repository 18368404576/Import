namespace WinImport
{
    partial class MRPRODUCT
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
            this.rbSql = new System.Windows.Forms.RichTextBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.dgError = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.dgRepet_excel = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.dgRepet = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.txtFile4 = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.btnSelect4 = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.btnSelect3 = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.opfDialog = new System.Windows.Forms.OpenFileDialog();
            this.txtFile1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnImport = new System.Windows.Forms.Button();
            this.btnSelect1 = new System.Windows.Forms.Button();
            this.btnCheck = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.panel3 = new System.Windows.Forms.Panel();
            this.label4 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.txtFile3 = new System.Windows.Forms.TextBox();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.tabPage5 = new System.Windows.Forms.TabPage();
            this.dgRepet_excel1 = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.dgRepet_excel2 = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgError)).BeginInit();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet_excel)).BeginInit();
            this.tabPage3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet)).BeginInit();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel1.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.tabPage5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet_excel1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet_excel2)).BeginInit();
            this.SuspendLayout();
            // 
            // rbSql
            // 
            this.rbSql.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.rbSql.Location = new System.Drawing.Point(0, 356);
            this.rbSql.Name = "rbSql";
            this.rbSql.Size = new System.Drawing.Size(819, 85);
            this.rbSql.TabIndex = 23;
            this.rbSql.Text = "";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Controls.Add(this.tabPage5);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(819, 232);
            this.tabControl1.TabIndex = 8;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.dgError);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(811, 206);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "错误数据";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // dgError
            // 
            this.dgError.AllowEditing = false;
            this.dgError.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgError.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgError.Location = new System.Drawing.Point(3, 3);
            this.dgError.Name = "dgError";
            this.dgError.Rows.DefaultSize = 20;
            this.dgError.Size = new System.Drawing.Size(805, 200);
            this.dgError.TabIndex = 0;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.dgRepet_excel);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(811, 206);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Excel重复1";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // dgRepet_excel
            // 
            this.dgRepet_excel.AllowEditing = false;
            this.dgRepet_excel.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgRepet_excel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgRepet_excel.Location = new System.Drawing.Point(3, 3);
            this.dgRepet_excel.Name = "dgRepet_excel";
            this.dgRepet_excel.Rows.DefaultSize = 20;
            this.dgRepet_excel.Size = new System.Drawing.Size(805, 200);
            this.dgRepet_excel.TabIndex = 1;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.dgRepet);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(811, 206);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "数据库重复";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // dgRepet
            // 
            this.dgRepet.AllowEditing = false;
            this.dgRepet.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgRepet.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgRepet.Location = new System.Drawing.Point(0, 0);
            this.dgRepet.Name = "dgRepet";
            this.dgRepet.Rows.DefaultSize = 20;
            this.dgRepet.Size = new System.Drawing.Size(811, 206);
            this.dgRepet.TabIndex = 2;
            // 
            // txtFile4
            // 
            this.txtFile4.Location = new System.Drawing.Point(170, 67);
            this.txtFile4.Name = "txtFile4";
            this.txtFile4.Size = new System.Drawing.Size(468, 21);
            this.txtFile4.TabIndex = 16;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(21, 71);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(137, 12);
            this.label8.TabIndex = 15;
            this.label8.Text = "Excel文件-物料保养项：";
            // 
            // btnSelect4
            // 
            this.btnSelect4.Location = new System.Drawing.Point(640, 66);
            this.btnSelect4.Name = "btnSelect4";
            this.btnSelect4.Size = new System.Drawing.Size(75, 23);
            this.btnSelect4.TabIndex = 17;
            this.btnSelect4.Text = "选择文件";
            this.btnSelect4.UseVisualStyleBackColor = true;
            this.btnSelect4.Click += new System.EventHandler(this.btnSelect_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(21, 45);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(137, 12);
            this.label6.TabIndex = 12;
            this.label6.Text = "Excel文件-物料校验项：";
            // 
            // btnSelect3
            // 
            this.btnSelect3.Location = new System.Drawing.Point(640, 40);
            this.btnSelect3.Name = "btnSelect3";
            this.btnSelect3.Size = new System.Drawing.Size(75, 23);
            this.btnSelect3.TabIndex = 14;
            this.btnSelect3.Text = "选择文件";
            this.btnSelect3.UseVisualStyleBackColor = true;
            this.btnSelect3.Click += new System.EventHandler(this.btnSelect_Click);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.tabControl1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 124);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(819, 232);
            this.panel2.TabIndex = 21;
            // 
            // opfDialog
            // 
            this.opfDialog.Filter = "Excel|*.xls;*.xlsx";
            // 
            // txtFile1
            // 
            this.txtFile1.Location = new System.Drawing.Point(170, 14);
            this.txtFile1.Name = "txtFile1";
            this.txtFile1.Size = new System.Drawing.Size(468, 21);
            this.txtFile1.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(21, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(149, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "Excel文件-物料基本信息：";
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(721, 92);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(75, 23);
            this.btnImport.TabIndex = 6;
            this.btnImport.Text = "开始导入";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // btnSelect1
            // 
            this.btnSelect1.Location = new System.Drawing.Point(640, 13);
            this.btnSelect1.Name = "btnSelect1";
            this.btnSelect1.Size = new System.Drawing.Size(75, 23);
            this.btnSelect1.TabIndex = 2;
            this.btnSelect1.Text = "选择文件";
            this.btnSelect1.UseVisualStyleBackColor = true;
            this.btnSelect1.Click += new System.EventHandler(this.btnSelect_Click);
            // 
            // btnCheck
            // 
            this.btnCheck.Location = new System.Drawing.Point(640, 92);
            this.btnCheck.Name = "btnCheck";
            this.btnCheck.Size = new System.Drawing.Size(75, 23);
            this.btnCheck.TabIndex = 5;
            this.btnCheck.Text = "数据检验";
            this.btnCheck.UseVisualStyleBackColor = true;
            this.btnCheck.Click += new System.EventHandler(this.btnCheck_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(21, 97);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(479, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "导入说明：职位信息、部门信息请预先维护完整,选择完文件检查数量与实际数量是否相符";
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.label4);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel3.Location = new System.Drawing.Point(0, 441);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(819, 37);
            this.panel3.TabIndex = 22;
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
            // panel1
            // 
            this.panel1.Controls.Add(this.txtFile4);
            this.panel1.Controls.Add(this.label8);
            this.panel1.Controls.Add(this.btnSelect4);
            this.panel1.Controls.Add(this.txtFile3);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.btnSelect3);
            this.panel1.Controls.Add(this.txtFile1);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.btnImport);
            this.panel1.Controls.Add(this.btnSelect1);
            this.panel1.Controls.Add(this.btnCheck);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(819, 124);
            this.panel1.TabIndex = 20;
            // 
            // txtFile3
            // 
            this.txtFile3.Location = new System.Drawing.Point(170, 41);
            this.txtFile3.Name = "txtFile3";
            this.txtFile3.Size = new System.Drawing.Size(468, 21);
            this.txtFile3.TabIndex = 13;
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.dgRepet_excel1);
            this.tabPage4.Location = new System.Drawing.Point(4, 22);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Size = new System.Drawing.Size(811, 206);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "Excel重复2";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // tabPage5
            // 
            this.tabPage5.Controls.Add(this.dgRepet_excel2);
            this.tabPage5.Location = new System.Drawing.Point(4, 22);
            this.tabPage5.Name = "tabPage5";
            this.tabPage5.Size = new System.Drawing.Size(811, 206);
            this.tabPage5.TabIndex = 4;
            this.tabPage5.Text = "Excel重复3";
            this.tabPage5.UseVisualStyleBackColor = true;
            // 
            // dgRepet_excel1
            // 
            this.dgRepet_excel1.AllowEditing = false;
            this.dgRepet_excel1.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgRepet_excel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgRepet_excel1.Location = new System.Drawing.Point(0, 0);
            this.dgRepet_excel1.Name = "dgRepet_excel1";
            this.dgRepet_excel1.Rows.DefaultSize = 20;
            this.dgRepet_excel1.Size = new System.Drawing.Size(811, 206);
            this.dgRepet_excel1.TabIndex = 2;
            // 
            // dgRepet_excel2
            // 
            this.dgRepet_excel2.AllowEditing = false;
            this.dgRepet_excel2.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgRepet_excel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgRepet_excel2.Location = new System.Drawing.Point(0, 0);
            this.dgRepet_excel2.Name = "dgRepet_excel2";
            this.dgRepet_excel2.Rows.DefaultSize = 20;
            this.dgRepet_excel2.Size = new System.Drawing.Size(811, 206);
            this.dgRepet_excel2.TabIndex = 3;
            // 
            // MRPRODUCT
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(819, 478);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.rbSql);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel1);
            this.Name = "MRPRODUCT";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "物料基本信息";
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgError)).EndInit();
            this.tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet_excel)).EndInit();
            this.tabPage3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.tabPage4.ResumeLayout(false);
            this.tabPage5.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet_excel1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet_excel2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.RichTextBox rbSql;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TextBox txtFile4;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button btnSelect4;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnSelect3;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.OpenFileDialog opfDialog;
        private System.Windows.Forms.TextBox txtFile1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.Button btnSelect1;
        private System.Windows.Forms.Button btnCheck;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Panel panel1;
        private C1.Win.C1FlexGrid.C1FlexGrid dgError;
        private C1.Win.C1FlexGrid.C1FlexGrid dgRepet_excel;
        private System.Windows.Forms.TextBox txtFile3;
        private System.Windows.Forms.TabPage tabPage3;
        private C1.Win.C1FlexGrid.C1FlexGrid dgRepet;
        private System.Windows.Forms.TabPage tabPage4;
        private C1.Win.C1FlexGrid.C1FlexGrid dgRepet_excel1;
        private System.Windows.Forms.TabPage tabPage5;
        private C1.Win.C1FlexGrid.C1FlexGrid dgRepet_excel2;
    }
}