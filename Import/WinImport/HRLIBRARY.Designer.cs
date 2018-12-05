namespace WinImport
{
    partial class HRLIBRARY
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnImport = new System.Windows.Forms.Button();
            this.btnCheck = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txtFile3 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.btnSelect3 = new System.Windows.Forms.Button();
            this.txtFile2 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btnSelect2 = new System.Windows.Forms.Button();
            this.txtFile1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnSelect1 = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.dgError1 = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.dgError2 = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.dgRepet2 = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.tabPage5 = new System.Windows.Forms.TabPage();
            this.dgError3 = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.tabPage6 = new System.Windows.Forms.TabPage();
            this.dgRepet3 = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.panel3 = new System.Windows.Forms.Panel();
            this.label4 = new System.Windows.Forms.Label();
            this.rbSql = new System.Windows.Forms.RichTextBox();
            this.opfDialog = new System.Windows.Forms.OpenFileDialog();
            this.tabPage7 = new System.Windows.Forms.TabPage();
            this.dgRepet1 = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.panel1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgError1)).BeginInit();
            this.tabPage3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgError2)).BeginInit();
            this.tabPage4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet2)).BeginInit();
            this.tabPage5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgError3)).BeginInit();
            this.tabPage6.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet3)).BeginInit();
            this.panel3.SuspendLayout();
            this.tabPage7.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btnImport);
            this.panel1.Controls.Add(this.btnCheck);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.txtFile3);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.btnSelect3);
            this.panel1.Controls.Add(this.txtFile2);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.btnSelect2);
            this.panel1.Controls.Add(this.txtFile1);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.btnSelect1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(979, 146);
            this.panel1.TabIndex = 39;
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(774, 108);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(75, 23);
            this.btnImport.TabIndex = 53;
            this.btnImport.Text = "开始导入";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // btnCheck
            // 
            this.btnCheck.Location = new System.Drawing.Point(693, 108);
            this.btnCheck.Name = "btnCheck";
            this.btnCheck.Size = new System.Drawing.Size(75, 23);
            this.btnCheck.TabIndex = 52;
            this.btnCheck.Text = "数据检验";
            this.btnCheck.UseVisualStyleBackColor = true;
            this.btnCheck.Click += new System.EventHandler(this.btnCheck_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(44, 113);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(647, 12);
            this.label2.TabIndex = 51;
            this.label2.Text = "导入说明：（题库关联上岗证不用选择）题库信息、上岗证信息请预先维护完整,选择完文件检查数量与实际数量是否相符";
            // 
            // txtFile3
            // 
            this.txtFile3.Location = new System.Drawing.Point(223, 80);
            this.txtFile3.Name = "txtFile3";
            this.txtFile3.Size = new System.Drawing.Size(468, 21);
            this.txtFile3.TabIndex = 46;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(44, 84);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(125, 12);
            this.label6.TabIndex = 45;
            this.label6.Text = "Excel文件-题库星级：";
            // 
            // btnSelect3
            // 
            this.btnSelect3.Location = new System.Drawing.Point(693, 79);
            this.btnSelect3.Name = "btnSelect3";
            this.btnSelect3.Size = new System.Drawing.Size(75, 23);
            this.btnSelect3.TabIndex = 47;
            this.btnSelect3.Text = "选择文件";
            this.btnSelect3.UseVisualStyleBackColor = true;
            this.btnSelect3.Click += new System.EventHandler(this.btnSelect1_Click);
            // 
            // txtFile2
            // 
            this.txtFile2.Location = new System.Drawing.Point(223, 51);
            this.txtFile2.Name = "txtFile2";
            this.txtFile2.Size = new System.Drawing.Size(468, 21);
            this.txtFile2.TabIndex = 43;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(44, 55);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(125, 12);
            this.label3.TabIndex = 42;
            this.label3.Text = "Excel文件-题库选项：";
            // 
            // btnSelect2
            // 
            this.btnSelect2.Location = new System.Drawing.Point(693, 50);
            this.btnSelect2.Name = "btnSelect2";
            this.btnSelect2.Size = new System.Drawing.Size(75, 23);
            this.btnSelect2.TabIndex = 44;
            this.btnSelect2.Text = "选择文件";
            this.btnSelect2.UseVisualStyleBackColor = true;
            this.btnSelect2.Click += new System.EventHandler(this.btnSelect1_Click);
            // 
            // txtFile1
            // 
            this.txtFile1.Location = new System.Drawing.Point(223, 24);
            this.txtFile1.Name = "txtFile1";
            this.txtFile1.Size = new System.Drawing.Size(468, 21);
            this.txtFile1.TabIndex = 40;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(44, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(101, 12);
            this.label1.TabIndex = 39;
            this.label1.Text = "Excel文件-题库：";
            // 
            // btnSelect1
            // 
            this.btnSelect1.Location = new System.Drawing.Point(693, 23);
            this.btnSelect1.Name = "btnSelect1";
            this.btnSelect1.Size = new System.Drawing.Size(75, 23);
            this.btnSelect1.TabIndex = 41;
            this.btnSelect1.Text = "选择文件";
            this.btnSelect1.UseVisualStyleBackColor = true;
            this.btnSelect1.Click += new System.EventHandler(this.btnSelect1_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage7);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Controls.Add(this.tabPage5);
            this.tabControl1.Controls.Add(this.tabPage6);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 146);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(979, 396);
            this.tabControl1.TabIndex = 40;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.dgError1);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(971, 370);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "错误数据1";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // dgError1
            // 
            this.dgError1.AllowEditing = false;
            this.dgError1.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgError1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgError1.Location = new System.Drawing.Point(3, 3);
            this.dgError1.Name = "dgError1";
            this.dgError1.Rows.DefaultSize = 20;
            this.dgError1.Size = new System.Drawing.Size(965, 364);
            this.dgError1.TabIndex = 0;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.dgError2);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(971, 370);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "错误数据2";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // dgError2
            // 
            this.dgError2.AllowEditing = false;
            this.dgError2.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgError2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgError2.Location = new System.Drawing.Point(0, 0);
            this.dgError2.Name = "dgError2";
            this.dgError2.Rows.DefaultSize = 20;
            this.dgError2.Size = new System.Drawing.Size(971, 370);
            this.dgError2.TabIndex = 1;
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.dgRepet2);
            this.tabPage4.Location = new System.Drawing.Point(4, 22);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Size = new System.Drawing.Size(971, 370);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "重复数据2";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // dgRepet2
            // 
            this.dgRepet2.AllowEditing = false;
            this.dgRepet2.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgRepet2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgRepet2.Location = new System.Drawing.Point(0, 0);
            this.dgRepet2.Name = "dgRepet2";
            this.dgRepet2.Rows.DefaultSize = 20;
            this.dgRepet2.Size = new System.Drawing.Size(971, 370);
            this.dgRepet2.TabIndex = 2;
            // 
            // tabPage5
            // 
            this.tabPage5.Controls.Add(this.dgError3);
            this.tabPage5.Location = new System.Drawing.Point(4, 22);
            this.tabPage5.Name = "tabPage5";
            this.tabPage5.Size = new System.Drawing.Size(971, 370);
            this.tabPage5.TabIndex = 4;
            this.tabPage5.Text = "错误数据3";
            this.tabPage5.UseVisualStyleBackColor = true;
            // 
            // dgError3
            // 
            this.dgError3.AllowEditing = false;
            this.dgError3.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgError3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgError3.Location = new System.Drawing.Point(0, 0);
            this.dgError3.Name = "dgError3";
            this.dgError3.Rows.DefaultSize = 20;
            this.dgError3.Size = new System.Drawing.Size(971, 370);
            this.dgError3.TabIndex = 3;
            // 
            // tabPage6
            // 
            this.tabPage6.Controls.Add(this.dgRepet3);
            this.tabPage6.Location = new System.Drawing.Point(4, 22);
            this.tabPage6.Name = "tabPage6";
            this.tabPage6.Size = new System.Drawing.Size(971, 370);
            this.tabPage6.TabIndex = 5;
            this.tabPage6.Text = "重复数据3";
            this.tabPage6.UseVisualStyleBackColor = true;
            // 
            // dgRepet3
            // 
            this.dgRepet3.AllowEditing = false;
            this.dgRepet3.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgRepet3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgRepet3.Location = new System.Drawing.Point(0, 0);
            this.dgRepet3.Name = "dgRepet3";
            this.dgRepet3.Rows.DefaultSize = 20;
            this.dgRepet3.Size = new System.Drawing.Size(971, 370);
            this.dgRepet3.TabIndex = 2;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.label4);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel3.Location = new System.Drawing.Point(0, 395);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(979, 37);
            this.panel3.TabIndex = 41;
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
            // rbSql
            // 
            this.rbSql.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.rbSql.Location = new System.Drawing.Point(0, 432);
            this.rbSql.Name = "rbSql";
            this.rbSql.Size = new System.Drawing.Size(979, 110);
            this.rbSql.TabIndex = 42;
            this.rbSql.Text = "";
            // 
            // opfDialog
            // 
            this.opfDialog.Filter = "Excel|*.xls;*.xlsx";
            // 
            // tabPage7
            // 
            this.tabPage7.Controls.Add(this.dgRepet1);
            this.tabPage7.Location = new System.Drawing.Point(4, 22);
            this.tabPage7.Name = "tabPage7";
            this.tabPage7.Size = new System.Drawing.Size(971, 370);
            this.tabPage7.TabIndex = 6;
            this.tabPage7.Text = "EXCEL重复1";
            this.tabPage7.UseVisualStyleBackColor = true;
            // 
            // dgRepet1
            // 
            this.dgRepet1.AllowEditing = false;
            this.dgRepet1.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgRepet1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgRepet1.Location = new System.Drawing.Point(0, 0);
            this.dgRepet1.Name = "dgRepet1";
            this.dgRepet1.Rows.DefaultSize = 20;
            this.dgRepet1.Size = new System.Drawing.Size(971, 370);
            this.dgRepet1.TabIndex = 2;
            // 
            // HRLIBRARY
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(979, 542);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.rbSql);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.panel1);
            this.Name = "HRLIBRARY";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "题库";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgError1)).EndInit();
            this.tabPage3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgError2)).EndInit();
            this.tabPage4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet2)).EndInit();
            this.tabPage5.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgError3)).EndInit();
            this.tabPage6.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet3)).EndInit();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.tabPage7.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.Button btnCheck;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtFile3;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnSelect3;
        private System.Windows.Forms.TextBox txtFile2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnSelect2;
        private System.Windows.Forms.TextBox txtFile1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnSelect1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private C1.Win.C1FlexGrid.C1FlexGrid dgError1;
        private System.Windows.Forms.TabPage tabPage3;
        private C1.Win.C1FlexGrid.C1FlexGrid dgError2;
        private System.Windows.Forms.TabPage tabPage4;
        private C1.Win.C1FlexGrid.C1FlexGrid dgRepet2;
        private System.Windows.Forms.TabPage tabPage5;
        private C1.Win.C1FlexGrid.C1FlexGrid dgError3;
        private System.Windows.Forms.TabPage tabPage6;
        private C1.Win.C1FlexGrid.C1FlexGrid dgRepet3;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.RichTextBox rbSql;
        private System.Windows.Forms.OpenFileDialog opfDialog;
        private System.Windows.Forms.TabPage tabPage7;
        private C1.Win.C1FlexGrid.C1FlexGrid dgRepet1;

    }
}