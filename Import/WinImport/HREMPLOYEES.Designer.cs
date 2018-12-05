namespace WinImport
{
    partial class HREMPLOYEES
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
            this.button2 = new System.Windows.Forms.Button();
            this.txtFileSon = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.btnImport = new System.Windows.Forms.Button();
            this.btnCheck = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.opfDialog0 = new System.Windows.Forms.OpenFileDialog();
            this.panel3 = new System.Windows.Forms.Panel();
            this.label4 = new System.Windows.Forms.Label();
            this.rbSql = new System.Windows.Forms.RichTextBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.dgError2 = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.dgRepet2 = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.panel1.SuspendLayout();
            this.panel3.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgError2)).BeginInit();
            this.tabPage4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet2)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.txtFileSon);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.btnImport);
            this.panel1.Controls.Add(this.btnCheck);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1257, 86);
            this.panel1.TabIndex = 19;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(829, 23);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 18;
            this.button2.Text = "计算技能分";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // txtFileSon
            // 
            this.txtFileSon.Location = new System.Drawing.Point(116, 23);
            this.txtFileSon.Name = "txtFileSon";
            this.txtFileSon.Size = new System.Drawing.Size(468, 21);
            this.txtFileSon.TabIndex = 10;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(21, 26);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(101, 12);
            this.label3.TabIndex = 9;
            this.label3.Text = "职工信息上岗证：";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(586, 22);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 11;
            this.button1.Text = "选择文件";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(748, 23);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(75, 23);
            this.btnImport.TabIndex = 6;
            this.btnImport.Text = "开始导入";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // btnCheck
            // 
            this.btnCheck.Location = new System.Drawing.Point(667, 23);
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
            this.label2.Location = new System.Drawing.Point(21, 60);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(293, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "导入说明：各厂部职位信息、部门信息请预先维护完整";
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
            this.panel3.Size = new System.Drawing.Size(1257, 37);
            this.panel3.TabIndex = 21;
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
            this.rbSql.Location = new System.Drawing.Point(0, 576);
            this.rbSql.Name = "rbSql";
            this.rbSql.Size = new System.Drawing.Size(1257, 110);
            this.rbSql.TabIndex = 22;
            this.rbSql.Text = "";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 86);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1257, 490);
            this.tabControl1.TabIndex = 23;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.dgError2);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(1249, 464);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "错误数据1";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // dgError2
            // 
            this.dgError2.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgError2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgError2.Location = new System.Drawing.Point(0, 0);
            this.dgError2.Name = "dgError2";
            this.dgError2.Rows.DefaultSize = 20;
            this.dgError2.Size = new System.Drawing.Size(1249, 464);
            this.dgError2.TabIndex = 0;
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.dgRepet2);
            this.tabPage4.Location = new System.Drawing.Point(4, 22);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Size = new System.Drawing.Size(1249, 464);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "重复数据1";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // dgRepet2
            // 
            this.dgRepet2.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgRepet2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgRepet2.Location = new System.Drawing.Point(0, 0);
            this.dgRepet2.Name = "dgRepet2";
            this.dgRepet2.Rows.DefaultSize = 20;
            this.dgRepet2.Size = new System.Drawing.Size(1249, 464);
            this.dgRepet2.TabIndex = 0;
            // 
            // HREMPLOYEES
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1257, 723);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.rbSql);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel1);
            this.Name = "HREMPLOYEES";
            this.Text = "HREMPLOYEES";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgError2)).EndInit();
            this.tabPage4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox txtFileSon;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.Button btnCheck;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.OpenFileDialog opfDialog0;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.RichTextBox rbSql;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage3;
        private C1.Win.C1FlexGrid.C1FlexGrid dgError2;
        private System.Windows.Forms.TabPage tabPage4;
        private C1.Win.C1FlexGrid.C1FlexGrid dgRepet2;
    }
}