namespace WinImport
{
    partial class TCModify
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
            this.label7 = new System.Windows.Forms.Label();
            this.txtFileMain = new System.Windows.Forms.TextBox();
            this.btnSelect = new System.Windows.Forms.Button();
            this.btnImport = new System.Windows.Forms.Button();
            this.btnCheck = new System.Windows.Forms.Button();
            this.rbSql = new System.Windows.Forms.RichTextBox();
            this.panel3 = new System.Windows.Forms.Panel();
            this.label4 = new System.Windows.Forms.Label();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.dgError = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.dgRepet = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.opfDialog0 = new System.Windows.Forms.OpenFileDialog();
            this.panel1.SuspendLayout();
            this.panel3.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgError)).BeginInit();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet)).BeginInit();
            this.SuspendLayout();
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
            this.panel1.Size = new System.Drawing.Size(1014, 51);
            this.panel1.TabIndex = 23;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(21, 19);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(29, 12);
            this.label7.TabIndex = 17;
            this.label7.Text = "CT：";
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
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(752, 14);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(75, 23);
            this.btnImport.TabIndex = 6;
            this.btnImport.Text = "开始导入";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // btnCheck
            // 
            this.btnCheck.Location = new System.Drawing.Point(671, 14);
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
            this.rbSql.Location = new System.Drawing.Point(0, 365);
            this.rbSql.Name = "rbSql";
            this.rbSql.Size = new System.Drawing.Size(1014, 321);
            this.rbSql.TabIndex = 26;
            this.rbSql.Text = "";
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.label4);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel3.Location = new System.Drawing.Point(0, 686);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(1014, 37);
            this.panel3.TabIndex = 25;
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
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 51);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1014, 314);
            this.tabControl1.TabIndex = 27;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.dgError);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1006, 288);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "错误数据1";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // dgError
            // 
            this.dgError.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgError.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgError.Location = new System.Drawing.Point(3, 3);
            this.dgError.Name = "dgError";
            this.dgError.Rows.DefaultSize = 20;
            this.dgError.Size = new System.Drawing.Size(1000, 282);
            this.dgError.TabIndex = 0;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.dgRepet);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1006, 288);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "EXCEL重复数据1";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // dgRepet
            // 
            this.dgRepet.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgRepet.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgRepet.Location = new System.Drawing.Point(3, 3);
            this.dgRepet.Name = "dgRepet";
            this.dgRepet.Rows.DefaultSize = 20;
            this.dgRepet.Size = new System.Drawing.Size(1000, 282);
            this.dgRepet.TabIndex = 0;
            // 
            // opfDialog0
            // 
            this.opfDialog0.Filter = "Excel|*.xls;*.xlsx";
            // 
            // TCModify
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1014, 723);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.rbSql);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel1);
            this.Name = "TCModify";
            this.Text = "CT修改";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgError)).EndInit();
            this.tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtFileMain;
        private System.Windows.Forms.Button btnSelect;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.Button btnCheck;
        private System.Windows.Forms.RichTextBox rbSql;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private C1.Win.C1FlexGrid.C1FlexGrid dgError;
        private System.Windows.Forms.TabPage tabPage2;
        private C1.Win.C1FlexGrid.C1FlexGrid dgRepet;
        private System.Windows.Forms.OpenFileDialog opfDialog0;
    }
}