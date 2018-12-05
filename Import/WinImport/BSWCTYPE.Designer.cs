namespace WinImport
{
    partial class BSWCTYPE
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
            this.label4 = new System.Windows.Forms.Label();
            this.panel3 = new System.Windows.Forms.Panel();
            this.opfDialog = new System.Windows.Forms.OpenFileDialog();
            this.panel2 = new System.Windows.Forms.Panel();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.dgError = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.dgRepet = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.txtFile1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnImport = new System.Windows.Forms.Button();
            this.btnSelect1 = new System.Windows.Forms.Button();
            this.btnCheck = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.rbSql = new System.Windows.Forms.RichTextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.dgRepetFromDB = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.panel3.SuspendLayout();
            this.panel2.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgError)).BeginInit();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet)).BeginInit();
            this.panel1.SuspendLayout();
            this.tabPage3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgRepetFromDB)).BeginInit();
            this.SuspendLayout();
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
            this.panel3.Location = new System.Drawing.Point(0, 331);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(974, 37);
            this.panel3.TabIndex = 26;
            // 
            // opfDialog
            // 
            this.opfDialog.Filter = "Excel|*.xls;*.xlsx";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.tabControl1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 214);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(974, 117);
            this.panel2.TabIndex = 25;
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
            this.tabControl1.Size = new System.Drawing.Size(974, 117);
            this.tabControl1.TabIndex = 8;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.dgError);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(966, 91);
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
            this.dgError.Size = new System.Drawing.Size(960, 85);
            this.dgError.TabIndex = 0;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.dgRepet);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(966, 91);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Excel数据重复";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // dgRepet
            // 
            this.dgRepet.AllowEditing = false;
            this.dgRepet.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgRepet.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgRepet.Location = new System.Drawing.Point(3, 3);
            this.dgRepet.Name = "dgRepet";
            this.dgRepet.Rows.DefaultSize = 20;
            this.dgRepet.Size = new System.Drawing.Size(960, 85);
            this.dgRepet.TabIndex = 1;
            // 
            // txtFile1
            // 
            this.txtFile1.Location = new System.Drawing.Point(176, 12);
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
            this.label1.Text = "Excel文件-工作中心类别：";
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(709, 171);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(75, 23);
            this.btnImport.TabIndex = 6;
            this.btnImport.Text = "开始导入";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // btnSelect1
            // 
            this.btnSelect1.Location = new System.Drawing.Point(646, 11);
            this.btnSelect1.Name = "btnSelect1";
            this.btnSelect1.Size = new System.Drawing.Size(75, 23);
            this.btnSelect1.TabIndex = 2;
            this.btnSelect1.Text = "选择文件";
            this.btnSelect1.UseVisualStyleBackColor = true;
            this.btnSelect1.Click += new System.EventHandler(this.btnSelect1_Click);
            // 
            // btnCheck
            // 
            this.btnCheck.Location = new System.Drawing.Point(628, 171);
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
            this.label2.Location = new System.Drawing.Point(21, 176);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "导入说明：";
            // 
            // rbSql
            // 
            this.rbSql.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.rbSql.Location = new System.Drawing.Point(0, 368);
            this.rbSql.Name = "rbSql";
            this.rbSql.Size = new System.Drawing.Size(974, 110);
            this.rbSql.TabIndex = 27;
            this.rbSql.Text = "";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.txtFile1);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.btnImport);
            this.panel1.Controls.Add(this.btnSelect1);
            this.panel1.Controls.Add(this.btnCheck);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(974, 214);
            this.panel1.TabIndex = 24;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.dgRepetFromDB);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(966, 91);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "数据库记录重复";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // dgRepetFromDB
            // 
            this.dgRepetFromDB.AllowEditing = false;
            this.dgRepetFromDB.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.dgRepetFromDB.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgRepetFromDB.Location = new System.Drawing.Point(0, 0);
            this.dgRepetFromDB.Name = "dgRepetFromDB";
            this.dgRepetFromDB.Rows.DefaultSize = 20;
            this.dgRepetFromDB.Size = new System.Drawing.Size(966, 91);
            this.dgRepetFromDB.TabIndex = 2;
            // 
            // BSWCTYPE
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(974, 478);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.rbSql);
            this.Name = "BSWCTYPE";
            this.Text = "工作中心类别";
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgError)).EndInit();
            this.tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgRepet)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgRepetFromDB)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.OpenFileDialog opfDialog;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private C1.Win.C1FlexGrid.C1FlexGrid dgError;
        private System.Windows.Forms.TabPage tabPage2;
        private C1.Win.C1FlexGrid.C1FlexGrid dgRepet;
        private System.Windows.Forms.TextBox txtFile1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.Button btnSelect1;
        private System.Windows.Forms.Button btnCheck;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.RichTextBox rbSql;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TabPage tabPage3;
        private C1.Win.C1FlexGrid.C1FlexGrid dgRepetFromDB;
    }
}