namespace InvertoryCheck
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnjisuankucun = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btnxitongkucun = new System.Windows.Forms.Button();
            this.btnshijikucun = new System.Windows.Forms.Button();
            this.lblxitongkucun = new System.Windows.Forms.Label();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.lstshijikucun = new System.Windows.Forms.ListBox();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.label1 = new System.Windows.Forms.Label();
            this.btnClear = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnjisuankucun
            // 
            this.btnjisuankucun.Location = new System.Drawing.Point(86, 180);
            this.btnjisuankucun.Name = "btnjisuankucun";
            this.btnjisuankucun.Size = new System.Drawing.Size(75, 49);
            this.btnjisuankucun.TabIndex = 0;
            this.btnjisuankucun.Text = "计算库存";
            this.btnjisuankucun.UseVisualStyleBackColor = true;
            this.btnjisuankucun.Click += new System.EventHandler(this.btnjisuankucun_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            // 
            // btnxitongkucun
            // 
            this.btnxitongkucun.Location = new System.Drawing.Point(240, 14);
            this.btnxitongkucun.Name = "btnxitongkucun";
            this.btnxitongkucun.Size = new System.Drawing.Size(122, 23);
            this.btnxitongkucun.TabIndex = 1;
            this.btnxitongkucun.Text = "选择系统库存";
            this.btnxitongkucun.UseVisualStyleBackColor = true;
            this.btnxitongkucun.Click += new System.EventHandler(this.btnxitongkucun_Click);
            // 
            // btnshijikucun
            // 
            this.btnshijikucun.Location = new System.Drawing.Point(240, 55);
            this.btnshijikucun.Name = "btnshijikucun";
            this.btnshijikucun.Size = new System.Drawing.Size(122, 23);
            this.btnshijikucun.TabIndex = 2;
            this.btnshijikucun.Text = "添加实际库存";
            this.btnshijikucun.UseVisualStyleBackColor = true;
            this.btnshijikucun.Click += new System.EventHandler(this.btnshijikucun_Click);
            // 
            // lblxitongkucun
            // 
            this.lblxitongkucun.Location = new System.Drawing.Point(12, 19);
            this.lblxitongkucun.Name = "lblxitongkucun";
            this.lblxitongkucun.Size = new System.Drawing.Size(149, 23);
            this.lblxitongkucun.TabIndex = 3;
            this.lblxitongkucun.Text = "label1";
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            // 
            // lstshijikucun
            // 
            this.lstshijikucun.BackColor = System.Drawing.SystemColors.Control;
            this.lstshijikucun.FormattingEnabled = true;
            this.lstshijikucun.ItemHeight = 12;
            this.lstshijikucun.Location = new System.Drawing.Point(14, 55);
            this.lstshijikucun.Name = "lstshijikucun";
            this.lstshijikucun.Size = new System.Drawing.Size(195, 100);
            this.lstshijikucun.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(241, 98);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(125, 12);
            this.label1.TabIndex = 5;
            this.label1.Text = "可以添加多个实际库存";
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(205, 180);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(76, 49);
            this.btnClear.TabIndex = 6;
            this.btnClear.Text = "清空";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(377, 262);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lstshijikucun);
            this.Controls.Add(this.lblxitongkucun);
            this.Controls.Add(this.btnshijikucun);
            this.Controls.Add(this.btnxitongkucun);
            this.Controls.Add(this.btnjisuankucun);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "库存检查";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnjisuankucun;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btnxitongkucun;
        private System.Windows.Forms.Button btnshijikucun;
        private System.Windows.Forms.Label lblxitongkucun;
        private System.Windows.Forms.OpenFileDialog openFileDialog2;
        private System.Windows.Forms.ListBox lstshijikucun;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnClear;
    }
}

