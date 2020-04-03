namespace WFA_ExcelValidate
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.button1 = new System.Windows.Forms.Button();
            this.resultGrid = new System.Windows.Forms.DataGridView();
            this.txtReportCode = new System.Windows.Forms.TextBox();
            this.txtValidateX = new System.Windows.Forms.TextBox();
            this.txtValidateY = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txtDiff = new System.Windows.Forms.TextBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.button2 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.resultGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(15, 261);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(279, 113);
            this.button1.TabIndex = 0;
            this.button1.Text = "Validate";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // resultGrid
            // 
            this.resultGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.resultGrid.Location = new System.Drawing.Point(300, 12);
            this.resultGrid.Name = "resultGrid";
            this.resultGrid.RowTemplate.Height = 24;
            this.resultGrid.Size = new System.Drawing.Size(1046, 699);
            this.resultGrid.TabIndex = 1;
            // 
            // txtReportCode
            // 
            this.txtReportCode.Location = new System.Drawing.Point(106, 12);
            this.txtReportCode.Multiline = true;
            this.txtReportCode.Name = "txtReportCode";
            this.txtReportCode.Size = new System.Drawing.Size(188, 77);
            this.txtReportCode.TabIndex = 2;
            // 
            // txtValidateX
            // 
            this.txtValidateX.Location = new System.Drawing.Point(106, 95);
            this.txtValidateX.Multiline = true;
            this.txtValidateX.Name = "txtValidateX";
            this.txtValidateX.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtValidateX.Size = new System.Drawing.Size(188, 77);
            this.txtValidateX.TabIndex = 3;
            // 
            // txtValidateY
            // 
            this.txtValidateY.Location = new System.Drawing.Point(106, 178);
            this.txtValidateY.Multiline = true;
            this.txtValidateY.Name = "txtValidateY";
            this.txtValidateY.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtValidateY.Size = new System.Drawing.Size(188, 77);
            this.txtValidateY.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(88, 17);
            this.label1.TabIndex = 5;
            this.label1.Text = "Report Code";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 98);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(72, 17);
            this.label2.TabIndex = 6;
            this.label2.Text = "Validate X";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 181);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(72, 17);
            this.label3.TabIndex = 7;
            this.label3.Text = "Validate Y";
            // 
            // txtDiff
            // 
            this.txtDiff.Location = new System.Drawing.Point(15, 566);
            this.txtDiff.Multiline = true;
            this.txtDiff.Name = "txtDiff";
            this.txtDiff.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtDiff.Size = new System.Drawing.Size(279, 145);
            this.txtDiff.TabIndex = 8;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // button2
            // 
            this.button2.Enabled = false;
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.Location = new System.Drawing.Point(12, 380);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(279, 113);
            this.button2.TabIndex = 9;
            this.button2.Text = "CONTINUE";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1358, 723);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.txtDiff);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtValidateY);
            this.Controls.Add(this.txtValidateX);
            this.Controls.Add(this.txtReportCode);
            this.Controls.Add(this.resultGrid);
            this.Controls.Add(this.button1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "Excel Validate by Meiio";
            ((System.ComponentModel.ISupportInitialize)(this.resultGrid)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.DataGridView resultGrid;
        private System.Windows.Forms.TextBox txtReportCode;
        private System.Windows.Forms.TextBox txtValidateX;
        private System.Windows.Forms.TextBox txtValidateY;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtDiff;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button button2;
    }
}