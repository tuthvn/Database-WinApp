namespace WinApp.Test
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
            this.DataGridView1 = new System.Windows.Forms.DataGridView();
            this.Button1 = new System.Windows.Forms.Button();
            this.cboSheetName = new System.Windows.Forms.ComboBox();
            this.loadData = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.DataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // DataGridView1
            // 
            this.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DataGridView1.Location = new System.Drawing.Point(15, 50);
            this.DataGridView1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.DataGridView1.Name = "DataGridView1";
            this.DataGridView1.Size = new System.Drawing.Size(718, 241);
            this.DataGridView1.TabIndex = 1;
            // 
            // Button1
            // 
            this.Button1.Location = new System.Drawing.Point(616, 14);
            this.Button1.Margin = new System.Windows.Forms.Padding(5, 3, 5, 3);
            this.Button1.Name = "Button1";
            this.Button1.Size = new System.Drawing.Size(117, 29);
            this.Button1.TabIndex = 0;
            this.Button1.Text = "Open";
            this.Button1.UseVisualStyleBackColor = true;
            this.Button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // cboSheetName
            // 
            this.cboSheetName.FormattingEnabled = true;
            this.cboSheetName.Location = new System.Drawing.Point(15, 12);
            this.cboSheetName.Name = "cboSheetName";
            this.cboSheetName.Size = new System.Drawing.Size(228, 23);
            this.cboSheetName.TabIndex = 2;
            this.cboSheetName.SelectedIndexChanged += new System.EventHandler(this.cboSheetName_SelectedIndexChanged);
            // 
            // loadData
            // 
            this.loadData.Location = new System.Drawing.Point(249, 11);
            this.loadData.Name = "loadData";
            this.loadData.Size = new System.Drawing.Size(106, 32);
            this.loadData.TabIndex = 3;
            this.loadData.Text = "Load";
            this.loadData.UseVisualStyleBackColor = true;
            this.loadData.Click += new System.EventHandler(this.loadData_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(748, 305);
            this.Controls.Add(this.loadData);
            this.Controls.Add(this.cboSheetName);
            this.Controls.Add(this.DataGridView1);
            this.Controls.Add(this.Button1);
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.Name = "Form1";
            this.Text = "Excel";
            ((System.ComponentModel.ISupportInitialize)(this.DataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.DataGridView DataGridView1;
        internal System.Windows.Forms.Button Button1;
        private ComboBox cboSheetName;
        private Button loadData;
    }
}