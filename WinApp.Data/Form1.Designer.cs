namespace WinApp.Data
{
    partial class ToolApp
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.loadButton = new System.Windows.Forms.Button();
            this.Data = new System.Windows.Forms.GroupBox();
            this.comboBoxSheet = new System.Windows.Forms.ComboBox();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.copyButton = new System.Windows.Forms.Button();
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.linkFile = new System.Windows.Forms.TextBox();
            this.Data.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // loadButton
            // 
            this.loadButton.Location = new System.Drawing.Point(12, 16);
            this.loadButton.Name = "loadButton";
            this.loadButton.Size = new System.Drawing.Size(117, 23);
            this.loadButton.TabIndex = 0;
            this.loadButton.Text = "Load";
            this.loadButton.UseVisualStyleBackColor = true;
            this.loadButton.Click += new System.EventHandler(this.loadButton_Click);
            // 
            // Data
            // 
            this.Data.Controls.Add(this.comboBoxSheet);
            this.Data.Controls.Add(this.button3);
            this.Data.Controls.Add(this.button2);
            this.Data.Controls.Add(this.copyButton);
            this.Data.Controls.Add(this.dataGridView);
            this.Data.Location = new System.Drawing.Point(12, 47);
            this.Data.Name = "Data";
            this.Data.Size = new System.Drawing.Size(776, 393);
            this.Data.TabIndex = 1;
            this.Data.TabStop = false;
            this.Data.Text = "Data";
            // 
            // comboBoxSheet
            // 
            this.comboBoxSheet.FormattingEnabled = true;
            this.comboBoxSheet.Location = new System.Drawing.Point(649, 22);
            this.comboBoxSheet.Name = "comboBoxSheet";
            this.comboBoxSheet.Size = new System.Drawing.Size(121, 23);
            this.comboBoxSheet.TabIndex = 4;
            this.comboBoxSheet.SelectedIndexChanged += new System.EventHandler(this.comboBoxSheet_SelectedIndexChanged);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(649, 109);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(121, 23);
            this.button3.TabIndex = 3;
            this.button3.Text = "button3";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(649, 80);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(121, 23);
            this.button2.TabIndex = 2;
            this.button2.Text = "button2";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // copyButton
            // 
            this.copyButton.Location = new System.Drawing.Point(649, 51);
            this.copyButton.Name = "copyButton";
            this.copyButton.Size = new System.Drawing.Size(121, 23);
            this.copyButton.TabIndex = 1;
            this.copyButton.Text = "Copy";
            this.copyButton.UseVisualStyleBackColor = true;
            this.copyButton.Click += new System.EventHandler(this.copyButton_Click);
            // 
            // dataGridView
            // 
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Location = new System.Drawing.Point(6, 22);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.RowTemplate.Height = 25;
            this.dataGridView.Size = new System.Drawing.Size(635, 365);
            this.dataGridView.TabIndex = 0;
            // 
            // linkFile
            // 
            this.linkFile.Location = new System.Drawing.Point(135, 16);
            this.linkFile.Name = "linkFile";
            this.linkFile.Size = new System.Drawing.Size(653, 23);
            this.linkFile.TabIndex = 0;
            // 
            // ToolApp
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 453);
            this.Controls.Add(this.linkFile);
            this.Controls.Add(this.Data);
            this.Controls.Add(this.loadButton);
            this.Name = "ToolApp";
            this.Text = "ToolApp";
            this.Data.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Button loadButton;
        private GroupBox Data;
        private Button button3;
        private Button button2;
        private Button copyButton;
        private DataGridView dataGridView;
        private TextBox linkFile;
        private ComboBox comboBoxSheet;
    }
}