namespace ExcelSplitter
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.txtRowCount = new System.Windows.Forms.TextBox();
            this.lblRowCount = new System.Windows.Forms.Label();
            this.btnProcess = new System.Windows.Forms.Button();
            this.lblStartRow = new System.Windows.Forms.Label();
            this.txtStartRow = new System.Windows.Forms.TextBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblSelectedFile = new System.Windows.Forms.Label();
            this.btnSelectFolder = new System.Windows.Forms.Button();
            this.lblSelectedFolder = new System.Windows.Forms.Label();
            this.SuspendLayout();

            // 
            // btnSelectFile
            // 
            this.btnSelectFile.Location = new System.Drawing.Point(12, 12);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(150, 23);
            this.btnSelectFile.TabIndex = 0;
            this.btnSelectFile.Text = "엑셀파일 선택하세요";
            this.btnSelectFile.UseVisualStyleBackColor = true;

            // 
            // txtRowCount
            // 
            this.txtRowCount.Location = new System.Drawing.Point(117, 50);
            this.txtRowCount.Name = "txtRowCount";
            this.txtRowCount.Size = new System.Drawing.Size(100, 20);
            this.txtRowCount.TabIndex = 1;

            // 
            // lblRowCount
            // 
            this.lblRowCount.AutoSize = true;
            this.lblRowCount.Location = new System.Drawing.Point(12, 53);
            this.lblRowCount.Name = "lblRowCount";
            this.lblRowCount.Size = new System.Drawing.Size(76, 13);
            this.lblRowCount.TabIndex = 2;
            this.lblRowCount.Text = "몇행수씩 나눌까요:";

            // 
            // btnProcess
            // 
            this.btnProcess.Location = new System.Drawing.Point(12, 210);
            this.btnProcess.Name = "btnProcess";
            this.btnProcess.Size = new System.Drawing.Size(150, 23);
            this.btnProcess.TabIndex = 3;
            this.btnProcess.Text = "작업 시작";
            this.btnProcess.UseVisualStyleBackColor = true;

            // 
            // lblStartRow
            // 
            this.lblStartRow.AutoSize = true;
            this.lblStartRow.Location = new System.Drawing.Point(12, 85);
            this.lblStartRow.Name = "lblStartRow";
            this.lblStartRow.Size = new System.Drawing.Size(82, 13);
            this.lblStartRow.TabIndex = 4;
            this.lblStartRow.Text = "데이터 시작 행 번호:";

            // 
            // txtStartRow
            // 
            this.txtStartRow.Location = new System.Drawing.Point(117, 82);
            this.txtStartRow.Name = "txtStartRow";
            this.txtStartRow.Size = new System.Drawing.Size(100, 20);
            this.txtStartRow.TabIndex = 5;

            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 150);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(205, 23);
            this.progressBar1.TabIndex = 6;

            // 
            // lblSelectedFile
            // 
            this.lblSelectedFile.AutoSize = true;
            this.lblSelectedFile.Location = new System.Drawing.Point(168, 17);
            this.lblSelectedFile.Name = "lblSelectedFile";
            this.lblSelectedFile.Size = new System.Drawing.Size(0, 13);
            this.lblSelectedFile.TabIndex = 7;

            // 
            // btnSelectFolder
            // 
            this.btnSelectFolder.Location = new System.Drawing.Point(12, 120);
            this.btnSelectFolder.Name = "btnSelectFolder";
            this.btnSelectFolder.Size = new System.Drawing.Size(150, 23);
            this.btnSelectFolder.TabIndex = 8;
            this.btnSelectFolder.Text = "파일 배포할 폴더선택";
            this.btnSelectFolder.UseVisualStyleBackColor = true;

            // 
            // lblSelectedFolder
            // 
            this.lblSelectedFolder.AutoSize = true;
            this.lblSelectedFolder.Location = new System.Drawing.Point(168, 125);
            this.lblSelectedFolder.Name = "lblSelectedFolder";
            this.lblSelectedFolder.Size = new System.Drawing.Size(0, 13);
            this.lblSelectedFolder.TabIndex = 9;

            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(470, 241);
            this.Controls.Add(this.lblSelectedFolder);
            this.Controls.Add(this.btnSelectFolder);
            this.Controls.Add(this.lblSelectedFile);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.txtStartRow);
            this.Controls.Add(this.lblStartRow);
            this.Controls.Add(this.btnProcess);
            this.Controls.Add(this.lblRowCount);
            this.Controls.Add(this.txtRowCount);
            this.Controls.Add(this.btnSelectFile);
            this.Name = "Form1";
            this.Text = "엑셀파일 분할 프로그램";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Button btnSelectFile;
        private System.Windows.Forms.TextBox txtRowCount;
        private System.Windows.Forms.Label lblRowCount;
        private System.Windows.Forms.Button btnProcess;
        private System.Windows.Forms.Label lblStartRow;
        private System.Windows.Forms.TextBox txtStartRow;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblSelectedFile;
        private System.Windows.Forms.Button btnSelectFolder;
        private System.Windows.Forms.Label lblSelectedFolder;
    }
}
