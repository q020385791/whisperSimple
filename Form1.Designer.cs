namespace WhisperSingle
{
    partial class Form1
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
            this.btnConvert = new System.Windows.Forms.Button();
            this.txtTargetFolder = new System.Windows.Forms.TextBox();
            this.labAudioFolder = new System.Windows.Forms.Label();
            this.btnAudioFilePath = new System.Windows.Forms.Button();
            this.btnSourceExcelPath = new System.Windows.Forms.Button();
            this.labSourceExcelPath = new System.Windows.Forms.Label();
            this.txtSourceExcelPath = new System.Windows.Forms.TextBox();
            this.txtMatchFilePath = new System.Windows.Forms.TextBox();
            this.txtNotmatchFilePath = new System.Windows.Forms.TextBox();
            this.labMatchFilePath = new System.Windows.Forms.Label();
            this.labNotMatchFilePath = new System.Windows.Forms.Label();
            this.btnMatchFilePath = new System.Windows.Forms.Button();
            this.btnNotMatchFilePath = new System.Windows.Forms.Button();
            this.labKeyWord = new System.Windows.Forms.Label();
            this.txtKeyWord = new System.Windows.Forms.TextBox();
            this.lblProcessedCount = new System.Windows.Forms.Label();
            this.lblTotalCount = new System.Windows.Forms.Label();
            this.ckOnlymatchExcel = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // btnConvert
            // 
            this.btnConvert.Location = new System.Drawing.Point(376, 192);
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Size = new System.Drawing.Size(125, 23);
            this.btnConvert.TabIndex = 0;
            this.btnConvert.Text = "轉換";
            this.btnConvert.UseVisualStyleBackColor = true;
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);
            // 
            // txtTargetFolder
            // 
            this.txtTargetFolder.Location = new System.Drawing.Point(121, 20);
            this.txtTargetFolder.Name = "txtTargetFolder";
            this.txtTargetFolder.Size = new System.Drawing.Size(257, 23);
            this.txtTargetFolder.TabIndex = 1;
            // 
            // labAudioFolder
            // 
            this.labAudioFolder.AutoSize = true;
            this.labAudioFolder.Location = new System.Drawing.Point(34, 23);
            this.labAudioFolder.Name = "labAudioFolder";
            this.labAudioFolder.Size = new System.Drawing.Size(79, 15);
            this.labAudioFolder.TabIndex = 2;
            this.labAudioFolder.Text = "音訊檔資料夾";
            // 
            // btnAudioFilePath
            // 
            this.btnAudioFilePath.Location = new System.Drawing.Point(389, 19);
            this.btnAudioFilePath.Name = "btnAudioFilePath";
            this.btnAudioFilePath.Size = new System.Drawing.Size(112, 23);
            this.btnAudioFilePath.TabIndex = 3;
            this.btnAudioFilePath.Text = "選擇音訊資料夾";
            this.btnAudioFilePath.UseVisualStyleBackColor = true;
            this.btnAudioFilePath.Click += new System.EventHandler(this.btnAudioFilePath_Click);
            // 
            // btnSourceExcelPath
            // 
            this.btnSourceExcelPath.Location = new System.Drawing.Point(389, 47);
            this.btnSourceExcelPath.Name = "btnSourceExcelPath";
            this.btnSourceExcelPath.Size = new System.Drawing.Size(112, 23);
            this.btnSourceExcelPath.TabIndex = 6;
            this.btnSourceExcelPath.Text = "csv資料夾";
            this.btnSourceExcelPath.UseVisualStyleBackColor = true;
            this.btnSourceExcelPath.Click += new System.EventHandler(this.btnSourceExcelPath_Click);
            // 
            // labSourceExcelPath
            // 
            this.labSourceExcelPath.AutoSize = true;
            this.labSourceExcelPath.Location = new System.Drawing.Point(29, 51);
            this.labSourceExcelPath.Name = "labSourceExcelPath";
            this.labSourceExcelPath.Size = new System.Drawing.Size(84, 15);
            this.labSourceExcelPath.TabIndex = 5;
            this.labSourceExcelPath.Text = "csv來源資料夾";
            // 
            // txtSourceExcelPath
            // 
            this.txtSourceExcelPath.Location = new System.Drawing.Point(121, 47);
            this.txtSourceExcelPath.Name = "txtSourceExcelPath";
            this.txtSourceExcelPath.Size = new System.Drawing.Size(257, 23);
            this.txtSourceExcelPath.TabIndex = 4;
            // 
            // txtMatchFilePath
            // 
            this.txtMatchFilePath.Location = new System.Drawing.Point(121, 105);
            this.txtMatchFilePath.Name = "txtMatchFilePath";
            this.txtMatchFilePath.Size = new System.Drawing.Size(250, 23);
            this.txtMatchFilePath.TabIndex = 7;
            // 
            // txtNotmatchFilePath
            // 
            this.txtNotmatchFilePath.Location = new System.Drawing.Point(121, 135);
            this.txtNotmatchFilePath.Name = "txtNotmatchFilePath";
            this.txtNotmatchFilePath.Size = new System.Drawing.Size(250, 23);
            this.txtNotmatchFilePath.TabIndex = 8;
            // 
            // labMatchFilePath
            // 
            this.labMatchFilePath.AutoSize = true;
            this.labMatchFilePath.Location = new System.Drawing.Point(24, 109);
            this.labMatchFilePath.Name = "labMatchFilePath";
            this.labMatchFilePath.Size = new System.Drawing.Size(91, 15);
            this.labMatchFilePath.TabIndex = 9;
            this.labMatchFilePath.Text = "符合關鍵字檔案";
            // 
            // labNotMatchFilePath
            // 
            this.labNotMatchFilePath.AutoSize = true;
            this.labNotMatchFilePath.Location = new System.Drawing.Point(12, 139);
            this.labNotMatchFilePath.Name = "labNotMatchFilePath";
            this.labNotMatchFilePath.Size = new System.Drawing.Size(103, 15);
            this.labNotMatchFilePath.TabIndex = 10;
            this.labNotMatchFilePath.Text = "不符合關鍵字檔案";
            // 
            // btnMatchFilePath
            // 
            this.btnMatchFilePath.Location = new System.Drawing.Point(377, 105);
            this.btnMatchFilePath.Name = "btnMatchFilePath";
            this.btnMatchFilePath.Size = new System.Drawing.Size(125, 23);
            this.btnMatchFilePath.TabIndex = 11;
            this.btnMatchFilePath.Text = "符合關鍵字資料夾";
            this.btnMatchFilePath.UseVisualStyleBackColor = true;
            this.btnMatchFilePath.Click += new System.EventHandler(this.btnMatchFilePath_Click);
            // 
            // btnNotMatchFilePath
            // 
            this.btnNotMatchFilePath.Location = new System.Drawing.Point(376, 135);
            this.btnNotMatchFilePath.Name = "btnNotMatchFilePath";
            this.btnNotMatchFilePath.Size = new System.Drawing.Size(125, 23);
            this.btnNotMatchFilePath.TabIndex = 12;
            this.btnNotMatchFilePath.Text = "不符合關鍵字資料夾";
            this.btnNotMatchFilePath.UseVisualStyleBackColor = true;
            this.btnNotMatchFilePath.Click += new System.EventHandler(this.btnNotMatchFilePath_Click);
            // 
            // labKeyWord
            // 
            this.labKeyWord.AutoSize = true;
            this.labKeyWord.Location = new System.Drawing.Point(72, 84);
            this.labKeyWord.Name = "labKeyWord";
            this.labKeyWord.Size = new System.Drawing.Size(43, 15);
            this.labKeyWord.TabIndex = 13;
            this.labKeyWord.Text = "關鍵字";
            // 
            // txtKeyWord
            // 
            this.txtKeyWord.Location = new System.Drawing.Point(121, 76);
            this.txtKeyWord.Name = "txtKeyWord";
            this.txtKeyWord.Size = new System.Drawing.Size(100, 23);
            this.txtKeyWord.TabIndex = 14;
            // 
            // lblProcessedCount
            // 
            this.lblProcessedCount.AutoSize = true;
            this.lblProcessedCount.Location = new System.Drawing.Point(24, 173);
            this.lblProcessedCount.Name = "lblProcessedCount";
            this.lblProcessedCount.Size = new System.Drawing.Size(79, 15);
            this.lblProcessedCount.TabIndex = 15;
            this.lblProcessedCount.Text = "已處理檔案數";
            // 
            // lblTotalCount
            // 
            this.lblTotalCount.AutoSize = true;
            this.lblTotalCount.Location = new System.Drawing.Point(24, 200);
            this.lblTotalCount.Name = "lblTotalCount";
            this.lblTotalCount.Size = new System.Drawing.Size(55, 15);
            this.lblTotalCount.TabIndex = 16;
            this.lblTotalCount.Text = "檔案總數";
            // 
            // ckOnlymatchExcel
            // 
            this.ckOnlymatchExcel.AutoSize = true;
            this.ckOnlymatchExcel.Location = new System.Drawing.Point(355, 76);
            this.ckOnlymatchExcel.Name = "ckOnlymatchExcel";
            this.ckOnlymatchExcel.Size = new System.Drawing.Size(146, 19);
            this.ckOnlymatchExcel.TabIndex = 17;
            this.ckOnlymatchExcel.Text = "不轉音訊，僅對應檔名";
            this.ckOnlymatchExcel.UseVisualStyleBackColor = true;
            this.ckOnlymatchExcel.Click += new System.EventHandler(this.ckOnlymatchExcel_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(526, 243);
            this.Controls.Add(this.ckOnlymatchExcel);
            this.Controls.Add(this.lblTotalCount);
            this.Controls.Add(this.lblProcessedCount);
            this.Controls.Add(this.txtKeyWord);
            this.Controls.Add(this.labKeyWord);
            this.Controls.Add(this.btnNotMatchFilePath);
            this.Controls.Add(this.btnMatchFilePath);
            this.Controls.Add(this.labNotMatchFilePath);
            this.Controls.Add(this.labMatchFilePath);
            this.Controls.Add(this.txtNotmatchFilePath);
            this.Controls.Add(this.txtMatchFilePath);
            this.Controls.Add(this.btnSourceExcelPath);
            this.Controls.Add(this.labSourceExcelPath);
            this.Controls.Add(this.txtSourceExcelPath);
            this.Controls.Add(this.btnAudioFilePath);
            this.Controls.Add(this.labAudioFolder);
            this.Controls.Add(this.txtTargetFolder);
            this.Controls.Add(this.btnConvert);
            this.Name = "Form1";
            this.Text = "音訊轉檔輔助程式";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Button btnConvert;
        private TextBox txtTargetFolder;
        private Label labAudioFolder;
        private Button btnAudioFilePath;
        private Button btnSourceExcelPath;
        private Label labSourceExcelPath;
        private TextBox txtSourceExcelPath;
        private TextBox txtMatchFilePath;
        private TextBox txtNotmatchFilePath;
        private Label labMatchFilePath;
        private Label labNotMatchFilePath;
        private Button btnMatchFilePath;
        private Button btnNotMatchFilePath;
        private Label labKeyWord;
        private TextBox txtKeyWord;
        private Label lblProcessedCount;
        private Label lblTotalCount;
        private CheckBox ckOnlymatchExcel;
    }
}