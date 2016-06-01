namespace PdfScriptTool
{
    partial class PdfScriptTool
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
            this.selectDocuments = new System.Windows.Forms.Button();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.fileView = new System.Windows.Forms.CheckedListBox();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.openCustomScriptDialog = new System.Windows.Forms.OpenFileDialog();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.addScript = new System.Windows.Forms.Button();
            this.selectTiming = new System.Windows.Forms.ComboBox();
            this.scriptSelector = new System.Windows.Forms.ComboBox();
            this.convertOnly = new System.Windows.Forms.Button();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.timeStampDefaultDay = new System.Windows.Forms.Button();
            this.timeStampDefaultMonth = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage2.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.SuspendLayout();
            // 
            // selectDocuments
            // 
            this.selectDocuments.Location = new System.Drawing.Point(8, 0);
            this.selectDocuments.Name = "selectDocuments";
            this.selectDocuments.Size = new System.Drawing.Size(128, 64);
            this.selectDocuments.TabIndex = 0;
            this.selectDocuments.Text = "Select Documents...";
            this.selectDocuments.UseVisualStyleBackColor = true;
            this.selectDocuments.Click += new System.EventHandler(this.selectDocuments_Click);
            // 
            // fileView
            // 
            this.fileView.FormattingEnabled = true;
            this.fileView.Location = new System.Drawing.Point(8, 72);
            this.fileView.Name = "fileView";
            this.fileView.Size = new System.Drawing.Size(768, 304);
            this.fileView.TabIndex = 3;
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(8, 384);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(768, 23);
            this.progressBar.TabIndex = 4;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.addScript);
            this.tabPage2.Controls.Add(this.selectTiming);
            this.tabPage2.Controls.Add(this.scriptSelector);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(624, 38);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Advanced";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // addScript
            // 
            this.addScript.Location = new System.Drawing.Point(248, 8);
            this.addScript.Name = "addScript";
            this.addScript.Size = new System.Drawing.Size(128, 23);
            this.addScript.TabIndex = 1;
            this.addScript.Text = "Add Script";
            this.addScript.UseVisualStyleBackColor = true;
            // 
            // selectTiming
            // 
            this.selectTiming.FormattingEnabled = true;
            this.selectTiming.Location = new System.Drawing.Point(128, 8);
            this.selectTiming.Name = "selectTiming";
            this.selectTiming.Size = new System.Drawing.Size(112, 21);
            this.selectTiming.TabIndex = 7;
            this.selectTiming.Text = "Script should run...";
            // 
            // scriptSelector
            // 
            this.scriptSelector.FormattingEnabled = true;
            this.scriptSelector.Location = new System.Drawing.Point(8, 8);
            this.scriptSelector.Name = "scriptSelector";
            this.scriptSelector.Size = new System.Drawing.Size(112, 21);
            this.scriptSelector.TabIndex = 6;
            this.scriptSelector.Text = "Select a script...";
            // 
            // convertOnly
            // 
            this.convertOnly.Location = new System.Drawing.Point(280, 8);
            this.convertOnly.Name = "convertOnly";
            this.convertOnly.Size = new System.Drawing.Size(128, 23);
            this.convertOnly.TabIndex = 10;
            this.convertOnly.Text = "Convert to PDF Only";
            this.convertOnly.UseVisualStyleBackColor = true;
            this.convertOnly.Click += new System.EventHandler(this.convertOnly_Click);
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.convertOnly);
            this.tabPage1.Controls.Add(this.timeStampDefaultDay);
            this.tabPage1.Controls.Add(this.timeStampDefaultMonth);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(624, 38);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "General";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // timeStampDefaultDay
            // 
            this.timeStampDefaultDay.Location = new System.Drawing.Point(8, 8);
            this.timeStampDefaultDay.Name = "timeStampDefaultDay";
            this.timeStampDefaultDay.Size = new System.Drawing.Size(128, 23);
            this.timeStampDefaultDay.TabIndex = 8;
            this.timeStampDefaultDay.Text = "Time Stamp 24 Hours";
            this.timeStampDefaultDay.UseVisualStyleBackColor = true;
            this.timeStampDefaultDay.Click += new System.EventHandler(this.timeStampDefaultDay_Click);
            // 
            // timeStampDefaultMonth
            // 
            this.timeStampDefaultMonth.Location = new System.Drawing.Point(144, 8);
            this.timeStampDefaultMonth.Name = "timeStampDefaultMonth";
            this.timeStampDefaultMonth.Size = new System.Drawing.Size(128, 23);
            this.timeStampDefaultMonth.TabIndex = 10;
            this.timeStampDefaultMonth.Text = "Time Stamp 30 Days";
            this.timeStampDefaultMonth.UseVisualStyleBackColor = true;
            this.timeStampDefaultMonth.Click += new System.EventHandler(this.timeStampDefaultMonth_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(144, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(632, 64);
            this.tabControl1.TabIndex = 12;
            // 
            // PdfScriptTool
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 412);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.selectDocuments);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.fileView);
            this.Name = "PdfScriptTool";
            this.Text = "PDF Script Tool";
            this.tabPage2.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button selectDocuments;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.CheckedListBox fileView;
        private System.Windows.Forms.OpenFileDialog openCustomScriptDialog;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button convertOnly;
        private System.Windows.Forms.Button addScript;
        private System.Windows.Forms.ComboBox selectTiming;
        private System.Windows.Forms.ComboBox scriptSelector;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Button timeStampDefaultDay;
        private System.Windows.Forms.Button timeStampDefaultMonth;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.ProgressBar progressBar;
    }
}

