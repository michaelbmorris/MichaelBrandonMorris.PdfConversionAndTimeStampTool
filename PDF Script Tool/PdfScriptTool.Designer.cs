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
            this.timeStampDocuments = new System.Windows.Forms.Button();
            this.documentsView = new System.Windows.Forms.CheckedListBox();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.progressLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // selectDocuments
            // 
            this.selectDocuments.Location = new System.Drawing.Point(8, 8);
            this.selectDocuments.Name = "selectDocuments";
            this.selectDocuments.Size = new System.Drawing.Size(128, 23);
            this.selectDocuments.TabIndex = 0;
            this.selectDocuments.Text = "Select Documents...";
            this.selectDocuments.UseVisualStyleBackColor = true;
            this.selectDocuments.Click += new System.EventHandler(this.selectDocuments_Click);
            // 
            // timeStampDocuments
            // 
            this.timeStampDocuments.Location = new System.Drawing.Point(648, 8);
            this.timeStampDocuments.Name = "timeStampDocuments";
            this.timeStampDocuments.Size = new System.Drawing.Size(128, 23);
            this.timeStampDocuments.TabIndex = 1;
            this.timeStampDocuments.Text = "Time Stamp Documents";
            this.timeStampDocuments.UseVisualStyleBackColor = true;
            this.timeStampDocuments.Click += new System.EventHandler(this.timeStampDocuments_Click);
            // 
            // documentsView
            // 
            this.documentsView.FormattingEnabled = true;
            this.documentsView.Location = new System.Drawing.Point(8, 40);
            this.documentsView.Name = "documentsView";
            this.documentsView.Size = new System.Drawing.Size(768, 304);
            this.documentsView.TabIndex = 3;
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(8, 384);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(768, 23);
            this.progressBar.TabIndex = 4;
            // 
            // progressLabel
            // 
            this.progressLabel.BackColor = System.Drawing.Color.Transparent;
            this.progressLabel.Location = new System.Drawing.Point(8, 360);
            this.progressLabel.Name = "progressLabel";
            this.progressLabel.Size = new System.Drawing.Size(768, 24);
            this.progressLabel.TabIndex = 5;
            this.progressLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // PdfScriptTool
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 412);
            this.Controls.Add(this.progressLabel);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.documentsView);
            this.Controls.Add(this.timeStampDocuments);
            this.Controls.Add(this.selectDocuments);
            this.Name = "PdfScriptTool";
            this.Text = "PDF Script Tool";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button selectDocuments;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.Button timeStampDocuments;
        private System.Windows.Forms.CheckedListBox documentsView;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label progressLabel;
    }
}

