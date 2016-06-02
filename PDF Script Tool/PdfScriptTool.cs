//-----------------------------------------------------------------------------------------------------------
// <copyright file="PdfScriptTool.cs" company="Michael Brandon Morris">
//     Copyright © Michael Brandon Morris 2016
// </copyright>
//-----------------------------------------------------------------------------------------------------------

namespace PdfScriptTool
{
    using System.Linq;

    /// <summary>
    /// The main application window.
    /// </summary>
    internal partial class PdfScriptTool : System.Windows.Forms.Form,
        System.IProgress<ProgressReport>
    {
        /// <summary>
        /// Whether or not files in the file view should be automatically
        /// checked (selected).
        /// </summary>
        private const bool FileViewFileIsChecked = true;

        /// <summary>
        /// Whether or not the open file dialog should allow selection of
        /// multiple files.
        /// </summary>
        private const bool OpenFileDialogAllowMultiple = true;

        /// <summary>
        /// The PDF Processor that does the back end work.
        /// </summary>
        private PdfProcessor pdfProcessor;

        /// <summary>
        /// Initializes a new instance of the <see cref="PdfScriptTool"/>
        /// class.
        /// </summary>
        internal PdfScriptTool()
        {
            this.InitializeComponent();
            this.InitializeOpenFileDialog();
            this.pdfProcessor = new PdfProcessor();
        }

        /// <summary>
        /// Reports the progress of the current task.
        /// </summary>
        /// <param name="progressReport">The progress report containing a
        /// current count, total count, and percent.</param>
        public void Report(ProgressReport progressReport)
        {
            if (this.InvokeRequired)
            {
                this.Invoke((System.Action)(() =>
                this.Report(progressReport)));
            }
            else
            {
                this.progressBar.Value = progressReport.Percent;
            }
        }

        /// <summary>
        /// Performs a specified task in the backend.
        /// </summary>
        /// <param name="function">The task to perform.</param>
        /// <returns>The completed task.</returns>
        internal async System.Threading.Tasks.Task PerformTask(
            System.Func<System.Threading.Tasks.Task> function)
        {
            if (fileView.CheckedItems.Count > 0)
            {
                this.Enabled = false;
                try
                {
                    this.pdfProcessor.Files =
                        this.fileView.CheckedItems.OfType<string>().ToList();
                    await function();
                }
                catch (System.Exception e)
                {
                    this.ShowException(e);
                }

                this.ShowMessage(Properties.Resources.FilesSavedInMessage
                    + PdfProcessor.OutputRootPath);
                this.progressBar.Value = 0;
                this.Enabled = true;
            }
            else
            {
                this.ShowMessage(Properties.Resources.NoFilesSelectedErrorMessage);
            }
        }

        /// <summary>
        /// Listener for the "Convert to PDF Only" button.
        /// </summary>
        /// <param name="sender">The object that triggered the event.</param>
        /// <param name="e">The event arguments.</param>
        private async void ConvertOnly_Click(object sender, System.EventArgs e)
        {
            await this.PerformTask(() =>
            this.pdfProcessor.ProcessPdfs(this));
        }

        /// <summary>
        /// Sets attributes for the open file dialog.
        /// </summary>
        private void InitializeOpenFileDialog()
        {
            this.openFileDialog.Filter =
                Properties.Resources.OpenFileDialogFilter;
            this.openFileDialog.Multiselect =
                OpenFileDialogAllowMultiple;
            this.openFileDialog.Title =
                Properties.Resources.OpenFileDialogTitle;
        }

        /// <summary>
        /// Listener for the "Select Files" button.
        /// </summary>
        /// <param name="sender">The object that triggered the event.</param>
        /// <param name="e">The event arguments.</param>
        private void SelectFiles_Click(object sender, System.EventArgs e)
        {
            var dialogResult = this.openFileDialog.ShowDialog();
            if (dialogResult == System.Windows.Forms.DialogResult.OK)
            {
                foreach (var file in this.openFileDialog.FileNames)
                {
                    this.fileView.Items.Add(file, FileViewFileIsChecked);
                }
            }
        }

        /// <summary>
        /// Shows an exception in a message box.
        /// </summary>
        /// <param name="e">The exception to show.</param>
        private void ShowException(System.Exception e)
        {
            this.ShowMessage(e.Message);
        }

        /// <summary>
        /// Shows a message in a message box.
        /// </summary>
        /// <param name="message">The message to show.</param>
        private void ShowMessage(string message)
        {
            System.Windows.Forms.MessageBox.Show(message);
        }

        /// <summary>
        /// Listener for the "Timestamp 24 Hours" button.
        /// </summary>
        /// <param name="sender">The object that triggered the event.</param>
        /// <param name="e">The event arguments.</param>
        private async void TimeStampDefaultDay_Click(
            object sender, System.EventArgs e)
        {
            await this.PerformTask(() =>
            this.pdfProcessor.ProcessPdfs(
                this,
                Field.DefaultTimeStampField,
                Script.TimeStampOnPrintDefaultDayScript));
        }

        /// <summary>
        /// Listener for the "Timestamp 30 Days" button.
        /// </summary>
        /// <param name="sender">The object that triggered the event.</param>
        /// <param name="e">The event arguments.</param>
        private async void TimeStampDefaultMonth_Click(
            object sender, System.EventArgs e)
        {
            await this.PerformTask(() =>
            this.pdfProcessor.ProcessPdfs(
                this,
                Field.DefaultTimeStampField,
                Script.TimeStampOnPrintDefaultMonthScript));
        }
    }
}