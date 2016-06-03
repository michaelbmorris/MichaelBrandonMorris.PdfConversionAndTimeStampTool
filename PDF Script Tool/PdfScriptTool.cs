//-----------------------------------------------------------------------------------------------------------
// <copyright file="PdfScriptTool.cs" company="Michael Brandon Morris">
//     Copyright © Michael Brandon Morris 2016
// </copyright>
//-----------------------------------------------------------------------------------------------------------

namespace PdfScriptTool
{
    using System.Linq;
    using Action = System.Action;
    using DialogResult = System.Windows.Forms.DialogResult;
    using EventArgs = System.EventArgs;
    using Exception = System.Exception;
    using Form = System.Windows.Forms.Form;
    using Func = System.Func<System.Threading.Tasks.Task>;
    using IProgress = System.IProgress<ProgressReport>;
    using MessageBox = System.Windows.Forms.MessageBox;
    using Resources = Properties.Resources;
    using Task = System.Threading.Tasks.Task;

    /// <summary>
    /// The main application window.
    /// </summary>
    internal partial class PdfScriptTool : Form, IProgress
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
            InitializeComponent();
            InitializeOpenFileDialog();
            pdfProcessor = new PdfProcessor();
        }

        /// <summary>
        /// Reports the progress of the current task.
        /// </summary>
        /// <param name="progressReport">The progress report containing a
        /// current count, total count, and percent.</param>
        public void Report(ProgressReport progressReport)
        {
            if (InvokeRequired)
            {
                Invoke((Action)(() => Report(progressReport)));
            }
            else
            {
                progressBar.Value = progressReport.Percent;
            }
        }

        /// <summary>
        /// Performs a specified task in the backend.
        /// </summary>
        /// <param name="function">The task to perform.</param>
        /// <returns>The completed task.</returns>
        internal async Task PerformTask(Func function)
        {
            if (fileView.CheckedItems.Count > 0)
            {
                Enabled = false;
                try
                {
                    pdfProcessor.Files =
                        fileView.CheckedItems.OfType<string>().ToList();
                    await function();
                }
                catch (Exception e)
                {
                    ShowException(e);
                }

                ShowMessage(Resources.FilesSavedInMessage +
                    PdfProcessor.OutputRootPath);
                progressBar.Value = 0;
                Enabled = true;
            }
            else
            {
                ShowMessage(Resources.NoFilesSelectedErrorMessage);
            }
        }

        /// <summary>
        /// Shows an exception in a message box.
        /// </summary>
        /// <param name="e">The exception to show.</param>
        private static void ShowException(Exception e)
        {
            ShowMessage(e.Message);
        }

        /// <summary>
        /// Shows a message in a message box.
        /// </summary>
        /// <param name="message">The message to show.</param>
        private static void ShowMessage(string message)
        {
            MessageBox.Show(message);
        }

        /// <summary>
        /// Listener for the "Convert to PDF Only" button.
        /// </summary>
        /// <param name="sender">The object that triggered the event.</param>
        /// <param name="e">The event arguments.</param>
        private async void ConvertOnly_Click(object sender, EventArgs e)
        {
            await PerformTask(() => pdfProcessor.ProcessPdfs(this));
        }

        /// <summary>
        /// Sets attributes for the open file dialog.
        /// </summary>
        private void InitializeOpenFileDialog()
        {
            openFileDialog.Filter = Resources.OpenFileDialogFilter;
            openFileDialog.Multiselect = OpenFileDialogAllowMultiple;
            openFileDialog.Title = Resources.OpenFileDialogTitle;
        }

        /// <summary>
        /// Listener for the "Select Files" button. Shows the select files
        /// dialog and adds all selected files to the files view, locking each
        /// file to prevent editing until released.
        /// </summary>
        /// <param name="sender">The object that triggered the event.</param>
        /// <param name="e">The event arguments.</param>
        private void SelectFiles_Click(object sender, EventArgs e)
        {
            var dialogResult = openFileDialog.ShowDialog();
            if (dialogResult == DialogResult.OK)
            {
                foreach (var filename in openFileDialog.FileNames)
                {
                    fileView.Items.Add(filename, FileViewFileIsChecked);
                }
            }
        }

        /// <summary>
        /// Listener for the "Timestamp 24 Hours" button.
        /// </summary>
        /// <param name="sender">The object that triggered the event.</param>
        /// <param name="e">The event arguments.</param>
        private async void TimeStampDefaultDay_Click(
            object sender, EventArgs e)
        {
            await PerformTask(() => pdfProcessor.ProcessPdfs(
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
            object sender, EventArgs e)
        {
            await PerformTask(() => pdfProcessor.ProcessPdfs(
                this,
                Field.DefaultTimeStampField,
                Script.TimeStampOnPrintDefaultMonthScript));
        }
    }
}