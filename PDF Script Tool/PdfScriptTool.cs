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
    using ListString = System.Collections.Generic.List<string>;
    using MessageBox = System.Windows.Forms.MessageBox;
    using Path = System.IO.Path;
    using Resources = Properties.Resources;
    using StringComparison = System.StringComparison;
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
                    PdfProcessor.OutputPath);
                PdfProcessor.ClearProcessing();
                fileView.Items.Clear();
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
            await PerformTask(() => pdfProcessor.ProcessFiles(this));
        }

        /// <summary>
        /// Checks whether or not a file was already selected.
        /// </summary>
        /// <param name="filename">The file to check.</param>
        /// <param name="selectedFilenames">The selected filenames.</param>
        /// <param name="filenameWithoutExtensionReturn">
        /// The filename without an extension, passed back if a duplicate was 
        /// found.
        /// </param>
        /// <returns>Whether or not the file was already selected.</returns>
        private bool FileIsAlreadySelected(
                    string filename,
                    ListString selectedFilenames,
                    out string filenameWithoutExtensionReturn)
        {
            filenameWithoutExtensionReturn = null;
            foreach (var selectedFilename in selectedFilenames)
            {
                var filenameWithoutExtension =
                    Path.GetFileNameWithoutExtension(filename);
                var selectedFilenameWithoutExtension =
                    Path.GetFileNameWithoutExtension(selectedFilename);
                if (string.Equals(
                    filenameWithoutExtension,
                    selectedFilenameWithoutExtension,
                    StringComparison.InvariantCultureIgnoreCase))
                {
                    filenameWithoutExtensionReturn = filenameWithoutExtension;
                    return true;
                }
            }

            return false;
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
                    string filenameWithoutExtension;
                    if (FileIsAlreadySelected(
                        filename,
                        fileView.CheckedItems.OfType<string>().ToList(),
                        out filenameWithoutExtension))
                    {
                        ShowMessage("A file with the name \"" +
                                filenameWithoutExtension +
                                "\" is already selected.");
                    }
                    else
                    {
                        fileView.Items.Add(
                            PdfProcessor.CopyFileToProcessing(filename),
                            FileViewFileIsChecked);
                    }
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
            await PerformTask(() => pdfProcessor.ProcessFiles(
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
            await PerformTask(() => pdfProcessor.ProcessFiles(
                this,
                Field.DefaultTimeStampField,
                Script.TimeStampOnPrintDefaultMonthScript));
        }
    }
}