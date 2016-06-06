//-----------------------------------------------------------------------------------------------------------
// <copyright file="PdfConversionAndTimeStampTool.cs" company="Michael Brandon Morris">
//     Copyright © Michael Brandon Morris 2016
// </copyright>
//-----------------------------------------------------------------------------------------------------------

namespace PdfConversionAndTimeStampTool
{
    using System.Linq;
    using Action = System.Action;
    using DialogResult = System.Windows.Forms.DialogResult;
    using EventArgs = System.EventArgs;
    using Exception = System.Exception;
    using Form = System.Windows.Forms.Form;
    using Func = System.Func<System.Threading.Tasks.Task>;
    using IProgressProgressReport = System.IProgress<ProgressReport>;
    using ListString = System.Collections.Generic.List<string>;
    using MessageBox = System.Windows.Forms.MessageBox;
    using Path = System.IO.Path;
    using Process = System.Diagnostics.Process;
    using Resources = Properties.Resources;
    using StringComparison = System.StringComparison;
    using Task = System.Threading.Tasks.Task;

    internal partial class PdfConversionAndTimeStampToolView : Form,
        IProgressProgressReport
    {
        private PdfProcessor pdfProcessor;

        internal PdfConversionAndTimeStampToolView()
        {
            InitializeComponent();
            InitializeOpenFileDialog();
            pdfProcessor = new PdfProcessor();
        }

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

        private static void ShowException(Exception e)
        {
            ShowMessage(e.Message);
        }

        private static void ShowMessage(string message)
        {
            MessageBox.Show(message);
        }

        private async void ConvertOnly_Click(object sender, EventArgs e)
        {
            await PerformTaskIfFilesSelected(() =>
            pdfProcessor.ProcessFiles(this));
        }

        private bool FileIsAlreadySelected(
                    string filename,
                    ListString selectedFilenames)
        {
            foreach (var selectedFilename in selectedFilenames)
            {
                if (FilenamesWithoutExtensionsAreEqual(
                    filename, selectedFilename))
                {
                    return true;
                }
            }

            return false;
        }

        private bool FilenamesWithoutExtensionsAreEqual(
            string filename1, string filename2)
        {
            var filename1WithoutExtension =
                    Path.GetFileNameWithoutExtension(filename1);
            var filename2WithoutExtension =
                Path.GetFileNameWithoutExtension(filename2);
            if (string.Equals(
                filename1WithoutExtension,
                filename2WithoutExtension,
                StringComparison.InvariantCultureIgnoreCase))
            {
                return true;
            }

            return false;
        }

        private void InitializeOpenFileDialog()
        {
            openFileDialog.Filter = Resources.OpenFileDialogFilter;
            openFileDialog.Multiselect = true;
            openFileDialog.Title = Resources.OpenFileDialogTitle;
        }

        private async Task PerformTask(Func function)
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
            Process.Start(PdfProcessor.OutputPath);
            PdfProcessor.ClearProcessing();
            fileView.Items.Clear();
            progressBar.Value = 0;
            Enabled = true;
        }

        private async Task PerformTaskIfFilesSelected(Func function)
        {
            if (fileView.CheckedItems.Count > 0)
            {
                await PerformTask(function);
            }  
            else
            {
                ShowMessage(Resources.NoFilesSelectedErrorMessage);
            }   
        }

        private void SelectFiles_Click(object sender, EventArgs e)
        {
            var dialogResult = openFileDialog.ShowDialog();
            if (dialogResult == DialogResult.OK)
            {
                foreach (var filename in openFileDialog.FileNames)
                {
                    if (FileIsAlreadySelected(
                        filename,
                        fileView.CheckedItems.OfType<string>().ToList()))
                    {
                        ShowMessage("A file with the name \"" +
                            Path.GetFileNameWithoutExtension(filename) +
                            "\" is already selected.");
                    }
                    else
                    {
                        fileView.Items.Add(
                            PdfProcessor.CopyFileToProcessing(filename),
                            isChecked:true);
                    }
                }
            }
        }

        private async void TimeStampDefaultDay_Click(
            object sender, EventArgs e)
        {
            await PerformTaskIfFilesSelected(() => pdfProcessor.ProcessFiles(
                this,
                Field.DefaultTimeStampField,
                Script.TimeStampOnPrintDefaultDayScript));
        }

        private async void TimeStampDefaultMonth_Click(
            object sender, EventArgs e)
        {
            await PerformTaskIfFilesSelected(() => pdfProcessor.ProcessFiles(
                this,
                Field.DefaultTimeStampField,
                Script.TimeStampOnPrintDefaultMonthScript));
        }
    }
}