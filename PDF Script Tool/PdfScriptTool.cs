using iTextSharp.text.exceptions;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.draw;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static PdfScriptTool.PdfScriptToolConstants;

namespace PdfScriptTool
{
    internal partial class PdfScriptTool : Form, IProgress<ProgressReport>
    {
        #region Folders

        private static string RootPath = Path.Combine(
            Environment.GetFolderPath(
                Environment.SpecialFolder.MyDocuments),
            RootFolderName);

        private static string ConfigurationPath = Path.Combine(
            RootPath,
            ConfigurationFolderName);

        private static string OutputRootPath = Path.Combine(
            RootPath,
            OutputFolderName);

        private static string ProcessingPath = Path.Combine(
            RootPath,
            ProcessingFolderName);

        private static string TimeStampScriptPath = Path.Combine(
            ConfigurationPath,
            TimeStampScriptFileName);

        #endregion Folders

        internal PdfScriptTool()
        {
            InitializeComponent();
            InitializeOpenFileDialog();
            Directory.CreateDirectory(RootPath);
            Directory.CreateDirectory(OutputRootPath);
            Directory.CreateDirectory(ConfigurationPath);
            Directory.CreateDirectory(ProcessingPath);
        }

        private static string TimeStampScript
        {
            get
            {
                var timeStampScript = string.Empty;
                if (File.Exists(TimeStampScriptPath))
                {
                    using (var reader = new StreamReader(TimeStampScriptPath))
                    {
                        timeStampScript = reader.ReadToEnd();
                    }
                }
                else
                {
                    timeStampScript = DefaultTimeStampScript;
                }
                return timeStampScript;
            }
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
                progressLabel.Text = progressReport.CurrentCount
                    + ProgressLabelDivider + progressReport.Total;
            }
        }

        private async Task PerformTask(Task task)
        {
            if(documentsView.CheckedItems.Count > 0)
            {
                try
                {
                    await task;
                }
                catch(Exception e)
                {
                    MessageBox.Show("Exception: " + e.Message);
                }
                progressBar.Value = 0;
                progressLabel.Text = string.Empty;
                Enabled = true;
            }
            else
            {
                MessageBox.Show("Please select at least one document.");
            }
        }

        private static string GetOutputPath(string inputPath)
        {
            return Path.Combine(OutputRootPath, Path.GetFileName(inputPath));
        }

        // TODO
        private string ConvertToPdf(string filename)
        {
            string pdfPath = Path.Combine(
                ProcessingPath,
                Path.GetFileNameWithoutExtension(filename) + ".pdf");
            return pdfPath;
        }

        private void InitializeOpenFileDialog()
        {
            openFileDialog.Filter = OpenFileDialogFilter;
            openFileDialog.Multiselect = OpenFileDialogAllowMultiple;
            openFileDialog.Title = OpenFileDialogTitle;
        }

        private void selectDocuments_Click(object sender, EventArgs e)
        {
            var dialogResult = openFileDialog.ShowDialog();
            if (dialogResult == DialogResult.OK)
            {
                foreach (var file in openFileDialog.FileNames)
                {
                    documentsView.Items.Add(file, DocumentsViewFileIsChecked);
                }
            }
        }

        private async void timeStampDocuments_Click(object sender, EventArgs e)
        {
            await PerformTask(TimeStampPdfs());
        }

        private Task TimeStampPdfs() => Task.Run(() =>
        {
            for (int i = 0; i < documentsView.CheckedItems.Count; i++)
            {
                if (!string.Equals(
                    Path.GetExtension(
                        documentsView.CheckedItems[i].ToString()),
                    ".pdf",
                    StringComparison.InvariantCultureIgnoreCase))
                {
                    documentsView.CheckedItems[i] = ConvertToPdf(
                        documentsView.CheckedItems[i].ToString());
                }
                TimeStampPdf(documentsView.CheckedItems[i].ToString());
                Report(new ProgressReport
                {
                    Total = documentsView.CheckedItems.Count,
                    CurrentCount = i + 1
                });
            }
        }).ContinueWith(t =>
        {
            MessageBox.Show("Timestamped all files.");
        },
            CancellationToken.None,
            TaskContinuationOptions.OnlyOnRanToCompletion,
            TaskScheduler.FromCurrentSynchronizationContext()
        );

        private void TimeStampPdf(string filename)
        {
            try
            {
                using (var pdfReader = new PdfReader(filename))
                {
                    using (var pdfStamper = new PdfStamper(
                            pdfReader,
                            new FileStream(
                                GetOutputPath(filename),
                                FileMode.Create)))
                    {
                        var parentField = PdfFormField.CreateTextField(
                                pdfStamper.Writer,
                                false,
                                false,
                                0);
                        parentField.FieldName = TimeStampFieldName;
                        var lineSeparator = new LineSeparator();
                        for (var pageNumber = PdfFirstPageNumber;
                            pageNumber <= pdfReader.NumberOfPages;
                            pageNumber++)
                        {
                            var pdfContentByte = pdfStamper.GetOverContent(
                                pageNumber);
                            var textField = new TextField(
                                    pdfStamper.Writer,
                                    new iTextSharp.text.Rectangle(
                                        TimeStampFieldTopLeftXCoordinate,
                                        TimeStampFieldTopLeftYCoordinate,
                                        TimeStampFieldBottomRightXCoordinate,
                                        TimeStampFieldBottomRightYCoordinate),
                                    null);
                            var childField = textField.GetTextField();
                            parentField.AddKid(childField);
                            childField.PlaceInPage = pageNumber;
                        }
                        pdfStamper.AddAnnotation(parentField, 1);
                        var pdfAction = PdfAction.JavaScript(
                                TimeStampScript,
                                pdfStamper.Writer);
                        pdfStamper.Writer.SetAdditionalAction(
                            PdfWriter.WILL_PRINT,
                            pdfAction);
                    }
                }
            }
            catch (InvalidPdfException e)
            {
                MessageBox.Show(e.Message + " " + filename);
            }
        }
    }
}