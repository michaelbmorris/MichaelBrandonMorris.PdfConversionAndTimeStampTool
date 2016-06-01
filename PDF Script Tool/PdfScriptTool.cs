using static PdfScriptTool.PdfScriptToolConstants;

namespace PdfScriptTool
{
    internal partial class PdfScriptTool : System.Windows.Forms.Form,
        System.IProgress<ProgressReport>
    {
        #region Folders

        #region RootFolder

        private static string RootPath = System.IO.Path.Combine(
            System.Environment.GetFolderPath(
                System.Environment.SpecialFolder.MyDocuments),
            RootFolderName);

        #endregion

        private static string ConfigurationPath = System.IO.Path.Combine(
            RootPath,
            ConfigurationFolderName);

        private static string OutputRootPath = System.IO.Path.Combine(
            RootPath,
            OutputFolderName);

        private static string ProcessingPath = System.IO.Path.Combine(
            RootPath,
            ProcessingFolderName);
        
        #endregion Folders

        internal PdfScriptTool()
        {
            InitializeComponent();
            InitializeOpenFileDialog();
            System.IO.Directory.CreateDirectory(RootPath);
            System.IO.Directory.CreateDirectory(OutputRootPath);
            System.IO.Directory.CreateDirectory(ConfigurationPath);
            System.IO.Directory.CreateDirectory(ProcessingPath);
        }

        public void Report(ProgressReport progressReport)
        {
            if (InvokeRequired)
            {
                Invoke((System.Action)(() => Report(progressReport)));
            }
            else
            {
                progressBar.Value = progressReport.Percent;
                progressBar.Text = progressReport.CurrentCount
                    + ProgressLabelDivider + progressReport.Total;
            }
        }

        private static string GetOutputPath(string inputPath)
        {
            return System.IO.Path.Combine(
                OutputRootPath,
                System.IO.Path.GetFileName(inputPath));
        }

        private async System.Threading.Tasks.Task AddScriptToMultiplePdfs(
            Script script)
        {
            await System.Threading.Tasks.Task.Run(() =>
            {
                for (int i = 0; i < documentsView.CheckedItems.Count; i++)
                {
                    var currentDocument
                        = documentsView.CheckedItems[i].ToString();
                    if (!string.Equals(
                        System.IO.Path.GetExtension(currentDocument),
                        ".pdf",
                        System.StringComparison.InvariantCultureIgnoreCase))
                    {
                        try
                        {
                            currentDocument = ConvertToPdf(currentDocument);
                        }
                        catch (System.Runtime.InteropServices.COMException e)
                        {
                            ShowException("File " + currentDocument + " could not be converted to PDF.");
                            continue;
                        }
                    }
                    AddScriptToSinglePdf(currentDocument, script);
                    Report(new ProgressReport
                    {
                        Total = documentsView.CheckedItems.Count,
                        CurrentCount = i + 1
                    });
                }
            });
            System.Windows.Forms.MessageBox.Show(
                "Files saved with time-stamp on print script in "
                + OutputRootPath);
        }

        private void ShowException(System.Exception e)
        {
            ShowException(e.Message);
        }

        private void ShowException(string message)
        {
            System.Windows.Forms.MessageBox.Show(message);
        }

        private void AddScriptToSinglePdf(string filename, Script script)
        {
            try
            {
                using (var pdfReader
                    = new iTextSharp.text.pdf.PdfReader(filename))
                {
                    using (var pdfStamper
                        = new iTextSharp.text.pdf.PdfStamper(
                            pdfReader,
                            new System.IO.FileStream(
                                GetOutputPath(filename),
                                System.IO.FileMode.Create)))
                    {
                        var parentField
                            = iTextSharp.text.pdf.PdfFormField.CreateTextField(
                                pdfStamper.Writer,
                                false,
                                false,
                                0);
                        parentField.FieldName = script.Field.Title;
                        for (var pageNumber = PdfFirstPageNumber;
                            pageNumber <= pdfReader.NumberOfPages;
                            pageNumber++)
                        {
                            var pdfContentByte = pdfStamper.GetOverContent(
                                pageNumber);
                            var textField = new iTextSharp.text.pdf.TextField(
                                    pdfStamper.Writer,
                                    new iTextSharp.text.Rectangle(
                                        script.Field.TopLeftX,
                                        script.Field.TopLeftY,
                                        script.Field.BottomRightX,
                                        script.Field.BottomRightY),
                                    null);
                            var childField = textField.GetTextField();
                            parentField.AddKid(childField);
                            childField.PlaceInPage = pageNumber;
                        }
                        pdfStamper.AddAnnotation(parentField, 1);
                        var pdfAction
                            = iTextSharp.text.pdf.PdfAction.JavaScript(
                                script.Text,
                                pdfStamper.Writer);
                        iTextSharp.text.pdf.PdfName actionType = null;
                        switch (script.Timing)
                        {
                            case ScriptTiming.DidPrint:
                                actionType
                                    = iTextSharp.text.pdf.PdfWriter.DID_PRINT;
                                break;

                            case ScriptTiming.DidSave:
                                actionType
                                    = iTextSharp.text.pdf.PdfWriter.DID_SAVE;
                                break;

                            case ScriptTiming.WillPrint:
                                actionType
                                    = iTextSharp.text.pdf.PdfWriter.WILL_PRINT;
                                break;

                            case ScriptTiming.WillSave:
                                actionType
                                    = iTextSharp.text.pdf.PdfWriter.WILL_SAVE;
                                break;
                        }
                        pdfStamper.Writer.SetAdditionalAction(
                            actionType,
                            pdfAction);
                    }
                }
            }
            catch (iTextSharp.text.exceptions.InvalidPdfException e)
            {
                System.Windows.Forms.MessageBox.Show(
                    e.Message + " " + filename);
            }
        }

        private string ConvertToPdf(string filename)
        {
            var pdfPath = System.IO.Path.Combine(
                ProcessingPath,
                System.IO.Path.GetFileNameWithoutExtension(filename) + ".pdf");
            var wordApplication
                = new Microsoft.Office.Interop.Word.Application();
            var wordDocument
                = wordApplication.Documents.Open(filename);
            wordDocument.ExportAsFixedFormat(
                pdfPath,
                Microsoft.Office.Interop.Word.WdExportFormat
                .wdExportFormatPDF);
            wordDocument.Close(false);
            wordApplication.Quit();
            return pdfPath;
        }

        private void InitializeOpenFileDialog()
        {
            openFileDialog.Filter = OpenFileDialogFilter;
            openFileDialog.Multiselect = OpenFileDialogAllowMultiple;
            openFileDialog.Title = OpenFileDialogTitle;
        }

        private async System.Threading.Tasks.Task PerformTask(
            System.Func<System.Threading.Tasks.Task> function)
        {
            if (documentsView.CheckedItems.Count > 0)
            {
                Enabled = false;
                try
                {
                    await function();
                }
                catch (System.Exception e)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Exception: " + e.Message);
                }
                progressBar.Value = 0;
                progressBar.Text = string.Empty;
                Enabled = true;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show(
                    "Please select at least one document.");
            }
        }
        private void selectDocuments_Click(object sender, System.EventArgs e)
        {
            var dialogResult = openFileDialog.ShowDialog();
            if (dialogResult == System.Windows.Forms.DialogResult.OK)
            {
                foreach (var file in openFileDialog.FileNames)
                {
                    documentsView.Items.Add(file, DocumentsViewFileIsChecked);
                }
            }
        }

        private async void timeStampDefaultDay_Click(object sender,
            System.EventArgs e)
        {
            await PerformTask(() => AddScriptToMultiplePdfs(
                Script.TimeStampOnPrintDefaultDay));
        }
        private async void timeStampDefaultMonth_Click(
            object sender, System.EventArgs e)
        {
            await PerformTask(() => AddScriptToMultiplePdfs(
                Script.TimeStampOnPrintDefaultMonth));
        }
    }
}