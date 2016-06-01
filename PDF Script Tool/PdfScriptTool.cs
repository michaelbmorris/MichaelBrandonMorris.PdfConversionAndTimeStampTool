using static PdfScriptTool.Properties.Resources;

namespace PdfScriptTool
{
    internal partial class PdfScriptTool : System.Windows.Forms.Form,
        System.IProgress<ProgressReport>
    {
        #region Constants

        #region Booleans

        private const bool OpenFileDialogAllowMultiple = true;
        private const bool FileViewFileIsChecked = true;

        #endregion Booleans

        #region Integers

        private const int PdfFirstPageNumber = 1;

        #endregion Integers

        #endregion Constants

        #region Folders

        #region RootFolder

        private static string RootPath = System.IO.Path.Combine(
            System.Environment.GetFolderPath(
                System.Environment.SpecialFolder.MyDocuments),
            RootFolderName);

        #endregion RootFolder

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
            }
        }

        private static string GetOutputPath(string inputPath)
        {
            return System.IO.Path.Combine(
                OutputRootPath,
                System.IO.Path.GetFileName(inputPath));
        }

        private async System.Threading.Tasks.Task ProcessPdfs(
            Script script)
        {
            await System.Threading.Tasks.Task.Run(() =>
            {
                for (int i = 0; i < fileView.CheckedItems.Count; i++)
                {
                    var currentFile = fileView.CheckedItems[i].ToString();
                    if (!IsPdf(currentFile))
                    {
                        try
                        {
                            currentFile = ConvertToPdf(currentFile);
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            ShowMessage(FileFailedToConvertToPdfErrorMessage
                                + Space + currentFile);
                            continue;
                        }
                    }
                    if (script != null)
                    {
                        AddScriptToPdf(currentFile, script);
                    }
                    Report(new ProgressReport
                    {
                        Total = fileView.CheckedItems.Count,
                        CurrentCount = i + 1
                    });
                }
            });
            ShowMessage(FilesSavedInMessage + Space + OutputRootPath);
        }

        private bool IsPdf(string filename)
        {
            return string.Equals(System.IO.Path.GetExtension(filename),
                PdfFileExtension,
                System.StringComparison.InvariantCultureIgnoreCase);
        }

        private void ShowException(System.Exception e)
        {
            ShowMessage(e.Message);
        }

        private void ShowMessage(string message)
        {
            System.Windows.Forms.MessageBox.Show(message);
        }

        private void AddScriptToPdf(string filename, Script script)
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
                ShowMessage(e.Message + Space + filename);
            }
        }

        private string ConvertToPdf(string filename)
        {
            var pdfPath = System.IO.Path.Combine(
                ProcessingPath,
                System.IO.Path.GetFileNameWithoutExtension(filename)
                + PdfFileExtension);
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
            if (fileView.CheckedItems.Count > 0)
            {
                Enabled = false;
                try
                {
                    await function();
                }
                catch (System.Exception e)
                {
                    ShowException(e);
                }
                progressBar.Value = 0;
                Enabled = true;
            }
            else
            {
                ShowMessage(NoFilesSelectedErrorMessage);
            }
        }

        private void selectDocuments_Click(object sender, System.EventArgs e)
        {
            var dialogResult = openFileDialog.ShowDialog();
            if (dialogResult == System.Windows.Forms.DialogResult.OK)
            {
                foreach (var file in openFileDialog.FileNames)
                {
                    fileView.Items.Add(file, FileViewFileIsChecked);
                }
            }
        }

        private async void timeStampDefaultDay_Click(object sender,
            System.EventArgs e)
        {
            await PerformTask(() => ProcessPdfs(
                Script.TimeStampOnPrintDefaultDay));
        }

        private async void timeStampDefaultMonth_Click(
            object sender, System.EventArgs e)
        {
            await PerformTask(() => ProcessPdfs(
                Script.TimeStampOnPrintDefaultMonth));
        }

        private async void convertOnly_Click(object sender, System.EventArgs e)
        {
            await PerformTask(() => ProcessPdfs(null));
        }
    }
}