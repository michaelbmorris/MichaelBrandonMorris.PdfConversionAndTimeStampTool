//-----------------------------------------------------------------------------------------------------------
// <copyright file="PdfScriptTool.cs" company="Michael Brandon Morris">
//     Copyright © Michael Brandon Morris 2016
// </copyright>
//-----------------------------------------------------------------------------------------------------------

namespace PdfScriptTool
{
    using static Field;
    using static Properties.Resources;
    using static Script;

    internal partial class PdfScriptTool : System.Windows.Forms.Form,
        System.IProgress<ProgressReport>
    {
        #region Constants

        #region Booleans

        private const bool FileViewFileIsChecked = true;
        private const bool OpenFileDialogAllowMultiple = true;

        #endregion Booleans

        #region Integers

        private const int EveryOtherPage = 2;
        private const int EveryPage = 1;
        private const int FirstPageNumber = 1;
        private const int SecondPageNumber = 2;

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

        #region UI Listeners

        private async void convertOnly_Click(object sender, System.EventArgs e)
        {
            await PerformTask(() => ProcessPdfs(null));
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
            await PerformTask(() => ProcessPdfs(DefaultTimeStampField,
                TimeStampOnPrintDefaultDayScript));
        }

        private async void timeStampDefaultMonth_Click(
            object sender, System.EventArgs e)
        {
            await PerformTask(() => ProcessPdfs(DefaultTimeStampField,
                TimeStampOnPrintDefaultMonthScript));
        }

        #endregion UI Listeners

        #region Helpers

        private void InitializeOpenFileDialog()
        {
            openFileDialog.Filter = OpenFileDialogFilter;
            openFileDialog.Multiselect = OpenFileDialogAllowMultiple;
            openFileDialog.Title = OpenFileDialogTitle;
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

        #endregion Helpers

        private void AddFieldToPage(Field field, int pageNumber,
            iTextSharp.text.pdf.PdfStamper pdfStamper,
            iTextSharp.text.pdf.PdfFormField parentField)
        {
            var textField = new iTextSharp.text.pdf.TextField(
                pdfStamper.Writer, new iTextSharp.text.Rectangle(
                    field.TopLeftX, field.TopLeftY, field.BottomRightX,
                    field.BottomRightY), null);

            var childField = textField.GetTextField();

            parentField.AddKid(childField);

            childField.PlaceInPage = pageNumber;
        }

        private void AddFieldToPdf(Field field,
            iTextSharp.text.pdf.PdfStamper pdfStamper, int numberOfPages)
        {
            var parentField = iTextSharp.text.pdf.PdfFormField.CreateTextField(
                pdfStamper.Writer, false, false, 0);

            parentField.FieldName = field.Title;

            int pageNumber = field.Pages == Pages.Last ?
                numberOfPages : FirstPageNumber;

            if (field.Pages == Pages.First || field.Pages == Pages.Last)
            {
                AddFieldToPage(field, pageNumber, pdfStamper, parentField);
            }
            else
            {
                int increment = field.Pages == Pages.All ?
                    EveryPage : EveryOtherPage;
                if (field.Pages == Pages.Even)
                {
                    pageNumber += 1;
                }
                for (; pageNumber <= numberOfPages; pageNumber += increment)
                {
                    AddFieldToPage(field, pageNumber, pdfStamper, parentField);
                }
            }

            pdfStamper.AddAnnotation(parentField, FirstPageNumber);
        }

        private void AddScriptToPdf(Script script,
            iTextSharp.text.pdf.PdfStamper pdfStamper)
        {
            var pdfAction = iTextSharp.text.pdf.PdfAction.JavaScript(
                script.ScriptText, pdfStamper.Writer);

            iTextSharp.text.pdf.PdfName actionType = null;

            switch (script.ScriptTiming)
            {
                case ScriptTiming.DidPrint:
                    actionType = iTextSharp.text.pdf.PdfWriter
                            .DID_PRINT;
                    break;

                case ScriptTiming.DidSave:
                    actionType = iTextSharp.text.pdf.PdfWriter
                            .DID_SAVE;
                    break;

                case ScriptTiming.WillPrint:
                    actionType = iTextSharp.text.pdf.PdfWriter
                            .WILL_PRINT;
                    break;

                case ScriptTiming.WillSave:
                    actionType = iTextSharp.text.pdf.PdfWriter
                            .WILL_SAVE;
                    break;
            }
            pdfStamper.Writer.SetAdditionalAction(
                actionType, pdfAction);
        }

        private string ConvertToPdf(string filename)
        {
            var pdfPath = System.IO.Path.Combine(ProcessingPath,
                System.IO.Path.GetFileNameWithoutExtension(filename)
                + PdfFileExtension);

            var wordApplication
                = new Microsoft.Office.Interop.Word.Application();

            var wordDocument = wordApplication.Documents.Open(filename);
            wordDocument.ExportAsFixedFormat(pdfPath,
                Microsoft.Office.Interop.Word.WdExportFormat
                .wdExportFormatPDF);
            wordDocument.Close(false);
            wordApplication.Quit();
            return pdfPath;
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

        private void ProcessPdf(string filename, Field field, Script script)
        {
            using (var pdfReader = new iTextSharp.text.pdf.PdfReader(filename))
            {
                using (var pdfStamper = new iTextSharp.text.pdf.PdfStamper(
                    pdfReader, new System.IO.FileStream(GetOutputPath(
                        filename), System.IO.FileMode.Create)))
                {
                    if (field != null)
                    {
                        AddFieldToPdf(field, pdfStamper, pdfReader.NumberOfPages);
                    }
                    if (script != null)
                    {
                        AddScriptToPdf(script, pdfStamper);
                    }
                }
            }
        }

        private async System.Threading.Tasks.Task ProcessPdfs(Field field)
        {
            await ProcessPdfs(field, null);
        }

        private async System.Threading.Tasks.Task ProcessPdfs(
            Field field,
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
                                + DefaultTimestampFieldTitle + currentFile);
                            continue;
                        }
                    }
                    ProcessPdf(currentFile, field, script);
                    Report(new ProgressReport
                    {
                        Total = fileView.CheckedItems.Count,
                        CurrentCount = i + 1
                    });
                }
            });
            ShowMessage(FilesSavedInMessage + DefaultTimestampFieldTitle
                + OutputRootPath);
        }
    }
}