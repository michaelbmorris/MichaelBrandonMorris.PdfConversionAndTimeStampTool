using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Threading;
using iTextSharp.text.pdf;
using MichaelBrandonMorris.Extensions.CollectionExtensions;
using MichaelBrandonMorris.Extensions.PrimitiveExtensions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using Action = System.Action;
using Progress =
    System.IProgress<(int Current, int Total, string CurrentFileName)>;
using WordApplication = Microsoft.Office.Interop.Word.Application;
using ExcelApplication = Microsoft.Office.Interop.Excel.Application;
using PowerPointApplication = Microsoft.Office.Interop.PowerPoint.Application;
using ActionTypeMapping =
    System.Collections.Generic.Dictionary<
        MichaelBrandonMorris.PdfTool.ScriptTiming,
        iTextSharp.text.pdf.PdfName>;
using Rectangle = iTextSharp.text.Rectangle;
using Task = System.Threading.Tasks.Task;

namespace MichaelBrandonMorris.PdfTool
{
    /// <summary>
    ///     Class FileProcessor.
    /// </summary>
    /// TODO Edit XML Comment Template for FileProcessor
    internal class FileProcessor
    {
        /// <summary>
        ///     The output folder path
        /// </summary>
        /// TODO Edit XML Comment Template for OutputFolderPath
        private static readonly string OutputFolderPath =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), FolderName);

        /// <summary>
        ///     The processing folder path
        /// </summary>
        /// TODO Edit XML Comment Template for ProcessingFolderPath
        private static readonly string ProcessingFolderPath =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), FolderName);

        /// <summary>
        ///     The action type mapping
        /// </summary>
        /// TODO Edit XML Comment Template for ActionTypeMapping
        private static readonly ActionTypeMapping ActionTypeMapping =
            new ActionTypeMapping
            {
                [ScriptTiming.DidPrint] = PdfWriter.DID_PRINT,
                [ScriptTiming.DidSave] = PdfWriter.DID_SAVE,
                [ScriptTiming.WillPrint] = PdfWriter.WILL_PRINT,
                [ScriptTiming.WillSave] = PdfWriter.WILL_SAVE
            };

        /// <summary>
        ///     The alternating page increment
        /// </summary>
        /// TODO Edit XML Comment Template for AlternatingPageIncrement
        private const int AlternatingPageIncrement = 2;

        /// <summary>
        ///     The every page increment
        /// </summary>
        /// TODO Edit XML Comment Template for EveryPageIncrement
        private const int EveryPageIncrement = 1;

        /// <summary>
        ///     The excel extension
        /// </summary>
        /// TODO Edit XML Comment Template for ExcelExtension
        private const string ExcelExtension = ".xls";

        /// <summary>
        ///     The excel XML extension
        /// </summary>
        /// TODO Edit XML Comment Template for ExcelXmlExtension
        private const string ExcelXmlExtension = ".xlsx";

        /// <summary>
        ///     The first page number
        /// </summary>
        /// TODO Edit XML Comment Template for FirstPageNumber
        private const int FirstPageNumber = 1;

        /// <summary>
        ///     The folder name
        /// </summary>
        /// TODO Edit XML Comment Template for FolderName
        private const string FolderName = "PDF Conversion And Time Stamp Tool";

        /// <summary>
        ///     The PDF extension
        /// </summary>
        /// TODO Edit XML Comment Template for PdfExtension
        private const string PdfExtension = ".pdf";

        /// <summary>
        ///     The power point extension
        /// </summary>
        /// TODO Edit XML Comment Template for PowerPointExtension
        private const string PowerPointExtension = ".ppt";

        /// <summary>
        ///     The power point XML extension
        /// </summary>
        /// TODO Edit XML Comment Template for PowerPointXmlExtension
        private const string PowerPointXmlExtension = ".pptx";

        /// <summary>
        ///     The word extension
        /// </summary>
        /// TODO Edit XML Comment Template for WordExtension
        private const string WordExtension = ".doc";

        /// <summary>
        ///     The word XML extension
        /// </summary>
        /// TODO Edit XML Comment Template for WordXmlExtension
        private const string WordXmlExtension = ".docx";

        /// <summary>
        ///     Initializes a new instance of the
        ///     <see cref="FileProcessor" /> class.
        /// </summary>
        /// <param name="fileNames">The file names.</param>
        /// <param name="progress">The progress.</param>
        /// <param name="field">The field.</param>
        /// <param name="script">The script.</param>
        /// TODO Edit XML Comment Template for #ctor
        internal FileProcessor(
            IList<string> fileNames,
            Progress progress,
            Field field = null,
            Script script = null)
        {
            FileNames = fileNames;
            Progress = progress;
            Field = field;
            Script = script;
            Directory.CreateDirectory(OutputFolderPath);
            Directory.CreateDirectory(ProcessingFolderPath);
        }

        /// <summary>
        ///     Gets the cancellation token.
        /// </summary>
        /// <value>The cancellation token.</value>
        /// TODO Edit XML Comment Template for CancellationToken
        private CancellationToken CancellationToken => CancellationTokenSource
            .Token;

        /// <summary>
        ///     Gets the cancellation token source.
        /// </summary>
        /// <value>The cancellation token source.</value>
        /// TODO Edit XML Comment Template for CancellationTokenSource
        private CancellationTokenSource CancellationTokenSource
        {
            get;
        } = new CancellationTokenSource();

        /// <summary>
        ///     Gets the field.
        /// </summary>
        /// <value>The field.</value>
        /// TODO Edit XML Comment Template for Field
        private Field Field
        {
            get;
        }

        /// <summary>
        ///     Gets the file names.
        /// </summary>
        /// <value>The file names.</value>
        /// TODO Edit XML Comment Template for FileNames
        private IList<string> FileNames
        {
            get;
        }

        /// <summary>
        ///     Gets the log.
        /// </summary>
        /// <value>The log.</value>
        /// TODO Edit XML Comment Template for Log
        private List<string> Log
        {
            get;
        } = new List<string>();

        /// <summary>
        ///     Gets the progress.
        /// </summary>
        /// <value>The progress.</value>
        /// TODO Edit XML Comment Template for Progress
        private Progress Progress
        {
            get;
        }

        /// <summary>
        ///     Gets the script.
        /// </summary>
        /// <value>The script.</value>
        /// TODO Edit XML Comment Template for Script
        private Script Script
        {
            get;
        }

        /// <summary>
        ///     Cancels this instance.
        /// </summary>
        /// TODO Edit XML Comment Template for Cancel
        internal void Cancel()
        {
            CancellationTokenSource.Cancel();
        }

        /// <summary>
        ///     Executes this instance.
        /// </summary>
        /// <returns>Task.</returns>
        /// TODO Edit XML Comment Template for Execute
        internal async Task Execute()
        {
            var task = Task.Run(
                () =>
                {
                    var count = 0;
                    foreach (var t in FileNames)
                    {
                        CancellationToken.ThrowIfCancellationRequested();
                        var currentFileName = string.Empty;
                        try
                        {
                            ClearProcessing();
                            currentFileName = CopyToProcessing(t);

                            if (!IsPdf(currentFileName))
                            {
                                currentFileName = ConvertToPdf(currentFileName);
                            }

                            if (Field != null
                                || Script != null)
                            {
                                ProcessPdf(currentFileName);
                            }
                            else
                            {
                                MoveToOutput(currentFileName);
                            }
                        }
                        catch (Exception e)
                        {
                            Log.Add(e.Message);
                        }

                        Progress.Report(
                            ( ++count, FileNames.Count, currentFileName));
                    }
                },
                CancellationToken);

            await task;
            Process.Start(OutputFolderPath);

            if (!Log.IsEmpty())
            {
                var now = DateTime.Now;
                var logFileName = $"Log - {now:yyyyMMddTHHmmss}.txt";
                var logFilePath = Path.Combine(OutputFolderPath, logFileName);
                File.WriteAllLines(logFilePath, Log);
            }
        }

        /// <summary>
        ///     Clears the processing.
        /// </summary>
        /// TODO Edit XML Comment Template for ClearProcessing
        private static void ClearProcessing()
        {
            foreach (var file in Directory.GetFiles(ProcessingFolderPath))
            {
                File.Delete(file);
            }
        }

        /// <summary>
        ///     Converts the excel to PDF.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="processingPath">The processing path.</param>
        /// <exception cref="Exception"></exception>
        /// TODO Edit XML Comment Template for ConvertExcelToPdf
        private static void ConvertExcelToPdf(
            string fileName,
            string processingPath)
        {
            var excelApplication = new ExcelApplication
            {
                ScreenUpdating = false,
                DisplayAlerts = false,
                AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable
            };

            var excelWorkbook = excelApplication.Workbooks.Open(fileName);

            if (excelWorkbook == null)
            {
                excelApplication.Quit();
                throw new Exception(
                    $"The file '{fileName}' could not be opened.");
            }

            try
            {
                excelWorkbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, processingPath);
            }
            finally
            {
                excelWorkbook.Close();
                excelApplication.Quit();
            }
        }

        /// <summary>
        ///     Converts the power point to PDF.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="processingPath">The processing path.</param>
        /// TODO Edit XML Comment Template for ConvertPowerPointToPdf
        private static void ConvertPowerPointToPdf(
            string fileName,
            string processingPath)
        {
            var powerPointApplication = new PowerPointApplication
            {
                DisplayAlerts = PpAlertLevel.ppAlertsNone,
                AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable,
                DisplayDocumentInformationPanel = false
            };

            var powerPointPresentation =
                powerPointApplication.Presentations.Open(
                    fileName,
                    WithWindow: MsoTriState.msoFalse);

            if (powerPointPresentation == null)
            {
                powerPointApplication.Quit();
                return;
            }

            try
            {
                powerPointPresentation.ExportAsFixedFormat(
                    processingPath,
                    PpFixedFormatType.ppFixedFormatTypePDF);
            }
            finally
            {
                powerPointPresentation.Close();
                powerPointApplication.Quit();
            }
        }

        /// <summary>
        ///     Converts to PDF.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <returns>System.String.</returns>
        /// <exception cref="Exception"></exception>
        /// TODO Edit XML Comment Template for ConvertToPdf
        [SuppressMessage("ReSharper", "ImplicitlyCapturedClosure")]
        private static string ConvertToPdf(string fileName)
        {
            var processingFileName = Path.GetFileNameWithoutExtension(fileName)
                                     + PdfExtension;

            var processingPath =
                Path.Combine(ProcessingFolderPath, processingFileName);

            var extension = Path.GetExtension(fileName)?.ToLower();

            if (extension == null || extension.IsWhiteSpace())
            {
                throw new Exception(
                    $"The file '{fileName}'does not have an extension.");
            }

            var extensionMapping = new Dictionary<string, Action>
            {
                [WordExtension] = () =>
                {
                    ConvertWordToPdf(fileName, processingPath);
                },
                [WordXmlExtension] = () =>
                {
                    ConvertWordToPdf(fileName, processingPath);
                },
                [ExcelExtension] = () =>
                {
                    ConvertExcelToPdf(fileName, processingPath);
                },
                [ExcelXmlExtension] = () =>
                {
                    ConvertExcelToPdf(fileName, processingPath);
                },
                [PowerPointExtension] = () =>
                {
                    ConvertPowerPointToPdf(fileName, processingPath);
                },
                [PowerPointXmlExtension] = () =>
                {
                    ConvertPowerPointToPdf(fileName, processingPath);
                }
            };

            extensionMapping[extension]();
            return processingPath;
        }

        /// <summary>
        ///     Converts the word to PDF.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="processingPath">The processing path.</param>
        /// TODO Edit XML Comment Template for ConvertWordToPdf
        private static void ConvertWordToPdf(
            string fileName,
            string processingPath)
        {
            var wordApplication = new WordApplication
            {
                AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable
            };

            wordApplication.Application.AutomationSecurity =
                MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            var wordDocument = wordApplication.Documents.Open(fileName);

            if (wordDocument == null)
            {
                wordApplication.Quit();
                return;
            }

            try
            {
                wordDocument.ExportAsFixedFormat(
                    processingPath,
                    WdExportFormat.wdExportFormatPDF);
            }
            finally
            {
                wordDocument.Close(false);
                wordApplication.Quit();
            }
        }

        /// <summary>
        ///     Copies to processing.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <returns>System.String.</returns>
        /// TODO Edit XML Comment Template for CopyToProcessing
        private static string CopyToProcessing(string fileName)
        {
            var processingPath = GetProcessingPath(fileName);
            File.Copy(fileName, processingPath);
            return processingPath;
        }

        /// <summary>
        ///     Gets the output path.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <returns>System.String.</returns>
        /// <exception cref="Exception"></exception>
        /// TODO Edit XML Comment Template for GetOutputPath
        private static string GetOutputPath(string fileName)
        {
            var outputFileName = Path.GetFileName(fileName);

            if (outputFileName == null)
            {
                throw new Exception(
                    $"The name of the file '{fileName}' "
                    + "is incorrectly formatted.");
            }

            return Path.Combine(OutputFolderPath, outputFileName);
        }

        /// <summary>
        ///     Gets the processing path.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <returns>System.String.</returns>
        /// <exception cref="Exception"></exception>
        /// TODO Edit XML Comment Template for GetProcessingPath
        private static string GetProcessingPath(string fileName)
        {
            var processingFileName = Path.GetFileName(fileName);

            if (processingFileName != null)
            {
                return Path.Combine(ProcessingFolderPath, processingFileName);
            }

            throw new Exception(
                $"The name of the file '{fileName}' "
                + "is incorrectly formatted.");
        }

        /// <summary>
        ///     Determines whether the specified file name is PDF.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <returns>
        ///     <c>true</c> if the specified file name is PDF;
        ///     otherwise, <c>false</c>.
        /// </returns>
        /// TODO Edit XML Comment Template for IsPdf
        private static bool IsPdf(string fileName)
        {
            return Path.GetExtension(fileName).EqualsOrdinalIgnoreCase(PdfExtension);
        }

        /// <summary>
        ///     Moves to output.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// TODO Edit XML Comment Template for MoveToOutput
        private static void MoveToOutput(string fileName)
        {
            var outputPath = GetOutputPath(fileName);

            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            File.Move(fileName, outputPath);
        }

        /// <summary>
        ///     Adds the field to page.
        /// </summary>
        /// <param name="pageNumber">The page number.</param>
        /// <param name="pdfStamper">The PDF stamper.</param>
        /// <param name="parentField">The parent field.</param>
        /// TODO Edit XML Comment Template for AddFieldToPage
        private void AddFieldToPage(
            int pageNumber,
            PdfStamper pdfStamper,
            PdfFormField parentField)
        {
            var rectangle = new Rectangle(
                Field.LeftX,
                Field.TopY,
                Field.RightX,
                Field.BottomY);

            var textField = new TextField(pdfStamper.Writer, rectangle, null);
            var childField = textField.GetTextField();
            parentField.AddKid(childField);
            childField.PlaceInPage = pageNumber;
        }

        /// <summary>
        ///     Adds the field to PDF.
        /// </summary>
        /// <param name="pdfStamper">The PDF stamper.</param>
        /// <param name="numberOfPages">The number of pages.</param>
        /// TODO Edit XML Comment Template for AddFieldToPdf
        private void AddFieldToPdf(PdfStamper pdfStamper, int numberOfPages)
        {
            var parentField =
                PdfFormField.CreateTextField(pdfStamper.Writer, false, false, 0);

            parentField.FieldName = Field.Name;

            var pageNumber = Field.Pages == FieldPages.Last
                ? numberOfPages
                : FirstPageNumber;

            // ReSharper disable once ConvertIfStatementToSwitchStatement
            if (Field.Pages == FieldPages.First
                || Field.Pages == FieldPages.Last)
            {
                AddFieldToPage(pageNumber, pdfStamper, parentField);
            }
            else if (Field.Pages == FieldPages.Custom)
            {
                foreach (var customPageNumber in Field.CustomPageNumbers)
                {
                    AddFieldToPage(customPageNumber, pdfStamper, parentField);
                }
            }
            else
            {
                var increment = Field.Pages == FieldPages.All
                    ? EveryPageIncrement
                    : AlternatingPageIncrement;

                if (Field.Pages == FieldPages.Even)
                {
                    pageNumber += 1;
                }

                for (; pageNumber <= numberOfPages; pageNumber += increment)
                {
                    AddFieldToPage(pageNumber, pdfStamper, parentField);
                }
            }

            pdfStamper.AddAnnotation(parentField, FirstPageNumber);
        }

        /// <summary>
        ///     Adds the script to PDF.
        /// </summary>
        /// <param name="pdfStamper">The PDF stamper.</param>
        /// TODO Edit XML Comment Template for AddScriptToPdf
        private void AddScriptToPdf(PdfStamper pdfStamper)
        {
            var pdfAction = PdfAction.JavaScript(Script.Text, pdfStamper.Writer);

            pdfStamper.Writer.SetAdditionalAction(
                ActionTypeMapping[Script.Timing],
                pdfAction);
        }

        /// <summary>
        ///     Processes the PDF.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// TODO Edit XML Comment Template for ProcessPdf
        private void ProcessPdf(string fileName)
        {
            using (var pdfReader = new PdfReader(fileName))
            using (var fileStream =
                new FileStream(GetOutputPath(fileName), FileMode.Create))
            using (var pdfStamper = new PdfStamper(pdfReader, fileStream))
            {
                if (Field != null)
                {
                    AddFieldToPdf(pdfStamper, pdfReader.NumberOfPages);
                }

                if (Script != null)
                {
                    AddScriptToPdf(pdfStamper);
                }
            }
        }
    }
}