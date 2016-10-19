using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using MichaelBrandonMorris.Extensions.CollectionExtensions;
using MichaelBrandonMorris.Extensions.PrimitiveExtensions;
using static System.Environment;
using static System.Environment.SpecialFolder;
using static System.IO.Path;
using static Microsoft.Office.Core.MsoAutomationSecurity;
using static System.IO.Directory;
using Progress = System.IProgress<System.Tuple<int, int, string>>;
using ProgressReport = System.Tuple<int, int, string>;
using static Microsoft.Office.Interop.Word.WdExportFormat;
using static Microsoft.Office.Interop.Excel.XlFixedFormatType;
using static Microsoft.Office.Interop.PowerPoint.PpFixedFormatType;
using WordApplication = Microsoft.Office.Interop.Word.Application;
using ExcelApplication = Microsoft.Office.Interop.Excel.Application;
using PowerPointApplication = Microsoft.Office.Interop.PowerPoint.Application;
using static System.IO.File;
using static System.Diagnostics.Debug;
using ActionTypeMapping = System.Collections.Generic.Dictionary
    <MichaelBrandonMorris.PdfConversionAndTimeStampTool.ScriptTiming,
        iTextSharp.text.pdf.PdfName>;
using static MichaelBrandonMorris.PdfConversionAndTimeStampTool.FieldPages;
using static MichaelBrandonMorris.PdfConversionAndTimeStampTool.ScriptTiming;
using static iTextSharp.text.pdf.PdfFormField;
using static Microsoft.Office.Interop.PowerPoint.PpAlertLevel;
using static iTextSharp.text.pdf.PdfAction;
using static Microsoft.Office.Core.MsoTriState;
using static iTextSharp.text.pdf.PdfWriter;
using System.Diagnostics;

namespace MichaelBrandonMorris.PdfConversionAndTimeStampTool
{
    internal class FileProcessor
    {
        private const int AlternatingPageIncrement = 2;
        private const string DateTimeFormat = "yyyyMMddTHHmmss";
        private const int EveryPageIncrement = 1;
        private const string ExcelExtension = ".xls";
        private const string ExcelXmlExtension = ".xlsx";
        private const int FirstPageNumber = 1;
        private const string FolderName = "PDF Conversion And Time Stamp Tool";
        private const string PdfExtension = ".pdf";
        private const string PowerPointExtension = ".ppt";
        private const string PowerPointXmlExtension = ".pptx";
        private const string WordExtension = ".doc";
        private const string WordXmlExtension = ".docx";

        private static readonly string OutputFolderPath = Combine(
            GetFolderPath(MyDocuments), FolderName);

        private static readonly string ProcessingFolderPath = Combine(
            GetFolderPath(ApplicationData), FolderName);

        private static readonly ActionTypeMapping ActionTypeMapping =
            new ActionTypeMapping
            {
                [DidPrint] = DID_PRINT,
                [DidSave] = DID_SAVE,
                [WillPrint] = WILL_PRINT,
                [WillSave] = WILL_SAVE
            };


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
            CreateDirectory(OutputFolderPath);
            CreateDirectory(ProcessingFolderPath);
        }

        private CancellationToken CancellationToken =>
            CancellationTokenSource.Token;

        private CancellationTokenSource CancellationTokenSource
        {
            get;
        } = new CancellationTokenSource();

        private Field Field
        {
            get;
        }

        private IList<string> FileNames
        {
            get;
        }

        private List<string> Log
        {
            get;
        } = new List<string>();

        private Progress Progress
        {
            get;
        }

        private Script Script
        {
            get;
        }

        internal void Cancel()
        {
            CancellationTokenSource.Cancel();
        }

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

                            if (Field != null || Script != null)
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
                            new ProgressReport(
                                ++count, FileNames.Count, currentFileName));
                    }
                },
                CancellationToken);

            await task;

            Process.Start(OutputFolderPath);

            if (!Log.IsEmpty())
            {
                var now = DateTime.Now;
                var logFileName = $"Log - {now.ToString(DateTimeFormat)}.txt";
                var logFilePath = Combine(OutputFolderPath, logFileName);
                WriteAllLines(logFilePath, Log);
            }
        }

        private static void ClearProcessing()
        {
            foreach (var file in GetFiles(ProcessingFolderPath))
            {
                File.Delete(file);
            }
        }

        private static void ConvertExcelToPdf(
            string fileName, string processingPath)
        {
            var excelApplication = new ExcelApplication
            {
                ScreenUpdating = false,
                DisplayAlerts = false,
                AutomationSecurity = msoAutomationSecurityForceDisable
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
                excelWorkbook.ExportAsFixedFormat(xlTypePDF, processingPath);
            }
            finally
            {
                excelWorkbook.Close();
                excelApplication.Quit();
            }
        }

        private static void ConvertPowerPointToPdf(
            string fileName, string processingPath)
        {
            var powerPointApplication = new PowerPointApplication
            {
                DisplayAlerts = ppAlertsNone,
                AutomationSecurity = msoAutomationSecurityForceDisable,
                DisplayDocumentInformationPanel = false
            };

            var powerPointPresentation =
                powerPointApplication.Presentations.Open(
                    fileName, WithWindow: msoFalse);

            if (powerPointPresentation == null)
            {
                powerPointApplication.Quit();
                return;
            }

            try
            {
                powerPointPresentation.ExportAsFixedFormat(
                    processingPath, ppFixedFormatTypePDF);
            }
            finally
            {
                powerPointPresentation.Close();
                powerPointApplication.Quit();
            }
        }

        [SuppressMessage("ReSharper", "ImplicitlyCapturedClosure")]
        private static string ConvertToPdf(string fileName)
        {
            var processingFileName = GetFileNameWithoutExtension(fileName) +
                                     PdfExtension;

            var processingPath = Combine(
                ProcessingFolderPath, processingFileName);

            var extension = GetExtension(fileName)?.ToLower();

            if (extension.IsNullOrWhiteSpace())
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

            Assert(extension != null, "extension != null");
            extensionMapping[extension]();
            return processingPath;
        }

        private static void ConvertWordToPdf(
            string fileName, string processingPath)
        {
            var wordApplication = new WordApplication
            {
                AutomationSecurity = msoAutomationSecurityForceDisable
            };

            wordApplication.Application.AutomationSecurity =
                msoAutomationSecurityForceDisable;

            var wordDocument = wordApplication.Documents.Open(fileName);

            if (wordDocument == null)
            {
                wordApplication.Quit();
                return;
            }

            try
            {
                wordDocument.ExportAsFixedFormat(
                    processingPath, wdExportFormatPDF);
            }
            finally
            {
                wordDocument.Close(false);
                wordApplication.Quit();
            }
        }

        private static string CopyToProcessing(string fileName)
        {
            var processingPath = GetProcessingPath(fileName);
            Copy(fileName, processingPath);
            return processingPath;
        }

        private static string GetOutputPath(string fileName)
        {
            var outputFileName = GetFileName(fileName);

            if (outputFileName == null)
            {
                throw new Exception(
                    $"The name of the file '{fileName}' " +
                    "is incorrectly formatted.");
            }

            return Combine(OutputFolderPath, outputFileName);
        }

        private static string GetProcessingPath(string fileName)
        {
            var processingFileName = GetFileName(fileName);

            if (processingFileName != null)
            {
                return Combine(ProcessingFolderPath, processingFileName);
            }

            throw new Exception(
                $"The name of the file '{fileName}' " +
                "is incorrectly formatted.");
        }

        private static bool IsPdf(string fileName)
        {
            return GetExtension(fileName).EqualsOrdinalIgnoreCase(
                PdfExtension);
        }

        private static void MoveToOutput(string fileName)
        {
            var outputPath = GetOutputPath(fileName);

            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            File.Move(fileName, outputPath);
        }

        private void AddFieldToPage(
            int pageNumber,
            PdfStamper pdfStamper,
            PdfFormField parentField)
        {
            var rectangle = new Rectangle(
                Field.LeftX, Field.TopY, Field.RightX, Field.BottomY);

            var textField = new TextField(pdfStamper.Writer, rectangle, null);
            var childField = textField.GetTextField();
            parentField.AddKid(childField);
            childField.PlaceInPage = pageNumber;
        }

        private void AddFieldToPdf(PdfStamper pdfStamper, int numberOfPages)
        {
            var parentField = CreateTextField(
                pdfStamper.Writer, false, false, 0);

            parentField.FieldName = Field.Name;

            var pageNumber = Field.Pages == Last
                ? numberOfPages
                : FirstPageNumber;

            // ReSharper disable once ConvertIfStatementToSwitchStatement
            if (Field.Pages == First ||
                Field.Pages == Last)
            {
                AddFieldToPage(pageNumber, pdfStamper, parentField);
            }
            else if (Field.Pages == Custom)
            {
                foreach (var customPageNumber in Field.CustomPageNumbers)
                {
                    AddFieldToPage(customPageNumber, pdfStamper, parentField);
                }
            }
            else
            {
                var increment = Field.Pages == All
                    ? EveryPageIncrement
                    : AlternatingPageIncrement;

                if (Field.Pages == Even)
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

        private void AddScriptToPdf(PdfStamper pdfStamper)
        {
            var pdfAction = JavaScript(Script.Text, pdfStamper.Writer);

            pdfStamper.Writer.SetAdditionalAction(
                ActionTypeMapping[Script.Timing], pdfAction);
        }

        private void ProcessPdf(string fileName)
        {
            using (var pdfReader = new PdfReader(fileName))
            using (var fileStream = new FileStream(
                GetOutputPath(fileName), FileMode.Create))
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